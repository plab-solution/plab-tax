/**
 * PLAB 청구서 (tax.plab.life) — Apps Script 백엔드
 *
 * 'PLAB 청구서' 구글 시트에 붙어 있는(컨테이너 바인딩) 스크립트입니다.
 * 시트 메뉴 [PLAB 청구서]에서 초기 설정 · 링크 만들기 · 거래처 집계를 할 수 있습니다.
 *
 * 스크립트 속성(프로젝트 설정 > 스크립트 속성)
 *   SLACK_WEBHOOK    : 문의·서류 제출 알림을 보낼 슬랙 웹훅 URL
 *   DOCS_FOLDER_ID   : 거래처가 올린 서류를 저장할 드라이브 폴더 ID (운영팀만 접근)
 *   SEAL_FILE_ID     : 직인 이미지(PNG) 드라이브 파일 ID
 *   PLAB_BIZ_FILE_ID : 프랩 사업자등록증 파일 ID (거래처 '프랩 서류 받기'용)
 *   PLAB_BANK_FILE_ID: 프랩 통장사본 파일 ID
 *
 * 거래처 구분
 *   신규 : 거래 1건 (사이트에서 서류 제출)
 *   기존 : 거래 2건 이상, 또는 등록 경로가 '기존 자료'(사이트 이전부터 거래하던 곳)
 *   서류 대기 : 사업자번호가 없거나 서류가 아직 없음
 */

var SITE_URL = 'https://tax.plab.life/';
var TZ = 'Asia/Seoul';

var SHEETS = {
  invoice: { name: '청구서', headers: ['링크ID', '문서번호', '발행일', '지급기한', '거래처 상호', '구분', '거래처 사업자번호', '거래처 대표자', '건명', '금액', '할인', '비고', '상태', '링크', '생성일'] },
  item:    { name: '품목',   headers: ['문서번호', '품목명', '수량', '단가'] },
  client:  { name: '거래처', headers: ['사업자번호', '상호', '대표자', '구분', '등록 경로', '사업자등록증', '통장사본', '첫 거래일', '거래 건수', '최근 청구일', '누적 청구액', '메모', '제출일', '청구서 문서번호'] },
  inquiry: { name: '문의',   headers: ['일시', '문서번호', '거래처', '구분', '내용', '처리'] }
};
// 스크립트가 채우는 칸(파란색으로 표시)
var AUTO_COLS = {
  invoice: ['링크ID', '문서번호', '발행일', '지급기한', '구분', '거래처 사업자번호', '거래처 대표자', '링크', '생성일'],
  client:  ['구분', '첫 거래일', '거래 건수', '최근 청구일', '누적 청구액', '제출일', '청구서 문서번호'],
  inquiry: ['일시', '문서번호', '거래처', '구분', '내용']
};

var MAX_FILE_BYTES = 10 * 1024 * 1024;
var ALLOWED_MIME = ['image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif', 'application/pdf'];

/* ───────────────────────── 메뉴 · 시트 준비 ───────────────────────── */

function onOpen() {
  try { ensureHeaders_(); applyFormats_(); } catch (e) {}
  SpreadsheetApp.getUi().createMenu('PLAB 청구서')
    .addItem('링크 만들기 (ID 없는 행)', 'makeLinks')
    .addItem('거래처 집계 새로고침', 'refreshAll')
    .addSeparator()
    .addItem('초기 설정 (시트 만들기)', 'setup')
    .addItem('슬랙 알림 테스트', 'testSlack')
    .addToUi();
}

/** 시트와 머리글을 만들고 예시 청구서를 한 건 넣습니다. 이미 있는 내용은 지우지 않습니다. */
function setup() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(function (k) {
    var def = SHEETS[k];
    if (!ss.getSheetByName(def.name)) {
      var sh = ss.insertSheet(def.name);
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  ensureHeaders_();
  var inv = sheet_('invoice');
  if (inv.getLastRow() < 2) {
    var today = new Date(); today.setHours(0, 0, 0, 0);
    appendByHeader_(inv, {
      '문서번호': 'PLAB-' + today.getFullYear() + '-0001', '발행일': today, '지급기한': new Date(today.getTime() + 7 * 864e5),
      '거래처 상호': '예시상회', '건명': 'AI 상담 챗봇 구축', '금액': 3000000, '할인': 0,
      '비고': '예시 청구서입니다. 지우고 쓰셔도 돼요.', '상태': '입금 대기'
    });
  }
  applyFormats_();
  makeLinks();
  refreshAll();
}

/** 예전에 만든 시트에 새 머리글이 없으면 정해진 순서 자리에 열을 끼워 넣습니다. */
function ensureHeaders_() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(function (k) {
    var def = SHEETS[k], sh = ss.getSheetByName(def.name);
    if (!sh) return;
    def.headers.forEach(function (name, i) {
      var h = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
      if (h.indexOf(name) >= 0) return;
      var prev = i > 0 ? h.indexOf(def.headers[i - 1]) : -1;
      var col;
      if (prev >= 0) { sh.insertColumnAfter(prev + 1); col = prev + 2; }
      else { col = sh.getLastColumn() + 1; }
      sh.getRange(1, col).setValue(name).setFontWeight('bold');
    });
  });
}

/** 머리글 색(자동 칸 파랑) · 날짜/금액 서식 · 드롭다운 */
function applyFormats_() {
  Object.keys(SHEETS).forEach(function (k) {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS[k].name);
    if (!sh || sh.getLastColumn() < 1) return;
    var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var auto = AUTO_COLS[k] || [];
    var rows = Math.max(1, sh.getMaxRows() - 1);
    h.forEach(function (name, i) {
      sh.getRange(1, i + 1).setBackground(auto.indexOf(name) >= 0 ? '#E8F0FE' : '#F8F9FA').setFontWeight('bold');
      if (/일$|일시$/.test(name) || name === '지급기한') sh.getRange(2, i + 1, rows, 1).setNumberFormat(name === '일시' ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd');
      if (['금액', '할인', '단가', '누적 청구액'].indexOf(name) >= 0) sh.getRange(2, i + 1, rows, 1).setNumberFormat('#,##0');
    });
  });

  var inv = sheet_('invoice'), cl = sheet_('client');
  var rows = Math.max(1, inv.getMaxRows() - 1);
  var st = colOf_(inv, '상태');
  if (st) inv.getRange(2, st, rows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['입금 대기', '입금 완료', '취소'], true).build());
  // 거래처 상호: 거래처 탭 목록에서 고르되, 목록에 없는 새 상호도 입력 가능
  var nm = colOf_(inv, '거래처 상호'), clName = colOf_(cl, '상호');
  if (nm && clName) inv.getRange(2, nm, rows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(cl.getRange(2, clName, Math.max(1, cl.getMaxRows() - 1), 1), true)
      .setAllowInvalid(true)
      .setHelpText('기존 거래처는 목록에서 고르고, 새 거래처는 상호를 그대로 입력하세요.')
      .build());
  var ep = colOf_(cl, '등록 경로');
  if (ep) cl.getRange(2, ep, Math.max(1, cl.getMaxRows() - 1), 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['사이트 제출', '기존 자료'], true).build());
}

/** 링크ID가 비어 있고 문서번호가 있는 행마다 추측할 수 없는 ID와 링크를 채웁니다. */
function makeLinks() {
  var sh = sheet_('invoice');
  var data = sh.getDataRange().getValues();
  var h = data[0];
  var cId = h.indexOf('링크ID'), cNo = h.indexOf('문서번호'), cLink = h.indexOf('링크'), cMade = h.indexOf('생성일'), cSt = h.indexOf('상태');
  var made = 0;
  for (var r = 1; r < data.length; r++) {
    if (!data[r][cNo] || data[r][cId]) continue;
    var id = newId_();
    sh.getRange(r + 1, cId + 1).setValue(id);
    sh.getRange(r + 1, cLink + 1).setValue(SITE_URL + '?id=' + id);
    if (!data[r][cMade]) sh.getRange(r + 1, cMade + 1).setValue(new Date());
    if (!data[r][cSt]) sh.getRange(r + 1, cSt + 1).setValue('입금 대기');
    made++;
  }
  try { SpreadsheetApp.getActive().toast(made ? made + '건의 링크를 만들었어요.' : '새로 만들 링크가 없어요.'); } catch (e) {}
}

function testSlack() {
  var ok = notifySlack_(':white_check_mark: PLAB 청구서 슬랙 알림 테스트입니다.');
  SpreadsheetApp.getUi().alert(ok ? '슬랙으로 보냈어요.' : 'SLACK_WEBHOOK 스크립트 속성을 확인해주세요.');
}

/* ───────────────────────── 편집 시 자동 채우기 ───────────────────────── */

/**
 * 단순 트리거 — 따로 설치할 필요 없음.
 * 청구서 탭: 거래처 상호를 적으면 문서번호·발행일·지급기한·상태·링크를 채우고,
 *            거래처 탭에 있는 상호면 사업자번호·대표자도 채웁니다. (이미 값이 있는 칸은 그대로)
 * 청구서·품목·거래처 탭이 바뀌면 구분과 거래처 집계를 다시 계산합니다.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var name = e.range.getSheet().getName();
  if (name === SHEETS.invoice.name) fillInvoiceRows_(e.range.getRow(), e.range.getLastRow());
  if ([SHEETS.invoice.name, SHEETS.item.name, SHEETS.client.name].indexOf(name) >= 0) refreshAll();
}

function fillInvoiceRows_(first, last) {
  first = Math.max(2, first);
  if (last < first) return;
  var sh = sheet_('invoice');
  var lastCol = sh.getLastColumn();
  var h = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var c = function (n) { return h.indexOf(n); };
  if (c('거래처 상호') < 0) return;

  var byName = {};
  rowsByHeader_(sheet_('client')).forEach(function (cl) {
    var n = String(cl['상호'] || '').trim();
    if (n && digits_(cl['사업자번호'])) byName[n] = cl;
  });

  var all = sh.getRange(1, 1, sh.getLastRow(), lastCol).getValues();
  for (var r = first; r <= last; r++) {
    var row = all[r - 1];
    var nm = String(row[c('거래처 상호')] || '').trim();
    if (!nm) continue;

    var known = byName[nm];
    if (known) {
      if (!digits_(row[c('거래처 사업자번호')])) setCell_(sh, r, c('거래처 사업자번호'), fmtBiz_(known['사업자번호']));
      if (!String(row[c('거래처 대표자')] || '').trim() && known['대표자']) setCell_(sh, r, c('거래처 대표자'), known['대표자']);
    }

    var issued = row[c('발행일')] instanceof Date ? row[c('발행일')] : null;
    if (!issued) { issued = new Date(); issued.setHours(0, 0, 0, 0); setCell_(sh, r, c('발행일'), issued); }
    if (!row[c('지급기한')]) setCell_(sh, r, c('지급기한'), new Date(issued.getTime() + 7 * 864e5));
    if (!row[c('상태')]) setCell_(sh, r, c('상태'), '입금 대기');
    if (!String(row[c('문서번호')] || '').trim()) {
      var no = nextDocNo_(all, c('문서번호'), issued.getFullYear());
      setCell_(sh, r, c('문서번호'), no);
      all[r - 1][c('문서번호')] = no;
    }
    if (!row[c('링크ID')]) {
      var id = newId_();
      setCell_(sh, r, c('링크ID'), id);
      setCell_(sh, r, c('링크'), SITE_URL + '?id=' + id);
      if (!row[c('생성일')]) setCell_(sh, r, c('생성일'), new Date());
    }
  }
}

/** 거래처 탭의 집계 칸과 청구서 탭의 구분 칸을 다시 계산합니다. */
function refreshAll() {
  var ctx = loadAll_();

  // 거래처 탭
  var cl = sheet_('client');
  if (ctx.clients.length) {
    var ch = cl.getRange(1, 1, 1, cl.getLastColumn()).getValues()[0];
    var cols = {};
    ['구분', '등록 경로', '첫 거래일', '거래 건수', '최근 청구일', '누적 청구액'].forEach(function (n) { cols[n] = ch.indexOf(n); });
    var out = {};
    Object.keys(cols).forEach(function (n) { out[n] = []; });
    ctx.clients.forEach(function (c) {
      var s = ctx.stats[c.biz] || { count: 0, first: '', last: '', total: 0 };
      var route = String(c.obj['등록 경로'] || '') || (c.hasDocs ? '기존 자료' : '');
      out['등록 경로'].push([route]);
      out['구분'].push([!c.biz ? '' : !c.hasDocs ? '서류 대기' : clientType_(route, s.count)]);
      out['첫 거래일'].push([s.first || '']);
      out['거래 건수'].push([s.count]);
      out['최근 청구일'].push([s.last || '']);
      out['누적 청구액'].push([s.total]);
    });
    Object.keys(cols).forEach(function (n) {
      if (cols[n] >= 0) cl.getRange(2, cols[n] + 1, out[n].length, 1).setValues(out[n]);
    });
  }

  // 청구서 탭 구분
  var inv = sheet_('invoice');
  var gc = colOf_(inv, '구분');
  if (gc && ctx.invoices.length) {
    inv.getRange(2, gc, ctx.invoices.length, 1).setValues(ctx.invoices.map(function (iv) {
      if (!String(iv.obj['거래처 상호'] || '').trim()) return [''];
      return [invoiceType_(iv, ctx)];
    }));
  }
}

/* ───────────────────────── 웹앱 API ───────────────────────── */

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'invoice') return json_(getInvoice_(p.id));
    if (p.action === 'asset') return json_(getAsset_(p.id, p.kind));
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (body.action === 'inquiry') return json_(postInquiry_(body));
    if (body.action === 'checkBiz') return json_(checkBiz_(body));
    if (body.action === 'register') return json_(postRegister_(body));
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/* ───────────────────────── 동작 ───────────────────────── */

function getInvoice_(id) {
  var ctx = loadAll_();
  var iv = ctx.byId[String(id || '').trim()];
  if (!iv || !/^[A-Za-z0-9_-]{16,64}$/.test(String(id))) return { ok: false, error: 'not_found' };
  var o = iv.obj;
  if (o['상태'] === '취소') return { ok: false, error: 'cancelled' };

  var cl = iv.biz ? ctx.clientByBiz[iv.biz] : null;
  var registered = !!(cl && cl.hasDocs);
  var s = (iv.biz && ctx.stats[iv.biz]) || { count: 1 };
  var type = registered ? (clientType_(String(cl.obj['등록 경로'] || '기존 자료'), s.count) === '기존' ? 'existing' : 'new') : 'new';

  var history = !iv.biz ? [] : ctx.invoices
    .filter(function (x) { return x.biz === iv.biz && x !== iv && x.obj['상태'] !== '취소' && x.obj['링크ID']; })
    .sort(function (a, b) { return String(fmtDate_(b.obj['발행일'])).localeCompare(String(fmtDate_(a.obj['발행일']))); })
    .slice(0, 6)
    .map(function (x) {
      return { id: String(x.obj['링크ID']), no: String(x.obj['문서번호']), title: x.title, issued: fmtDate_(x.obj['발행일']), total: x.total, status: String(x.obj['상태'] || '입금 대기') };
    });

  return {
    ok: true,
    invoice: {
      no: String(o['문서번호']),
      issued: fmtDate_(o['발행일']),
      due: fmtDate_(o['지급기한']),
      title: iv.title,
      memo: String(o['비고'] || ''),
      status: String(o['상태'] || '입금 대기'),
      client: { name: String(o['거래처 상호'] || ''), ceo: String(o['거래처 대표자'] || (cl && cl.obj['대표자']) || ''), bizNo: fmtBiz_(iv.biz) },
      items: iv.items, supply: iv.supply, discount: iv.discount, vat: iv.vat, total: iv.total
    },
    registered: registered,
    client: { type: type, count: s.count || 1, history: history }
  };
}

/** 직인·프랩 서류는 공개 레포에 두지 않고, 유효한 청구서 링크로만 내려줍니다. */
function getAsset_(id, kind) {
  if (!findInvoice_(id)) return { ok: false, error: 'not_found' };
  var key = { seal: 'SEAL_FILE_ID', biz: 'PLAB_BIZ_FILE_ID', bank: 'PLAB_BANK_FILE_ID' }[kind];
  if (!key) return { ok: false, error: 'bad_kind' };
  var fileId = prop_(key);
  if (!fileId) return { ok: false, error: 'not_configured' };
  var blob = DriveApp.getFileById(fileId).getBlob();
  return { ok: true, name: blob.getName(), mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}

function postInquiry_(b) {
  var row = findInvoice_(b.id);
  if (!row) return { ok: false, error: 'not_found' };
  var msg = String(b.message || '').trim().slice(0, 2000);
  if (!msg) return { ok: false, error: 'empty' };
  var o = row.obj;
  var kind = String(o['구분'] || '신규');
  appendByHeader_(sheet_('inquiry'), { '일시': new Date(), '문서번호': o['문서번호'], '거래처': o['거래처 상호'], '구분': kind, '내용': msg, '처리': '' });
  notifySlack_(tag_(kind) + ' :speech_balloon: *청구서 문의* — ' + o['거래처 상호'] + ' (' + o['문서번호'] + ')\n>' + msg.replace(/\n/g, '\n>'));
  return { ok: true };
}

function checkBiz_(b) {
  if (!findInvoice_(b.id)) return { ok: false, error: 'not_found' };
  var d = digits_(b.bizNo);
  if (!validBiz_(d)) return { ok: false, error: 'invalid_biz' };
  return { ok: true, registered: isRegistered_(d) };
}

function postRegister_(b) {
  var row = findInvoice_(b.id);
  if (!row) return { ok: false, error: 'not_found' };
  var d = digits_(b.bizNo);
  if (!validBiz_(d)) return { ok: false, error: 'invalid_biz' };
  var ceo = String(b.ceo || '').trim().slice(0, 40);
  var o = row.obj;
  var inv = sheet_('invoice');
  var linkInvoice = function (ceoName) {
    if (!digits_(o['거래처 사업자번호'])) inv.getRange(row.rowIndex, colOf_(inv, '거래처 사업자번호')).setValue(fmtBiz_(d));
    if (ceoName && !String(o['거래처 대표자'] || '').trim()) inv.getRange(row.rowIndex, colOf_(inv, '거래처 대표자')).setValue(ceoName);
  };

  // 이미 서류가 있는 거래처면 청구서에 연결만
  var existing = rowsByHeader_(sheet_('client')).filter(function (c) { return digits_(c['사업자번호']) === d; })[0];
  if (existing && existing['사업자등록증'] && existing['통장사본']) {
    linkInvoice(existing['대표자'] || ceo);
    refreshAll();
    return { ok: true, already: true };
  }

  var files = b.files || [];
  var byKind = {};
  files.forEach(function (f) { byKind[f.kind] = f; });
  if (!byKind.biz || !byKind.bank) return { ok: false, error: 'missing_files' };

  var root = DriveApp.getFolderById(prop_('DOCS_FOLDER_ID'));
  var folderName = (o['거래처 상호'] || '거래처') + '_' + fmtBiz_(d);
  var it = root.getFoldersByName(folderName);
  var folder = it.hasNext() ? it.next() : root.createFolder(folderName);

  var urls = {};
  ['biz', 'bank'].forEach(function (k) {
    var f = byKind[k];
    var mime = String(f.mime || '').toLowerCase();
    if (ALLOWED_MIME.indexOf(mime) < 0) throw new Error('bad_type');
    var bytes = Utilities.base64Decode(String(f.data || ''));
    if (!bytes.length || bytes.length > MAX_FILE_BYTES) throw new Error('bad_size');
    var ext = (String(f.name || '').match(/\.[a-z0-9]{2,5}$/i) || [''])[0];
    var label = k === 'biz' ? '사업자등록증' : '통장사본';
    urls[k] = folder.createFile(Utilities.newBlob(bytes, mime, label + '_' + fmtBiz_(d) + ext)).getUrl();
  });

  var cl = sheet_('client');
  if (existing) {
    // 거래처 탭에 사업자번호만 있고 서류가 없던 줄 → 채우기
    var rIdx = rowIndexByBiz_(cl, d);
    var put = function (n, v) { var col = colOf_(cl, n); if (col && v !== '' && v != null) cl.getRange(rIdx, col).setValue(v); };
    put('사업자등록증', urls.biz); put('통장사본', urls.bank); put('제출일', new Date());
    put('등록 경로', '사이트 제출'); if (!existing['대표자']) put('대표자', ceo); if (!existing['상호']) put('상호', o['거래처 상호']);
  } else {
    appendByHeader_(cl, {
      '사업자번호': fmtBiz_(d), '상호': o['거래처 상호'], '대표자': ceo, '등록 경로': '사이트 제출',
      '사업자등록증': urls.biz, '통장사본': urls.bank, '제출일': new Date(), '청구서 문서번호': o['문서번호']
    });
  }
  linkInvoice(ceo);
  refreshAll();

  notifySlack_(':new: *신규 거래처 서류 제출* — ' + o['거래처 상호'] + (ceo ? ' (대표 ' + ceo + ')' : '') + ' · ' + fmtBiz_(d) + '\n'
    + '청구서 ' + o['문서번호'] + ' · <' + folder.getUrl() + '|서류 폴더 열기>');
  return { ok: true };
}

/* ───────────────────────── 계산 ───────────────────────── */

/** 청구서·품목·거래처를 한 번에 읽어 금액과 거래처 통계를 계산합니다. */
function loadAll_() {
  var itemsByNo = {};
  rowsByHeader_(sheet_('item')).forEach(function (it) {
    var no = String(it['문서번호'] || '').trim();
    if (!no || !it['품목명']) return;
    var q = num_(it['수량']) || 1, pr = num_(it['단가']);
    (itemsByNo[no] = itemsByNo[no] || []).push({ name: String(it['품목명']), qty: q, price: pr, amount: q * pr });
  });

  var invoices = rowsByHeader_(sheet_('invoice')).map(function (o) {
    var no = String(o['문서번호'] || '').trim();
    var items = itemsByNo[no] || [];
    if (!items.length && num_(o['금액'])) {
      var amt = num_(o['금액']);
      items = [{ name: String(o['건명'] || '용역'), qty: 1, price: amt, amount: amt }];
    }
    var supply = items.reduce(function (s, it) { return s + it.amount; }, 0);
    var discount = num_(o['할인']);
    var vat = Math.round((supply - discount) * 0.1);
    return {
      obj: o, biz: digits_(o['거래처 사업자번호']), items: items,
      title: String(o['건명'] || (items[0] && items[0].name) || ''),
      supply: supply, discount: discount, vat: vat, total: supply - discount + vat
    };
  });

  var stats = {}, byId = {};
  invoices.forEach(function (iv) {
    if (iv.obj['링크ID']) byId[String(iv.obj['링크ID'])] = iv;
    if (!iv.biz || iv.obj['상태'] === '취소' || !String(iv.obj['문서번호'] || '').trim()) return;
    var d = fmtDate_(iv.obj['발행일']);
    var s = stats[iv.biz] = stats[iv.biz] || { count: 0, first: '', last: '', total: 0 };
    s.count++; s.total += iv.total;
    if (d && (!s.first || d < s.first)) s.first = d;
    if (d && (!s.last || d > s.last)) s.last = d;
  });

  var clientByBiz = {};
  var clients = rowsByHeader_(sheet_('client')).map(function (o) {
    var c = { obj: o, biz: digits_(o['사업자번호']), hasDocs: !!(o['사업자등록증'] && o['통장사본']) };
    if (c.biz) clientByBiz[c.biz] = c;
    return c;
  });

  return { invoices: invoices, stats: stats, byId: byId, clients: clients, clientByBiz: clientByBiz };
}

function clientType_(route, count) {
  return route === '기존 자료' || count >= 2 ? '기존' : '신규';
}

function invoiceType_(iv, ctx) {
  if (!iv.biz) return '신규 · 서류 대기';
  var cl = ctx.clientByBiz[iv.biz];
  if (!cl || !cl.hasDocs) return '서류 대기';
  var s = ctx.stats[iv.biz] || { count: 1 };
  return clientType_(String(cl.obj['등록 경로'] || '기존 자료'), s.count);
}

function tag_(kind) { return /기존/.test(kind) ? '[기존]' : '[신규]'; }

/* ───────────────────────── 도우미 ───────────────────────── */

function sheet_(key) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS[key].name);
  if (!sh) throw new Error('sheet_missing:' + SHEETS[key].name);
  return sh;
}

function rowsByHeader_(sh) {
  var v = sh.getDataRange().getValues();
  var h = v.shift() || [];
  return v.map(function (r) { var o = {}; h.forEach(function (k, i) { o[k] = r[i]; }); return o; });
}

function appendByHeader_(sh, obj) {
  var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(h.map(function (k) { return obj.hasOwnProperty(k) ? obj[k] : ''; }));
}

function colOf_(sh, header) {
  if (sh.getLastColumn() < 1) return 0;
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf(header) + 1;
}

function setCell_(sh, r, colIdx, value) {
  if (colIdx < 0) return;
  sh.getRange(r, colIdx + 1).setValue(value);
}

function rowIndexByBiz_(sh, d) {
  var c = colOf_(sh, '사업자번호');
  var v = sh.getRange(1, c, sh.getLastRow(), 1).getValues();
  for (var i = 1; i < v.length; i++) if (digits_(v[i][0]) === d) return i + 1;
  return -1;
}

function findInvoice_(id) {
  id = String(id || '').trim();
  if (!/^[A-Za-z0-9_-]{16,64}$/.test(id)) return null;
  var sh = sheet_('invoice');
  var v = sh.getDataRange().getValues();
  var h = v[0], c = h.indexOf('링크ID');
  for (var r = 1; r < v.length; r++) {
    if (String(v[r][c]) === id) {
      var o = {}; h.forEach(function (k, i) { o[k] = v[r][i]; });
      return { rowIndex: r + 1, obj: o };
    }
  }
  return null;
}

function isRegistered_(d) {
  return rowsByHeader_(sheet_('client')).some(function (c) {
    return digits_(c['사업자번호']) === d && c['사업자등록증'] && c['통장사본'];
  });
}

/** 그 해의 PLAB-연도-NNNN 중 가장 큰 번호 + 1 */
function nextDocNo_(rows, col, year) {
  var re = new RegExp('^PLAB-' + year + '-(\\d+)$'), max = 0;
  rows.forEach(function (row) {
    var m = String(row[col]).trim().match(re);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return 'PLAB-' + year + '-' + ('000' + (max + 1)).slice(-4);
}

function newId_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 12) + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

function validBiz_(d) {
  if (!/^\d{10}$/.test(d)) return false;
  var w = [1, 3, 7, 1, 3, 7, 1, 3, 5], s = 0;
  for (var i = 0; i < 9; i++) s += Number(d[i]) * w[i];
  s += Math.floor(Number(d[8]) * 5 / 10);
  return (10 - (s % 10)) % 10 === Number(d[9]);
}

function digits_(v) { return String(v || '').replace(/\D/g, ''); }
function fmtBiz_(d) { d = digits_(d); return d.length === 10 ? d.slice(0, 3) + '-' + d.slice(3, 5) + '-' + d.slice(5) : d; }
function num_(v) { var n = Number(String(v === '' || v == null ? 0 : v).replace(/[^\d.-]/g, '')); return isNaN(n) ? 0 : n; }
function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v || '');
}
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }

function notifySlack_(text) {
  var url = prop_('SLACK_WEBHOOK');
  if (!url) return false;
  try {
    UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ text: text }), muteHttpExceptions: true });
    return true;
  } catch (e) { return false; }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
