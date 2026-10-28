/**
 * PLAB 청구서 (tax.plab.life) — Apps Script 백엔드
 *
 * 이 스크립트는 'PLAB 청구서' 구글 시트에 붙어 있는(컨테이너 바인딩) 스크립트입니다.
 * 시트 메뉴 [PLAB 청구서]에서 초기 설정 · 링크 만들기를 할 수 있습니다.
 *
 * 스크립트 속성(프로젝트 설정 > 스크립트 속성)
 *   SLACK_WEBHOOK   : 문의·서류 제출 알림을 보낼 슬랙 웹훅 URL
 *   DOCS_FOLDER_ID  : 거래처가 올린 서류를 저장할 드라이브 폴더 ID (운영팀만 접근)
 *   SEAL_FILE_ID    : 직인 이미지(PNG) 드라이브 파일 ID
 *   PLAB_BIZ_FILE_ID: 프랩 사업자등록증 파일 ID (거래처 '프랩 서류 받기'용)
 *   PLAB_BANK_FILE_ID: 프랩 통장사본 파일 ID
 */

var SITE_URL = 'https://tax.plab.life/';
var TZ = 'Asia/Seoul';

var SHEETS = {
  invoice: { name: '청구서', headers: ['링크ID', '문서번호', '발행일', '지급기한', '거래처 상호', '거래처 대표자', '거래처 사업자번호', '건명', '금액', '할인', '비고', '상태', '링크', '생성일'] },
  item:    { name: '품목',   headers: ['문서번호', '품목명', '수량', '단가'] },
  client:  { name: '거래처', headers: ['사업자번호', '상호', '제출일', '사업자등록증', '통장사본', '청구서 문서번호'] },
  inquiry: { name: '문의',   headers: ['일시', '문서번호', '거래처', '내용', '처리'] }
};

var MAX_FILE_BYTES = 10 * 1024 * 1024;
var ALLOWED_MIME = ['image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif', 'application/pdf'];

/* ───────────────────────── 메뉴 ───────────────────────── */

function onOpen() {
  try { ensureHeaders_(); } catch (e) {}
  SpreadsheetApp.getUi().createMenu('PLAB 청구서')
    .addItem('링크 만들기 (ID 없는 행)', 'makeLinks')
    .addSeparator()
    .addItem('초기 설정 (시트 만들기)', 'setup')
    .addItem('슬랙 알림 테스트', 'testSlack')
    .addToUi();
}

/** 시트 4개와 머리글을 만들고 예시 청구서를 한 건 넣습니다. 이미 있는 시트는 건드리지 않습니다. */
function setup() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(function (k) {
    var def = SHEETS[k];
    var sh = ss.getSheetByName(def.name);
    if (!sh) {
      sh = ss.insertSheet(def.name);
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold').setBackground('#EEF1F7');
      sh.setFrozenRows(1);
    }
  });
  ensureHeaders_();
  var inv = ss.getSheetByName(SHEETS.invoice.name);
  if (inv.getLastRow() < 2) {
    var today = new Date();
    var due = new Date(today.getTime() + 7 * 864e5);
    appendByHeader_(inv, {
      '문서번호': 'PLAB-2026-0001', '발행일': today, '지급기한': due,
      '거래처 상호': '예시상회', '거래처 대표자': '홍길동', '거래처 사업자번호': '',
      '건명': 'AI 상담 챗봇 구축', '금액': 3000000, '할인': 0, '비고': '예시 청구서입니다. 지우고 쓰셔도 돼요.', '상태': '입금 대기'
    });
    inv.getRange(2, 3, 1, 2).setNumberFormat('yyyy-mm-dd');
  }
  var st = inv.getRange(2, colOf_(inv, '상태'), inv.getMaxRows() - 1, 1);
  st.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['입금 대기', '입금 완료', '취소'], true).build());
  makeLinks();
}

/** 예전에 만든 시트에 새로 생긴 머리글(예: 금액)이 없으면, 정해진 순서 자리에 열을 끼워 넣습니다. */
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
      sh.getRange(1, col).setValue(name).setFontWeight('bold').setBackground('#EEF1F7');
    });
  });
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

/**
 * 청구서 탭에서 거래처 상호를 적으면 자동으로 채웁니다. (단순 트리거 — 따로 설치할 필요 없음)
 *   문서번호(PLAB-연도-순번) · 발행일(오늘) · 지급기한(발행일+7일) · 상태(입금 대기) · 링크ID · 링크 · 생성일
 * 이미 값이 있는 칸은 건드리지 않습니다.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== SHEETS.invoice.name) return;
  var first = Math.max(2, e.range.getRow()), last = e.range.getLastRow();
  if (last < 2) return;

  var lastCol = sh.getLastColumn();
  var h = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var c = function (name) { return h.indexOf(name); };
  if (c('거래처 상호') < 0 || c('문서번호') < 0) return;

  var all = sh.getRange(1, 1, sh.getLastRow(), lastCol).getValues();
  for (var r = first; r <= last; r++) {
    var row = all[r - 1];
    if (!String(row[c('거래처 상호')]).trim()) continue;

    var issued = row[c('발행일')] instanceof Date ? row[c('발행일')] : null;
    if (!issued) {
      issued = new Date(); issued.setHours(0, 0, 0, 0);
      setCell_(sh, r, c('발행일'), issued, 'yyyy-mm-dd');
    }
    if (!row[c('지급기한')]) setCell_(sh, r, c('지급기한'), new Date(issued.getTime() + 7 * 864e5), 'yyyy-mm-dd');
    if (!row[c('상태')]) setCell_(sh, r, c('상태'), '입금 대기');

    if (!String(row[c('문서번호')]).trim()) {
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

function setCell_(sh, r, colIdx, value, fmt) {
  if (colIdx < 0) return;
  var cell = sh.getRange(r, colIdx + 1);
  cell.setValue(value);
  if (fmt) cell.setNumberFormat(fmt);
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

function testSlack() {
  var ok = notifySlack_(':white_check_mark: PLAB 청구서 슬랙 알림 테스트입니다.');
  SpreadsheetApp.getUi().alert(ok ? '슬랙으로 보냈어요.' : 'SLACK_WEBHOOK 스크립트 속성을 확인해주세요.');
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
  var row = findInvoice_(id);
  if (!row) return { ok: false, error: 'not_found' };
  var o = row.obj;
  if (o['상태'] === '취소') return { ok: false, error: 'cancelled' };

  var items = rowsByHeader_(sheet_('item'))
    .filter(function (it) { return String(it['문서번호']).trim() === String(o['문서번호']).trim() && it['품목명']; })
    .map(function (it) {
      var q = num_(it['수량']) || 1, pr = num_(it['단가']);
      return { name: String(it['품목명']), qty: q, price: pr, amount: q * pr };
    });
  // 품목 탭에 줄이 없으면 청구서 탭의 건명·금액을 품목 1개로 씁니다.
  if (!items.length && num_(o['금액'])) {
    var amt = num_(o['금액']);
    items = [{ name: String(o['건명'] || '용역'), qty: 1, price: amt, amount: amt }];
  }
  var supply = items.reduce(function (s, it) { return s + it.amount; }, 0);
  var discount = num_(o['할인']);
  var vat = Math.round((supply - discount) * 0.1);
  var bizNo = digits_(o['거래처 사업자번호']);

  return {
    ok: true,
    invoice: {
      no: String(o['문서번호']),
      issued: fmtDate_(o['발행일']),
      due: fmtDate_(o['지급기한']),
      title: String(o['건명'] || (items[0] && items[0].name) || ''),
      memo: String(o['비고'] || ''),
      status: String(o['상태'] || '입금 대기'),
      client: { name: String(o['거래처 상호'] || ''), ceo: String(o['거래처 대표자'] || ''), bizNo: fmtBiz_(bizNo) },
      items: items, supply: supply, discount: discount, vat: vat, total: supply - discount + vat
    },
    registered: bizNo ? isRegistered_(bizNo) : false
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
  appendByHeader_(sheet_('inquiry'), { '일시': new Date(), '문서번호': o['문서번호'], '거래처': o['거래처 상호'], '내용': msg, '처리': '' });
  notifySlack_(':speech_balloon: *청구서 문의* — ' + o['거래처 상호'] + ' (' + o['문서번호'] + ')\n>' + msg.replace(/\n/g, '\n>'));
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
  var o = row.obj;
  var inv = sheet_('invoice');

  // 이미 서류가 있는 거래처면 청구서에 사업자번호만 연결
  if (isRegistered_(d)) {
    if (!digits_(o['거래처 사업자번호'])) inv.getRange(row.rowIndex, colOf_(inv, '거래처 사업자번호')).setValue(fmtBiz_(d));
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
    var file = folder.createFile(Utilities.newBlob(bytes, mime, label + '_' + fmtBiz_(d) + ext));
    urls[k] = file.getUrl();
  });

  appendByHeader_(sheet_('client'), {
    '사업자번호': fmtBiz_(d), '상호': o['거래처 상호'], '제출일': new Date(),
    '사업자등록증': urls.biz, '통장사본': urls.bank, '청구서 문서번호': o['문서번호']
  });
  if (!digits_(o['거래처 사업자번호'])) inv.getRange(row.rowIndex, colOf_(inv, '거래처 사업자번호')).setValue(fmtBiz_(d));

  notifySlack_(':page_facing_up: *기초 정보 제출* — ' + o['거래처 상호'] + ' (' + fmtBiz_(d) + ')\n'
    + '청구서 ' + o['문서번호'] + ' · <' + folder.getUrl() + '|서류 폴더 열기>');
  return { ok: true };
}

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
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf(header) + 1;
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
