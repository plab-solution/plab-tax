(function () {
  'use strict';

  var CFG = window.PLAB_CONFIG || {};
  var CO = CFG.company || {};
  var params = new URLSearchParams(location.search);
  var ID = (params.get('id') || '').trim();
  var DEMO = !CFG.API_URL || params.get('demo') === '1';

  var state = { inv: null, registered: false, client: { type: 'new', count: 1, history: [] }, biz: '', ceo: '', files: {}, seal: null };
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };

  /* ───────── API ───────── */

  function apiGet(qs) {
    if (DEMO) return demo('get', qs);
    var url = CFG.API_URL + (CFG.API_URL.indexOf('?') < 0 ? '?' : '&') + new URLSearchParams(qs).toString();
    return fetch(url, { redirect: 'follow' }).then(function (r) { return r.json(); });
  }
  function apiPost(body) {
    if (DEMO) return demo('post', body);
    return fetch(CFG.API_URL, {
      method: 'POST', redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body)
    }).then(function (r) { return r.json(); });
  }

  /* ───────── 형식 ───────── */

  function won(n) { return Number(n || 0).toLocaleString('ko-KR'); }
  function dot(d) { return d ? d.replace(/-/g, '.') : ''; }
  function dotSpaced(d) { if (!d) return ''; var p = d.split('-'); return p[0] + '. ' + Number(p[1]) + '. ' + Number(p[2]) + '.'; }
  function koDate(d) { if (!d) return ''; var p = d.split('-'); return p[0] + '년 ' + Number(p[1]) + '월 ' + Number(p[2]) + '일'; }
  function fmtBiz(d) {
    d = String(d || '').replace(/\D/g, '').slice(0, 10);
    if (d.length > 5) return d.slice(0, 3) + '-' + d.slice(3, 5) + '-' + d.slice(5);
    if (d.length > 3) return d.slice(0, 3) + '-' + d.slice(3);
    return d;
  }
  function validBiz(d) {
    if (!/^\d{10}$/.test(d)) return false;
    var w = [1, 3, 7, 1, 3, 7, 1, 3, 5], s = 0;
    for (var i = 0; i < 9; i++) s += +d[i] * w[i];
    s += Math.floor(+d[8] * 5 / 10);
    return (10 - (s % 10)) % 10 === +d[9];
  }
  // 3300000 → 삼백삼십만
  function koNum(n) {
    n = Math.floor(Math.abs(n));
    if (!n) return '영';
    var digit = ['', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구'];
    var small = ['', '십', '백', '천'], big = ['', '만', '억', '조'];
    var out = '', g = 0;
    while (n > 0) {
      var chunk = n % 10000, s = '';
      for (var i = 0; i < 4; i++) {
        var dgt = Math.floor(chunk / Math.pow(10, i)) % 10;
        if (dgt) s = digit[dgt] + small[i] + s;
      }
      if (s) out = s + big[g] + out;
      n = Math.floor(n / 10000); g++;
    }
    return out;
  }
  function daysLeft(due) {
    if (!due) return null;
    var t = new Date(); t.setHours(0, 0, 0, 0);
    var d = new Date(due + 'T00:00:00');
    return Math.round((d - t) / 864e5);
  }

  /* ───────── 화면 ───────── */

  var VIEWS = ['loading', 'error', 'home', 'invoice', 'biz', 'docs', 'done'];
  function show(v) {
    VIEWS.forEach(function (k) { $('#v-' + k).hidden = k !== v; });
    try { window.scrollTo(0, 0); } catch (e) {}
    if (v === 'invoice') loadSeal();
    if (v === 'biz') {
      // 시트에 이미 있는 사업자번호·대표자는 미리 채워둡니다.
      var c = state.inv && state.inv.client;
      if (c && c.bizNo && !$('#biz-input').value) { $('#biz-input').value = c.bizNo; state.biz = c.bizNo.replace(/\D/g, ''); }
      if (c && c.ceo && !$('#ceo-input').value) { $('#ceo-input').value = c.ceo; state.ceo = c.ceo; }
      $('#biz-next').disabled = !(state.biz.length === 10 && state.ceo.length >= 2);
      setTimeout(function () { $('#biz-input').focus(); }, 50);
    }
  }
  function toast(msg) {
    var t = $('#toast'); t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(function () { t.hidden = true; }, 2200);
  }
  function fail(title, desc) {
    $('#err-title').textContent = title;
    if (desc) $('#err-desc').textContent = desc;
    show('error');
  }

  function fillCompany() {
    CO.bankShort = String(CO.bank || '').replace(/은행$/, '');
    $$('[data-co]').forEach(function (el) { el.textContent = CO[el.getAttribute('data-co')] || ''; });
    $$('[data-co-href="tel"]').forEach(function (a) { a.href = 'tel:' + String(CO.phone || '').replace(/\D/g, ''); });
    $$('[data-co-href="mail"]').forEach(function (a) { a.href = 'mailto:' + (CO.email || ''); });
  }

  var STEPS = ['청구서 확인', '입금', '서류 제출', '세금계산서 발행'];

  function render() {
    var inv = state.inv;
    var c = inv.client;
    var info = [c.name !== (c.legalName || c.name) ? c.name : '', c.ceo ? '대표 ' + c.ceo : '', c.bizNo].filter(Boolean).join(' · ');
    var f = {
      no: inv.no, title: inv.title, memo: inv.memo,
      clientName: c.name, clientLegal: c.legalName || c.name, clientInfo: info,
      discountFmt: won(inv.discount), vatFmt: won(inv.vat), totalFmt: won(inv.total),
      totalKo: koNum(inv.total), issuedDot: dotSpaced(inv.issued), dueDot: dotSpaced(inv.due), issuedKo: koDate(inv.issued)
    };
    $$('[data-f]').forEach(function (el) { var k = el.getAttribute('data-f'); el.textContent = f[k] == null ? '' : f[k]; });
    document.title = '청구서 ' + inv.no + ' · ' + (CO.name || '');

    // 단계 상태
    var paid = inv.status === '입금 완료';
    var docs = state.registered;
    var tax = !!inv.taxIssued;
    var done = [true, paid, docs, tax];
    var cur = done.indexOf(false); // -1 이면 모두 완료
    renderTracker(done, cur);

    $('#now-pay').hidden = cur !== 1;
    $('#now-docs').hidden = cur !== 2;
    $('#now-tax').hidden = cur !== 3;
    $('#now-done').hidden = cur !== -1;
    $('#also-docs').hidden = !(cur === 1 && !docs);
    if (tax) $('#tax-date').textContent = koDate(inv.taxIssued);

    var left = daysLeft(inv.due);
    var due = inv.due ? inv.due.split('-') : null;
    var wd = inv.due ? '일월화수목금토'[new Date(inv.due + 'T00:00:00').getDay()] : '';
    $('#due-line').textContent = !due ? '' : Number(due[1]) + '월 ' + Number(due[2]) + '일(' + wd + ')까지'
      + (left == null ? '' : left > 0 ? ' · ' + left + '일 남았어요' : left === 0 ? ' · 오늘까지예요' : ' · 기한이 ' + (-left) + '일 지났어요');

    // 기존 거래처
    var existing = state.client.type === 'existing';
    $('#hello').hidden = !existing;
    $('#hello-sub').textContent = (state.client.count || 1) > 1 ? '다시 찾아주셔서 감사해요 · ' + state.client.count + '번째 거래' : '늘 함께해주셔서 감사해요';
    if (existing) {
      $('#reg-desc').textContent = '세금계산서 발행용 사업자등록증과 통장 사본을 한 번만 보내주세요. 다음 청구부터는 이 단계가 나타나지 않아요.';
    }

    // 내역
    var lines = $('#item-lines'); lines.innerHTML = '';
    inv.items.forEach(function (it) {
      var r = document.createElement('div'); r.className = 'row';
      var a = document.createElement('span'); a.textContent = it.name + (it.qty > 1 ? ' × ' + it.qty : '');
      var b = document.createElement('span'); b.className = 'num'; b.textContent = won(it.amount);
      r.appendChild(a); r.appendChild(b); lines.appendChild(r);
    });
    $('#item-count').textContent = inv.items.length + '개 품목';
    $('#disc-row').hidden = !inv.discount;
    if (!inv.discount) $('#vat-row').classList.add('dashed'); else $('#vat-row').classList.remove('dashed');
    $('#memo-box').hidden = !inv.memo;
    $('#p-memo-wrap').hidden = !inv.memo;

    // 원본 표
    var tb = $('#tbl-body'); tb.innerHTML = '';
    inv.items.forEach(function (it) {
      tb.appendChild(tr([[it.name], [won(it.qty), 'r'], [won(it.price), 'r'], [won(it.amount), 'r']]));
    });
    if (inv.discount) tb.appendChild(tr([['할인', '', 3], ['−' + won(inv.discount), 'r']], 'sub'));
    tb.appendChild(tr([['공급가액', '', 3], [won(inv.supply - inv.discount), 'r']], 'sub'));
    tb.appendChild(tr([['부가세 10%', '', 3], [won(inv.vat), 'r']], 'sub'));
    tb.appendChild(tr([['합계', '', 3], [won(inv.total) + '원', 'r']], 'grand'));

    renderHistory();
  }

  function renderTracker(done, cur) {
    var ol = $('#tracker'); ol.innerHTML = '';
    var last = cur === -1 ? STEPS.length - 1 : cur;
    var fill = document.createElement('span'); fill.className = 'fill';
    fill.style.width = (last / (STEPS.length - 1) * 75) + '%';
    ol.appendChild(fill);
    STEPS.forEach(function (label, i) {
      var li = document.createElement('li');
      var isDone = done[i] && i !== cur, isNow = i === cur;
      li.className = isDone ? 'is-done' : isNow ? 'is-now' : '';
      if (isNow) li.setAttribute('aria-current', 'step');
      var dot = document.createElement('span'); dot.className = 'dot';
      if (isDone) dot.innerHTML = '<svg aria-hidden="true" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12l5 5 9-10"/></svg>';
      else dot.textContent = i + 1;
      var lab = document.createElement('span'); lab.className = 'lab'; lab.textContent = label;
      li.appendChild(dot); li.appendChild(lab);
      var sr = document.createElement('span'); sr.className = 'sr'; sr.textContent = isDone ? ' 완료' : isNow ? ' 진행 중' : ' 대기';
      li.appendChild(sr);
      ol.appendChild(li);
    });
  }

  function renderHistory() {
    var list = state.client.history || [];
    $('#history').hidden = !list.length;
    if (!list.length) return;
    $('#history-n').textContent = list.length + '건';
    var box = $('#history-list'); box.innerHTML = '';
    list.forEach(function (h) {
      var a = document.createElement('a'); a.className = 'hist'; a.href = '?id=' + encodeURIComponent(h.id) + (DEMO ? '&demo=1&as=existing' : '');
      var l = document.createElement('span'); l.className = 'stack g2';
      var meta = document.createElement('span'); meta.className = 'meta num'; meta.textContent = h.no + ' · ' + dot(h.issued);
      var t = document.createElement('span'); t.className = 't'; t.textContent = h.title;
      l.appendChild(meta); l.appendChild(t);
      var r = document.createElement('span'); r.className = 'r';
      var amt = document.createElement('span'); amt.className = 'amt num'; amt.textContent = won(h.total) + '원';
      var paid = h.status === '입금 완료';
      var st = document.createElement('span'); st.className = 'st ' + (paid ? 'paid' : 'wait'); st.textContent = paid ? '입금 완료' : '입금 전';
      r.appendChild(amt); r.appendChild(st);
      a.appendChild(l); a.appendChild(r); box.appendChild(a);
    });
  }
  function tr(cells, cls) {
    var row = document.createElement('tr'); if (cls) row.className = cls;
    cells.forEach(function (c) {
      var td = document.createElement('td'); td.textContent = c[0];
      if (c[1]) td.className = c[1]; if (c[2]) td.colSpan = c[2];
      row.appendChild(td);
    });
    return row;
  }

  /* ───────── 직인·서류 ───────── */

  function loadSeal() {
    if (state.seal) return Promise.resolve();
    return apiGet({ action: 'asset', id: ID, kind: 'seal' }).then(function (r) {
      if (!r.ok) return;
      state.seal = 'data:' + r.mime + ';base64,' + r.data;
      var img = $('#seal'); img.src = state.seal; img.hidden = false;
    }).catch(function () {});
  }

  function downloadAsset(kind, btn) {
    var label = btn.textContent;
    btn.disabled = true; btn.textContent = '받는 중…';
    apiGet({ action: 'asset', id: ID, kind: kind }).then(function (r) {
      if (!r.ok) throw new Error(r.error);
      var bin = atob(r.data), u8 = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
      var url = URL.createObjectURL(new Blob([u8], { type: r.mime }));
      var a = document.createElement('a'); a.href = url; a.download = r.name || (label + '.pdf');
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
    }).catch(function () {
      toast('서류를 받지 못했어요. 문의하기로 요청해주세요.');
    }).then(function () { btn.disabled = false; btn.textContent = label; });
  }

  /* ───────── 파일 처리 ───────── */

  var MAX = 10 * 1024 * 1024;
  function readFile(file) {
    var isImg = /^image\/(jpeg|png|webp)$/.test(file.type);
    if (isImg && file.size > 1.5 * 1024 * 1024) return shrink(file);
    return new Promise(function (res, rej) {
      if (file.size > MAX) return rej(new Error('size'));
      var fr = new FileReader();
      fr.onload = function () { res({ name: file.name, mime: file.type || 'application/octet-stream', data: String(fr.result).split(',')[1] }); };
      fr.onerror = rej; fr.readAsDataURL(file);
    });
  }
  function shrink(file) {
    return new Promise(function (res, rej) {
      var img = new Image(), url = URL.createObjectURL(file);
      img.onload = function () {
        var s = Math.min(1, 2200 / Math.max(img.width, img.height));
        var c = document.createElement('canvas'); c.width = Math.round(img.width * s); c.height = Math.round(img.height * s);
        var g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(img, 0, 0, c.width, c.height);
        URL.revokeObjectURL(url);
        var data = c.toDataURL('image/jpeg', 0.85).split(',')[1];
        res({ name: file.name.replace(/\.[^.]+$/, '') + '.jpg', mime: 'image/jpeg', data: data });
      };
      img.onerror = function () { URL.revokeObjectURL(url); rej(new Error('image')); };
      img.src = url;
    });
  }

  /* ───────── 이벤트 ───────── */

  function bind() {
    $$('[data-go]').forEach(function (b) { b.addEventListener('click', function () { show(b.getAttribute('data-go')); }); });
    $$('[data-asset]').forEach(function (b) { b.addEventListener('click', function () { downloadAsset(b.getAttribute('data-asset'), b); }); });

    $('#btn-copy').addEventListener('click', function () {
      var text = (CO.bank || '') + ' ' + (CO.account || '');
      var done = function () { $('#copy-label').textContent = '복사됨'; toast('계좌번호를 복사했어요'); setTimeout(function () { $('#copy-label').textContent = '계좌 복사'; }, 1600); };
      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(done, function () { legacyCopy(text); done(); });
      else { legacyCopy(text); done(); }
    });

    $('#btn-print').addEventListener('click', function () {
      loadSeal().then(function () { setTimeout(function () { window.print(); }, 150); });
    });

    // 문의
    var ta = $('#ask-text'), send = $('#ask-send');
    var toggle = function (btn, body) {
      btn.addEventListener('click', function () {
        var open = body.hidden;
        body.hidden = !open; btn.setAttribute('aria-expanded', String(open));
        if (open && body === $('#ask-box')) { $('#ask-sent').hidden = true; ta.focus(); }
      });
    };
    toggle($('#plab-docs-open'), $('#plab-docs'));
    toggle($('#ask-open'), $('#ask-box'));
    $('#ask-cancel').addEventListener('click', function () { $('#ask-box').hidden = true; $('#ask-open').setAttribute('aria-expanded', 'false'); });
    ta.addEventListener('input', function () { send.disabled = !ta.value.trim(); });
    send.addEventListener('click', function () {
      send.disabled = true; send.textContent = '보내는 중…';
      apiPost({ action: 'inquiry', id: ID, message: ta.value.trim() }).then(function (r) {
        if (!r.ok) throw new Error(r.error);
        ta.value = ''; $('#ask-box').hidden = true; $('#ask-sent').hidden = false; $('#ask-open').setAttribute('aria-expanded', 'false');
      }).catch(function () {
        toast('보내지 못했어요. 잠시 후 다시 시도해주세요.');
        send.disabled = false;
      }).then(function () { send.textContent = '보내기'; });
    });

    // 1단계
    var bi = $('#biz-input'), ci = $('#ceo-input'), next = $('#biz-next');
    var step1Ready = function () { next.disabled = !(state.biz.length === 10 && state.ceo.length >= 2); };
    bi.addEventListener('input', function () {
      bi.value = fmtBiz(bi.value);
      state.biz = bi.value.replace(/\D/g, '');
      $('#biz-err').textContent = '';
      step1Ready();
    });
    ci.addEventListener('input', function () { state.ceo = ci.value.trim(); step1Ready(); });
    ci.addEventListener('keydown', function (e) { if (e.key === 'Enter' && !next.disabled) next.click(); });
    bi.addEventListener('keydown', function (e) { if (e.key === 'Enter' && !next.disabled) next.click(); });
    next.addEventListener('click', function () {
      if (!validBiz(state.biz)) { $('#biz-err').textContent = '사업자등록번호가 올바르지 않아요. 다시 확인해주세요.'; return; }
      next.disabled = true; next.textContent = '확인 중…';
      apiPost({ action: 'checkBiz', id: ID, bizNo: state.biz }).then(function (r) {
        if (!r.ok) throw new Error(r.error);
        if (r.registered) {
          // 이미 서류가 있는 거래처 → 이 청구서에 연결만
          return apiPost({ action: 'register', id: ID, bizNo: state.biz, ceo: state.ceo }).then(function () { finish(true); });
        }
        show('docs');
      }).catch(function (e) {
        $('#biz-err').textContent = e.message === 'invalid_biz' ? '사업자등록번호가 올바르지 않아요.' : '확인하지 못했어요. 잠시 후 다시 시도해주세요.';
      }).then(function () { next.textContent = '계속'; step1Ready(); });
    });

    // 2단계
    $$('#v-docs input[type=file]').forEach(function (inp) {
      inp.addEventListener('change', function () {
        var file = inp.files && inp.files[0]; if (!file) return;
        var kind = inp.getAttribute('data-kind'), box = $('#up-' + kind);
        $('#docs-err').textContent = '';
        box.querySelector('.fname').textContent = '불러오는 중…';
        readFile(file).then(function (f) {
          if (atob(f.data).length > MAX) throw new Error('size');
          f.kind = kind; state.files[kind] = f;
          box.querySelector('.fname').textContent = file.name;
          box.classList.add('done');
        }).catch(function (e) {
          delete state.files[kind]; box.classList.remove('done');
          box.querySelector('.fname').textContent = '눌러서 촬영하거나 파일 고르기';
          $('#docs-err').textContent = e.message === 'size' ? '10MB 이하 파일만 올릴 수 있어요.' : '파일을 읽지 못했어요. 다른 파일로 시도해주세요.';
        }).then(updateSubmit);
      });
    });
    var sub = $('#docs-submit');
    sub.addEventListener('click', function () {
      sub.disabled = true; sub.textContent = '보내는 중…';
      apiPost({ action: 'register', id: ID, bizNo: state.biz, ceo: state.ceo, files: [state.files.biz, state.files.bank] }).then(function (r) {
        if (!r.ok) throw new Error(r.error);
        finish(false);
      }).catch(function (e) {
        var m = { bad_type: '이미지나 PDF 파일만 올릴 수 있어요.', bad_size: '10MB 이하 파일만 올릴 수 있어요.' }[e.message];
        $('#docs-err').textContent = m || '제출하지 못했어요. 잠시 후 다시 시도해주세요.';
        updateSubmit();
      }).then(function () { sub.textContent = '보내기'; });
    });
  }
  function updateSubmit() { $('#docs-submit').disabled = !(state.files.biz && state.files.bank); }
  function finish(already) {
    state.registered = true;
    if (!state.inv.client.bizNo) state.inv.client.bizNo = fmtBiz(state.biz);
    if (!state.inv.client.ceo && state.ceo) state.inv.client.ceo = state.ceo;
    $('#done-biz').textContent = fmtBiz(state.biz);
    $('#done-title').textContent = already ? '이미 등록된 사업자예요' : '서류를 받았어요';
    render(); show('done');
  }
  function legacyCopy(text) {
    var t = document.createElement('textarea'); t.value = text; t.style.position = 'fixed'; t.style.opacity = '0';
    document.body.appendChild(t); t.select(); try { document.execCommand('copy'); } catch (e) {} t.remove();
  }

  /* ───────── 예시 데이터 (API_URL이 없거나 ?demo=1) ───────── */

  function demo(kind, q) {
    return new Promise(function (res) {
      setTimeout(function () {
        if (kind === 'get' && q.action === 'invoice') {
          var due = new Date(Date.now() + 7 * 864e5).toISOString().slice(0, 10);
          var today = new Date().toISOString().slice(0, 10);
          var ex = params.get('as') === 'existing' || params.get('as') === 'nodocs'; var nodocs = params.get('as') === 'nodocs';
          res({ ok: true, registered: ex && !nodocs,
            client: ex ? { type: 'existing', count: 3, history: [
              { id: 'demo-0010-aaaaaaaaaaaa', no: 'PLAB-2026-0010', title: '월 유지관리 (9월)', issued: '2026-09-30', total: 330000, status: '입금 완료' },
              { id: 'demo-0008-aaaaaaaaaaaa', no: 'PLAB-2026-0008', title: '월 유지관리 (8월)', issued: '2026-08-29', total: 330000, status: '입금 완료' }
            ] } : { type: 'new', count: 1, history: [] },
            invoice: {
            no: 'PLAB-2026-0012', issued: today, due: due, title: 'AI 상담 챗봇 구축', status: params.get('paid') ? '입금 완료' : '입금 대기', taxIssued: params.get('tax') ? today : '',
            memo: '1차 구축분 청구서입니다. 세금계산서는 입금 확인 후 발행됩니다.',
            client: { name: '한결푸드', ceo: ex ? '홍길동' : '', bizNo: ex ? '155-81-02082' : '' },
            items: [{ name: 'AI 상담 챗봇 구축', qty: 1, price: 3000000, amount: 3000000 }, { name: '월 유지관리 (11월)', qty: 1, price: 300000, amount: 300000 }],
            supply: 3300000, discount: 300000, vat: 300000, total: 3300000
          } });
        } else if (kind === 'get') res({ ok: false, error: 'demo' });
        else if (q.action === 'checkBiz') res({ ok: true, registered: false });
        else res({ ok: true });
      }, 300);
    });
  }

  /* ───────── 시작 ───────── */

  fillCompany();
  bind();
  if (!ID && !DEMO) { fail('청구서 링크가 필요해요', '받으신 청구서 링크로 다시 접속해주세요.'); return; }
  apiGet({ action: 'invoice', id: ID }).then(function (r) {
    if (!r.ok) {
      if (r.error === 'cancelled') return fail('취소된 청구서예요', '자세한 내용은 ' + (CO.phone || '') + '로 문의해주세요.');
      return fail('청구서를 찾을 수 없어요', '링크가 정확한지 확인해주세요. 문의: ' + (CO.phone || ''));
    }
    state.inv = r.invoice; state.registered = !!r.registered; if (r.client) state.client = r.client;
    render(); show('home');
  }).catch(function () {
    fail('잠시 연결이 원활하지 않아요', '잠시 후 새로고침해주세요. 문의: ' + (CO.phone || ''));
  });
})();
