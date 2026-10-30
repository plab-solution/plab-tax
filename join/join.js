(function () {
  'use strict';

  var CFG = window.PLAB_CONFIG || {};
  var CO = CFG.company || {};
  var params = new URLSearchParams(location.search);
  var DEMO = !CFG.API_URL || params.get('demo') === '1';
  var MAX = 10 * 1024 * 1024;
  var $ = function (s) { return document.querySelector(s); };
  var $$ = function (s) { return Array.prototype.slice.call(document.querySelectorAll(s)); };
  var state = { type: '', files: {} };

  function apiGet(qs) {
    if (DEMO) return new Promise(function (res) { setTimeout(function () { res({ ok: false, error: 'demo' }); }, 300); });
    return fetch(CFG.API_URL + '?' + new URLSearchParams(qs).toString(), { redirect: 'follow' }).then(function (r) { return r.json(); });
  }
  function apiPost(body) {
    if (DEMO) return new Promise(function (res) { setTimeout(function () { res({ ok: true }); }, 500); });
    return fetch(CFG.API_URL, { method: 'POST', redirect: 'follow', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(body) })
      .then(function (r) { return r.json(); });
  }

  function toast(msg) {
    var t = $('#toast'); t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(function () { t.hidden = true; }, 2400);
  }
  function fmtBiz(d) {
    d = String(d || '').replace(/\D/g, '').slice(0, 10);
    if (d.length > 5) return d.slice(0, 3) + '-' + d.slice(3, 5) + '-' + d.slice(5);
    if (d.length > 3) return d.slice(0, 3) + '-' + d.slice(3);
    return d;
  }
  function fmtPhone(v) {
    var d = String(v || '').replace(/\D/g, '').slice(0, 11);
    if (/^02/.test(d)) {
      if (d.length > 9) return d.slice(0, 2) + '-' + d.slice(2, 6) + '-' + d.slice(6);
      if (d.length > 5) return d.slice(0, 2) + '-' + d.slice(2, 5) + '-' + d.slice(5);
      if (d.length > 2) return d.slice(0, 2) + '-' + d.slice(2);
      return d;
    }
    if (d.length > 10) return d.slice(0, 3) + '-' + d.slice(3, 7) + '-' + d.slice(7);
    if (d.length > 6) return d.slice(0, 3) + '-' + d.slice(3, 6) + '-' + d.slice(6);
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
  function validEmail(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(v || '').trim()); }
  function val(id) { return $('#' + id).value.trim(); }

  /* ───────── 회사 정보 ───────── */
  function fillCompany() {
    $$('[data-co]').forEach(function (el) { el.textContent = CO[el.getAttribute('data-co')] || ''; });
    $$('[data-co-href="tel"]').forEach(function (a) { a.href = 'tel:' + String(CO.phone || '').replace(/\D/g, ''); });
    $$('[data-co-href="mail"]').forEach(function (a) { a.href = 'mailto:' + (CO.email || ''); });
  }

  /* ───────── 거래 유형에 따라 필수 항목 ───────── */
  function needBank() { return state.type === 'receive'; }
  function applyType() {
    $('#bank-opt').textContent = needBank() ? '' : '(선택)';
    $('#bank-hint').hidden = needBank();
    check();
  }

  function check() {
    var d = val('biz').replace(/\D/g, '');
    var ok = !!state.type && d.length === 10 && val('legal') && val('ceo').length >= 2 && val('contact')
      && val('phone').replace(/\D/g, '').length >= 9
      && validEmail(val('email'))
      && state.files.biz && (!needBank() || state.files.bank)
      && $('#agree').checked;
    $('#submit').disabled = !ok;
    return ok;
  }

  /* ───────── 파일 ───────── */
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
        res({ name: file.name.replace(/\.[^.]+$/, '') + '.jpg', mime: 'image/jpeg', data: c.toDataURL('image/jpeg', 0.85).split(',')[1] });
      };
      img.onerror = function () { URL.revokeObjectURL(url); rej(new Error('image')); };
      img.src = url;
    });
  }

  function downloadAsset(kind, btn) {
    if (DEMO) { toast('예시 화면에서는 내려받을 수 없어요.'); return; }
    var label = btn.querySelector('b').textContent;
    btn.disabled = true;
    apiGet({ action: 'joinAsset', kind: kind }).then(function (r) {
      if (!r.ok) throw new Error(r.error);
      var bin = atob(r.data), u8 = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
      var url = URL.createObjectURL(new Blob([u8], { type: r.mime }));
      var a = document.createElement('a'); a.href = url; a.download = r.name || ('프랩_' + label);
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
    }).catch(function () {
      toast('내려받지 못했어요. ' + (CO.phone || '') + '로 연락 주세요.');
    }).then(function () { btn.disabled = false; });
  }

  function copy(text) {
    var done = function () { toast('복사했어요'); };
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(done, done);
    else {
      var t = document.createElement('textarea'); t.value = text; t.style.position = 'fixed'; t.style.opacity = '0';
      document.body.appendChild(t); t.select(); try { document.execCommand('copy'); } catch (e) {} t.remove(); done();
    }
  }

  /* ───────── 이벤트 ───────── */
  function bind() {
    $$('[data-asset]').forEach(function (b) { b.addEventListener('click', function () { downloadAsset(b.getAttribute('data-asset'), b); }); });
    $('#copy-mail').addEventListener('click', function () { copy(CO.taxEmail || ''); });

    $$('input[name=type]').forEach(function (r) {
      r.addEventListener('change', function () { state.type = r.value; applyType(); });
    });
    $('#biz').addEventListener('input', function (e) { e.target.value = fmtBiz(e.target.value); $('#form-err').textContent = ''; check(); });
    $('#phone').addEventListener('input', function (e) { e.target.value = fmtPhone(e.target.value); check(); });
    ['legal', 'brand', 'ceo', 'contact', 'email'].forEach(function (id) { $('#' + id).addEventListener('input', check); });
    $('#agree').addEventListener('change', check);

    $$('input[type=file]').forEach(function (inp) {
      inp.addEventListener('change', function () {
        var file = inp.files && inp.files[0]; if (!file) return;
        var kind = inp.getAttribute('data-kind'), box = $('#up-' + kind);
        $('#form-err').textContent = '';
        box.querySelector('.fname').textContent = '불러오는 중…';
        readFile(file).then(function (f) {
          if (atob(f.data).length > MAX) throw new Error('size');
          f.kind = kind; state.files[kind] = f;
          box.querySelector('.fname').textContent = file.name;
          box.classList.add('done');
        }).catch(function (e) {
          delete state.files[kind]; box.classList.remove('done');
          box.querySelector('.fname').textContent = '눌러서 촬영하거나 파일 고르기';
          $('#form-err').textContent = e.message === 'size' ? '10MB 이하 파일만 올릴 수 있어요.' : '파일을 읽지 못했어요. 다른 파일로 시도해주세요.';
        }).then(check);
      });
    });

    $('#join-form').addEventListener('submit', function (e) {
      e.preventDefault();
      if (!check()) return;
      var d = val('biz').replace(/\D/g, '');
      if (!validBiz(d)) { $('#form-err').textContent = '사업자등록번호가 올바르지 않아요. 다시 확인해주세요.'; $('#biz').focus(); return; }
      var btn = $('#submit');
      btn.disabled = true; btn.textContent = '보내는 중…';
      var files = [state.files.biz];
      if (state.files.bank) files.push(state.files.bank);
      apiPost({
        action: 'join', type: state.type, bizNo: d, legalName: val('legal'), brand: val('brand'), ceo: val('ceo'),
        contact: val('contact'), phone: val('phone'), email: val('email'), agree: $('#agree').checked,
        website: $('#website').value, files: files
      }).then(function (r) {
        if (!r.ok) throw new Error(r.error);
        $('#done-name').textContent = (val('brand') || val('legal')) + ' · ' + fmtBiz(d);
        $('#v-form').hidden = true; $('#v-done').hidden = false;
        try { window.scrollTo(0, 0); } catch (e2) {}
      }).catch(function (err) {
        var m = {
          invalid_biz: '사업자등록번호가 올바르지 않아요.', invalid_email: '이메일 주소를 다시 확인해주세요.',
          missing_email: '세금계산서 받을 이메일을 적어주세요.', missing_files: '필요한 서류를 올려주세요.',
          missing_fields: '빠진 칸이 있어요. 다시 확인해주세요.', bad_choice: '거래 방식을 골라주세요.', bad_type: '이미지나 PDF 파일만 올릴 수 있어요.',
          bad_size: '10MB 이하 파일만 올릴 수 있어요.',
          rate_limited: '잠시 후 다시 시도해주세요. 계속 안 되면 ' + (CO.phone || '') + '로 연락 주세요.'
        }[err.message];
        $('#form-err').textContent = m || '보내지 못했어요. 잠시 후 다시 시도해주세요.';
        check();
      }).then(function () { btn.textContent = '등록하기'; });
    });
  }

  fillCompany();
  bind();
  applyType();
})();
