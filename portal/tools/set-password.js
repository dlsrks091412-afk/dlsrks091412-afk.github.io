(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  function b64(bytes) { return btoa(String.fromCharCode.apply(null, new Uint8Array(bytes))); }
  function hex(bytes) { return Array.prototype.map.call(bytes, function (b) { return b.toString(16).padStart(2, '0'); }).join(''); }

  $('f').addEventListener('submit', async function (e) {
    e.preventDefault();
    var p1 = $('p1').value, p2 = $('p2').value, msg = $('msg');
    if (p1 !== p2) { msg.textContent = '두 비밀번호가 달라요.'; msg.className = 'msg bad'; return; }
    if (p1.length < 8) { msg.textContent = '8자 이상으로 정해 주세요. (12자 이상 권장)'; msg.className = 'msg bad'; return; }
    $('go').disabled = true;
    msg.textContent = '계산 중… (몇 초 걸려요)'; msg.className = 'msg';
    var salt = crypto.getRandomValues(new Uint8Array(16));
    var iv = crypto.getRandomValues(new Uint8Array(12));
    var iterations = 600000;
    var marker = 'portal-ok:' + hex(crypto.getRandomValues(new Uint8Array(6)));
    var base = await crypto.subtle.importKey('raw', new TextEncoder().encode(p1), 'PBKDF2', false, ['deriveKey']);
    var key = await crypto.subtle.deriveKey({ name: 'PBKDF2', salt: salt, iterations: iterations, hash: 'SHA-256' },
      base, { name: 'AES-GCM', length: 256 }, false, ['encrypt']);
    var ct = await crypto.subtle.encrypt({ name: 'AES-GCM', iv: iv }, key, new TextEncoder().encode(marker));
    $('out').value =
      '/* 자동 생성 파일 — tools/set-password.html 에서 새로 만들어 통째로 교체하세요. */\n' +
      'window.PORTAL_AUTH = {\n' +
      '  "salt": "' + b64(salt) + '",\n' +
      '  "iv": "' + b64(iv) + '",\n' +
      '  "iterations": ' + iterations + ',\n' +
      '  "marker": "' + marker + '",\n' +
      '  "check": "' + b64(ct) + '"\n' +
      '};\n';
    $('p1').value = $('p2').value = '';
    $('copy').disabled = false;
    $('go').disabled = false;
    msg.textContent = '완료! 아래 내용을 auth-config.js 에 붙여 넣으세요.'; msg.className = 'msg good';
  });

  $('copy').addEventListener('click', function () {
    $('out').select();
    (navigator.clipboard ? navigator.clipboard.writeText($('out').value) : Promise.reject())
      .then(function () { $('msg').textContent = '복사했어요.'; })
      .catch(function () { document.execCommand('copy'); $('msg').textContent = '복사했어요.'; });
  });
})();
