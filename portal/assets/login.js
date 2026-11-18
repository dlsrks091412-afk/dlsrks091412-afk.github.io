(function () {
  'use strict';
  if (window.Portal.isLoggedIn()) { location.replace(window.Portal.nextPage()); return; }

  var form = document.getElementById('loginForm');
  var pw = document.getElementById('password');
  var remember = document.getElementById('remember');
  var btn = document.getElementById('submitBtn');
  var msg = document.getElementById('msg');
  var timer = null;

  if (!window.Portal.configured()) {
    document.getElementById('setupHint').hidden = false;
    btn.disabled = true;
  }

  function setMsg(text, kind) {
    msg.textContent = text;
    msg.className = 'msg' + (kind ? ' ' + kind : '');
  }

  function showLock() {
    clearInterval(timer);
    var tick = function () {
      var s = window.Portal.lockRemainingSec();
      if (s <= 0) { clearInterval(timer); btn.disabled = false; setMsg(''); return; }
      btn.disabled = true;
      var m = Math.floor(s / 60), r = s % 60;
      setMsg('실패가 많아 잠시 잠겼어요. ' + (m ? m + '분 ' : '') + r + '초 후 다시 시도하세요.', 'bad');
    };
    tick();
    timer = setInterval(tick, 1000);
  }

  if (window.Portal.lockRemainingSec() > 0) showLock();

  form.addEventListener('submit', async function (e) {
    e.preventDefault();
    if (!pw.value) { setMsg('비밀번호를 입력해 주세요.', 'bad'); pw.focus(); return; }
    btn.disabled = true;
    setMsg('확인 중…');
    var res = await window.Portal.login(pw.value, remember.checked);
    if (res.ok) {
      setMsg('어서 오세요, 대표님!', 'good');
      form.classList.add('ok');
      setTimeout(function () { location.replace(window.Portal.nextPage()); }, 450);
      return;
    }
    pw.value = '';
    pw.focus();
    form.classList.remove('shake'); void form.offsetWidth; form.classList.add('shake');
    if (res.locked > 0) { showLock(); return; }
    btn.disabled = false;
    setMsg('비밀번호가 맞지 않아요.' + (res.left ? ' (잠금까지 ' + res.left + '회 남음)' : ''), 'bad');
  });
})();
