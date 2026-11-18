(function () {
  'use strict';
  if (window.Portal.isLoggedIn()) { location.replace(window.Portal.nextPage()); return; }

  /* ---------- 배경 일러스트 채우기 (별, 도시, 꽃, 떠다니는 꽃잎/반딧불) ---------- */
  var NS = 'http://www.w3.org/2000/svg';
  var seed = 7;
  function rnd() { seed = (seed * 16807) % 2147483647; return (seed - 1) / 2147483646; }
  function el(tag, attrs, parent) {
    var n = document.createElementNS(NS, tag);
    for (var k in attrs) n.setAttribute(k, attrs[k]);
    parent.appendChild(n);
    return n;
  }

  var stars = document.getElementById('stars');
  for (var i = 0; i < 70; i++) el('circle', { cx: (rnd() * 1600).toFixed(1), cy: (rnd() * 520).toFixed(1), r: (0.8 + rnd() * 1.8).toFixed(1) }, stars);

  var city = document.getElementById('city');
  var x = -20;
  while (x < 1640) {
    var w = 40 + Math.floor(rnd() * 60);
    var h = 70 + Math.floor(rnd() * 150);
    var isTower = x > 980 && x < 1080 && !city.querySelector('.tower');
    if (isTower) { w = 96; h = 420; }
    var top = 760 - h;
    el('rect', { x: x, y: top, width: w, height: h + 60, rx: isTower ? 10 : 4, class: isTower ? 'tower' : 'bld' + (rnd() < 0.5 ? ' b2' : '') }, city);
    if (isTower) {
      el('rect', { x: x + 12, y: top + 18, width: w - 24, height: 22, rx: 5, class: 'mywin' }, city);
      el('rect', { x: x + w / 2 - 2, y: top - 26, width: 4, height: 26, class: 'bld' }, city);
      el('circle', { cx: x + w / 2, cy: top - 28, r: 5, class: 'beacon' }, city);
    }
    for (var wy = top + (isTower ? 56 : 12); wy < 760; wy += 16) {
      for (var wx = x + 8; wx < x + w - 12; wx += 14) {
        if (rnd() < 0.5) continue;
        var on = rnd() < 0.55;
        el('rect', { x: wx, y: wy, width: 6, height: 8, rx: 1.5, class: 'win' + (on ? '' : ' off') + (on && rnd() < 0.2 ? ' tw' : '') }, city);
      }
    }
    x += w + 6 + Math.floor(rnd() * 14);
  }

  var flowers = document.getElementById('flowers');
  for (var f = 0; f < 140; f++) {
    var fx = rnd() * 1600, fy = 860 + rnd() * 140;
    el('circle', { cx: fx.toFixed(1), cy: fy.toFixed(1), r: (2 + rnd() * 4.5).toFixed(1), class: 'f' + (1 + Math.floor(rnd() * 5)) }, flowers);
  }

  var floaters = document.getElementById('floaters');
  for (var p = 0; p < 16; p++) {
    var c = el('circle', { cx: (rnd() * 1500).toFixed(0), cy: (700 + rnd() * 280).toFixed(0), r: (2.5 + rnd() * 3).toFixed(1) }, floaters);
    // CSP 때문에 style 속성 대신 DOM API로 지연시간 지정
    c.style.animationDelay = (rnd() * 9).toFixed(2) + 's';
    c.style.animationDuration = (8 + rnd() * 6).toFixed(2) + 's';
  }

  /* ---------- 로그인 폼 ---------- */
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
      setMsg('잠시 쉬어갈게요. ' + (m ? m + '분 ' : '') + r + '초 후에 다시 시도해 주세요.', 'bad');
    };
    tick();
    timer = setInterval(tick, 1000);
  }

  if (window.Portal.lockRemainingSec() > 0) showLock();

  form.addEventListener('submit', async function (e) {
    e.preventDefault();
    if (!pw.value) { setMsg('비밀번호를 입력해 주세요.', 'bad'); pw.focus(); return; }
    btn.disabled = true;
    setMsg('확인하고 있어요…');
    var res = await window.Portal.login(pw.value, remember.checked);
    if (res.ok) {
      setMsg('어서 오세요, 대표님!', 'good');
      setTimeout(function () {
        form.classList.add('ok');
        setTimeout(function () { window.PortalNav.go(window.Portal.nextPage(), true); }, 380);
      }, 350);
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
