/* 보호 페이지 공통: 로그인 확인 + 로그아웃 버튼 + 상단 시계 */
(function () {
  'use strict';
  if (!window.Portal.requireAuth()) return;
  document.addEventListener('DOMContentLoaded', function () {
    document.querySelectorAll('[data-logout]').forEach(function (b) {
      b.addEventListener('click', function () { window.Portal.logout(); });
    });
    var clock = document.querySelector('[data-clock]');
    if (clock) {
      var tick = function () {
        clock.textContent = new Date().toLocaleString('ko-KR', { month: 'long', day: 'numeric', weekday: 'short', hour: '2-digit', minute: '2-digit' });
      };
      tick(); setInterval(tick, 15000);
    }
    /* 로그인 유효시간이 끝나면 자동으로 로그인 화면으로 */
    setInterval(function () { if (!window.Portal.isLoggedIn()) window.Portal.requireAuth(); }, 60000);
  });
})();
