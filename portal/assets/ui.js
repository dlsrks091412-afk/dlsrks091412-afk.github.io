/*
 * 공통 UI: 낮/밤 테마 + 부드러운 페이지 전환
 * <head> 맨 앞쪽에서 불러와야 화면이 깜빡이지 않습니다.
 *
 * 테마 설정값: 'auto'(시간에 따라: 06~18시 낮) | 'day' | 'night'
 */
(function () {
  'use strict';
  var KEY = 'portal.theme';
  var root = document.documentElement;
  var reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  function readPref() {
    try { return localStorage.getItem(KEY) || 'auto'; } catch (e) { return 'auto'; }
  }
  function resolve(pref) {
    if (pref === 'day' || pref === 'night') return pref;
    var h = new Date().getHours();
    return h >= 6 && h < 18 ? 'day' : 'night';
  }

  var LABEL = { auto: '자동', day: '낮', night: '밤' };

  function syncButtons() {
    var pref = root.getAttribute('data-theme-pref');
    var theme = root.getAttribute('data-theme');
    document.querySelectorAll('[data-theme-toggle]').forEach(function (b) {
      b.setAttribute('aria-label', '테마: ' + LABEL[pref] + ' (눌러서 바꾸기)');
      b.setAttribute('title', '테마: ' + LABEL[pref]);
      var t = b.querySelector('.theme-text');
      if (t) t.textContent = LABEL[pref];
      b.dataset.mode = theme;
    });
  }

  function apply() {
    var pref = readPref();
    var theme = resolve(pref);
    var changed = root.getAttribute('data-theme') !== theme;
    root.setAttribute('data-theme', theme);
    root.setAttribute('data-theme-pref', pref);
    syncButtons();
    if (changed) window.dispatchEvent(new CustomEvent('themechange', { detail: { theme: theme, pref: pref } }));
  }

  function setPref(pref) {
    try { localStorage.setItem(KEY, pref); } catch (e) { /* 저장 불가 */ }
    if (document.startViewTransition && !reduceMotion) {
      root.classList.add('theme-swap');
      var vt = document.startViewTransition(apply);
      vt.finished.finally(function () { root.classList.remove('theme-swap'); });
    } else {
      apply();
    }
  }

  window.Theme = {
    get: function () { return root.getAttribute('data-theme'); },
    pref: readPref,
    set: setPref,
    cycle: function () {
      var order = ['auto', 'day', 'night'];
      setPref(order[(order.indexOf(readPref()) + 1) % order.length]);
    }
  };

  apply();
  setInterval(apply, 60000);   // 자동 모드에서 06시/18시가 되면 저절로 바뀜

  /* ---------- 페이지 전환 ----------
   * 최신 크롬/엣지/사파리: CSS @view-transition 으로 문서 간 전환 (CSS 참고)
   * 그 외 브라우저: 들어올 때 페이드인, 링크를 누르면 페이드아웃 후 이동
   */
  var nativeVT = 'onpagereveal' in window;
  if (!nativeVT && !reduceMotion) root.classList.add('fade-nav');

  window.Portal = window.Portal || {};
  window.PortalNav = {
    go: function (url, replace) {
      if (nativeVT || reduceMotion) { replace ? location.replace(url) : (location.href = url); return; }
      root.classList.add('leaving');
      setTimeout(function () { replace ? location.replace(url) : (location.href = url); }, 260);
    }
  };

  document.addEventListener('DOMContentLoaded', function () {
    syncButtons();
    document.querySelectorAll('[data-theme-toggle]').forEach(function (b) {
      b.addEventListener('click', function () { window.Theme.cycle(); });
    });
    if (nativeVT || reduceMotion) return;
    document.addEventListener('click', function (e) {
      var a = e.target.closest && e.target.closest('a[href]');
      if (!a || e.defaultPrevented || e.button !== 0 || e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return;
      if (a.target && a.target !== '_self') return;
      var url = new URL(a.href, location.href);
      if (url.origin !== location.origin || url.hash && url.pathname === location.pathname) return;
      e.preventDefault();
      window.PortalNav.go(url.href);
    });
  });
  // 뒤로가기로 돌아왔을 때 페이드 상태 초기화
  window.addEventListener('pageshow', function () { root.classList.remove('leaving'); });
})();
