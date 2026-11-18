/*
 * Portal 인증 모듈 (정적 사이트용 클라이언트 측 로그인)
 * ---------------------------------------------------------------
 * - 비밀번호는 저장소에 저장되지 않습니다. 대신 PBKDF2(SHA-256)로 만든 키로
 *   암호화된 "검증용 문장"만 auth-config.js 에 들어 있습니다.
 * - 로그인 시 입력한 비밀번호로 키를 다시 만들어 그 문장을 복호화할 수 있으면 통과.
 * - 5회 이상 실패하면 점점 길어지는 잠금 시간이 적용됩니다.
 * - 세션은 기본적으로 탭을 닫으면 사라지고(sessionStorage), "로그인 유지"를
 *   선택하면 이 기기에서 7일간 유지됩니다(localStorage).
 * 한계: GitHub Pages는 정적 호스팅이라 HTML/JS 파일 자체는 누구나 내려받을 수
 *   있습니다. 진짜 비공개가 필요한 데이터는 README의 "보안" 항목을 참고하세요.
 */
(function () {
  'use strict';

  var SESSION_KEY = 'portal.session';
  var FAIL_KEY = 'portal.fail';
  var SESSION_HOURS = 2;          // 일반 로그인 유지 시간
  var PERSIST_DAYS = 7;           // "로그인 유지" 체크 시
  var FREE_TRIES = 5;             // 잠금 없이 허용되는 실패 횟수
  var MAX_LOCK_SEC = 15 * 60;     // 최대 잠금 15분

  function cfg() { return window.PORTAL_AUTH || null; }

  function b64ToBytes(s) {
    var bin = atob(s), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  function store(persist) { return persist ? window.localStorage : window.sessionStorage; }

  function safeGet(storage, key) {
    try { return JSON.parse(storage.getItem(key) || 'null'); } catch (e) { return null; }
  }
  function safeSet(storage, key, val) {
    try { storage.setItem(key, JSON.stringify(val)); } catch (e) { /* 저장 불가 환경 */ }
  }
  function safeDel(storage, key) {
    try { storage.removeItem(key); } catch (e) { /* noop */ }
  }

  /* 설정이 바뀌면(비밀번호 변경) 기존 세션이 모두 무효가 되도록 지문을 남깁니다. */
  function fingerprint() {
    var c = cfg();
    return c ? (c.salt + ':' + c.check).slice(0, 40) : '';
  }

  async function deriveKey(password, c) {
    var base = await crypto.subtle.importKey(
      'raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey(
      { name: 'PBKDF2', salt: b64ToBytes(c.salt), iterations: c.iterations, hash: 'SHA-256' },
      base, { name: 'AES-GCM', length: 256 }, false, ['decrypt']);
  }

  async function verifyPassword(password) {
    var c = cfg();
    if (!c || !c.salt) return false;
    try {
      var key = await deriveKey(password, c);
      var plain = await crypto.subtle.decrypt(
        { name: 'AES-GCM', iv: b64ToBytes(c.iv) }, key, b64ToBytes(c.check));
      return new TextDecoder().decode(plain) === c.marker;
    } catch (e) {
      return false;
    }
  }

  function failState() {
    return safeGet(localStorage, FAIL_KEY) || { count: 0, lockUntil: 0 };
  }

  function lockRemainingSec() {
    var f = failState();
    return Math.max(0, Math.ceil((f.lockUntil - Date.now()) / 1000));
  }

  function registerFail() {
    var f = failState();
    f.count += 1;
    if (f.count >= FREE_TRIES) {
      var sec = Math.min(MAX_LOCK_SEC, 30 * Math.pow(2, f.count - FREE_TRIES));
      f.lockUntil = Date.now() + sec * 1000;
    }
    safeSet(localStorage, FAIL_KEY, f);
    return f;
  }

  function randomToken() {
    var a = new Uint8Array(16);
    crypto.getRandomValues(a);
    return Array.prototype.map.call(a, function (b) { return b.toString(16).padStart(2, '0'); }).join('');
  }

  async function login(password, persist) {
    var wait = lockRemainingSec();
    if (wait > 0) return { ok: false, locked: wait };
    var ok = await verifyPassword(password);
    if (!ok) {
      var f = registerFail();
      return { ok: false, tries: f.count, locked: lockRemainingSec(), left: Math.max(0, FREE_TRIES - f.count) };
    }
    safeDel(localStorage, FAIL_KEY);
    var ms = persist ? PERSIST_DAYS * 864e5 : SESSION_HOURS * 36e5;
    var sess = { t: randomToken(), fp: fingerprint(), exp: Date.now() + ms };
    safeDel(localStorage, SESSION_KEY);
    safeDel(sessionStorage, SESSION_KEY);
    safeSet(store(persist), SESSION_KEY, sess);
    return { ok: true };
  }

  function currentSession() {
    var s = safeGet(sessionStorage, SESSION_KEY) || safeGet(localStorage, SESSION_KEY);
    if (!s || s.fp !== fingerprint() || !(s.exp > Date.now())) return null;
    return s;
  }

  function logout() {
    safeDel(localStorage, SESSION_KEY);
    safeDel(sessionStorage, SESSION_KEY);
    if (window.PortalNav) window.PortalNav.go('index.html', true);
    else location.replace('index.html');
  }

  /* 보호 페이지 <head> 에서 호출: 로그인 안 되어 있으면 로그인 화면으로 */
  function requireAuth() {
    if (currentSession()) {
      document.documentElement.classList.remove('auth-pending');
      return true;
    }
    var page = location.pathname.split('/').pop() || 'home.html';
    location.replace('index.html?next=' + encodeURIComponent(page));
    return false;
  }

  /* 로그인 후 이동할 페이지: 같은 폴더의 .html 파일만 허용 (오픈 리다이렉트 방지) */
  function nextPage() {
    var n = new URLSearchParams(location.search).get('next') || '';
    return /^[a-z0-9_-]+\.html$/i.test(n) && n !== 'index.html' ? n : 'home.html';
  }

  window.Portal = {
    login: login,
    logout: logout,
    requireAuth: requireAuth,
    isLoggedIn: function () { return !!currentSession(); },
    lockRemainingSec: lockRemainingSec,
    nextPage: nextPage,
    configured: function () { var c = cfg(); return !!(c && c.salt && c.check); }
  };
})();
