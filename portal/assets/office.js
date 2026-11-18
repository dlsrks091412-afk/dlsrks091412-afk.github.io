/*
 * AI 에이전트 사무실 — 픽셀 오피스 시뮬레이션
 * ---------------------------------------------------------------
 * 모든 그림은 이 파일에서 코드로 그립니다(이미지 파일 없음).
 * 논리 해상도 400x256 픽셀을 정수 배율로 키워서 또렷한 도트 느낌을 냅니다.
 *
 * 외부에서 쓰는 API (브라우저 콘솔이나 다른 스크립트에서):
 *   Office.checkIn({ id, name, dept: 'records', task })
 *   Office.progress(id, { task, percent, docs })
 *   Office.requestApproval(id, { title, detail })   → Promise<{decision, comment}>
 *   Office.checkOut(id)
 *   Office.dispatch(event)   // 위 동작을 이벤트 객체 하나로 (feed와 같은 형식)
 *   Office.onDecision(fn)    // 승인/반려 결과 구독
 */
(function () {
  'use strict';

  /* ================================================================
   * 1. 기본 상수와 레이아웃
   * ================================================================ */
  var W = 400, H = 256, WALL = 46;
  var FONT = '"Galmuri9", "Galmuri11", "DungGeunMo", monospace';
  var SMALL_FONT = '"Galmuri7", "Galmuri9", monospace';
  var SPEED = 46;               // 걷는 속도 (논리 px / 초)

  var DOOR = { x: 22, y: 47 };
  var HALL = { x: 22, y: 62 };
  var BOSS_LANE_X = 226;        // 부서 ↔ 대표실을 잇는 세로 통로
  var BOSS_LANE_Y = 134;        // 대표 책상 앞 대기 줄
  var BOSS_DESK = { x: 288, y: 92, w: 60 };
  var BOSS_SEAT = { x: 318, y: 96 };
  var BOSS_SLOTS = [318, 336, 300, 282, 264];

  /* 부서 정의 — 새 부서는 여기에 추가합니다. (README 참고) */
  var DEPTS = {
    records: {
      name: '기록부',
      desc: '문서 작업 담당',
      rug: { x: 44, y: 64, w: 162, h: 132, color: '#79a985', alt: '#72a17e', border: '#4f7a5b' },
      desks: [
        { x: 54, y: 96 }, { x: 104, y: 96 }, { x: 154, y: 96 },
        { x: 54, y: 160 }, { x: 104, y: 160 }, { x: 154, y: 160 }
      ],
      deskW: 36,
      lines: ['문서 정리 중…', '오탈자 체크!', '목차 다듬는 중', '표 서식 맞추는 중', '요약본 작성 중', '참고자료 확인 중'],
      unit: '문서'
    }
  };

  /* 아직 없는 부서 자리 (바닥 표시만) */
  var FUTURE_ZONE = { x: 232, y: 166, w: 162, h: 78 };

  /* 좌석 계산 */
  Object.keys(DEPTS).forEach(function (key) {
    var d = DEPTS[key];
    d.seats = d.desks.map(function (desk, i) {
      return {
        dept: key, index: i,
        desk: desk,
        x: desk.x + Math.round(d.deskW / 2),
        y: desk.y + 4,
        aisleY: desk.y - 18,
        agent: null,
        docs: 0
      };
    });
  });

  /* ================================================================
   * 2. 캔버스 준비
   * ================================================================ */
  var canvas = document.getElementById('office');
  var ctx = canvas.getContext('2d');
  var S = 3;
  var bg = document.createElement('canvas');
  bg.width = W; bg.height = H;
  var bgx = bg.getContext('2d');

  function resize() {
    var cssW = canvas.clientWidth || W * 2;
    var dpr = window.devicePixelRatio || 1;
    S = Math.max(2, Math.min(8, Math.ceil(cssW * dpr / W)));
    canvas.width = W * S;
    canvas.height = H * S;
  }
  window.addEventListener('resize', resize);

  function r(g, x, y, w, h, c) { g.fillStyle = c; g.fillRect(x, y, w, h); }

  /* ================================================================
   * 3. 배경 (벽, 바닥, 창문, 가구 중 캐릭터가 지나가지 않는 것들)
   * ================================================================ */
  function skyColors() {
    var h = new Date().getHours();
    if (h >= 7 && h < 17) return { top: '#8fd3f4', bottom: '#c8ecfb', night: false };
    if (h >= 17 && h < 19) return { top: '#f09a5b', bottom: '#f7d08a', night: false };
    if (h >= 5 && h < 7) return { top: '#a7a8e0', bottom: '#f5c7a9', night: false };
    return { top: '#1f2650', bottom: '#3a3f78', night: true };
  }

  function drawWindow(g, x, y, w, h, sky) {
    r(g, x - 2, y - 2, w + 4, h + 4, '#8a6a4c');
    r(g, x, y, w, h, sky.top);
    r(g, x, y + Math.floor(h / 2), w, Math.ceil(h / 2), sky.bottom);
    if (sky.night) {
      r(g, x + 5, y + 4, 1, 1, '#fff7c2'); r(g, x + w - 7, y + 7, 1, 1, '#fff7c2');
      r(g, x + Math.floor(w / 2), y + 3, 1, 1, '#fff7c2');
    } else {
      r(g, x + 4, y + 5, 8, 2, 'rgba(255,255,255,.7)'); r(g, x + 6, y + 4, 4, 1, 'rgba(255,255,255,.7)');
    }
    // 창살
    r(g, x + Math.floor(w / 2), y, 1, h, '#fdfaf2');
    r(g, x, y + Math.floor(h / 2), w, 1, '#fdfaf2');
    // 블라인드 느낌의 위쪽 그림자
    r(g, x, y, w, 2, 'rgba(0,0,0,.08)');
    // 창틀
    r(g, x - 3, y + h + 2, w + 6, 2, '#fdfaf2');
  }

  function drawPlant(g, x, y, big) {
    // x,y = 화분 바닥 중앙
    var leaf = '#4f9a4a', leaf2 = '#3a7a39', leaf3 = '#6dbb5e';
    var s = big ? 1 : 0;
    r(g, x - 4 - s, y - 7, 8 + s * 2, 7, '#c56a3d');
    r(g, x - 5 - s, y - 8, 10 + s * 2, 2, '#a85530');
    r(g, x - 1, y - 12, 2, 4, leaf2);
    r(g, x - 6 - s, y - 15 - s, 5, 4, leaf); r(g, x + 1, y - 16 - s, 6 + s, 4, leaf);
    r(g, x - 3, y - 20 - s * 3, 6, 6, leaf3); r(g, x - 8 - s, y - 19 - s, 4, 3, leaf2);
    r(g, x + 4, y - 21 - s * 2, 4, 4, leaf); r(g, x - 1, y - 24 - s * 4, 3, 4, leaf2);
    if (big) { r(g, x - 9, y - 25, 4, 4, leaf3); r(g, x + 5, y - 27, 4, 4, leaf3); }
  }

  function drawRug(g, z) {
    r(g, z.x, z.y, z.w, z.h, z.border);
    r(g, z.x + 2, z.y + 2, z.w - 4, z.h - 4, z.color);
    for (var yy = z.y + 2; yy < z.y + z.h - 2; yy += 8) {
      for (var xx = z.x + 2; xx < z.x + z.w - 2; xx += 8) {
        if (((xx - z.x) / 8 + (yy - z.y) / 8) % 2 === 0) {
          r(g, xx, yy, Math.min(8, z.x + z.w - 2 - xx), Math.min(8, z.y + z.h - 2 - yy), z.alt);
        }
      }
    }
    r(g, z.x + 2, z.y + z.h - 3, z.w - 4, 1, 'rgba(0,0,0,.12)');
  }

  function drawBackground() {
    var g = bgx;
    g.clearRect(0, 0, W, H);
    var sky = skyColors();

    // 바닥: 나무 마루
    r(g, 0, WALL, W, H - WALL, '#c08a5a');
    var row = 0;
    for (var y = WALL; y < H; y += 8, row++) {
      r(g, 0, y, W, 4, row % 2 ? '#c48e5e' : '#bd8657');
      r(g, 0, y + 7, W, 1, '#a8754a');
      for (var x = (row % 3) * 22; x < W; x += 66) r(g, x, y, 1, 7, '#a8754a');
    }

    // 뒷벽
    r(g, 0, 0, W, WALL, '#f1e6cf');
    r(g, 0, 0, W, 5, '#5b3a29');
    r(g, 0, 5, W, 1, '#7a4e33');
    r(g, 0, 33, W, 9, '#e2d2b0');
    r(g, 0, 33, W, 1, '#cbb68f');
    r(g, 0, 42, W, 4, '#7a4e33');
    r(g, 0, 46, W, 2, 'rgba(0,0,0,.18)');

    // 옆벽 (얇은 기둥)
    r(g, 0, 0, 4, H, '#5b3a29'); r(g, W - 4, 0, 4, H, '#5b3a29');
    r(g, 4, WALL, 2, H - WALL, 'rgba(0,0,0,.12)');

    // 출입문
    r(g, 10, 10, 24, 36, '#6b4430');
    r(g, 12, 12, 20, 34, '#3d2a22');
    r(g, 12, 12, 20, 34, 'rgba(255,240,200,.08)');
    r(g, 13, 13, 6, 33, '#8a5a3b');            // 열린 문짝
    r(g, 17, 28, 1, 2, '#f2c14e');            // 손잡이
    r(g, 12, 46, 20, 2, '#d8c39a');           // 문턱
    // 출입 매트
    r(g, 12, 50, 20, 8, '#5a6b8c'); r(g, 13, 51, 18, 6, '#6c7fa3');

    // 창문들
    drawWindow(g, 66, 10, 30, 18, sky);
    drawWindow(g, 108, 10, 30, 18, sky);
    drawWindow(g, 262, 10, 30, 18, sky);
    drawWindow(g, 304, 10, 30, 18, sky);

    // 화이트보드
    r(g, 150, 9, 48, 24, '#8d8f99');
    r(g, 151, 10, 46, 22, '#fbfbf7');
    r(g, 155, 14, 18, 1, '#3b6cff'); r(g, 155, 17, 26, 1, '#3b6cff'); r(g, 155, 20, 14, 1, '#d65a4f');
    r(g, 180, 22, 12, 7, '#4fa862'); r(g, 182, 24, 8, 1, '#fff'); r(g, 182, 26, 6, 1, '#fff');
    r(g, 156, 24, 2, 5, '#f2a33a'); r(g, 159, 22, 2, 7, '#f2a33a'); r(g, 162, 25, 2, 4, '#f2a33a');
    r(g, 150, 33, 48, 2, '#6d6f78');

    // 자판기
    r(g, 38, 16, 18, 42, '#3a2f2a');
    r(g, 39, 17, 16, 40, '#d64a4a');
    r(g, 41, 19, 12, 18, '#d9f1fb');
    for (var i = 0; i < 3; i++) {
      r(g, 42, 20 + i * 6, 3, 4, ['#f2a33a', '#4fa862', '#3b6cff'][i]);
      r(g, 46, 20 + i * 6, 3, 4, ['#3b6cff', '#f2a33a', '#d64a4a'][i]);
      r(g, 50, 20 + i * 6, 2, 4, ['#4fa862', '#d64a4a', '#f2a33a'][i]);
    }
    r(g, 42, 41, 8, 3, '#2b2230'); r(g, 51, 40, 2, 6, '#fff3');
    r(g, 41, 49, 12, 5, '#2b2230'); r(g, 38, 58, 18, 2, 'rgba(0,0,0,.2)');

    // 정수기
    r(g, 60, 34, 8, 4, '#bfe6f7'); r(g, 59, 30, 10, 5, '#8fd3f4'); r(g, 60, 38, 8, 18, '#e6e6ea');
    r(g, 61, 42, 2, 2, '#d64a4a'); r(g, 65, 42, 2, 2, '#3b6cff'); r(g, 59, 56, 10, 2, 'rgba(0,0,0,.2)');

    // 벽 화분
    drawPlant(g, 214, 60, false);
    drawPlant(g, 248, 60, false);

    // 대표실 책장 (벽 앞)
    r(g, 352, 18, 40, 40, '#5b3a29');
    r(g, 354, 20, 36, 36, '#8a5a3b');
    var colors = ['#d64a4a', '#3b6cff', '#4fa862', '#f2a33a', '#7b5cff', '#e9c46a'];
    for (var sh = 0; sh < 3; sh++) {
      r(g, 354, 31 + sh * 12, 36, 1, '#5b3a29');
      for (var b = 0; b < 8; b++) {
        if ((b + sh) % 5 === 4) continue;
        var hh = 7 + ((b * 7 + sh * 3) % 3);
        r(g, 356 + b * 4, 31 + sh * 12 - hh, 3, hh, colors[(b + sh * 2) % colors.length]);
      }
    }
    r(g, 352, 58, 40, 2, 'rgba(0,0,0,.2)');

    // 러그(부서 구역)
    Object.keys(DEPTS).forEach(function (k) { drawRug(g, DEPTS[k].rug); });
    drawRug(g, { x: 250, y: 56, w: 144, h: 92, color: '#a85454', alt: '#a04e4e', border: '#7a3434' });
    // 아직 비어 있는 부서 자리 (점선)
    var fz = FUTURE_ZONE;
    r(g, fz.x, fz.y, fz.w, fz.h, 'rgba(255,255,255,.07)');
    for (var dx = 0; dx < fz.w; dx += 6) { r(g, fz.x + dx, fz.y, 3, 1, '#8a6040'); r(g, fz.x + dx, fz.y + fz.h - 1, 3, 1, '#8a6040'); }
    for (var dy = 0; dy < fz.h; dy += 6) { r(g, fz.x, fz.y + dy, 1, 3, '#8a6040'); r(g, fz.x + fz.w - 1, fz.y + dy, 1, 3, '#8a6040'); }
    // 빈 책상 실루엣
    for (var e = 0; e < 3; e++) {
      r(g, fz.x + 18 + e * 46, fz.y + 30, 32, 14, 'rgba(90,60,40,.18)');
    }

    // 아래쪽 휴게 공간: 소파, 탁자, 복사기, 화분
    r(g, 58, 216, 54, 6, '#3d5a8c'); r(g, 58, 222, 54, 12, '#4a6ea8');
    r(g, 56, 218, 4, 16, '#3d5a8c'); r(g, 110, 218, 4, 16, '#3d5a8c');
    r(g, 62, 224, 22, 6, '#5b82c2'); r(g, 86, 224, 22, 6, '#5b82c2');
    r(g, 58, 234, 54, 2, 'rgba(0,0,0,.2)');
    r(g, 70, 240, 30, 6, '#8a5a3b'); r(g, 70, 246, 30, 3, '#6b4430');
    r(g, 76, 238, 4, 3, '#fff'); r(g, 88, 239, 6, 2, '#e9c46a');

    r(g, 140, 210, 26, 24, '#d9dbe2'); r(g, 140, 208, 26, 4, '#9ea2ad');
    r(g, 143, 214, 20, 3, '#3a2f2a'); r(g, 146, 220, 8, 2, '#4fa862'); r(g, 140, 234, 26, 2, 'rgba(0,0,0,.2)');
    r(g, 168, 222, 8, 12, '#3b6cff'); r(g, 168, 220, 8, 3, '#2a51cc');   // 휴지통

    drawPlant(g, 14, 240, true);
    drawPlant(g, 30, 156, false);
    drawPlant(g, 200, 246, true);
    drawPlant(g, 386, 160, false);
  }

  /* ================================================================
   * 4. 캐릭터 그리기
   * ================================================================ */
  var LOOKS = [
    { skin: '#f6cfa8', hair: '#3b2a22', style: 'short', shirt: '#4a7fd1', pants: '#2f3a52' },
    { skin: '#f3c39b', hair: '#7a3f1d', style: 'long', shirt: '#e0707a', pants: '#3a3550' },
    { skin: '#eab48c', hair: '#1f1d24', style: 'spiky', shirt: '#4fa862', pants: '#2f3a52' },
    { skin: '#f6d2b0', hair: '#c98a2b', style: 'bob', shirt: '#f2a33a', pants: '#4b3b2f' },
    { skin: '#d9a07a', hair: '#2b2230', style: 'bun', shirt: '#7b5cff', pants: '#2b2b38' },
    { skin: '#f6cfa8', hair: '#8c8c96', style: 'short', shirt: '#2fb3b3', pants: '#2f3a52' }
  ];
  var BOSS_LOOK = { skin: '#f6cfa8', hair: '#2b2230', style: 'short', shirt: '#2f3446', pants: '#23263a', tie: '#d64a4a' };

  function hashStr(s) {
    var h = 0;
    for (var i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
    return Math.abs(h);
  }

  function drawHair(g, x, top, look, dir) {
    var c = look.hair;
    if (dir === 'up') {
      r(g, x - 5, top, 10, 8, c);
      if (look.style === 'long') r(g, x - 5, top + 8, 10, 4, c);
      if (look.style === 'bun') r(g, x - 2, top - 3, 4, 3, c);
      return;
    }
    // 앞/옆 공통: 정수리
    r(g, x - 5, top, 10, 3, c);
    r(g, x - 4, top - 1, 8, 1, c);
    if (dir === 'down') {
      r(g, x - 5, top + 3, 2, 3, c); r(g, x + 3, top + 3, 2, 3, c);
      if (look.style === 'short') r(g, x - 3, top + 3, 3, 1, c);
      if (look.style === 'spiky') { r(g, x - 4, top - 2, 1, 1, c); r(g, x - 1, top - 2, 1, 1, c); r(g, x + 2, top - 2, 1, 1, c); r(g, x - 2, top + 3, 4, 1, c); }
      if (look.style === 'long') { r(g, x - 6, top + 2, 2, 9, c); r(g, x + 4, top + 2, 2, 9, c); r(g, x - 1, top + 3, 4, 1, c); }
      if (look.style === 'bob') { r(g, x - 6, top + 2, 2, 6, c); r(g, x + 4, top + 2, 2, 6, c); r(g, x - 4, top + 3, 8, 1, c); }
      if (look.style === 'bun') { r(g, x - 2, top - 4, 4, 3, c); r(g, x - 3, top + 3, 2, 1, c); }
    } else {
      var back = dir === 'left' ? 1 : -1;      // 머리 뒤쪽 방향
      r(g, back > 0 ? x : x - 5, top + 3, 5, 4, c);
      if (look.style === 'long') r(g, back > 0 ? x + 1 : x - 5, top + 3, 4, 9, c);
      if (look.style === 'bob') r(g, back > 0 ? x + 1 : x - 5, top + 3, 4, 6, c);
      if (look.style === 'bun') r(g, back > 0 ? x + 3 : x - 6, top - 1, 3, 3, c);
      if (look.style === 'spiky') r(g, back > 0 ? x + 4 : x - 6, top, 2, 2, c);
    }
  }

  /*
   * p = { look, dir, moving, frame, carrying, talking, typing, blink, alpha }
   * (x, y) = 발 아래 중앙
   */
  function drawPerson(g, p, x, y) {
    x = Math.round(x); y = Math.round(y);
    var L = p.look, dir = p.dir || 'down', f = p.frame || 0;
    var walking = !!p.moving;
    var bob = walking && (f % 2 === 1) ? -1 : 0;
    if (p.typing) bob = (f % 2) ? -1 : 0;

    // 그림자
    r(g, x - 5, y - 1, 10, 2, 'rgba(0,0,0,.22)');

    // 다리 + 신발
    var shoes = '#2b2230';
    if (dir === 'down' || dir === 'up') {
      var l = walking && f === 1 ? -1 : 0, rr = walking && f === 3 ? -1 : 0;
      r(g, x - 3, y - 6 + l, 3, 5, L.pants); r(g, x, y - 6 + rr, 3, 5, L.pants);
      r(g, x - 3, y - 2 + l, 3, 2, shoes); r(g, x, y - 2 + rr, 3, 2, shoes);
    } else {
      var s = dir === 'left' ? -1 : 1;
      if (walking && (f === 1 || f === 3)) {
        r(g, x - 3, y - 6, 2, 4, L.pants); r(g, x + 1, y - 6, 2, 4, L.pants);
        r(g, x - 4, y - 2, 3, 2, shoes); r(g, x + 1, y - 2, 3, 2, shoes);
      } else {
        r(g, x - 2, y - 6, 4, 5, L.pants);
        r(g, s > 0 ? x - 2 : x - 3, y - 2, 5, 2, shoes);
      }
    }

    // 몸통
    var bodyTop = y - 12 + bob;
    r(g, x - 4, bodyTop, 8, 7, L.shirt);
    r(g, x - 4, bodyTop + 6, 8, 1, 'rgba(0,0,0,.18)');
    if (dir === 'down') {
      r(g, x - 1, bodyTop, 2, 1, '#fff');              // 깃
      if (L.tie) r(g, x, bodyTop + 1, 1, 4, L.tie);
    }

    // 팔
    var swing = walking ? (f === 1 ? 1 : f === 3 ? -1 : 0) : 0;
    if (dir === 'down' || dir === 'up') {
      if (p.carrying && dir === 'down') {
        r(g, x - 5, bodyTop + 1, 1, 4, L.shirt); r(g, x + 4, bodyTop + 1, 1, 4, L.shirt);
      } else {
        r(g, x - 5, bodyTop + 1 + swing, 1, 5, L.shirt); r(g, x + 4, bodyTop + 1 - swing, 1, 5, L.shirt);
        r(g, x - 5, bodyTop + 5 + swing, 1, 1, L.skin); r(g, x + 4, bodyTop + 5 - swing, 1, 1, L.skin);
      }
    } else {
      var hx = dir === 'left' ? x - 1 - swing : x + swing;
      r(g, hx, bodyTop + 1, 2, 5, L.shirt);
      r(g, hx, bodyTop + 5, 2, 1, L.skin);
    }

    // 머리
    var top = bodyTop - 9;
    r(g, x - 5, top + 1, 10, 8, L.skin);
    r(g, x - 4, top + 9, 8, 0, L.skin);
    drawHair(g, x, top, L, dir);
    if (dir === 'down') {
      if (!p.blink) { r(g, x - 3, top + 5, 1, 2, '#2b2230'); r(g, x + 2, top + 5, 1, 2, '#2b2230'); }
      else { r(g, x - 3, top + 6, 1, 1, '#2b2230'); r(g, x + 2, top + 6, 1, 1, '#2b2230'); }
      r(g, x - 4, top + 7, 1, 1, '#f3a3a3'); r(g, x + 3, top + 7, 1, 1, '#f3a3a3');
      if (p.talking && f % 2) r(g, x - 1, top + 8, 2, 1, '#9c3a31');
    } else if (dir === 'left') {
      r(g, x - 4, top + 5, 1, 2, '#2b2230'); r(g, x - 4, top + 7, 1, 1, '#f3a3a3');
    } else if (dir === 'right') {
      r(g, x + 3, top + 5, 1, 2, '#2b2230'); r(g, x + 3, top + 7, 1, 1, '#f3a3a3');
    }

    // 결재 파일
    if (p.carrying) {
      if (dir === 'down') {
        r(g, x - 4, bodyTop + 1, 8, 6, '#b8892e'); r(g, x - 3, bodyTop + 1, 7, 5, '#e9c46a');
        r(g, x - 2, bodyTop, 5, 1, '#fffdf6'); r(g, x - 2, bodyTop + 3, 4, 1, '#d65a4f');
      } else if (dir === 'left') {
        r(g, x - 6, bodyTop + 1, 3, 6, '#e9c46a'); r(g, x - 6, bodyTop + 6, 3, 1, '#b8892e');
      } else if (dir === 'right') {
        r(g, x + 3, bodyTop + 1, 3, 6, '#e9c46a'); r(g, x + 3, bodyTop + 6, 3, 1, '#b8892e');
      } else {
        r(g, x - 6, bodyTop + 2, 2, 5, '#e9c46a'); r(g, x + 4, bodyTop + 2, 2, 5, '#e9c46a');
      }
    }
  }

  /* ================================================================
   * 5. 가구 (캐릭터와 앞뒤가 겹치는 것들 — 매 프레임 깊이 정렬)
   * ================================================================ */
  function drawChair(g, x, y, occupied) {
    // x = 의자 중앙, y = 책상 윗면 y
    r(g, x - 6, y - 14, 12, 11, '#2f3446');
    r(g, x - 5, y - 13, 10, 9, '#3f4760');
    r(g, x - 5, y - 13, 10, 1, '#55607f');
    if (!occupied) { r(g, x - 6, y - 4, 12, 3, '#2f3446'); }
  }

  function drawDesk(g, seat, now) {
    var d = DEPTS[seat.dept], x = seat.desk.x, y = seat.desk.y, w = d.deskW;
    // 윗면
    r(g, x, y, w, 10, '#e8dcc4');
    r(g, x, y, w, 1, '#f6eedc');
    r(g, x, y + 9, w, 1, '#cbbd9f');
    // 앞판
    r(g, x, y + 10, w, 8, '#a77b52');
    r(g, x, y + 17, w, 1, '#7a5638');
    r(g, x + 3, y + 12, w - 6, 1, '#8f6643');
    r(g, x + Math.floor(w / 2) - 2, y + 14, 4, 1, '#e9c46a');
    r(g, x + 1, y + 18, w - 2, 2, 'rgba(0,0,0,.18)');

    // 서류 더미 (작성한 문서 수에 따라 높아짐)
    var stack = Math.min(6, Math.ceil(seat.docs / 2));
    for (var i = 0; i < stack; i++) {
      r(g, x + 2, y + 6 - i * 2, 8, 2, i % 2 ? '#fffdf6' : '#ece6d6');
      r(g, x + 2, y + 7 - i * 2, 8, 1, '#cfc6b0');
    }
    // 머그컵
    r(g, x + w - 7, y + 3, 4, 4, '#fffdf6'); r(g, x + w - 3, y + 4, 1, 2, '#fffdf6');
    r(g, x + w - 7, y + 3, 4, 1, '#6b4430');

    // 노트북 (뒷면이 보임) — 사람이 앉아 있으면 화면 빛이 새어 나옴
    var cx = seat.x;
    var on = seat.agent && seat.agent.state === 'working';
    r(g, cx - 7, y - 2, 14, 9, '#4b515c');
    r(g, cx - 6, y - 1, 12, 7, '#5d6470');
    r(g, cx - 1, y + 2, 2, 2, '#9aa3b2');
    r(g, cx - 8, y + 7, 16, 2, '#3d424a');
    if (on) {
      var glow = (Math.floor(now / 300) % 2) ? '#bff3ff' : '#a2e8ff';
      r(g, cx - 7, y - 3, 14, 1, glow);
    }
  }

  function drawBossDesk(g, now) {
    var x = BOSS_DESK.x, y = BOSS_DESK.y, w = BOSS_DESK.w;
    r(g, x, y, w, 12, '#9c6440');
    r(g, x, y, w, 1, '#b97c54');
    r(g, x, y + 11, w, 1, '#7b4a2e');
    r(g, x, y + 12, w, 10, '#7b4a2e');
    r(g, x, y + 21, w, 1, '#5b3a29');
    r(g, x + 4, y + 14, 14, 6, '#6b4027'); r(g, x + w - 18, y + 14, 14, 6, '#6b4027');
    r(g, x + 10, y + 16, 2, 1, '#e9c46a'); r(g, x + w - 12, y + 16, 2, 1, '#e9c46a');
    r(g, x + 1, y + 22, w - 2, 2, 'rgba(0,0,0,.2)');
    // 명패
    r(g, x + 22, y + 13, 16, 6, '#e9c46a'); r(g, x + 23, y + 14, 14, 4, '#f6dc8f');
    // 모니터 (뒷면, 얼굴을 가리지 않게 왼쪽)
    r(g, x + 3, y - 6, 18, 12, '#3d424a'); r(g, x + 4, y - 5, 16, 10, '#4b515c');
    r(g, x + 10, y + 6, 4, 3, '#3d424a');
    r(g, x + 3, y - 7, 18, 1, (Math.floor(now / 400) % 2) ? '#bff3ff' : '#a2e8ff');
    // 키보드
    r(g, x + 23, y + 5, 14, 3, '#d9dbe2'); r(g, x + 23, y + 7, 14, 1, '#9ea2ad');
    // 결재함 (승인한 서류가 쌓임)
    r(g, x + 40, y + 2, 12, 7, '#6b4430'); r(g, x + 41, y + 3, 10, 5, '#8a5a3b');
    var approved = Math.min(4, state.approvedPile);
    for (var i = 0; i < approved; i++) r(g, x + 42, y + 6 - i, 8, 1, i % 2 ? '#e9c46a' : '#fffdf6');
  }

  function drawBossChair(g) {
    var x = BOSS_SEAT.x, y = BOSS_DESK.y;
    r(g, x - 8, y - 18, 16, 15, '#1f2028');
    r(g, x - 7, y - 17, 14, 13, '#2c2e3a');
    r(g, x - 7, y - 17, 14, 2, '#3e4152');
  }

  /* ================================================================
   * 6. 말풍선, 이름표, 이펙트
   * ================================================================ */
  function textWidth(text, font) {
    ctx.font = font;
    return Math.ceil(ctx.measureText(text).width);
  }

  function drawBubble(x, y, text, opts) {
    opts = opts || {};
    var font = '8px ' + FONT;
    var tw = textWidth(text, font);
    var w = tw + 8, h = 13;
    var bx = Math.round(Math.max(3, Math.min(W - w - 3, x - w / 2)));
    var by = Math.round(y - h - 4);
    var fill = opts.fill || '#ffffff', line = opts.line || '#2b2230';
    ctx.fillStyle = line;
    ctx.fillRect(bx + 1, by, w - 2, h);
    ctx.fillRect(bx, by + 1, w, h - 2);
    ctx.fillStyle = fill;
    ctx.fillRect(bx + 1, by + 1, w - 2, h - 2);
    // 꼬리
    var tx = Math.round(Math.max(bx + 3, Math.min(bx + w - 6, x - 1)));
    ctx.fillStyle = line; ctx.fillRect(tx, by + h, 3, 1); ctx.fillRect(tx + 1, by + h + 1, 1, 2);
    ctx.fillStyle = fill; ctx.fillRect(tx + 1, by + h - 1, 1, 1);
    ctx.fillStyle = opts.color || '#2b2230';
    ctx.font = font;
    ctx.textBaseline = 'middle';
    ctx.fillText(text, bx + 4, by + h / 2 + 0.5);
  }

  function drawName(x, y, text, color) {
    var font = '7px ' + SMALL_FONT;
    var tw = textWidth(text, font);
    var bx = Math.round(x - tw / 2 - 2), by = Math.round(y);
    ctx.fillStyle = 'rgba(30,24,36,.72)';
    ctx.fillRect(bx, by, tw + 4, 9);
    ctx.fillStyle = color || '#fff';
    ctx.font = font;
    ctx.textBaseline = 'middle';
    ctx.fillText(text, bx + 2, by + 5);
  }

  function drawProgress(x, y, pct) {
    var w = 18, bx = Math.round(x - w / 2), by = Math.round(y);
    r(ctx, bx - 1, by - 1, w + 2, 4, '#2b2230');
    r(ctx, bx, by, w, 2, '#5b4a52');
    r(ctx, bx, by, Math.round(w * Math.max(0, Math.min(1, pct / 100))), 2, '#7fe08a');
  }

  function drawDocIcon(g, x, y) {
    r(g, x, y, 6, 7, '#2b2230'); r(g, x + 1, y + 1, 4, 5, '#fffdf6');
    r(g, x + 2, y + 2, 2, 1, '#3b6cff'); r(g, x + 2, y + 4, 2, 1, '#3b6cff');
  }

  function drawWallSigns() {
    // 부서 이름 팻말 (바닥 러그 위)
    Object.keys(DEPTS).forEach(function (k) {
      var d = DEPTS[k], z = d.rug;
      var label = d.name;
      var font = '8px ' + FONT, tw = textWidth(label, font);
      var ly = z.y + z.h - 17;
      r(ctx, z.x + 4, ly, tw + 10, 13, '#3a2f2a');
      r(ctx, z.x + 5, ly + 1, tw + 8, 11, '#fbf3df');
      ctx.fillStyle = '#2b2230'; ctx.font = font; ctx.textBaseline = 'middle';
      ctx.fillText(label, z.x + 9, ly + 7);
      ctx.fillStyle = 'rgba(30,50,35,.55)'; ctx.font = '7px ' + SMALL_FONT;
      ctx.fillText(d.desc, z.x + tw + 18, ly + 7);
    });
    var bossLabel = '대표실', bf = '8px ' + FONT, bw = textWidth(bossLabel, bf);
    r(ctx, 254, 60, bw + 10, 13, '#3a2f2a'); r(ctx, 255, 61, bw + 8, 11, '#f6dc8f');
    ctx.fillStyle = '#2b2230'; ctx.font = bf; ctx.textBaseline = 'middle'; ctx.fillText(bossLabel, 259, 67);

    ctx.fillStyle = 'rgba(70,45,30,.75)'; ctx.font = '8px ' + FONT;
    ctx.fillText('신규 부서 준비 중', FUTURE_ZONE.x + 8, FUTURE_ZONE.y + 10);

    // 벽시계
    var now = new Date();
    var cx = 236, cy = 20;
    r(ctx, cx - 7, cy - 7, 14, 14, '#3a2f2a'); r(ctx, cx - 6, cy - 6, 12, 12, '#fffdf6');
    r(ctx, cx - 7, cy - 5, 1, 10, '#3a2f2a'); r(ctx, cx + 6, cy - 5, 1, 10, '#3a2f2a');
    var hA = ((now.getHours() % 12) + now.getMinutes() / 60) / 12 * Math.PI * 2;
    var mA = now.getMinutes() / 60 * Math.PI * 2;
    for (var i = 1; i <= 3; i++) r(ctx, Math.round(cx + Math.sin(hA) * i) - 0, Math.round(cy - Math.cos(hA) * i), 1, 1, '#2b2230');
    for (var j = 1; j <= 5; j++) r(ctx, Math.round(cx + Math.sin(mA) * j), Math.round(cy - Math.cos(mA) * j), 1, 1, '#d65a4f');
  }

  /* ================================================================
   * 7. 상태
   * ================================================================ */
  var state = {
    agents: [],
    particles: [],
    approvals: [],          // {id, agentId, title, detail, status: 'coming'|'arrived'|'approved'|'rejected', ...}
    approvedPile: 0,
    docsToday: 0,
    boss: { look: BOSS_LOOK, dir: 'down', frame: 0, blink: false },
    listeners: [],
    resolvers: {},
    selected: null
  };

  /* ================================================================
   * 8. 에이전트 동작
   * ================================================================ */
  function say(agent, text, ms) {
    agent.bubble = { text: text, until: performance.now() + (ms || 2600) };
  }

  function setPath(agent, points, onArrive) {
    agent.path = points.map(function (p) { return { x: p.x, y: p.y }; });
    agent.onArrive = onArrive || null;
    agent.moving = true;
  }

  function freeSeat(dept) {
    var d = DEPTS[dept];
    if (!d) return null;
    for (var i = 0; i < d.seats.length; i++) if (!d.seats[i].agent) return d.seats[i];
    return null;
  }

  function freeSlot() {
    for (var i = 0; i < BOSS_SLOTS.length; i++) {
      var taken = state.agents.some(function (a) { return a.slot === i; });
      if (!taken) return i;
    }
    return BOSS_SLOTS.length - 1;
  }

  function pathDoorToSeat(seat) {
    return [HALL, { x: HALL.x, y: seat.aisleY }, { x: seat.x, y: seat.aisleY }, { x: seat.x, y: seat.y }];
  }
  function pathSeatToDoor(seat) {
    return [{ x: seat.x, y: seat.aisleY }, { x: HALL.x, y: seat.aisleY }, HALL, DOOR];
  }
  function pathSeatToBoss(seat, slot) {
    var sx = BOSS_SLOTS[slot];
    return [{ x: seat.x, y: seat.aisleY }, { x: BOSS_LANE_X, y: seat.aisleY }, { x: BOSS_LANE_X, y: BOSS_LANE_Y }, { x: sx, y: BOSS_LANE_Y }];
  }
  function pathBossToSeat(agent, seat) {
    return [{ x: agent.x, y: BOSS_LANE_Y }, { x: BOSS_LANE_X, y: BOSS_LANE_Y }, { x: BOSS_LANE_X, y: seat.aisleY }, { x: seat.x, y: seat.aisleY }, { x: seat.x, y: seat.y }];
  }

  function findAgent(id) {
    for (var i = 0; i < state.agents.length; i++) if (state.agents[i].id === id) return state.agents[i];
    return null;
  }

  function checkIn(opts) {
    opts = opts || {};
    var id = String(opts.id || opts.name || ('agent-' + Date.now()));
    var dept = DEPTS[opts.dept] ? opts.dept : 'records';
    var existing = findAgent(id);
    if (existing) {
      if (opts.task) existing.task = opts.task;
      return existing;
    }
    var seat = freeSeat(dept);
    if (!seat) {
      log('⚠', (opts.name || id) + ' — ' + DEPTS[dept].name + '에 빈 자리가 없어요.');
      return null;
    }
    var agent = {
      id: id,
      name: opts.name || id,
      dept: dept,
      look: opts.look || LOOKS[hashStr(id) % LOOKS.length],
      x: DOOR.x, y: DOOR.y, dir: 'down',
      frame: 0, frameT: 0,
      path: [], moving: false,
      state: 'entering',
      seat: seat,
      slot: null,
      task: opts.task || '문서 작업',
      percent: 0,
      docs: 0,
      queue: [],
      bubble: null,
      nextDoc: performance.now() + 4000 + Math.random() * 3000,
      nextChat: performance.now() + 9000 + Math.random() * 6000,
      fireUntil: 0,
      alpha: 0,
      carrying: false,
      approval: null,
      since: new Date()
    };
    seat.agent = agent;
    state.agents.push(agent);
    say(agent, '출근했습니다~', 3000);
    log('🟢', agent.name + ' 출근 · ' + DEPTS[dept].name + ' (' + agent.task + ')');
    setPath(agent, pathDoorToSeat(seat), function () {
      agent.state = 'working';
      agent.dir = 'down';
      say(agent, '업무 시작! ' + shorten(agent.task, 12), 2800);
    });
    renderPanels();
    return agent;
  }

  function shorten(s, n) { s = String(s || ''); return s.length > n ? s.slice(0, n - 1) + '…' : s; }

  function progress(id, opts) {
    var a = findAgent(id);
    if (!a) return;
    opts = opts || {};
    if (opts.task && opts.task !== a.task) {
      a.task = opts.task;
      log('📝', a.name + ' — 새 업무: ' + a.task);
      if (a.state === 'working') say(a, '다음 업무: ' + shorten(a.task, 12));
    }
    if (typeof opts.percent === 'number') {
      var jump = opts.percent - a.percent;
      a.percent = Math.max(0, Math.min(100, opts.percent));
      if (jump >= 25 && a.state === 'working') ignite(a);
    }
    if (typeof opts.docs === 'number' && opts.docs > a.docs) {
      for (var i = a.docs; i < opts.docs; i++) produceDoc(a);
    }
    renderPanels();
  }

  function ignite(a) {
    a.fireUntil = performance.now() + 3500;
    say(a, '몰입 모드!', 1800);
  }

  function produceDoc(a) {
    a.docs += 1;
    a.seat.docs += 1;
    state.docsToday += 1;
    state.particles.push({ kind: 'doc', x: a.x - 6, y: a.y - 30, vy: -10, life: 1.6, max: 1.6, text: '+1' });
  }

  function requestApproval(id, opts) {
    var a = findAgent(id);
    opts = opts || {};
    if (!a) return Promise.reject(new Error('출근하지 않은 에이전트입니다: ' + id));
    var ap = {
      id: String(opts.approvalId || opts.id || ('ap-' + Date.now() + '-' + Math.floor(Math.random() * 1000))),
      agentId: a.id,
      agentName: a.name,
      dept: a.dept,
      title: opts.title || '결재 요청',
      detail: opts.detail || '',
      status: 'coming',
      createdAt: new Date()
    };
    state.approvals.push(ap);
    a.queue.push({ type: 'approval', approval: ap });
    log('📂', a.name + ' — 결재 요청: ' + ap.title);
    renderPanels();
    return new Promise(function (resolve) { state.resolvers[ap.id] = resolve; });
  }

  function checkOut(id) {
    var a = findAgent(id);
    if (!a) return;
    if (!a.queue.some(function (q) { return q.type === 'checkout'; })) a.queue.push({ type: 'checkout' });
  }

  /* 근무 중일 때만 대기열의 다음 일을 처리 */
  function processQueue(a) {
    if (a.state !== 'working' || a.moving || !a.queue.length) return;
    var job = a.queue.shift();
    if (job.type === 'approval') {
      a.state = 'toBoss';
      a.carrying = true;
      a.approval = job.approval;
      a.slot = freeSlot();
      say(a, '결재 받으러 갑니다!', 2200);
      setPath(a, pathSeatToBoss(a.seat, a.slot), function () {
        a.state = 'waiting';
        a.dir = 'up';
        job.approval.status = 'arrived';
        say(a, '대표님, 결재 부탁드립니다!', 4000);
        log('🔔', a.name + ' 도착 — "' + job.approval.title + '" 결재 대기');
        renderPanels();
        notify(job.approval);
        maybeOpenDialog(job.approval);
      });
      renderPanels();
    } else if (job.type === 'checkout') {
      a.state = 'leaving';
      say(a, '퇴근하겠습니다~ 수고하셨습니다!', 3200);
      log('🌙', a.name + ' 퇴근 (오늘 ' + DEPTS[a.dept].unit + ' ' + a.docs + '건)');
      setPath(a, pathSeatToDoor(a.seat), function () {
        a.state = 'gone';
      });
      renderPanels();
    }
  }

  function decide(approvalId, decision, comment) {
    var ap = null;
    state.approvals.forEach(function (x) { if (x.id === approvalId) ap = x; });
    if (!ap || (ap.status !== 'arrived' && ap.status !== 'coming')) return false;
    if (ap.status === 'coming') return false;   // 아직 도착 전
    ap.status = decision === 'approved' ? 'approved' : 'rejected';
    ap.comment = comment || '';
    ap.decidedAt = new Date();
    var a = findAgent(ap.agentId);
    if (a) {
      a.approval = null;
      if (ap.status === 'approved') {
        a.carrying = false;
        state.approvedPile += 1;
        say(a, '감사합니다! 바로 진행할게요', 2600);
        state.particles.push({ kind: 'stamp', x: a.x, y: a.y - 34, vy: -6, life: 1.4, max: 1.4, text: '승인' });
      } else {
        say(a, '넵, 수정해서 다시 올릴게요…', 2800);
        state.particles.push({ kind: 'stamp', x: a.x, y: a.y - 34, vy: -6, life: 1.4, max: 1.4, text: '반려', bad: true });
      }
      a.state = 'returning';
      setTimeout(function () {
        a.dir = 'left';
        setPath(a, pathBossToSeat(a, a.seat), function () {
          a.state = 'working';
          a.carrying = false;
          a.slot = null;
          a.dir = 'down';
        });
        a.slot = null;
      }, 900);
    }
    log(ap.status === 'approved' ? '✅' : '↩️', '"' + ap.title + '" ' + (ap.status === 'approved' ? '승인' : '반려') + (ap.comment ? ' — ' + ap.comment : ''));
    var result = { approvalId: ap.id, agentId: ap.agentId, decision: ap.status, comment: ap.comment, decidedAt: ap.decidedAt.toISOString() };
    if (state.resolvers[ap.id]) { state.resolvers[ap.id](result); delete state.resolvers[ap.id]; }
    state.listeners.forEach(function (fn) { try { fn(result, ap); } catch (e) { console.error(e); } });
    sendDecision(result);
    renderPanels();
    return true;
  }

  /* ================================================================
   * 9. 매 프레임 업데이트
   * ================================================================ */
  function update(dt, now) {
    state.agents.forEach(function (a) {
      // 입장/퇴장 페이드
      if (a.state === 'gone') a.alpha -= dt * 3; else a.alpha = Math.min(1, a.alpha + dt * 3);

      if (a.moving && a.path.length) {
        var t = a.path[0];
        var dx = t.x - a.x, dy = t.y - a.y;
        var dist = Math.sqrt(dx * dx + dy * dy);
        var step = SPEED * dt;
        if (Math.abs(dx) > Math.abs(dy)) a.dir = dx < 0 ? 'left' : 'right';
        else if (dist > 0.01) a.dir = dy < 0 ? 'up' : 'down';
        if (dist <= step) {
          a.x = t.x; a.y = t.y; a.path.shift();
          if (!a.path.length) {
            a.moving = false;
            var cb = a.onArrive; a.onArrive = null;
            if (cb) cb();
            renderPanels();
          }
        } else {
          a.x += dx / dist * step; a.y += dy / dist * step;
        }
        a.frameT += dt;
        if (a.frameT > 0.13) { a.frameT = 0; a.frame = (a.frame + 1) % 4; }
      } else if (a.state === 'working') {
        a.frameT += dt;
        if (a.frameT > 0.22) { a.frameT = 0; a.frame = (a.frame + 1) % 4; }
        if (now > a.nextDoc) {
          a.nextDoc = now + 6000 + Math.random() * 6000;
          produceDoc(a);
          if (a.percent < 95) a.percent = Math.min(95, a.percent + 4 + Math.round(Math.random() * 6));
          renderPanels();
        }
        if (now > a.nextChat) {
          a.nextChat = now + 12000 + Math.random() * 12000;
          var lines = DEPTS[a.dept].lines;
          if (Math.random() < 0.25) ignite(a); else say(a, lines[Math.floor(Math.random() * lines.length)], 2400);
        }
        processQueue(a);
      } else {
        a.frame = 0;
      }

      a.blink = (now % 3200) < 140 + (hashStr(a.id) % 5) * 10 && ((now + hashStr(a.id) * 37) % 3200) < 140;

      if (a.fireUntil > now && Math.random() < dt * 22) {
        state.particles.push({
          kind: 'fire', x: a.x - 6 + Math.random() * 12, y: a.y - 6 - Math.random() * 10,
          vy: -18 - Math.random() * 12, life: 0.7, max: 0.7, size: 2 + Math.floor(Math.random() * 2)
        });
      }
    });

    // 퇴근 완료된 에이전트 정리
    state.agents = state.agents.filter(function (a) {
      if (a.state === 'gone' && a.alpha <= 0) {
        if (a.seat.agent === a) a.seat.agent = null;
        if (state.selected === a.id) state.selected = null;
        renderPanels();
        return false;
      }
      return true;
    });

    state.particles.forEach(function (p) { p.life -= dt; p.y += p.vy * dt; });
    state.particles = state.particles.filter(function (p) { return p.life > 0; });

    state.boss.blink = (now % 4100) < 130;
    var waiting = state.approvals.some(function (x) { return x.status === 'arrived'; });
    state.boss.typing = !waiting;
    state.boss.frameT = (state.boss.frameT || 0) + dt;
    if (state.boss.frameT > 0.3) { state.boss.frameT = 0; state.boss.frame = (state.boss.frame + 1) % 4; }
  }

  /* ================================================================
   * 10. 그리기
   * ================================================================ */
  function render(now) {
    ctx.setTransform(S, 0, 0, S, 0, 0);
    ctx.imageSmoothingEnabled = false;
    ctx.drawImage(bg, 0, 0, W, H);
    drawWallSigns();

    // 깊이 정렬할 그리기 목록
    var list = [];
    Object.keys(DEPTS).forEach(function (k) {
      DEPTS[k].seats.forEach(function (seat) {
        var seated = seat.agent && !seat.agent.moving && (seat.agent.state === 'working' || seat.agent.state === 'entering');
        list.push({ y: seat.desk.y + 2, draw: function () { drawChair(ctx, seat.x, seat.desk.y, seated); } });
        list.push({ y: seat.desk.y + 10, draw: function () { drawDesk(ctx, seat, now); } });
      });
    });
    list.push({ y: BOSS_DESK.y - 4, draw: function () { drawBossChair(ctx); } });
    list.push({ y: BOSS_SEAT.y, draw: function () { drawPerson(ctx, state.boss, BOSS_SEAT.x, BOSS_SEAT.y); } });
    list.push({ y: BOSS_DESK.y + 12, draw: function () { drawBossDesk(ctx, now); } });

    state.agents.forEach(function (a) {
      list.push({
        y: a.y + 0.5, draw: function () {
          ctx.globalAlpha = Math.max(0, Math.min(1, a.alpha));
          drawPerson(ctx, {
            look: a.look, dir: a.dir, moving: a.moving, frame: a.frame,
            carrying: a.carrying, typing: a.state === 'working' && !a.moving,
            talking: !!(a.bubble && a.bubble.until > now), blink: a.blink
          }, a.x, a.y);
          if (state.selected === a.id) {
            r(ctx, Math.round(a.x) - 1, Math.round(a.y) - 31, 3, 2, '#f2c14e');
            r(ctx, Math.round(a.x), Math.round(a.y) - 29, 1, 1, '#f2c14e');
          }
          ctx.globalAlpha = 1;
        }
      });
    });
    list.sort(function (a, b) { return a.y - b.y; });
    list.forEach(function (item) { item.draw(); });

    // 이펙트
    state.particles.forEach(function (p) {
      var k = p.life / p.max;
      ctx.globalAlpha = Math.min(1, k * 1.6);
      if (p.kind === 'fire') {
        var c = k > 0.66 ? '#ffd23f' : k > 0.33 ? '#ff8c1a' : '#ff4b1f';
        r(ctx, Math.round(p.x), Math.round(p.y), p.size, p.size + 1, c);
      } else if (p.kind === 'doc') {
        drawDocIcon(ctx, Math.round(p.x), Math.round(p.y));
        ctx.font = '8px ' + FONT; ctx.textBaseline = 'middle';
        ctx.fillStyle = '#2b2230'; ctx.fillText(p.text, Math.round(p.x) + 8, Math.round(p.y) + 5);
        ctx.fillStyle = '#ffffff'; ctx.fillText(p.text, Math.round(p.x) + 7, Math.round(p.y) + 4);
      } else if (p.kind === 'stamp') {
        var col = p.bad ? '#d65a4f' : '#e0453a';
        ctx.font = '8px ' + FONT; ctx.textBaseline = 'middle';
        var tw = textWidth(p.text, ctx.font);
        r(ctx, Math.round(p.x - tw / 2 - 3), Math.round(p.y - 6), tw + 6, 12, col);
        r(ctx, Math.round(p.x - tw / 2 - 2), Math.round(p.y - 5), tw + 4, 10, '#fff6f2');
        ctx.fillStyle = col; ctx.fillText(p.text, Math.round(p.x - tw / 2), Math.round(p.y));
      }
      ctx.globalAlpha = 1;
    });

    // 이름표, 진행률, 말풍선 (항상 맨 위)
    state.agents.forEach(function (a) {
      if (a.alpha <= 0.05) return;
      ctx.globalAlpha = Math.max(0, Math.min(1, a.alpha));
      var headTop = a.y - 23;
      drawName(a.x, headTop - 10, a.name, a.state === 'waiting' ? '#f6dc8f' : '#ffffff');
      if (a.state === 'working' && !a.moving) drawProgress(a.x, headTop - 15, a.percent);
      if (a.state === 'waiting' && Math.floor(now / 400) % 2) {
        r(ctx, Math.round(a.x) + 7, Math.round(headTop) - 2, 3, 6, '#d65a4f'); r(ctx, Math.round(a.x) + 7, Math.round(headTop) + 5, 3, 2, '#d65a4f');
      }
      if (a.bubble && a.bubble.until > now) {
        drawBubble(a.x, headTop - (a.state === 'working' && !a.moving ? 17 : 11), a.bubble.text, a.state === 'waiting' ? { fill: '#fff6d6' } : null);
      }
      ctx.globalAlpha = 1;
    });
    var waitingCount = state.approvals.filter(function (x) { return x.status === 'arrived'; }).length;
    drawName(BOSS_SEAT.x, BOSS_SEAT.y - 33, '나 (대표)', '#f6dc8f');
    if (waitingCount) drawBubble(BOSS_SEAT.x, BOSS_SEAT.y - 34, '결재 ' + waitingCount + '건 확인할게요', { fill: '#fff6d6' });
  }

  var last = performance.now();
  var lastBg = 0;
  function loop(now) {
    var dt = Math.min(0.05, (now - last) / 1000);
    last = now;
    if (now - lastBg > 60000) { drawBackground(); lastBg = now; }
    update(dt, now);
    render(now);
    requestAnimationFrame(loop);
  }

  /* ================================================================
   * 11. 화면 패널 (결재함, 근무 현황, 업무 일지, 결재 창)
   * ================================================================ */
  var $ = function (sel) { return document.querySelector(sel); };
  var el = {
    approvals: $('#approvals'), approvalCount: $('#approvalCount'),
    roster: $('#roster'), log: $('#log'),
    statDocs: $('#statDocs'), statWorking: $('#statWorking'), statWaiting: $('#statWaiting'),
    dialog: $('#approvalDialog'), dTitle: $('#dTitle'), dFrom: $('#dFrom'), dDetail: $('#dDetail'),
    dComment: $('#dComment'), dApprove: $('#dApprove'), dReject: $('#dReject'), dStamp: $('#dStamp'),
    dClose: $('#dClose'), toast: $('#toast')
  };

  var STATE_LABEL = {
    entering: ['출근 중', 'move'], working: ['업무 중', 'work'], toBoss: ['결재 가는 중', 'move'],
    waiting: ['결재 대기', 'wait'], returning: ['자리로 복귀', 'move'], leaving: ['퇴근 중', 'off'], gone: ['퇴근', 'off']
  };

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  var panelQueued = false;
  function renderPanels() {
    if (panelQueued) return;
    panelQueued = true;
    requestAnimationFrame(function () { panelQueued = false; drawPanels(); });
  }

  function drawPanels() {
    var pending = state.approvals.filter(function (x) { return x.status === 'coming' || x.status === 'arrived'; });
    el.approvalCount.textContent = pending.length;
    el.approvalCount.hidden = !pending.length;
    if (!pending.length) {
      el.approvals.innerHTML = '<li class="empty">대기 중인 결재가 없어요.</li>';
    } else {
      el.approvals.innerHTML = pending.map(function (ap) {
        var arrived = ap.status === 'arrived';
        return '<li class="ap-item' + (arrived ? ' arrived' : '') + '">' +
          '<div class="ap-head"><span class="file-ico" aria-hidden="true"></span><strong>' + esc(ap.title) + '</strong></div>' +
          '<div class="ap-meta">' + esc(ap.agentName) + ' · ' + esc(DEPTS[ap.dept].name) + ' · ' +
          (arrived ? '<span class="ok">자리 앞에서 대기 중</span>' : '<span class="muted">서류 들고 오는 중…</span>') + '</div>' +
          '<div class="ap-actions">' +
          '<button class="btn small" data-open="' + esc(ap.id) + '"' + (arrived ? '' : ' disabled') + '>서류 열기</button>' +
          '</div></li>';
      }).join('');
    }

    if (!state.agents.length) {
      el.roster.innerHTML = '<li class="empty">아직 출근한 에이전트가 없어요.</li>';
    } else {
      el.roster.innerHTML = state.agents.map(function (a) {
        var st = STATE_LABEL[a.state] || ['?', ''];
        return '<li class="ro-item' + (state.selected === a.id ? ' selected' : '') + '" data-agent="' + esc(a.id) + '">' +
          '<span class="avatar"></span>' +
          '<span class="ro-body"><strong>' + esc(a.name) + '</strong><span>' + esc(DEPTS[a.dept].name) + ' · ' + esc(shorten(a.task, 18)) + '</span></span>' +
          '<span class="chip ' + st[1] + '">' + st[0] + '</span>' +
          '<span class="ro-stat">' + a.docs + '건 · ' + a.percent + '%</span></li>';
      }).join('');
      // CSP(인라인 style 금지) 때문에 아바타 색은 DOM API로 지정
      Array.prototype.forEach.call(el.roster.querySelectorAll('.avatar'), function (av, i) {
        var lk = state.agents[i].look;
        av.style.setProperty('--shirt', lk.shirt);
        av.style.setProperty('--hair', lk.hair);
        av.style.setProperty('--skin', lk.skin);
      });
    }

    var working = state.agents.filter(function (a) { return a.state !== 'gone' && a.state !== 'leaving'; }).length;
    el.statDocs.textContent = state.docsToday;
    el.statWorking.textContent = working;
    el.statWaiting.textContent = pending.length;
  }

  var logItems = [];
  function log(icon, text) {
    var t = new Date().toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' });
    logItems.unshift({ t: t, icon: icon, text: text });
    logItems = logItems.slice(0, 60);
    if (el.log) {
      el.log.innerHTML = logItems.map(function (l) {
        return '<li><time>' + l.t + '</time><span class="ic">' + l.icon + '</span><span>' + esc(l.text) + '</span></li>';
      }).join('');
    }
  }

  var openApproval = null;
  function openDialog(ap) {
    if (!ap || ap.status !== 'arrived') return;
    openApproval = ap;
    el.dTitle.textContent = ap.title;
    el.dFrom.textContent = ap.agentName + ' · ' + DEPTS[ap.dept].name + ' · ' + ap.createdAt.toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' });
    el.dDetail.textContent = ap.detail || '(상세 내용 없음)';
    el.dComment.value = '';
    el.dStamp.className = 'stamp';
    el.dStamp.textContent = '';
    el.dApprove.disabled = el.dReject.disabled = false;
    if (!el.dialog.open) el.dialog.showModal();
  }

  function maybeOpenDialog(ap) {
    if (!el.dialog.open && !document.hidden) openDialog(ap);
  }

  function finishDialog(decision) {
    if (!openApproval) return;
    var ap = openApproval;
    el.dApprove.disabled = el.dReject.disabled = true;
    el.dStamp.textContent = decision === 'approved' ? '승인' : '반려';
    el.dStamp.className = 'stamp show ' + (decision === 'approved' ? 'ok' : 'no');
    decide(ap.id, decision, el.dComment.value.trim());
    setTimeout(function () {
      el.dialog.close();
      openApproval = null;
      var next = state.approvals.filter(function (x) { return x.status === 'arrived'; })[0];
      if (next) setTimeout(function () { openDialog(next); }, 400);
    }, 750);
  }

  el.dApprove.addEventListener('click', function () { finishDialog('approved'); });
  el.dReject.addEventListener('click', function () { finishDialog('rejected'); });
  el.dClose.addEventListener('click', function () { el.dialog.close(); openApproval = null; });
  el.dialog.addEventListener('cancel', function () { openApproval = null; });

  el.approvals.addEventListener('click', function (e) {
    var b = e.target.closest('[data-open]');
    if (!b) return;
    var id = b.getAttribute('data-open');
    openDialog(state.approvals.filter(function (x) { return x.id === id; })[0]);
  });

  el.roster.addEventListener('click', function (e) {
    var li = e.target.closest('[data-agent]');
    if (!li) return;
    var id = li.getAttribute('data-agent');
    state.selected = state.selected === id ? null : id;
    renderPanels();
  });

  /* 캔버스 클릭: 에이전트 선택 / 결재 대기 중이면 서류 열기 */
  canvas.addEventListener('click', function (e) {
    var rect = canvas.getBoundingClientRect();
    var mx = (e.clientX - rect.left) / rect.width * W;
    var my = (e.clientY - rect.top) / rect.height * H;
    var hit = null;
    state.agents.forEach(function (a) {
      if (mx > a.x - 7 && mx < a.x + 7 && my > a.y - 26 && my < a.y + 1) hit = a;
    });
    if (hit) {
      if (hit.state === 'waiting' && hit.approval) { openDialog(hit.approval); return; }
      state.selected = hit.id;
      say(hit, hit.state === 'working' ? shorten(hit.task, 14) + ' (' + hit.percent + '%)' : (STATE_LABEL[hit.state] || [''])[0], 2000);
    } else if (Math.abs(mx - BOSS_SEAT.x) < 10 && my > BOSS_SEAT.y - 26 && my < BOSS_SEAT.y) {
      var next = state.approvals.filter(function (x) { return x.status === 'arrived'; })[0];
      if (next) openDialog(next);
    } else {
      state.selected = null;
    }
    renderPanels();
  });
  canvas.addEventListener('mousemove', function (e) {
    var rect = canvas.getBoundingClientRect();
    var mx = (e.clientX - rect.left) / rect.width * W;
    var my = (e.clientY - rect.top) / rect.height * H;
    var over = state.agents.some(function (a) { return mx > a.x - 7 && mx < a.x + 7 && my > a.y - 26 && my < a.y + 1; });
    canvas.style.cursor = over ? 'pointer' : 'default';
  });

  function toast(text) {
    if (!el.toast) return;
    el.toast.textContent = text;
    el.toast.classList.add('show');
    clearTimeout(toast.t);
    toast.t = setTimeout(function () { el.toast.classList.remove('show'); }, 3200);
  }

  function notify(ap) {
    toast('📂 ' + ap.agentName + ': "' + ap.title + '" 결재 요청');
    if (document.hidden && 'Notification' in window && Notification.permission === 'granted') {
      try { new Notification('결재 요청', { body: ap.agentName + ' — ' + ap.title }); } catch (e) { /* noop */ }
    }
  }

  /* ================================================================
   * 12. 외부 연결 (이벤트 피드 / 결정 전송)
   * ================================================================ */
  var CONFIG = window.OFFICE_CONFIG || {};
  var seenEvents = {};
  var feedStart = 0;
  var pendingTimed = [];

  function dispatch(ev) {
    if (!ev || typeof ev !== 'object') return;
    var type = ev.type;
    var id = ev.agentId || ev.agent || ev.id;
    if (type === 'checkin') return checkIn({ id: ev.agentId || ev.agent, name: ev.name, dept: ev.dept, task: ev.task });
    if (type === 'progress') return progress(id, { task: ev.task, percent: ev.percent, docs: ev.docs });
    if (type === 'approval') return requestApproval(id, { approvalId: ev.approvalId, title: ev.title, detail: ev.detail });
    if (type === 'checkout') return checkOut(id);
    if (type === 'say') { var a = findAgent(id); if (a) say(a, String(ev.text || '').slice(0, 30)); return; }
    console.warn('[Office] 알 수 없는 이벤트', ev);
  }

  async function pollFeed() {
    if (!CONFIG.feedUrl) return;
    try {
      var res = await fetch(CONFIG.feedUrl + (CONFIG.feedUrl.indexOf('?') < 0 ? '?' : '&') + '_=' + Date.now(), { cache: 'no-store' });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      var data = await res.json();
      var events = Array.isArray(data) ? data : (data.events || []);
      if (!feedStart) feedStart = performance.now();
      events.forEach(function (ev) {
        var key = ev.eventId || ev.id && ev.type + ':' + ev.id || JSON.stringify(ev);
        if (seenEvents[key]) return;
        seenEvents[key] = true;
        if (ev.delayMs) pendingTimed.push({ at: feedStart + ev.delayMs, ev: ev });
        else dispatch(ev);
      });
    } catch (e) {
      console.warn('[Office] 피드를 읽지 못했어요:', e.message);
    }
  }

  setInterval(function () {
    var now = performance.now();
    pendingTimed = pendingTimed.filter(function (p) {
      if (p.at <= now) { dispatch(p.ev); return false; }
      return true;
    });
  }, 250);

  function sendDecision(result) {
    if (!CONFIG.decisionUrl) return;
    fetch(CONFIG.decisionUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(result)
    }).catch(function (e) { console.warn('[Office] 결정 전송 실패:', e.message); });
  }

  /* ================================================================
   * 13. 데모
   * ================================================================ */
  var DEMO_NAMES = ['기록이', '서기봇', '문서왕', '메모리', '타자봇', '정리왕'];
  var DEMO_TASKS = ['주간 회의록 정리', '프로젝트 기획서 초안', '블로그 글 교정', '보고서 요약본 작성', '매뉴얼 목차 구성', '독서 노트 정리'];
  var DEMO_APPROVALS = [
    { title: '주간 보고서 v1 배포', detail: '이번 주 회의 3건을 요약한 보고서입니다.\n- 결정 사항 5건, 할 일 8건 정리\n- 노션 "주간 보고" 페이지에 게시 예정\n\n게시해도 될까요?' },
    { title: '기획서 초안 공유', detail: '프로젝트 기획서 초안(4쪽)을 팀 드라이브에 공유하려고 합니다.\n외부 공유는 하지 않습니다.' },
    { title: '블로그 글 발행', detail: '"AI 에이전트와 일하는 법" 글의 교정을 마쳤어요.\n맞춤법 23곳 수정, 문단 2개 재구성.\n지금 발행할까요?' },
    { title: '문서 템플릿 변경', detail: '회의록 템플릿에 "다음 액션" 칸을 추가하려고 합니다.\n기존 문서에는 영향이 없습니다.' }
  ];
  var demoN = 0;

  function demoCheckIn() {
    var free = freeSeat('records');
    if (!free) { toast('기록부 자리가 모두 찼어요.'); return null; }
    var i = demoN++ % DEMO_NAMES.length;
    var id = 'demo-' + i + '-' + Date.now().toString(36);
    return checkIn({ id: id, name: DEMO_NAMES[i], dept: 'records', task: DEMO_TASKS[i] });
  }

  function pickWorking() {
    var list = state.agents.filter(function (a) { return a.state === 'working' || a.state === 'entering'; });
    return list[Math.floor(Math.random() * list.length)] || null;
  }

  function demoApproval() {
    var a = pickWorking() || demoCheckIn();
    if (!a) return;
    var spec = DEMO_APPROVALS[Math.floor(Math.random() * DEMO_APPROVALS.length)];
    requestApproval(a.id, spec);
  }

  function demoCheckOut() {
    var list = state.agents.filter(function (a) { return a.state !== 'leaving' && a.state !== 'gone'; });
    var a = list[0];
    if (!a) { toast('퇴근할 에이전트가 없어요.'); return; }
    checkOut(a.id);
  }

  function demoScenario() {
    var a = demoCheckIn();
    if (!a) return;
    setTimeout(function () { progress(a.id, { percent: 40 }); }, 9000);
    setTimeout(function () { requestApproval(a.id, DEMO_APPROVALS[0]); }, 14000);
    setTimeout(function () { var b = demoCheckIn(); if (b) setTimeout(function () { progress(b.id, { percent: 30 }); }, 8000); }, 5000);
  }

  document.querySelectorAll('[data-demo]').forEach(function (b) {
    b.addEventListener('click', function () {
      var k = b.getAttribute('data-demo');
      if (k === 'checkin') demoCheckIn();
      if (k === 'approval') demoApproval();
      if (k === 'checkout') demoCheckOut();
      if (k === 'scenario') demoScenario();
      if (k === 'notify' && 'Notification' in window) Notification.requestPermission().then(function (p) { toast(p === 'granted' ? '알림을 켰어요.' : '알림이 허용되지 않았어요.'); });
    });
  });

  /* ================================================================
   * 14. 시작
   * ================================================================ */
  window.Office = {
    dispatch: dispatch,
    checkIn: checkIn,
    progress: progress,
    requestApproval: requestApproval,
    checkOut: checkOut,
    decide: decide,
    onDecision: function (fn) { state.listeners.push(fn); },
    agents: function () {
      return state.agents.map(function (a) {
        return { id: a.id, name: a.name, dept: a.dept, state: a.state, task: a.task, percent: a.percent, docs: a.docs };
      });
    },
    departments: DEPTS
  };

  resize();
  drawBackground();
  lastBg = performance.now();
  drawPanels();
  log('🏢', '사무실 문을 열었어요.');
  requestAnimationFrame(loop);
  // 폰트가 늦게 도착하면 한 번 더 그리기
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(function () { renderPanels(); });

  if (CONFIG.feedUrl) {
    pollFeed();
    setInterval(pollFeed, Math.max(2000, CONFIG.pollMs || 5000));
    log('🔌', '이벤트 피드 연결: ' + CONFIG.feedUrl);
  }
  if (CONFIG.autoDemo && !CONFIG.feedUrl) setTimeout(demoScenario, 800);
})();
