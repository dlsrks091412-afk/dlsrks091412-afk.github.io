/*
 * AI 에이전트 사무실 — 구름 위 42층 사무실 (아이소메트릭 일러스트)
 * ---------------------------------------------------------------
 * 모든 그림은 이 파일에서 캔버스 코드로 그립니다(이미지 파일 없음).
 * 사무실은 "월드 좌표"(타일 단위, x: 오른쪽 아래 / y: 왼쪽 아래)로 배치하고
 * iso() 함수로 화면 좌표(논리 500 x 352)로 바꿔 그립니다.
 *
 * 낮 테마: 창밖 파스텔 하늘·구름·새, 창으로 들어오는 햇살과 먼지, 가끔 바람이 불어 커튼·꽃잎·식물이 흔들림
 * 밤 테마: 창밖 도시 야경·별·달·별똥별, 실내는 어둡게 + 스탠드·모니터·네온 조명이 은은하게 빛남
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
   * 1. 기본 상수, 좌표계, 레이아웃
   * ================================================================ */
  var W = 424, H = 322;                  // 논리 화면 크기
  var HX = 18, HY = 9;                   // 타일 반쪽 크기 (2:1 아이소메트릭)
  var RX = 12, RY = 10;                  // 방 크기 (타일)
  var OX = 194, OY = 98;                 // 방 안쪽 꼭짓점(뒤쪽 모서리)의 화면 위치
  var WH = 82;                           // 벽 높이
  var WT = 0.35;                         // 벽 두께 (타일)
  var FONT = '"Gowun Dodum", "Pretendard Variable", Pretendard, "Apple SD Gothic Neo", sans-serif';
  var SPEED = 2.3;                       // 걷는 속도 (타일 / 초)
  var TAU = Math.PI * 2;

  function iso(x, y, z) { return { x: OX + (x - y) * HX, y: OY + (x + y) * HY - (z || 0) }; }

  /* 벽 안쪽 면의 "벽 좌표" (h: 벽을 따라 가로 px, v: 벽 위쪽에서 아래로 px) */
  var WALLS = {
    right: { len: RX * HX, m: [1, 0.5, 0, 1, OX, OY - WH] },                      // y = 0 인 벽 (오른쪽 위)
    left: { len: RY * HX, m: [1, -0.5, 0, 1, OX - RY * HX, OY + RY * HY - WH] }   // x = 0 인 벽 (왼쪽 위), h=0 이 앞쪽 끝
  };
  function wallPt(wall, h, v) { var m = WALLS[wall].m; return { x: m[4] + h, y: m[5] + m[1] * h + v }; }

  /* 창문 (벽 좌표) */
  var WIN_TOP = 14, WIN_BOTTOM = 62;
  var WINDOWS = [
    { wall: 'right', h0: 12, h1: 64 },
    { wall: 'right', h0: 72, h1: 124 },
    { wall: 'right', h0: 152, h1: 192 },
    { wall: 'left', h0: 68, h1: 120 }
  ];
  var PILLAR_H = 138;                    // 오른쪽 벽 기둥 (시계, 네온 달)
  var ELEV = { h0: 132, h1: 162, v0: 22 };  // 왼쪽 벽 엘리베이터 (월드 y ≈ 1.0 ~ 2.7)
  var BOARD = { h0: 22, h1: 56, v0: 22, v1: 48 }; // 왼쪽 벽 화이트보드

  /* 동선 (월드 좌표) */
  var DOOR = { x: 0.25, y: 1.85 };
  var HALL = { x: 1.4, y: 1.85 };
  var LANE_X = 1.4;                      // 왼쪽 세로 통로
  var CROSS_X = 7.45;                    // 가운데 세로 통로
  var BOSS_Y = 3.1;                      // 대표 책상 앞 대기 줄

  var BOSS_DESK = { x: 8.35, y: 1.5, w: 2.2, d: 1.0 };
  var BOSS_SEAT = { x: 9.45, y: 1.0 };
  var BOSS_SLOTS = [9.45, 8.65, 10.25, 7.85, 11.05];
  var BOSS_RUG = [7.7, 0.6, 11.8, 3.9];

  /* 부서 정의 — 새 부서는 여기에 추가합니다. (README 참고)
   * 책상은 두 줄로 마주 보게 붙어 있습니다: 뒷줄(row 0)은 이쪽을 보고, 앞줄(row 1)은 등을 보입니다. */
  var DEPTS = {
    records: {
      name: '기록부',
      desc: '문서 작업 담당',
      x0: 2.4, y0: 4.4, deskW: 1.4, deskD: 0.9, cols: 3,
      sign: { x: 1.9, y: 5.3 },
      neon: '#7ce7c8',
      chair: '#9fc9b0',
      lines: ['문서 정리 중…', '오탈자 꼼꼼히 확인 중', '목차를 다듬고 있어요', '표 모양 맞추는 중', '요약본 쓰는 중', '킁킁, 참고자료 찾는 중', '꼬리 살랑~ 거의 다 했어요'],
      unit: '문서'
    }
  };
  var FUTURE_ZONE = [8.3, 5.4, 11.6, 9.4];

  Object.keys(DEPTS).forEach(function (key) {
    var d = DEPTS[key];
    d.seats = [];
    for (var row = 0; row < 2; row++) {
      for (var c = 0; c < d.cols; c++) {
        var dx = d.x0 + c * d.deskW, dy = d.y0 + row * d.deskD;
        d.seats.push({
          dept: key, index: d.seats.length, row: row,
          desk: { x: dx, y: dy, w: d.deskW, d: d.deskD },
          x: dx + d.deskW / 2,
          y: row === 0 ? dy - 0.5 : dy + d.deskD + 0.5,
          aisleY: row === 0 ? d.y0 - 1.3 : d.y0 + 2 * d.deskD + 1.3,
          face: row === 0 ? 'SW' : 'NE',
          agent: null, docs: 0
        });
      }
    }
  });

  /* 식물 (흔들림) */
  var PLANTS = [
    { x: 0.55, y: 0.55, kind: 'monstera', s: 1.15 },
    { x: 7.75, y: 0.45, kind: 'snake', s: 1.0 },
    { x: 11.5, y: 4.95, kind: 'monstera', s: 1.0 },
    { x: 7.85, y: 9.55, kind: 'bush', s: 1.1 },
    { x: 6.6, y: 9.5, kind: 'monstera', s: 0.95 },
    { x: 0.5, y: 3.3, kind: 'snake', s: 0.9 },
    { x: 7.0, y: 4.0, kind: 'bush', s: 0.7 }
  ];
  PLANTS.forEach(function (p, i) { p.ph = i * 1.7 + 0.3; });
  var HANGING = [{ wall: 'right', h: 68, ph: 0.4, len: 22 }, { wall: 'left', h: 62, ph: 2.1, len: 18 }];

  /* ================================================================
   * 2. 캔버스 준비 & 그리기 도우미
   * ================================================================ */
  var canvas = document.getElementById('office');
  var ctx = canvas.getContext('2d');
  var S = 2.5, TS = 1;
  var backdrop = document.createElement('canvas');   // 방 바깥 하늘
  var room = document.createElement('canvas');       // 벽·바닥 (정적)
  var skyStrip = { right: document.createElement('canvas'), left: document.createElement('canvas') };
  var shade = document.createElement('canvas');
  var fade = document.createElement('canvas');
  var fadeAlpha = 0;
  var theme = document.documentElement.getAttribute('data-theme') === 'night' ? 'night' : 'day';
  var NIGHT = theme === 'night';

  function resize() {
    var cssW = canvas.clientWidth || 960;
    var dpr = Math.min(2.5, window.devicePixelRatio || 1);
    S = Math.max(1, cssW * dpr / W);
    TS = cssW < 560 ? 1.45 : cssW < 860 ? 1.18 : 1;
    [canvas, backdrop, room, shade, fade].forEach(function (c) { c.width = Math.round(W * S); c.height = Math.round(H * S); });
    ['right', 'left'].forEach(function (k) { skyStrip[k].width = Math.round(WALLS[k].len * S); skyStrip[k].height = Math.round(WIN_BOTTOM * S); });
    buildStatic();
  }

  function rrPath(g, x, y, w, h, r) {
    r = Math.min(r, w / 2, h / 2);
    g.beginPath();
    g.moveTo(x + r, y);
    g.arcTo(x + w, y, x + w, y + h, r);
    g.arcTo(x + w, y + h, x, y + h, r);
    g.arcTo(x, y + h, x, y, r);
    g.arcTo(x, y, x + w, y, r);
    g.closePath();
  }
  function rr(g, x, y, w, h, r, fill) { rrPath(g, x, y, w, h, r); g.fillStyle = fill; g.fill(); }
  function ell(g, x, y, rx, ry, fill, rot) { g.beginPath(); g.ellipse(x, y, Math.max(0.01, rx), Math.max(0.01, ry), rot || 0, 0, TAU); g.fillStyle = fill; g.fill(); }
  function circ(g, x, y, r, fill) { ell(g, x, y, r, r, fill); }
  function lin(g, x0, y0, x1, y1, stops) {
    var gr = g.createLinearGradient(x0, y0, x1, y1);
    stops.forEach(function (s) { gr.addColorStop(s[0], s[1]); });
    return gr;
  }
  function rad(g, x, y, r, stops) {
    var gr = g.createRadialGradient(x, y, 0, x, y, r);
    stops.forEach(function (s) { gr.addColorStop(s[0], s[1]); });
    return gr;
  }
  function stroke(g, w, c) { g.lineWidth = w; g.strokeStyle = c; g.lineCap = 'round'; g.lineJoin = 'round'; g.stroke(); }
  function poly(g, pts, fill) {
    g.beginPath();
    pts.forEach(function (p, i) { g[i ? 'lineTo' : 'moveTo'](p.x, p.y); });
    g.closePath();
    if (fill) { g.fillStyle = fill; g.fill(); }
  }
  function floorPoly(g, x0, y0, x1, y1, fill, z) {
    poly(g, [iso(x0, y0, z), iso(x1, y0, z), iso(x1, y1, z), iso(x0, y1, z)], fill);
  }
  function floorEll(g, x, y, r, fill, z) {
    var c = iso(x, y, z || 0);
    ell(g, c.x, c.y, r * HX * 1.414, r * HY * 1.414, fill);
  }
  function softShadow(g, x, y, rx, ry, a) {
    var c = iso(x, y, 0);
    g.save(); g.translate(c.x, c.y); g.scale(1, ry / rx);
    g.fillStyle = rad(g, 0, 0, rx, [[0, 'rgba(110,70,60,' + a + ')'], [1, 'rgba(110,70,60,0)']]);
    g.beginPath(); g.arc(0, 0, rx, 0, TAU); g.fill(); g.restore();
  }
  function hexToRgb(h) {
    if (h.charAt(0) !== '#') { var m = h.match(/[\d.]+/g); return [+m[0], +m[1], +m[2]]; }
    var n = parseInt(h.slice(1), 16); return [n >> 16 & 255, n >> 8 & 255, n & 255];
  }
  function shadeCol(hex, k) {      // k < 0: 어둡게, k > 0: 밝게
    var c = hexToRgb(hex);
    var f = function (v) { return Math.round(k < 0 ? v * (1 + k) : v + (255 - v) * k); };
    return 'rgb(' + f(c[0]) + ',' + f(c[1]) + ',' + f(c[2]) + ')';
  }
  /* 상자: 윗면 + 왼쪽 앞면(y+d) + 오른쪽 앞면(x+w) */
  function box(g, x, y, z, w, d, h, col, opt) {
    opt = opt || {};
    var top = opt.top || col, left = opt.left || shadeCol(col, -0.1), right = opt.right || shadeCol(col, -0.2);
    poly(g, [iso(x, y + d, z + h), iso(x + w, y + d, z + h), iso(x + w, y + d, z), iso(x, y + d, z)], left);
    poly(g, [iso(x + w, y, z + h), iso(x + w, y + d, z + h), iso(x + w, y + d, z), iso(x + w, y, z)], right);
    poly(g, [iso(x, y, z + h), iso(x + w, y, z + h), iso(x + w, y + d, z + h), iso(x, y + d, z + h)], top);
    if (!opt.noEdge) {
      g.beginPath(); var a = iso(x, y + d, z + h), b = iso(x + w, y + d, z + h), c = iso(x + w, y, z + h);
      g.moveTo(a.x, a.y); g.lineTo(b.x, b.y); g.lineTo(c.x, c.y);
      stroke(g, 0.5, 'rgba(255,255,255,.45)');
    }
  }
  function wallTransform(g, wall) { var m = WALLS[wall].m; g.transform(m[0], m[1], m[2], m[3], m[4], m[5]); }

  var seedN = 1;
  function srand(s) { seedN = s; }
  function rnd() { seedN = (seedN * 16807) % 2147483647; return (seedN - 1) / 2147483646; }
  function hashStr(s) { var h = 0; for (var i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0; return Math.abs(h); }
  function clamp(v, a, b) { return Math.max(a, Math.min(b, v)); }

  /* ================================================================
   * 3. 정적인 그림 (바깥 하늘, 벽·바닥, 창밖 풍경) — 테마/크기 바뀔 때만 다시 그림
   * ================================================================ */
  var CITY = { right: { tw: [], av: [] }, left: { tw: [], av: [] }, stars: [] };

  function buildStatic() {
    CITY.stars = [];
    drawBackdrop();
    drawRoom();
    drawSkyStrip('right');
    drawSkyStrip('left');
  }

  function drawBackdrop() {
    var g = backdrop.getContext('2d');
    g.setTransform(S, 0, 0, S, 0, 0);
    g.clearRect(0, 0, W, H);
    srand(3);
    if (!NIGHT) {
      g.fillStyle = lin(g, 0, 0, 0, H, [[0, '#bfe0f3'], [0.55, '#f5e1e8'], [1, '#fbe9df']]);
      g.fillRect(0, 0, W, H);
      g.fillStyle = rad(g, W - 40, 30, 110, [[0, 'rgba(255,248,220,.9)'], [1, 'rgba(255,248,220,0)']]);
      g.fillRect(0, 0, W, H);
    } else {
      g.fillStyle = lin(g, 0, 0, 0, H, [[0, '#0b0f2e'], [0.6, '#221950'], [1, '#4b2b63']]);
      g.fillRect(0, 0, W, H);
      for (var i = 0; i < 120; i++) circ(g, rnd() * W, rnd() * H * 0.75, 0.3 + rnd() * 0.5, 'rgba(255,250,230,' + (0.3 + rnd() * 0.6).toFixed(2) + ')');
    }
    // 아주 멀리 아래쪽 도시
    var base = H;
    var x = -5;
    while (x < W) {
      var bw = 8 + rnd() * 14, top = base - 18 - rnd() * 30;
      g.fillStyle = NIGHT ? '#1f1b45' : 'rgba(200,192,226,.75)';
      g.fillRect(x, top, bw, base - top);
      for (var wy = top + 3; wy < base - 2; wy += 4) for (var wx = x + 2; wx < x + bw - 2; wx += 3.2) {
        if (rnd() < (NIGHT ? 0.45 : 0.25)) { g.fillStyle = NIGHT ? (rnd() < 0.8 ? 'rgba(255,214,140,.85)' : 'rgba(150,225,255,.8)') : 'rgba(255,255,255,.6)'; g.fillRect(wx, wy, 1.2, 1.6); }
      }
      x += bw + 1 + rnd() * 3;
    }
    g.fillStyle = lin(g, 0, H - 70, 0, H, NIGHT ? [[0, 'rgba(120,70,140,0)'], [1, 'rgba(160,90,160,.45)']] : [[0, 'rgba(255,245,245,0)'], [1, 'rgba(255,245,245,.7)']]);
    g.fillRect(0, H - 70, W, 70);
  }

  function drawRoom() {
    var g = room.getContext('2d');
    g.setTransform(S, 0, 0, S, 0, 0);
    g.clearRect(0, 0, W, H);

    // 바닥판 두께 + 그림자
    var slab = 12;
    var sc = iso(RX * 0.62, RY * 0.62, -slab - 8);
    g.save(); g.translate(sc.x, sc.y); g.scale(1, 0.32);
    g.fillStyle = rad(g, 0, 0, 230, [[0, NIGHT ? 'rgba(0,0,0,.35)' : 'rgba(150,100,120,.2)'], [1, 'rgba(150,100,120,0)']]);
    g.beginPath(); g.arc(0, 0, 230, 0, TAU); g.fill(); g.restore();
    poly(g, [iso(-WT, RY, 0), iso(RX, RY, 0), iso(RX, RY, -slab), iso(-WT, RY, -slab)], '#e7cdb8');
    poly(g, [iso(RX, -WT, 0), iso(RX, RY, 0), iso(RX, RY, -slab), iso(RX, -WT, -slab)], '#d4b49c');
    poly(g, [iso(-WT, RY, -slab + 2), iso(RX, RY, -slab + 2), iso(RX, RY, -slab), iso(-WT, RY, -slab)], '#cfae95');
    poly(g, [iso(RX, -WT, -slab + 2), iso(RX, RY, -slab + 2), iso(RX, RY, -slab), iso(RX, -WT, -slab)], '#bf9c84');

    // 바닥: 원목 마루 (x 방향으로 긴 판)
    floorPoly(g, 0, 0, RX, RY, '#ecd2b4');
    srand(9);
    var plank = 0.5;
    for (var y = 0; y < RY; y += plank) {
      var x = 0;
      var shift = rnd() * 2;
      while (x < RX) {
        var len = Math.min(RX - x, (x === 0 ? shift : 0) + 1.6 + rnd() * 1.8);
        var tone = 0.9 + rnd() * 0.12;
        var c = hexToRgb('#ecd2b4');
        floorPoly(g, x, y, x + len, y + plank, 'rgb(' + Math.round(c[0] * tone) + ',' + Math.round(c[1] * tone) + ',' + Math.round(c[2] * tone) + ')');
        var a = iso(x, y, 0), b = iso(x, y + plank, 0);
        g.beginPath(); g.moveTo(a.x, a.y); g.lineTo(b.x, b.y); stroke(g, 0.35, 'rgba(150,100,70,.18)');
        x += len;
      }
      var p0 = iso(0, y, 0), p1 = iso(RX, y, 0);
      g.beginPath(); g.moveTo(p0.x, p0.y); g.lineTo(p1.x, p1.y); stroke(g, 0.35, 'rgba(150,100,70,.16)');
    }
    // 벽 아래 그림자
    poly(g, [iso(0, 0, 0), iso(RX, 0, 0), iso(RX, 0.6, 0), iso(0.6, 0.6, 0), iso(0.6, RY, 0), iso(0, RY, 0)], 'rgba(120,80,60,.10)');

    // 벽
    poly(g, [iso(0, 0, 0), iso(RX, 0, 0), iso(RX, 0, WH), iso(0, 0, WH)], '#f7e7dc');
    poly(g, [iso(0, 0, 0), iso(0, RY, 0), iso(0, RY, WH), iso(0, 0, WH)], '#efdccd');
    ['right', 'left'].forEach(function (wl) {
      g.save(); wallTransform(g, wl);
      var L = WALLS[wl].len;
      g.fillStyle = wl === 'right' ? '#f1ddcf' : '#e8d2c2'; g.fillRect(0, 64, L, WH - 64);     // 아래쪽 패널
      g.fillStyle = 'rgba(255,255,255,.55)'; g.fillRect(0, 64, L, 0.8);
      g.fillStyle = '#dcbfa8'; g.fillRect(0, WH - 4, L, 4);                                   // 걸레받이
      g.fillStyle = 'rgba(255,255,255,.35)'; g.fillRect(0, 2, L, 1);
      g.restore();
    });
    // 벽 두께 (윗면, 앞쪽 끝)
    poly(g, [iso(0, 0, WH), iso(RX, 0, WH), iso(RX, -WT, WH), iso(-WT, -WT, WH)], '#e2c8b6');
    poly(g, [iso(0, 0, WH), iso(0, RY, WH), iso(-WT, RY, WH), iso(-WT, -WT, WH)], '#e9d2c1');
    poly(g, [iso(RX, -WT, 0), iso(RX, 0, 0), iso(RX, 0, WH), iso(RX, -WT, WH)], '#d9bca8');
    poly(g, [iso(-WT, RY, 0), iso(0, RY, 0), iso(0, RY, WH), iso(-WT, RY, WH)], '#e2c6b2');

    // 기둥 장식 판
    g.save(); wallTransform(g, 'right');
    rr(g, PILLAR_H - 13, 12, 26, 52, 6, 'rgba(255,255,255,.35)');
    g.restore();

    // 화이트보드
    g.save(); wallTransform(g, 'left');
    rr(g, BOARD.h0 - 1.5, BOARD.v0 - 1.5, BOARD.h1 - BOARD.h0 + 3, BOARD.v1 - BOARD.v0 + 3, 2, '#c9c3c8');
    rr(g, BOARD.h0, BOARD.v0, BOARD.h1 - BOARD.h0, BOARD.v1 - BOARD.v0, 1.5, '#fdfcfa');
    g.fillStyle = '#9cc3e6'; g.fillRect(BOARD.h0 + 4, BOARD.v0 + 5, 16, 1); g.fillRect(BOARD.h0 + 4, BOARD.v0 + 9, 22, 1);
    g.fillStyle = '#f0a7b2'; g.fillRect(BOARD.h0 + 4, BOARD.v0 + 13, 12, 1);
    [[0, 6, '#a6d1b0'], [5, 10, '#f3cf86'], [10, 4, '#9cc3e6']].forEach(function (b) { g.fillStyle = b[2]; g.fillRect(BOARD.h0 + 22 + b[0] * 0.8, BOARD.v1 - 4 - b[1], 3, b[1]); });
    rr(g, BOARD.h0 + 4, BOARD.v1 + 1.5, BOARD.h1 - BOARD.h0 - 8, 1.6, 0.8, '#c9c3c8');
    g.restore();

    // 러그 (바닥에 붙어 있는 것들)
    var br = BOSS_RUG;
    floorPoly(g, br[0], br[1], br[2], br[3], '#efc7bc');
    floorPoly(g, br[0] + 0.15, br[1] + 0.15, br[2] - 0.15, br[3] - 0.15, '#f7dcd2');
    g.save(); g.setLineDash([2, 2]);
    floorPoly(g, br[0] + 0.4, br[1] + 0.4, br[2] - 0.4, br[3] - 0.4); stroke(g, 0.6, 'rgba(255,255,255,.8)');
    g.restore();
    // 휴게 공간 러그
    floorEll(g, 2.2, 8.7, 1.6, '#f3e3d3');
    floorEll(g, 2.2, 8.7, 1.38, '#f8ece0');
    // 새 부서 자리 (점선)
    var fz = FUTURE_ZONE;
    floorPoly(g, fz[0], fz[1], fz[2], fz[3], 'rgba(255,255,255,.2)');
    g.save(); g.setLineDash([3, 3]);
    floorPoly(g, fz[0], fz[1], fz[2], fz[3]); stroke(g, 0.7, 'rgba(170,120,100,.45)');
    g.restore();
    // 부서 자리 아래 카펫 타일 (은은하게)
    Object.keys(DEPTS).forEach(function (k) {
      var d = DEPTS[k];
      var x0 = d.x0 - 0.5, y0 = d.y0 - 1.0, x1 = d.x0 + d.cols * d.deskW + 0.5, y1 = d.y0 + 2 * d.deskD + 1.0;
      floorPoly(g, x0, y0, x1, y1, 'rgba(160,200,170,.22)');
      floorPoly(g, x0, y0, x1, y1); stroke(g, 0.5, 'rgba(120,170,135,.3)');
    });
  }

  /* 창밖 풍경 띠 (벽 좌표) */
  function drawSkyStrip(wall) {
    var cv = skyStrip[wall], g = cv.getContext('2d');
    var L = WALLS[wall].len;
    g.setTransform(S, 0, 0, S, 0, 0);
    g.clearRect(0, 0, L, WIN_BOTTOM);
    srand(wall === 'right' ? 42 : 17);
    var info = CITY[wall]; info.tw = []; info.av = [];
    if (!NIGHT) {
      g.fillStyle = lin(g, 0, WIN_TOP, 0, WIN_BOTTOM, [[0, '#a9d3f0'], [0.55, '#d6e8f4'], [0.85, '#fbe2e3'], [1, '#fde9df']]);
      g.fillRect(0, 0, L, WIN_BOTTOM);
      if (wall === 'right') {
        g.fillStyle = rad(g, 176, 24, 46, [[0, 'rgba(255,250,225,.95)'], [0.2, 'rgba(255,240,200,.6)'], [1, 'rgba(255,240,200,0)']]);
        g.fillRect(0, 0, L, WIN_BOTTOM);
        circ(g, 176, 24, 5, 'rgba(255,253,240,.95)');
      }
    } else {
      g.fillStyle = lin(g, 0, WIN_TOP, 0, WIN_BOTTOM, [[0, '#0d1131'], [0.55, '#271d57'], [0.85, '#5a3470'], [1, '#7a4277']]);
      g.fillRect(0, 0, L, WIN_BOTTOM);
      for (var i = 0; i < 40; i++) {
        var sx = rnd() * L, sy = WIN_TOP + rnd() * 26;
        circ(g, sx, sy, 0.25 + rnd() * 0.4, 'rgba(255,250,230,' + (0.4 + rnd() * 0.5).toFixed(2) + ')');
        if (rnd() < 0.35) CITY.stars.push({ wall: wall, h: sx, v: sy, ph: rnd() * TAU });
      }
      if (wall === 'left') {
        g.fillStyle = rad(g, 98, 26, 26, [[0, 'rgba(230,220,255,.5)'], [1, 'rgba(230,220,255,0)']]);
        g.fillRect(0, 0, L, WIN_BOTTOM);
        circ(g, 98, 26, 5.5, '#fff8e4');
        circ(g, 101, 24, 5, '#211a50');
      }
    }
    // 먼 산
    g.fillStyle = NIGHT ? '#2b2457' : 'rgba(214,196,226,.85)';
    g.beginPath(); g.moveTo(0, 46);
    for (var mx = 0; mx <= L; mx += 12) g.lineTo(mx, 44 - Math.sin(mx * 0.03) * 3 - Math.sin(mx * 0.08) * 1.5);
    g.lineTo(L, WIN_BOTTOM); g.lineTo(0, WIN_BOTTOM); g.fill();
    // 도시 세 겹
    var layers = NIGHT
      ? [{ top: 44, v: 8, col: '#29265a', win: 0.3 }, { top: 50, v: 9, col: '#211e4b', win: 0.42 }, { top: 56, v: 6, col: '#191739', win: 0.5 }]
      : [{ top: 44, v: 8, col: 'rgba(196,192,224,.9)', win: 0 }, { top: 50, v: 9, col: '#b9bbdc', win: 0.3 }, { top: 56, v: 6, col: '#aab0d1', win: 0.4 }];
    layers.forEach(function (Ly, li) {
      var x = -3;
      while (x < L) {
        var bw = 4 + rnd() * 9, top = Ly.top - rnd() * Ly.v;
        if (rnd() < 0.08) top -= 6;
        g.fillStyle = Ly.col; g.fillRect(x, top, bw, WIN_BOTTOM - top);
        if (li === 0 && top < Ly.top - 6) info.av.push({ h: x + bw / 2, v: top - 1, ph: rnd() * TAU });
        for (var wy = top + 2; wy < WIN_BOTTOM - 1; wy += 2.6) {
          for (var wx = x + 1.2; wx < x + bw - 1.2; wx += 2.2) {
            if (rnd() > Ly.win) continue;
            if (NIGHT) {
              var r = rnd();
              g.fillStyle = r < 0.62 ? 'rgba(255,214,140,.9)' : r < 0.85 ? 'rgba(255,240,205,.85)' : 'rgba(150,225,255,.8)';
              if (rnd() < 0.05) info.tw.push({ h: wx, v: wy, ph: rnd() * TAU });
            } else g.fillStyle = 'rgba(255,255,255,.55)';
            g.fillRect(wx, wy, 0.9, 1.2);
          }
        }
        x += bw + 0.6 + rnd() * 1.6;
      }
    });
    g.fillStyle = lin(g, 0, 50, 0, WIN_BOTTOM, NIGHT ? [[0, 'rgba(120,70,140,0)'], [1, 'rgba(120,70,140,.35)']] : [[0, 'rgba(255,240,240,0)'], [1, 'rgba(255,240,240,.5)']]);
    g.fillRect(0, 50, L, WIN_BOTTOM - 50);
  }

  /* ================================================================
   * 4. 창문, 커튼, 엘리베이터, 벽 장식 (매 프레임)
   * ================================================================ */
  var clouds = [];          // 창밖 구름 (u: 왼쪽 벽 0~216, 오른쪽 벽 216~468)
  var outerClouds = [];     // 방 바깥 아래쪽 구름
  srand(77);
  for (var ci = 0; ci < 8; ci++) clouds.push({ u: rnd() * 400, v: 18 + rnd() * 20, s: 0.4 + rnd() * 0.45, sp: 2 + rnd() * 3 });
  for (var oc = 0; oc < 8; oc++) outerClouds.push({ x: rnd() * (W + 120) - 60, y: H * 0.6 + rnd() * H * 0.38, s: 0.8 + rnd() * 1.1, sp: 2 + rnd() * 4 });

  function drawCloud(g, x, y, s, a, col) {
    g.save(); g.translate(x, y); g.scale(s, s);
    g.globalAlpha = a;
    col = col || '#ffffff';
    ell(g, 0, 4, 22, 5.5, col);
    circ(g, -9, 1, 6.5, col); circ(g, 1, -2, 8.5, col); circ(g, 11, 1, 6, col);
    ell(g, 0, 7, 20, 2.5, NIGHT ? 'rgba(40,30,80,.4)' : 'rgba(230,215,235,.7)');
    g.restore(); g.globalAlpha = 1;
  }

  function drawOuterClouds(t, dt) {
    var g = ctx;
    outerClouds.forEach(function (c) {
      c.x += c.sp * dt * (1 + state.wind);
      if (c.x > W + 70) { c.x = -70; c.y = H * 0.6 + Math.random() * H * 0.38; }
      drawCloud(g, c.x, c.y, c.s, NIGHT ? 0.22 : 0.85, NIGHT ? '#8f80c8' : '#ffffff');
    });
  }

  function windowsOn(wall) { return WINDOWS.filter(function (w) { return w.wall === wall; }); }

  function drawWindows(t, dt) {
    var g = ctx;
    clouds.forEach(function (c) {
      c.u += c.sp * dt * (1 + state.wind * 2);
      if (c.u > 420) { c.u = -30; c.v = 18 + Math.random() * 20; }
    });
    ['left', 'right'].forEach(function (wl) {
      var off = wl === 'left' ? 0 : WALLS.left.len;
      g.save(); wallTransform(g, wl);
      g.beginPath();
      windowsOn(wl).forEach(function (w) { g.rect(w.h0, WIN_TOP, w.h1 - w.h0, WIN_BOTTOM - WIN_TOP); });
      g.clip();
      g.drawImage(skyStrip[wl], 0, 0, WALLS[wl].len, WIN_BOTTOM);
      if (!NIGHT) {
        clouds.forEach(function (c) { var h = c.u - off; if (h > -30 && h < WALLS[wl].len + 30) drawCloud(g, h, c.v, c.s, 0.92); });
        if (state.birds) drawBirds(g, t, off);
      } else {
        CITY.stars.forEach(function (s) {
          if (s.wall !== wl) return;
          var a = 0.35 + 0.65 * Math.max(0, Math.sin(t * 1.3 + s.ph));
          circ(g, s.h, s.v, 0.5, 'rgba(255,250,230,' + a.toFixed(2) + ')');
        });
        CITY[wl].tw.forEach(function (w) { if (Math.sin(t * 0.7 + w.ph) > 0.6) { g.fillStyle = 'rgba(40,30,70,.9)'; g.fillRect(w.h, w.v, 0.9, 1.2); } });
        CITY[wl].av.forEach(function (p) {
          if (Math.sin(t * 3 + p.ph) > 0.4) { circ(g, p.h, p.v, 1.8, 'rgba(255,80,100,.25)'); circ(g, p.h, p.v, 0.6, '#ff6a7c'); }
        });
        if (state.shooting && state.shooting.wall === wl) drawShootingStar(g);
      }
      // 유리 반사
      windowsOn(wl).forEach(function (w) {
        g.fillStyle = 'rgba(255,255,255,' + (NIGHT ? 0.04 : 0.15) + ')';
        g.beginPath(); g.moveTo(w.h0 + 6, WIN_BOTTOM); g.lineTo(w.h0 + 22, WIN_TOP); g.lineTo(w.h0 + 30, WIN_TOP); g.lineTo(w.h0 + 14, WIN_BOTTOM); g.fill();
      });
      g.restore();

      // 창틀
      g.save(); wallTransform(g, wl);
      var frame = NIGHT ? '#e3d8e4' : '#fffaf4';
      windowsOn(wl).forEach(function (w) {
        g.strokeStyle = frame; g.lineWidth = 2.2;
        g.strokeRect(w.h0 - 1, WIN_TOP - 1, w.h1 - w.h0 + 2, WIN_BOTTOM - WIN_TOP + 2);
        g.fillStyle = frame;
        g.fillRect((w.h0 + w.h1) / 2 - 0.8, WIN_TOP, 1.6, WIN_BOTTOM - WIN_TOP);
        g.fillRect(w.h0, WIN_TOP + 30, w.h1 - w.h0, 1.2);
        rr(g, w.h0 - 3, WIN_BOTTOM + 1, w.h1 - w.h0 + 6, 2.6, 1.3, '#f2e2d4');
      });
      g.restore();
    });
  }

  function drawBirds(g, t, off) {
    var b = state.birds;
    for (var i = 0; i < b.n; i++) {
      var bx = b.u - off - Math.abs(i - (b.n - 1) / 2) * 5, by = b.v + i * 1.8 + Math.sin(t * 2 + i) * 0.6;
      var flap = Math.sin(t * 12 + i * 1.3) * 1.3;
      g.beginPath(); g.moveTo(bx - 2.4, by - flap); g.quadraticCurveTo(bx - 1.2, by - 0.8, bx, by); g.quadraticCurveTo(bx + 1.2, by - 0.8, bx + 2.4, by - flap);
      stroke(g, 0.6, 'rgba(90,80,110,.75)');
    }
  }

  function drawShootingStar(g) {
    var s = state.shooting, k = s.life / s.max;
    var tx = s.h - s.vh * 0.3, ty = s.v - s.vv * 0.3;
    g.strokeStyle = lin(g, s.h, s.v, tx, ty, [[0, 'rgba(255,250,235,' + k + ')'], [1, 'rgba(255,250,235,0)']]);
    g.lineWidth = 0.8; g.lineCap = 'round';
    g.beginPath(); g.moveTo(s.h, s.v); g.lineTo(tx, ty); g.stroke();
  }

  /* 커튼: 바람(state.wind)이 불면 아래쪽부터 펄럭임 (벽 좌표에서 그림) */
  function drawCurtains(t) {
    var g = ctx;
    var fabric = NIGHT ? 'rgba(170,160,220,.42)' : 'rgba(255,249,244,.9)';
    var fold = NIGHT ? 'rgba(90,80,150,.35)' : 'rgba(232,196,190,.6)';
    var tie = NIGHT ? 'rgba(200,190,240,.55)' : 'rgba(240,200,195,.95)';
    WINDOWS.forEach(function (w, wi) {
      g.save(); wallTransform(g, w.wall);
      g.fillStyle = '#d2b29a'; g.fillRect(w.h0 - 6, 9, w.h1 - w.h0 + 12, 1.3);
      circ(g, w.h0 - 6, 9.6, 1.4, '#c79f84'); circ(g, w.h1 + 6, 9.6, 1.4, '#c79f84');
      [0, 1].forEach(function (side) {
        var cw = 12;
        var x0 = side === 0 ? w.h0 - 4.5 : w.h1 + 4.5 - cw;
        var ph = wi * 1.9 + side * 0.7;
        var gust = state.wind;
        function sway(v) {
          var k = Math.pow((v - 10) / 56, 1.6);
          return (Math.sin(t * 0.9 + ph) * 0.7 + gust * (7 + 3 * Math.sin(t * 4.6 + ph + v * 0.1))) * k;
        }
        var bottom = function (xx) { return 66 - gust * 5 * (0.6 + 0.4 * Math.sin(t * 5 + xx * 0.5 + ph)); };
        g.beginPath();
        g.moveTo(x0, 10);
        for (var v = 10; v <= 66; v += 4) g.lineTo(x0 + sway(v), v);
        for (var i = 0; i <= 5; i++) {
          var xx = x0 + (cw * i) / 5, bv = bottom(xx);
          g.lineTo(xx + sway(bv), bv + Math.sin(i * 1.7 + t * 2) * 0.6);
        }
        for (var v2 = 66; v2 >= 10; v2 -= 4) g.lineTo(x0 + cw + sway(v2) * 1.05, v2);
        g.closePath();
        g.fillStyle = fabric; g.fill();
        for (var f = 1; f < 4; f++) {
          var fx = x0 + (cw * f) / 4;
          g.beginPath(); g.moveTo(fx, 11);
          for (var v3 = 15; v3 <= 64; v3 += 5) g.lineTo(fx + sway(v3) + Math.sin(v3 * 0.18 + f) * 0.4, v3);
          stroke(g, 0.5, fold);
        }
        rr(g, x0 + sway(44) + 1.5, 44, cw - 3, 1.6, 0.8, tie);
      });
      g.restore();
    });
  }

  function drawElevator() {
    var g = ctx, e = ELEV;
    g.save(); wallTransform(g, 'left');
    var w = e.h1 - e.h0, h = WH - e.v0;
    rr(g, e.h0 - 3, e.v0 - 7, w + 6, h + 7, 2, '#e5d4c8');
    g.fillStyle = NIGHT ? '#fff0d4' : '#fff6e6'; g.fillRect(e.h0, e.v0, w, h);
    g.fillStyle = lin(g, 0, e.v0, 0, WH, [[0, 'rgba(255,225,180,.7)'], [1, 'rgba(255,225,180,0)']]); g.fillRect(e.h0, e.v0, w, h);
    var o = state.elev.open, lw = (w / 2) * (1 - o);
    var door = lin(g, e.h0, 0, e.h1, 0, [[0, '#d6d0d6'], [0.5, '#ebe6e9'], [1, '#cdc6ce']]);
    if (lw > 0.3) {
      g.fillStyle = door; g.fillRect(e.h0, e.v0, lw, h); g.fillRect(e.h1 - lw, e.v0, lw, h);
      g.fillStyle = 'rgba(140,120,135,.4)'; g.fillRect(e.h0 + lw - 0.4, e.v0, 0.5, h); g.fillRect(e.h1 - lw, e.v0, 0.5, h);
    }
    rr(g, e.h0 + 8, e.v0 - 5.5, w - 16, 4, 2, '#5d5068');
    var lit = o > 0.05 || state.elev.target > 0;
    g.fillStyle = lit ? '#ffd27a' : '#9c8fa6'; g.font = '3px ' + FONT; g.textAlign = 'center'; g.textBaseline = 'middle';
    g.fillText('42F', (e.h0 + e.h1) / 2, e.v0 - 3.4); g.textAlign = 'left';
    rr(g, e.h1 + 3, e.v0 + 24, 3, 6, 1.5, '#e3d6cc');
    circ(g, e.h1 + 4.5, e.v0 + 27, 0.8, lit ? '#ffcf7a' : '#c9b8ae');
    g.restore();
  }

  function drawClock() {
    var g = ctx;
    g.save(); wallTransform(g, 'right');
    var cx = PILLAR_H, cy = 24;
    circ(g, cx, cy, 8, '#d8b694'); circ(g, cx, cy, 6.8, '#fffaf2');
    for (var i = 0; i < 12; i++) { var a = i / 12 * TAU; circ(g, cx + Math.sin(a) * 5.5, cy - Math.cos(a) * 5.5, i % 3 ? 0.3 : 0.5, '#c8a98e'); }
    var now = new Date();
    var hA = ((now.getHours() % 12) + now.getMinutes() / 60) / 12 * TAU, mA = (now.getMinutes() + now.getSeconds() / 60) / 60 * TAU;
    g.beginPath(); g.moveTo(cx, cy); g.lineTo(cx + Math.sin(hA) * 3.3, cy - Math.cos(hA) * 3.3); stroke(g, 0.9, '#6b5560');
    g.beginPath(); g.moveTo(cx, cy); g.lineTo(cx + Math.sin(mA) * 5, cy - Math.cos(mA) * 5); stroke(g, 0.6, '#e8848c');
    circ(g, cx, cy, 0.7, '#6b5560');
    g.restore();
  }

  /* 네온 튜브 (낮엔 꺼진 유리관, 밤엔 은은하게 빛남) */
  function drawNeon(t, lit) {
    var g = ctx;
    var tubes = [
      { kind: 'cove', col: null, w: 1.2 },
      { kind: 'moon', col: '#ff9fd6', w: 1.1 },
      { kind: 'star', col: '#ffe08a', w: 0.9 }
    ];
    tubes.forEach(function (tb, i) {
      g.save();
      if (tb.kind === 'cove') {
        var a = iso(0, RY - 0.2, WH - 3), b = iso(0, 0, WH - 3), c = iso(RX - 0.2, 0, WH - 3);
        g.beginPath(); g.moveTo(a.x, a.y); g.lineTo(b.x, b.y); g.lineTo(c.x, c.y);
      } else {
        wallTransform(g, 'right');
        g.beginPath();
        if (tb.kind === 'moon') {
          g.arc(PILLAR_H - 1, 48, 7, Math.PI * 0.55, Math.PI * 1.85);
          g.arc(PILLAR_H + 2.5, 46, 5.4, Math.PI * 1.75, Math.PI * 0.62, true);
        } else {
          var cx = PILLAR_H + 7, cy = 40, r1 = 2.6, r2 = 1.1;
          for (var k = 0; k < 10; k++) { var an = -Math.PI / 2 + k * Math.PI / 5, r = k % 2 ? r2 : r1; g[k ? 'lineTo' : 'moveTo'](cx + Math.cos(an) * r, cy + Math.sin(an) * r); }
          g.closePath();
        }
      }
      var col = tb.col || lin(g, a ? a.x : 0, 0, c ? c.x : W, 0, [[0, '#7ce7e0'], [0.45, '#b49bff'], [1, '#ff9fd6']]);
      if (!lit) { stroke(g, tb.w, 'rgba(240,210,225,.75)'); g.restore(); return; }
      var flick = 0.85 + 0.15 * Math.sin(t * 2.2 + i * 2) + (Math.random() < 0.004 ? -0.4 : 0);
      g.globalAlpha = clamp(flick, 0.3, 1);
      g.shadowColor = tb.col || '#c7a8ff'; g.shadowBlur = 9 * S;
      stroke(g, tb.w * 1.6, col);
      g.shadowBlur = 0;
      stroke(g, tb.w * 0.55, 'rgba(255,255,255,.85)');
      g.restore();
    });
  }

  /* 벽에 매달린 화분 */
  function drawHanging(g, hp, t) {
    var anchor = wallPt(hp.wall, hp.h, 4);
    var a = Math.sin(t * 0.9 + hp.ph) * 0.05 + state.wind * 0.16 * Math.sin(t * 4 + hp.ph);
    g.save(); g.translate(anchor.x, anchor.y); g.rotate(a);
    g.beginPath(); g.moveTo(-3.5, hp.len - 5); g.lineTo(0, 0); g.lineTo(3.5, hp.len - 5); stroke(g, 0.35, '#b89b88');
    rr(g, -4.5, hp.len - 5, 9, 5.5, 2.6, '#fbf3ea');
    ell(g, 0, hp.len - 5, 4.5, 1.1, '#e9d8c8');
    [-3.5, -1.2, 1.2, 3.5].forEach(function (vx, i) {
      var len = 12 + (i % 2) * 7;
      var sw = Math.sin(t * 1.5 + i + hp.ph) * 1.2 + state.wind * 4 * Math.sin(t * 6 + i);
      g.beginPath(); g.moveTo(vx, hp.len - 2.5); g.quadraticCurveTo(vx + sw * 0.5, hp.len + len * 0.5, vx + sw, hp.len + len);
      stroke(g, 0.45, '#6fa57d');
      for (var k = 1; k <= 4; k++) {
        var f = k / 4, lx = vx + sw * f * f, ly = hp.len - 2.5 + len * f;
        ell(g, lx + (k % 2 ? 1 : -1), ly, 1.5, 0.95, k % 2 ? '#8cc39a' : '#a5d4ad', k % 2 ? 0.5 : -0.5);
      }
    });
    g.restore();
  }

  /* ================================================================
   * 5. 가구와 식물
   * ================================================================ */
  function plantSway(p, t) {
    return Math.sin(t * 1.1 + p.ph) * 0.03 + state.wind * 0.12 * Math.sin(t * 5.2 + p.ph) * (p.kind === 'snake' ? 0.5 : 1);
  }

  function drawPlant(g, p, t) {
    var c = iso(p.x, p.y, 0), a = plantSway(p, t), s = p.s * 0.9;
    softShadow(g, p.x, p.y, 8 * s, 3 * s, 0.22);
    g.save(); g.translate(c.x, c.y); g.scale(s, s);
    g.beginPath(); g.moveTo(-6, -10); g.lineTo(6, -10); g.lineTo(4.6, 0); g.lineTo(-4.6, 0); g.closePath();
    g.fillStyle = p.kind === 'snake' ? '#f6efe7' : '#e7a684'; g.fill();
    g.fillStyle = 'rgba(0,0,0,.06)'; g.fillRect(0, -10, 6, 10);
    rr(g, -7, -12, 14, 3, 1.5, p.kind === 'snake' ? '#e9ddd0' : '#d99375');
    g.translate(0, -11); g.rotate(a);
    if (p.kind === 'monstera') {
      [[-0.95, 16, '#7ab38a'], [-0.35, 20, '#8cc39a'], [0.35, 18, '#7ab38a'], [0.95, 14, '#a5d4ad'], [0.05, 24, '#6fa57d']].forEach(function (L, i) {
        g.save(); g.rotate(L[0] + state.wind * 0.12 * Math.sin(t * 7 + i));
        g.beginPath(); g.moveTo(0, 0); g.lineTo(0, -L[1] + 5); stroke(g, 0.7, '#6a9c76');
        ell(g, 0, -L[1] + 2, 5.2, 6.2, L[2]);
        g.strokeStyle = 'rgba(255,255,255,.35)'; g.lineWidth = 0.4;
        g.beginPath(); g.moveTo(0, -L[1] - 3.2); g.lineTo(0, -L[1] + 6.5); g.stroke();
        g.restore();
      });
    } else if (p.kind === 'snake') {
      [[-4.5, 17, -0.15], [-1.8, 23, -0.05], [1.4, 26, 0.04], [4.2, 19, 0.14]].forEach(function (b, i) {
        g.save(); g.translate(b[0], 0); g.rotate(b[2] + state.wind * 0.05 * Math.sin(t * 6 + i));
        g.beginPath(); g.moveTo(-1.7, 0); g.quadraticCurveTo(-2.1, -b[1] * 0.6, 0, -b[1]); g.quadraticCurveTo(2.1, -b[1] * 0.6, 1.7, 0); g.closePath();
        g.fillStyle = i % 2 ? '#7fb08a' : '#94c49f'; g.fill();
        g.beginPath(); g.moveTo(0, -1); g.lineTo(0, -b[1] + 3); stroke(g, 0.35, 'rgba(240,235,180,.6)');
        g.restore();
      });
    } else {
      [[-6, -4, 6], [5, -5, 6.5], [0, -10, 7], [-3, -15, 5], [4, -14, 5]].forEach(function (b, i) { circ(g, b[0] + state.wind * Math.sin(t * 6 + i) * 0.6, b[1], b[2], i % 2 ? '#8cc39a' : '#a5d4ad'); });
      circ(g, -2, -11, 1.2, '#f7b6c6'); circ(g, 3, -7, 1.1, '#fff1f4'); circ(g, -6, -5, 1, '#f7b6c6');
    }
    g.restore();
  }

  function drawSteam(g, x, y, t, seed) {
    for (var i = 0; i < 2; i++) {
      var p = ((t * 0.5 + i * 0.5 + seed) % 1);
      g.beginPath();
      for (var k = 0; k <= 6; k++) {
        var yy = y - k * 1.2 - p * 3, xx = x + i * 1.4 - 0.7 + Math.sin(k * 0.9 + t * 3 + i) * 0.7;
        g[k ? 'lineTo' : 'moveTo'](xx, yy);
      }
      stroke(g, 0.45, (NIGHT ? 'rgba(255,255,255,' : 'rgba(190,150,140,') + (0.55 * (1 - p)).toFixed(2) + ')');
    }
  }

  var DESK_H = 12;

  function drawChair(g, seat, occupied) {
    var d = DEPTS[seat.dept], col = d.chair;
    var back = seat.row === 0 ? seat.y - 0.32 : seat.y + 0.22;      // 등받이 위치
    var c = iso(seat.x, seat.y, 0);
    // 다리 + 바퀴
    g.beginPath(); g.moveTo(c.x, c.y - 7); g.lineTo(c.x, c.y - 1.5); stroke(g, 1, '#8c8494');
    ell(g, c.x, c.y - 0.8, 5, 1.6, 'rgba(120,110,130,.6)');
    box(g, seat.x - 0.32, seat.y - 0.3, 6.5, 0.64, 0.6, 2.5, col, { noEdge: true });
    // 등받이
    box(g, seat.x - 0.34, back, 8.5, 0.68, 0.12, seat.row === 0 ? 13 : 5, shadeCol(col, -0.05), { noEdge: true, left: shadeCol(col, 0.08), top: shadeCol(col, 0.12) });
    void occupied;
  }

  function drawDesk(g, seat, t) {
    var dk = seat.desk, x = dk.x, y = dk.y, w = dk.w, d = dk.d;
    var working = seat.agent && seat.agent.state === 'working' && !seat.agent.moving;
    var gap = 0.04;
    // 다리 + 상판
    box(g, x + gap, y + gap, 0, 0.08, 0.08, DESK_H - 1, '#c9a684', { noEdge: true });
    box(g, x + w - 0.12, y + d - 0.12, 0, 0.08, 0.08, DESK_H - 1, '#c9a684', { noEdge: true });
    box(g, x + gap, y + d - 0.12, 0, 0.08, 0.08, DESK_H - 1, '#c9a684', { noEdge: true });
    box(g, x + gap, y + gap, DESK_H - 7, w - gap * 2, d - gap * 2, 6, '#e3c8a8', { noEdge: true });   // 서랍/가림판
    box(g, x + gap, y + gap, DESK_H - 1, w - gap * 2, d - gap * 2, 1.2, '#f4e4cf', { left: '#dcbf9d', right: '#cfae8b' });

    var z = DESK_H + 0.2;
    var mx = x + w / 2;
    // 모니터 (두 줄 책상이 맞닿는 가운데 쪽)
    var my = seat.row === 0 ? y + d - 0.22 : y + 0.1;
    var mon = iso(mx, my, z);
    g.beginPath(); g.moveTo(mon.x, mon.y); g.lineTo(mon.x, mon.y - 3.5); stroke(g, 1.2, '#b9bdc8');
    ell(g, mon.x, mon.y, 2.6, 0.9, '#b9bdc8');
    var face = seat.row === 1;   // 앞줄: 화면이 보임
    box(g, mx - 0.38, my - 0.05, z + 3, 0.76, 0.1, 9.5, '#dfe2ea', {
      left: face ? (working ? (NIGHT ? '#bfe8ff' : '#d8f0ff') : '#7e8597') : '#d5d9e2',
      right: '#c4c8d3', top: '#eef0f5', noEdge: true
    });
    if (face && working) {
      var s0 = iso(mx - 0.3, my + 0.05, z + 10.5);
      for (var i = 0; i < 3; i++) {
        var a = iso(mx - 0.28, my + 0.06, z + 10 - i * 2.2), b = iso(mx + 0.18 - (i === 2 ? 0.2 : 0), my + 0.06, z + 10 - i * 2.2);
        g.beginPath(); g.moveTo(a.x, a.y); g.lineTo(b.x, b.y); stroke(g, 0.5, i === 0 ? '#f0a7b2' : '#9cc3e6');
      }
      void s0;
    }
    if (!face) { var lg = iso(mx, my + 0.05, z + 8); circ(g, lg.x, lg.y, 0.9, '#f0a7b2'); }
    // 키보드
    var ky = seat.row === 0 ? y + 0.28 : y + d - 0.42;
    box(g, mx - 0.32, ky, z, 0.64, 0.16, 0.8, '#f7f7fa', { noEdge: true });
    // 서류 더미
    var stack = Math.min(6, Math.ceil(seat.docs / 2));
    var px = seat.row === 0 ? x + 0.12 : x + w - 0.52;
    for (var k = 0; k < stack; k++) box(g, px + (k % 2) * 0.02, y + 0.3, z + k * 1.3, 0.4, 0.3, 1.2, k % 2 ? '#fffdf8' : '#f3ece0', { noEdge: true });
    // 머그컵 + 김
    var cupX = seat.row === 0 ? x + w - 0.3 : x + 0.25;
    var cup = iso(cupX, y + 0.35, z);
    rr(g, cup.x - 1.6, cup.y - 4, 3.2, 4, 1.2, '#fffaf2');
    ell(g, cup.x, cup.y - 4, 1.6, 0.6, '#c99a74');
    if (working) drawSteam(g, cup.x, cup.y - 4.5, t, seat.index * 0.37);
    // 스탠드 (책상 끝)
    if (seat.index % DEPTS[seat.dept].cols === 0) {
      var lb = iso(x + 0.2, seat.row === 0 ? y + 0.2 : y + d - 0.2, z);
      g.beginPath(); g.moveTo(lb.x, lb.y); g.lineTo(lb.x - 1, lb.y - 11); g.lineTo(lb.x + 3, lb.y - 14); stroke(g, 0.8, '#c9b2a6');
      ell(g, lb.x, lb.y, 2.2, 0.8, '#c9b2a6');
      g.beginPath(); g.moveTo(lb.x + 1.5, lb.y - 16); g.lineTo(lb.x + 6.5, lb.y - 13.5); g.lineTo(lb.x + 4.5, lb.y - 11); g.closePath();
      g.fillStyle = '#a9cdb5'; g.fill();
      if (seat.agent || NIGHT) circ(g, lb.x + 4.6, lb.y - 12.4, 0.9, '#fff2c4');
    }
  }

  function drawDeptSign(g, d, t) {
    var p = iso(d.sign.x, d.sign.y, 0);
    g.beginPath(); g.moveTo(p.x, p.y); g.lineTo(p.x, p.y - 22); stroke(g, 1, '#c9b2a6');
    ell(g, p.x, p.y, 3, 1.1, '#c9b2a6');
    var font = (6.5 * TS) + 'px ' + FONT;
    g.font = font;
    var tw = g.measureText(d.name).width, w = tw + 9, h = 6.5 * TS + 4;
    var bx = p.x - w / 2, by = p.y - 22 - h;
    rr(g, bx, by, w, h, h / 2, NIGHT ? 'rgba(20,24,40,.85)' : '#fffaf4');
    g.textBaseline = 'middle';
    if (NIGHT) {
      g.save(); g.shadowColor = d.neon; g.shadowBlur = 8 * S; g.fillStyle = d.neon;
      g.fillText(d.name, bx + 4.5, by + h / 2 + 0.3); g.restore();
      g.fillStyle = '#eafff8';
    } else g.fillStyle = '#5b7a62';
    g.fillText(d.name, bx + 4.5, by + h / 2 + 0.3);
  }

  function drawBossDesk(g, t) {
    var b = BOSS_DESK, x = b.x, y = b.y, w = b.w, d = b.d, h = 13;
    softShadow(g, x + w / 2, y + d / 2 + 0.3, 30, 9, 0.18);
    box(g, x, y, 0, w, d, h, '#c8946c', { top: '#d9a983', left: '#c18c64', right: '#a97852' });
    // 앞면 서랍 + 명패
    var f0 = iso(x + 0.25, y + d, 9), f1 = iso(x + 0.95, y + d, 9);
    g.beginPath(); g.moveTo(f0.x, f0.y); g.lineTo(f1.x, f1.y); stroke(g, 0.5, 'rgba(255,255,255,.35)');
    var np = iso(x + 1.5, y + d + 0.02, 7.5);
    g.save(); g.translate(np.x, np.y); g.transform(1, -0.5, 0, 1, 0, 0);
    rr(g, -6, -2.6, 12, 5.2, 1, '#f3cf86');
    g.fillStyle = '#7a5640'; g.font = '3.6px ' + FONT; g.textAlign = 'center'; g.textBaseline = 'middle'; g.fillText('대표', 0, 0.2); g.textAlign = 'left';
    g.restore();
    var z = h + 0.2;
    // 모니터 (대표 쪽을 향함 → 뒷면이 보임)
    var mon = iso(x + 0.65, y + 0.35, z);
    g.beginPath(); g.moveTo(mon.x, mon.y); g.lineTo(mon.x, mon.y - 3); stroke(g, 1.2, '#b9bdc8');
    box(g, x + 0.2, y + 0.3, z + 3, 0.9, 0.1, 10, '#dfe2ea', { left: '#d5d9e2', right: '#c4c8d3', top: '#eef0f5', noEdge: true });
    var lg = iso(x + 0.65, y + 0.42, z + 8); circ(g, lg.x, lg.y, 0.9, '#9cc3e6');
    // 결재함 (승인한 서류가 쌓임)
    box(g, x + 1.55, y + 0.35, z, 0.55, 0.4, 2.5, '#b58660', { noEdge: true });
    var approved = Math.min(5, state.approvedPile);
    for (var i = 0; i < approved; i++) box(g, x + 1.6, y + 0.4, z + 1 + i * 1.1, 0.45, 0.3, 1, i % 2 ? '#f6d38d' : '#fffdf8', { noEdge: true });
    // 꽃병
    var v = iso(x + 1.35, y + 0.7, z);
    rr(g, v.x - 1.5, v.y - 5, 3, 5, 1.4, '#cfe3f3');
    [[-2, -3.5, '#f7b6c6'], [1.4, -4.8, '#fff1f4'], [2.6, -2, '#f0a7b2']].forEach(function (fl) {
      g.beginPath(); g.moveTo(v.x, v.y - 4.5); g.lineTo(v.x + fl[0], v.y - 4.5 + fl[1]); stroke(g, 0.35, '#7ab38a');
      circ(g, v.x + fl[0], v.y - 4.5 + fl[1], 1.3, fl[2]);
    });
    // 찻잔
    var cup = iso(x + 1.95, y + 0.85, z);
    ell(g, cup.x, cup.y, 2.6, 0.9, '#f6efe7');
    rr(g, cup.x - 1.5, cup.y - 3, 3, 3, 1.2, '#fffaf2');
    drawSteam(g, cup.x, cup.y - 3.5, t, 0.5);
  }

  function drawBossChair(g) {
    var x = BOSS_SEAT.x, y = BOSS_SEAT.y;
    var c = iso(x, y, 0);
    g.beginPath(); g.moveTo(c.x, c.y - 7); g.lineTo(c.x, c.y - 1.5); stroke(g, 1.1, '#8c8494');
    ell(g, c.x, c.y - 0.8, 6, 1.9, 'rgba(120,110,130,.6)');
    box(g, x - 0.42, y - 0.35, 6.5, 0.84, 0.7, 3, '#d58f8d', { noEdge: true });
    box(g, x - 0.45, y - 0.48, 9, 0.9, 0.14, 17, '#cc8584', { left: '#e7a9a5', top: '#eab2ae', noEdge: true });
  }

  function drawCabinet(g, t) {
    // 오른쪽 벽 창문 아래 낮은 수납장
    box(g, 0.75, 0.02, 0, 6.2, 0.55, 15, '#f3e6d8', { left: '#e8d5c2', right: '#d9c3ad', top: '#fbf3ea' });
    for (var i = 1; i < 5; i++) {
      var a = iso(0.75 + i * 1.24, 0.57, 1), b = iso(0.75 + i * 1.24, 0.57, 14);
      g.beginPath(); g.moveTo(a.x, a.y); g.lineTo(b.x, b.y); stroke(g, 0.4, 'rgba(160,120,90,.3)');
      var k = iso(0.75 + i * 1.24 - 0.15, 0.57, 8); circ(g, k.x, k.y, 0.6, '#c9a684');
    }
    // 위에 올린 소품: 책, 작은 화분, 액자, 프린터
    var cols = ['#f0a7b2', '#9cc3e6', '#a6d1b0', '#f3cf86'];
    for (var j = 0; j < 4; j++) box(g, 1.0 + j * 0.13, 0.12, 15, 0.11, 0.38, 5 + (j % 2) * 1.5, cols[j], { noEdge: true });
    var pot = iso(2.3, 0.3, 15);
    rr(g, pot.x - 2, pot.y - 3, 4, 3, 1, '#e7a684');
    var sw = Math.sin(t * 1.3) * 0.6 + state.wind * 1.5 * Math.sin(t * 6);
    circ(g, pot.x - 1.2 + sw * 0.5, pot.y - 4.5, 2, '#8cc39a'); circ(g, pot.x + 1.4 + sw * 0.6, pot.y - 5.2, 2.2, '#a5d4ad'); circ(g, pot.x + sw, pot.y - 7, 1.8, '#7ab38a');
    var fr = iso(3.6, 0.2, 15);
    g.save(); g.translate(fr.x, fr.y); g.transform(1, 0.5, 0, 1, 0, 0);
    rr(g, -3.5, -8, 7, 8, 1, '#d8b694'); rr(g, -2.6, -7, 5.2, 6, 0.6, '#cfe6f5'); circ(g, 0, -3.2, 1.3, '#f3cf86');
    g.restore();
    box(g, 5.5, 0.08, 15, 1.2, 0.45, 5, '#e9e8ee', { left: '#dcdbe3', right: '#cfcdd7', top: '#f6f5f9' });   // 프린터
    box(g, 5.65, 0.15, 20, 0.8, 0.3, 0.6, '#fffdf8', { noEdge: true });
    var led = iso(6.55, 0.53, 17); circ(g, led.x, led.y, 0.5, NIGHT ? '#7ce7c8' : '#6fae8f');
  }

  function drawBookshelf(g) {
    var x = 10.6, y = 0.02, w = 1.3, d = 0.55, h = 50;
    box(g, x, y, 0, w, d, h, '#e2c3a4', { left: '#d9b896', right: '#c9a684', top: '#efd7bf' });
    var cols = ['#f0a7b2', '#9cc3e6', '#a6d1b0', '#f3cf86', '#c3b2ea', '#f6c1a4'];
    srand(5);
    for (var s = 0; s < 3; s++) {
      var z0 = 4 + s * 15;
      poly(g, [iso(x + 0.06, y + d, z0 + 12), iso(x + w - 0.06, y + d, z0 + 12), iso(x + w - 0.06, y + d, z0), iso(x + 0.06, y + d, z0)], '#bb9775');
      var bx = x + 0.1;
      while (bx < x + w - 0.15) {
        var bw = 0.07 + rnd() * 0.06, bh = 7 + rnd() * 4;
        if (rnd() < 0.15) { bx += 0.12; continue; }
        poly(g, [iso(bx, y + d, z0 + bh), iso(bx + bw, y + d, z0 + bh), iso(bx + bw, y + d, z0), iso(bx, y + d, z0)], cols[Math.floor(rnd() * cols.length)]);
        bx += bw + 0.01;
      }
    }
  }

  function drawLounge(g, t) {
    // 소파 (왼쪽 벽에 붙어서 방 쪽을 봄)
    softShadow(g, 1.0, 8.7, 26, 8, 0.2);
    box(g, 0.25, 7.75, 0, 1.05, 1.9, 6, '#eeb7b8', { top: '#f6c6c6', left: '#e3a6a8', right: '#d9999c' });  // 좌석
    box(g, 0.08, 7.75, 0, 0.32, 1.9, 15, '#e3a6a8', { top: '#f0bcbd', left: '#d99a9e', right: '#cf8f93' }); // 등받이
    box(g, 0.25, 7.6, 0, 1.05, 0.25, 10, '#e3a6a8', { top: '#f0bcbd', left: '#d99a9e', right: '#cf8f93' });  // 팔걸이
    box(g, 0.25, 9.6, 0, 1.05, 0.25, 10, '#e3a6a8', { top: '#f0bcbd', left: '#d99a9e', right: '#cf8f93' });
    box(g, 0.45, 8.0, 6, 0.3, 0.55, 5, '#fbe3c8', { noEdge: true });     // 쿠션
    box(g, 0.45, 8.95, 6, 0.3, 0.55, 5, '#cfe3f3', { noEdge: true });
    // 커피 테이블
    var tc = iso(2.4, 8.75, 0);
    softShadow(g, 2.4, 8.75, 14, 5, 0.18);
    g.beginPath(); g.moveTo(tc.x, tc.y - 1); g.lineTo(tc.x, tc.y - 6); stroke(g, 1.2, '#c99a74');
    ell(g, tc.x, tc.y - 6.5, 13, 6, '#c99a74'); ell(g, tc.x, tc.y - 7.3, 13, 6, '#e2b892');
    rr(g, tc.x - 4, tc.y - 10.5, 3.2, 3.2, 1.2, '#fffaf2'); drawSteam(g, tc.x - 2.4, tc.y - 11, t, 0.2);
    ell(g, tc.x + 3, tc.y - 7.8, 3, 1.2, '#f3cf86');
    // 플로어 스탠드
    var fl = iso(0.5, 7.1, 0);
    ell(g, fl.x, fl.y, 4, 1.4, '#b89b88');
    g.beginPath(); g.moveTo(fl.x, fl.y); g.lineTo(fl.x, fl.y - 34); stroke(g, 1, '#b89b88');
    g.beginPath(); g.moveTo(fl.x - 6, fl.y - 32); g.lineTo(fl.x + 6, fl.y - 32); g.lineTo(fl.x + 3.5, fl.y - 41); g.lineTo(fl.x - 3.5, fl.y - 41); g.closePath();
    g.fillStyle = '#fbe7cf'; g.fill();
    // 정수기 + 커피 카트
    box(g, 3.8, 9.3, 0, 0.5, 0.5, 20, '#eceaf0', { top: '#f6f5f9', left: '#e1dfe7', right: '#d4d1db' });
    var wt = iso(4.05, 9.55, 20); rr(g, wt.x - 3, wt.y - 9, 6, 9, 3, 'rgba(160,210,240,.85)');
    box(g, 4.6, 9.2, 0, 1.2, 0.6, 13, '#f4e6d6', { top: '#fbf3ea', left: '#e8d5c2', right: '#d9c3ad' });
    box(g, 4.7, 9.3, 13, 0.4, 0.4, 5, '#9cc3e6', { noEdge: true });
    var mg = iso(5.45, 9.45, 13); rr(g, mg.x - 1.4, mg.y - 3, 2.8, 3, 1, '#f0a7b2');
  }

  function drawBoxes(g) {
    // 새 부서 자리에 쌓인 이삿짐 상자
    var fz = FUTURE_ZONE;
    var bx = fz[2] - 1.35, by = fz[1] + 0.35;
    [[0, 0, 0, 0.9, 0.8, 9], [0.05, 1.0, 0, 0.8, 0.7, 8], [0.1, 0.1, 9, 0.7, 0.6, 7]].forEach(function (b) {
      box(g, bx + b[0], by + b[1], b[2], b[3], b[4], b[5], '#e7c39b', { top: '#f0d3b0', left: '#dcb388', right: '#cfa47a' });
      var a = iso(bx + b[0] + b[3] / 2, by + b[1], b[2] + b[5]), c = iso(bx + b[0] + b[3] / 2, by + b[1] + b[4], b[2] + b[5]);
      g.beginPath(); g.moveTo(a.x, a.y); g.lineTo(c.x, c.y); stroke(g, 0.8, 'rgba(255,255,255,.5)');
    });
  }

  /* ================================================================
   * 6. 캐릭터
   * ================================================================ */
  var LOOKS = [
    { skin: '#f8dac3', hair: '#5a3d32', style: 'short', shirt: '#9cc3e6', pants: '#66708f' },
    { skin: '#f6d2b7', hair: '#8a4b36', style: 'long', shirt: '#f0a7b2', pants: '#6f5f80' },
    { skin: '#eec4a4', hair: '#33293a', style: 'curly', shirt: '#a6d1b0', pants: '#5a6680' },
    { skin: '#f8dcc6', hair: '#c48a52', style: 'bob', shirt: '#f3cf86', pants: '#7d6a5b' },
    { skin: '#e4b393', hair: '#3a2b30', style: 'bun', shirt: '#c3b2ea', pants: '#544e6c' },
    { skin: '#f8dac3', hair: '#aaa3bb', style: 'short', shirt: '#8fd0cb', pants: '#66708f' }
  ];
  var BOSS_LOOK = { skin: '#f8dac3', hair: '#3b2c33', style: 'bob', shirt: '#cf6f6f', pants: '#4e4560', scarf: '#fff1e6' };

  function limb(g, x0, y0, x1, y1, w, col) { g.beginPath(); g.moveTo(x0, y0); g.lineTo(x1, y1); stroke(g, w, col); }

  function hairBack(g, L, x, hy) {
    if (L.style === 'long') ell(g, x, hy + 4, 8.2, 10, L.hair);
    if (L.style === 'bob') ell(g, x, hy + 1.5, 8.6, 8, L.hair);
  }

  function hairFront(g, L, x, hy, front, turn) {
    var c = L.hair;
    if (!front) {
      circ(g, x, hy - 0.5, 7.8, c);
      if (L.style === 'long') rr(g, x - 7.5, hy, 15, 12, 6, c);
      if (L.style === 'bob') rr(g, x - 7.9, hy - 1, 15.8, 8, 5, c);
      if (L.style === 'bun') circ(g, x, hy - 8.5, 3.5, c);
      if (L.style === 'curly') for (var j = 0; j < 6; j++) circ(g, x - 7 + j * 2.8, hy - 5 + Math.abs(j - 2.5) * 0.7, 2.5, c);
      return;
    }
    var o = turn * 1.2;
    g.beginPath();
    g.moveTo(x - 7.8, hy + 1.5);
    g.bezierCurveTo(x - 9, hy - 10.5, x + 9, hy - 10.5, x + 7.8, hy + 1.5);
    g.quadraticCurveTo(x + 5 + o, hy - 4.2, x + 1.5 + o, hy - 3.4);
    g.quadraticCurveTo(x - 3 + o, hy - 5, x - 7.8, hy + 1.5);
    g.fillStyle = c; g.fill();
    if (L.style === 'long' || L.style === 'bob') { rr(g, x - 8.2, hy - 2, 3, L.style === 'long' ? 13 : 8, 1.5, c); rr(g, x + 5.2, hy - 2, 3, L.style === 'long' ? 13 : 8, 1.5, c); }
    if (L.style === 'bun') circ(g, x - turn * 1.5, hy - 9, 3.5, c);
    if (L.style === 'curly') for (var i = 0; i < 5; i++) circ(g, x - 6 + i * 3, hy - 6.4 + Math.abs(i - 2) * 0.9, 2.5, c);
    g.fillStyle = 'rgba(255,255,255,.18)'; g.beginPath(); g.ellipse(x - 2.5, hy - 5.5, 2.5, 1, -0.3, 0, TAU); g.fill();
  }

  /*
   * p = { look, dir(SE|SW|NE|NW), moving, phase, carrying, talking, typing, blink, happy, sitting, seed }
   * (x, y) = 발 아래 중앙 (화면 좌표)
   */
  function drawPerson(g, p, x, y, t) {
    var L = p.look, dir = p.dir || 'SW';
    var front = dir === 'SE' || dir === 'SW';
    var turn = dir === 'SE' || dir === 'NE' ? 1 : -1;
    var sw = p.moving ? Math.sin(p.phase) : 0;
    var bob = p.moving ? -Math.abs(Math.cos(p.phase)) * 1.1 : p.typing ? (Math.sin(t * 10 + p.seed) > 0.3 ? -0.35 : 0) : Math.sin(t * 1.7 + p.seed) * 0.35;
    var sit = p.sitting ? 4 : 0;
    var hy = y - 26 + bob + sit;
    var by = y - 20 + bob + sit;
    var shoe = '#6b5560';

    if (!p.sitting) {
      var sh = g.globalAlpha; g.globalAlpha = sh * 0.9;
      ell(g, x, y, 6.5, 2.2, 'rgba(110,70,60,.18)');
      g.globalAlpha = sh;
    }
    if (front) hairBack(g, L, x, hy);

    // 다리
    if (!p.sitting) {
      var hip = y - 9.5 + bob * 0.4;
      var l1 = Math.max(0, sw) * 2.2, l2 = Math.max(0, -sw) * 2.2;
      var lx = turn * sw * 0.8;
      rr(g, x - 4.2 + lx, hip, 3.6, 8.5 - l1, 1.8, L.pants); rr(g, x + 0.6 - lx, hip, 3.6, 8.5 - l2, 1.8, L.pants);
      ell(g, x - 2.4 + lx + turn * 0.6, y - 0.8 - l1, 2.3, 1.3, shoe); ell(g, x + 2.4 - lx + turn * 0.6, y - 0.8 - l2, 2.3, 1.3, shoe);
    } else {
      rr(g, x - 4.6, by + 9.5, 9.2, 3.5, 1.7, L.pants);
    }

    // 몸통
    rr(g, x - 6, by, 12, 11.5, 4.8, L.shirt);
    g.fillStyle = 'rgba(80,50,60,.07)'; rrPath(g, x - 6, by + 7.5, 12, 4, 2); g.fill();
    if (front) {
      if (L.scarf) rr(g, x - 4.5, by - 0.5, 9, 3, 1.5, L.scarf);
      else { g.beginPath(); g.moveTo(x - 2.2 + turn, by); g.lineTo(x + turn, by + 2.4); g.lineTo(x + 2.2 + turn, by); g.fillStyle = 'rgba(255,255,255,.7)'; g.fill(); }
    }

    // 팔
    var armCol = shadeCol(L.shirt, -0.04);
    if (p.carrying && front) {
      limb(g, x - 5.4, by + 2.5, x - 4, by + 7, 2.8, armCol); limb(g, x + 5.4, by + 2.5, x + 4, by + 7, 2.8, armCol);
    } else if (p.typing) {
      var tp = Math.sin(t * 14 + p.seed) * 0.6;
      if (front) { limb(g, x - 5.4, by + 2.5, x - 3.2, by + 8 + tp, 2.8, armCol); limb(g, x + 5.4, by + 2.5, x + 3.2, by + 8 - tp, 2.8, armCol); }
      else { limb(g, x - 5.4, by + 2.5, x - 6.5, by + 6.5 + tp, 2.8, armCol); limb(g, x + 5.4, by + 2.5, x + 6.5, by + 6.5 - tp, 2.8, armCol); }
    } else {
      limb(g, x - 5.4, by + 2.5, x - 6.6, by + 8.6 + sw * 1.4, 2.8, armCol); limb(g, x + 5.4, by + 2.5, x + 6.6, by + 8.6 - sw * 1.4, 2.8, armCol);
      circ(g, x - 6.6, by + 9.6 + sw * 1.4, 1.5, L.skin); circ(g, x + 6.6, by + 9.6 - sw * 1.4, 1.5, L.skin);
    }
    if (p.carrying && !front) { rr(g, x - 7.5, by + 3, 2, 6, 1, '#f4cf7a'); rr(g, x + 5.5, by + 3, 2, 6, 1, '#f4cf7a'); }

    // 머리
    rr(g, x - 1.5, by - 2.2, 3, 3, 1, L.skin);
    if (front) { circ(g, x - 7.1 + turn * 0.8, hy + 1, 1.4, L.skin); circ(g, x + 7.1 + turn * 0.8, hy + 1, 1.4, L.skin); }
    circ(g, x, hy, 7.5, L.skin);
    hairFront(g, L, x, hy, front, turn);

    if (front) {
      var ex = turn * 1.1, eye = '#4a3438';
      if (p.blink) { limb(g, x - 3.6 + ex, hy + 1, x - 1.8 + ex, hy + 1, 0.55, eye); limb(g, x + 1.8 + ex, hy + 1, x + 3.6 + ex, hy + 1, 0.55, eye); }
      else if (p.happy) {
        g.beginPath(); g.arc(x - 2.7 + ex, hy + 1.4, 1.1, Math.PI, 0); stroke(g, 0.6, eye);
        g.beginPath(); g.arc(x + 2.7 + ex, hy + 1.4, 1.1, Math.PI, 0); stroke(g, 0.6, eye);
      } else {
        ell(g, x - 2.7 + ex, hy + 0.9, 0.95, 1.25, eye); ell(g, x + 2.7 + ex, hy + 0.9, 0.95, 1.25, eye);
        circ(g, x - 2.4 + ex, hy + 0.4, 0.35, '#fff'); circ(g, x + 3 + ex, hy + 0.4, 0.35, '#fff');
      }
      ell(g, x - 4.4 + ex, hy + 3, 1.6, 1, 'rgba(255,140,150,.45)'); ell(g, x + 4.4 + ex, hy + 3, 1.6, 1, 'rgba(255,140,150,.45)');
      if (p.talking && Math.sin(t * 14) > 0) ell(g, x + ex, hy + 3.9, 0.9, 0.7, '#c2646c');
      else { g.beginPath(); g.arc(x + ex, hy + 3.2, 0.9, 0.2, Math.PI - 0.2); stroke(g, 0.5, '#b0606a'); }
    }

    // 결재 파일
    if (p.carrying && front) {
      g.save(); g.translate(x + turn, by + 5.5);
      rr(g, -5.5, -4.5, 11, 9, 1.8, '#e9b85f');
      rr(g, -4.5, -5.5, 9, 3, 1, '#fffdf8');
      rr(g, -5.5, -3.5, 11, 8, 1.8, '#f6d38d');
      circ(g, 0, 0.5, 1.6, '#e8848c');
      g.restore();
      circ(g, x - 4.8, by + 7.4, 1.5, L.skin); circ(g, x + 4.8, by + 7.4, 1.5, L.skin);
    }
  }

  /* ================================================================
   * 6-1. 강아지 직원들
   * ================================================================
   * 에이전트는 셔츠를 입은 강아지로 그립니다. 견종마다 털색·귀·꼬리·무늬가 달라요.
   * checkIn({ breed: 'samoyed' }) 처럼 견종을 정할 수 있고, 안 정하면 id에 따라 자동으로 골라집니다.
   */
  var BREEDS = {
    pomeranian: { name: '포메라니안', fur: '#f2a65a', light: '#fcdcae', ear: 'point', fluffy: true, ruff: true, tail: 'plume' },
    samoyed: { name: '사모예드', fur: '#fbf8f2', light: '#ffffff', edge: '#e6ddd0', ear: 'point', fluffy: true, ruff: true, tail: 'plume' },
    shiba: { name: '시바견', fur: '#e59a5c', light: '#fdf0de', ear: 'point', cheeks: true, tail: 'curl' },
    corgi: { name: '웰시코기', fur: '#eba15d', light: '#fff6ea', ear: 'big', blaze: true, tail: 'stub', shortLegs: true },
    maltese: { name: '말티즈', fur: '#fffcf6', light: '#ffffff', edge: '#ece3d6', ear: 'longflop', tail: 'plume', bow: '#f7a8c2' },
    poodle: { name: '푸들', fur: '#b8794e', light: '#d39c70', ear: 'curlflop', curly: true, tail: 'pom' },
    beagle: { name: '비글', fur: '#cf9150', light: '#fffaf2', earCol: '#7a5038', ear: 'flop', blaze: true, tail: 'straight' },
    husky: { name: '허스키', fur: '#8d94a4', light: '#ffffff', ear: 'point', mask: true, tail: 'plume' },
    golden: { name: '골든 리트리버', fur: '#e8b56a', light: '#f8ddb0', ear: 'flop', tail: 'plume' }
  };
  var BREED_KEYS = Object.keys(BREEDS);
  var SHIRTS = ['#9cc3e6', '#f0a7b2', '#a6d1b0', '#f3cf86', '#c3b2ea', '#8fd0cb', '#f6c1a4'];

  function dogLook(id, breed) {
    var h = hashStr(String(id));
    return {
      breed: BREEDS[breed] ? breed : BREED_KEYS[h % BREED_KEYS.length],
      shirt: SHIRTS[(h >> 4) % SHIRTS.length],
      tie: (h >> 7) % 3 === 0
    };
  }

  function fluffCircle(g, x, y, r, col, bumps, br) {
    circ(g, x, y, r, col);
    for (var i = 0; i < bumps; i++) {
      var a = i / bumps * TAU;
      circ(g, x + Math.cos(a) * (r - 0.6), y + Math.sin(a) * (r - 0.6), br, col);
    }
  }

  function dogTail(g, B, bx, by, wag, side, t) {
    // side: 꼬리가 뻗는 방향 (-1 왼쪽, 1 오른쪽, 0 위쪽)
    g.save(); g.translate(bx, by); g.rotate(wag);
    var c = B.fur;
    if (B.tail === 'stub') {
      circ(g, side * 1.5, -1.5, 2.2, c);
    } else if (B.tail === 'curl') {
      g.beginPath(); g.arc(side * 2.6, -4.2, 2.8, 0, TAU); stroke(g, 2.4, c);
      circ(g, side * 2.6, -4.2, 1, B.light);
    } else if (B.tail === 'pom') {
      g.beginPath(); g.moveTo(0, 0); g.quadraticCurveTo(side * 3, -2, side * 3.5, -6); stroke(g, 1.4, c);
      fluffCircle(g, side * 3.6, -7.2, 2.4, B.light, 6, 1.1);
    } else if (B.tail === 'straight') {
      g.beginPath(); g.moveTo(0, 0); g.quadraticCurveTo(side * 3.5, -3, side * 4.5, -8); stroke(g, 2.2, c);
      circ(g, side * 4.5, -8.2, 1.3, B.light);
    } else {
      // 풍성한 꼬리
      var pts = [[0, 0], [side * 2.6, -2.6], [side * 4.2, -5.6], [side * 4, -8.8], [side * 2.4, -10.6]];
      pts.forEach(function (p, i) { circ(g, p[0], p[1], 2 + Math.sin(i / 4 * Math.PI) * 1.6, c); });
      if (B.edge) pts.forEach(function (p, i) { if (i > 1) circ(g, p[0] - side * 0.6, p[1] - 0.4, 1, 'rgba(255,255,255,.6)'); });
    }
    g.restore();
  }

  function dogEars(g, B, x, hy, front, turn, behind) {
    var c = B.earCol || B.fur, inner = '#f6b3b8';
    if (B.ear === 'point' || B.ear === 'big') {
      if (!behind) return;
      var es = B.ear === 'big' ? 1.35 : 1;
      [-1, 1].forEach(function (s) {
        var ox = turn * 0.6;
        var b1 = { x: x + s * 1.6 + ox, y: hy - 6.6 }, b2 = { x: x + s * 7.4 + ox, y: hy - 2.6 }, tip = { x: x + s * (5.8 + es) + ox, y: hy - 7.5 - 5.5 * es };
        g.beginPath(); g.moveTo(b1.x, b1.y); g.quadraticCurveTo(tip.x - s * 1.2, tip.y + 1, tip.x, tip.y); g.quadraticCurveTo(tip.x + s * 1, tip.y + 2.5, b2.x, b2.y); g.closePath();
        g.fillStyle = c; g.fill();
        if (front) {
          var m = { x: (b1.x + b2.x) / 2, y: (b1.y + b2.y) / 2 };
          g.beginPath(); g.moveTo(m.x - s * 1.6, m.y - 0.8); g.lineTo(tip.x - s * 0.3, tip.y + 2.4); g.lineTo(m.x + s * 1.4, m.y + 0.4); g.closePath();
          g.fillStyle = inner; g.fill();
        }
      });
      return;
    }
    if (behind) return;
    [-1, 1].forEach(function (s) {
      var ex = x + s * 7 + turn * 0.4;
      if (B.ear === 'curlflop') {
        [[0, -1], [0.6, 1.8], [0, 4.4], [0.4, 6.6]].forEach(function (p) { circ(g, ex + s * p[0], hy + p[1], 2.4, c); });
      } else if (B.ear === 'longflop') {
        ell(g, ex, hy + 3, 2.8, 7.2, c, s * 0.18);
        if (B.edge) ell(g, ex + s * 0.8, hy + 4, 1, 5.5, B.edge, s * 0.18);
      } else {
        ell(g, ex, hy + 1.4, 3, 5.6, c, s * 0.28);
      }
    });
  }

  function dogHead(g, B, x, hy, front, turn, p) {
    var R = 8;
    dogEars(g, B, x, hy, front, turn, true);
    if (B.fluffy) { fluffCircle(g, x, hy, R, B.fur, 12, 2.4); if (B.edge) { g.beginPath(); g.arc(x, hy, R + 1.5, Math.PI * 0.15, Math.PI * 0.85); stroke(g, 0.6, B.edge); } }
    else if (B.curly) { fluffCircle(g, x, hy, R - 0.5, B.fur, 11, 2.3); fluffCircle(g, x, hy - 7.5, 3.2, B.light, 6, 1.4); }
    else ell(g, x, hy, R + 0.3, R - 0.3, B.fur);

    if (!front) {
      if (B.mask) ell(g, x, hy + 4.5, 5.5, 3, 'rgba(255,255,255,.6)');
      dogEars(g, B, x, hy, front, turn, false);
      if (B.bow) drawBow(g, x - 3.5, hy - 6.4, B.bow);
      return;
    }

    var ex = turn * 1.1;
    // 무늬
    if (B.mask) {
      ell(g, x + ex * 0.5, hy + 2.6, 6.4, 5.2, B.light);
      circ(g, x - 2.7 + ex, hy - 2.4, 1.05, B.light); circ(g, x + 2.7 + ex, hy - 2.4, 1.05, B.light);
    }
    if (B.cheeks) { ell(g, x - 3.8 + ex, hy + 3, 3.4, 2.7, B.light); ell(g, x + 3.8 + ex, hy + 3, 3.4, 2.7, B.light); }
    if (B.blaze) rr(g, x - 1 + ex, hy - 7, 2, 8, 1, B.light);
    ell(g, x + ex * 0.9, hy + 3.3, 3.8, 2.8, B.light);

    // 눈
    var eye = '#2b2026';
    if (p.blink) {
      limb(g, x - 3.8 + ex, hy + 0.3, x - 1.9 + ex, hy + 0.3, 0.6, eye); limb(g, x + 1.9 + ex, hy + 0.3, x + 3.8 + ex, hy + 0.3, 0.6, eye);
    } else if (p.happy) {
      g.beginPath(); g.arc(x - 2.9 + ex, hy + 0.9, 1.2, Math.PI, 0); stroke(g, 0.7, eye);
      g.beginPath(); g.arc(x + 2.9 + ex, hy + 0.9, 1.2, Math.PI, 0); stroke(g, 0.7, eye);
    } else {
      ell(g, x - 2.9 + ex, hy + 0.2, 1.2, 1.45, eye); ell(g, x + 2.9 + ex, hy + 0.2, 1.2, 1.45, eye);
      circ(g, x - 2.5 + ex, hy - 0.4, 0.45, '#fff'); circ(g, x + 3.3 + ex, hy - 0.4, 0.45, '#fff');
    }
    // 코와 입
    ell(g, x + ex, hy + 2.5, 1.4, 1, '#3a2a2e');
    circ(g, x + ex - 0.4, hy + 2.2, 0.35, 'rgba(255,255,255,.7)');
    if (p.happy || (p.talking && Math.sin(p.t * 14) > 0)) {
      ell(g, x + ex, hy + 4.6, 1.5, 1.3, '#a8505c');
      ell(g, x + ex, hy + 5.2, 1, 0.8, '#ff8fa3');
    } else {
      g.beginPath();
      g.moveTo(x + ex - 1.6, hy + 3.7); g.quadraticCurveTo(x + ex - 0.8, hy + 4.6, x + ex, hy + 3.5);
      g.quadraticCurveTo(x + ex + 0.8, hy + 4.6, x + ex + 1.6, hy + 3.7);
      stroke(g, 0.5, '#7a4a52');
    }
    ell(g, x - 5 + ex, hy + 2.6, 1.6, 1, 'rgba(255,140,150,.45)'); ell(g, x + 5 + ex, hy + 2.6, 1.6, 1, 'rgba(255,140,150,.45)');
    dogEars(g, B, x, hy, front, turn, false);
    if (B.bow) drawBow(g, x - 3.5, hy - 6.4, B.bow);
  }

  function drawBow(g, x, y, col) {
    g.beginPath(); g.moveTo(x, y); g.lineTo(x - 3, y - 2); g.lineTo(x - 3, y + 2); g.closePath(); g.fillStyle = col; g.fill();
    g.beginPath(); g.moveTo(x, y); g.lineTo(x + 3, y - 2); g.lineTo(x + 3, y + 2); g.closePath(); g.fill();
    circ(g, x, y, 1.1, shadeCol(col, -0.15));
  }

  /*
   * 강아지 그리기. p 는 drawPerson 과 같은 형식 (look 은 { breed, shirt, tie })
   */
  function drawDog(g, p, x, y, t) {
    var L = p.look, B = BREEDS[L.breed] || BREEDS.pomeranian;
    var dir = p.dir || 'SW';
    var front = dir === 'SE' || dir === 'SW';
    var turn = dir === 'SE' || dir === 'NE' ? 1 : -1;
    var sw = p.moving ? Math.sin(p.phase) : 0;
    var bob = p.moving ? -Math.abs(Math.cos(p.phase)) * 1.1 : p.typing ? (Math.sin(t * 10 + p.seed) > 0.3 ? -0.35 : 0) : Math.sin(t * 1.7 + p.seed) * 0.35;
    var sit = p.sitting ? 4 : 0;
    var legLen = B.shortLegs ? 6 : 8;
    var by = y - legLen - 11 + bob + sit;
    var hy = by - 6.5;
    var wag = Math.sin(t * (p.happy || p.moving ? 15 : 5) + p.seed) * (p.happy ? 0.55 : p.moving ? 0.35 : 0.2);
    p.t = t;

    if (!p.sitting) ell(g, x, y, 6.5, 2.2, 'rgba(110,70,60,.18)');
    if (front) dogTail(g, B, x - turn * 4.5, by + 8.5, wag, -turn, t);

    // 다리 (털 + 발바닥)
    if (!p.sitting) {
      var hip = by + 10.5;
      var l1 = Math.max(0, sw) * 2, l2 = Math.max(0, -sw) * 2;
      var lx = turn * sw * 0.8;
      rr(g, x - 4.2 + lx, hip - 1, 3.6, legLen + 1 - l1, 1.8, B.fur); rr(g, x + 0.6 - lx, hip - 1, 3.6, legLen + 1 - l2, 1.8, B.fur);
      ell(g, x - 2.4 + lx + turn * 0.5, y - 0.9 - l1, 2.3, 1.3, B.light); ell(g, x + 2.4 - lx + turn * 0.5, y - 0.9 - l2, 2.3, 1.3, B.light);
      if (B.edge) { ell(g, x - 2.4 + lx + turn * 0.5, y - 0.4 - l1, 2.1, 0.5, B.edge); ell(g, x + 2.4 - lx + turn * 0.5, y - 0.4 - l2, 2.1, 0.5, B.edge); }
    } else {
      rr(g, x - 4.6, by + 9.5, 9.2, 3.5, 1.7, B.fur);
    }

    // 셔츠
    rr(g, x - 6, by, 12, 11.5, 4.8, L.shirt);
    g.fillStyle = 'rgba(80,50,60,.07)'; rrPath(g, x - 6, by + 7.5, 12, 4, 2); g.fill();
    if (!front) dogTail(g, B, x, by + 9, wag, turn * 0.6, t);
    if (front) {
      g.beginPath(); g.moveTo(x - 3 + turn, by - 0.2); g.lineTo(x + turn, by + 2.6); g.lineTo(x + 3 + turn, by - 0.2); g.fillStyle = 'rgba(255,255,255,.85)'; g.fill();
      if (L.tie) { g.beginPath(); g.moveTo(x + turn - 0.9, by + 2); g.lineTo(x + turn + 0.9, by + 2); g.lineTo(x + turn + 1.1, by + 6.5); g.lineTo(x + turn, by + 7.6); g.lineTo(x + turn - 1.1, by + 6.5); g.closePath(); g.fillStyle = shadeCol(L.shirt, -0.35); g.fill(); }
    }
    if (B.ruff) for (var r = -2; r <= 2; r++) circ(g, x + r * 2.4 + turn * 0.5, by + 0.6 + Math.abs(r) * 0.3, 2.3, B.fur);

    // 팔 (소매 + 앞발)
    var armCol = shadeCol(L.shirt, -0.04);
    var paw = B.fur;
    if (p.carrying && front) {
      limb(g, x - 5.4, by + 2.5, x - 4, by + 7, 2.8, armCol); limb(g, x + 5.4, by + 2.5, x + 4, by + 7, 2.8, armCol);
    } else if (p.typing) {
      var tp = Math.sin(t * 14 + p.seed) * 0.6;
      if (front) { limb(g, x - 5.4, by + 2.5, x - 3.2, by + 8 + tp, 2.8, armCol); limb(g, x + 5.4, by + 2.5, x + 3.2, by + 8 - tp, 2.8, armCol); circ(g, x - 3.2, by + 8.6 + tp, 1.6, paw); circ(g, x + 3.2, by + 8.6 - tp, 1.6, paw); }
      else { limb(g, x - 5.4, by + 2.5, x - 6.5, by + 6.5 + tp, 2.8, armCol); limb(g, x + 5.4, by + 2.5, x + 6.5, by + 6.5 - tp, 2.8, armCol); }
    } else {
      limb(g, x - 5.4, by + 2.5, x - 6.6, by + 8.4 + sw * 1.4, 2.8, armCol); limb(g, x + 5.4, by + 2.5, x + 6.6, by + 8.4 - sw * 1.4, 2.8, armCol);
      circ(g, x - 6.6, by + 9.5 + sw * 1.4, 1.7, paw); circ(g, x + 6.6, by + 9.5 - sw * 1.4, 1.7, paw);
    }
    if (p.carrying && !front) { rr(g, x - 7.5, by + 3, 2, 6, 1, '#f4cf7a'); rr(g, x + 5.5, by + 3, 2, 6, 1, '#f4cf7a'); }

    dogHead(g, B, x, hy, front, turn, p);

    if (p.carrying && front) {
      g.save(); g.translate(x + turn, by + 5.5);
      rr(g, -5.5, -4.5, 11, 9, 1.8, '#e9b85f');
      rr(g, -4.5, -5.5, 9, 3, 1, '#fffdf8');
      rr(g, -5.5, -3.5, 11, 8, 1.8, '#f6d38d');
      circ(g, 0, 0.5, 1.6, '#e8848c');
      g.restore();
      circ(g, x - 4.8, by + 7.4, 1.7, paw); circ(g, x + 4.8, by + 7.4, 1.7, paw);
    }
  }

  /* 근무 현황 목록에 쓰는 동그란 얼굴 사진 */
  var avatarCache = {};
  function dogAvatar(look) {
    var key = look.breed + look.shirt + (look.tie ? 't' : '');
    if (avatarCache[key]) return avatarCache[key];
    var c = document.createElement('canvas');
    c.width = c.height = 96;
    var g = c.getContext('2d');
    g.scale(4, 4);
    circ(g, 12, 12, 12, '#fdf1e8');
    g.save(); g.beginPath(); g.arc(12, 12, 12, 0, TAU); g.clip();
    drawDog(g, { look: look, dir: 'SE', seed: 0, happy: false }, 12.5, 38, 0);
    g.restore();
    avatarCache[key] = c.toDataURL();
    return avatarCache[key];
  }

  /* ================================================================
   * 7. 말풍선, 이름표
   * ================================================================ */
  function textWidth(text, font) { ctx.font = font; return ctx.measureText(text).width; }

  function drawBubble(x, y, text, opts) {
    opts = opts || {};
    var g = ctx;
    var fs = 6.6 * TS, font = fs + 'px ' + FONT;
    var tw = textWidth(text, font);
    var w = tw + 9 * TS, h = fs + 6 * TS;
    var bx = clamp(x - w / 2, 4, W - w - 4), by = y - h - 4;
    var fill = opts.fill || (NIGHT ? 'rgba(42,38,88,.94)' : 'rgba(255,253,249,.97)');
    var ink = opts.color || (NIGHT ? '#f3eeff' : '#4b3a3f');
    g.save();
    g.shadowColor = NIGHT ? 'rgba(0,0,0,.35)' : 'rgba(170,110,110,.25)'; g.shadowBlur = 6 * S; g.shadowOffsetY = 1.2 * S;
    rrPath(g, bx, by, w, h, h / 2);
    var tx = clamp(x, bx + 6, bx + w - 6);
    g.fillStyle = fill; g.fill();
    g.beginPath(); g.moveTo(tx - 2.5, by + h - 0.5); g.lineTo(tx, by + h + 3.2); g.lineTo(tx + 2.5, by + h - 0.5); g.fill();
    g.restore();
    g.fillStyle = ink; g.font = font; g.textBaseline = 'middle';
    g.fillText(text, bx + 4.5 * TS, by + h / 2 + 0.4);
  }

  function drawName(x, y, text, highlight) {
    var g = ctx, fs = 5.2 * TS, font = fs + 'px ' + FONT;
    var tw = textWidth(text, font), w = tw + 6 * TS, h = fs + 3 * TS;
    var bx = x - w / 2;
    rr(g, bx, y, w, h, h / 2, highlight ? (NIGHT ? 'rgba(255,143,200,.85)' : 'rgba(232,132,140,.9)') : (NIGHT ? 'rgba(30,28,70,.72)' : 'rgba(255,255,255,.85)'));
    g.fillStyle = highlight ? '#fff' : (NIGHT ? '#e9e3ff' : '#6b5559');
    g.font = font; g.textBaseline = 'middle';
    g.fillText(text, bx + 3 * TS, y + h / 2 + 0.3);
    return h;
  }

  function drawProgress(x, y, pct) {
    var g = ctx, w = 20;
    rr(g, x - w / 2, y, w, 2.4, 1.2, NIGHT ? 'rgba(255,255,255,.18)' : 'rgba(120,90,90,.15)');
    rr(g, x - w / 2, y, Math.max(2.4, w * clamp(pct / 100, 0, 1)), 2.4, 1.2, NIGHT ? '#7ce7c8' : '#86c29e');
  }

  function drawFloorText() {
    var g = ctx, fz = FUTURE_ZONE;
    var c = iso(fz[0] + 1.3, fz[1] + 2.3, 0);
    g.font = (5.6 * TS) + 'px ' + FONT; g.textAlign = 'center'; g.textBaseline = 'middle';
    g.fillStyle = NIGHT ? 'rgba(220,210,255,.5)' : 'rgba(150,110,100,.65)';
    g.fillText('새 부서가 들어올 자리', c.x, c.y);
    var b = iso(BOSS_RUG[0] + 0.5, BOSS_RUG[3] - 0.45, 0);
    g.font = (6 * TS) + 'px ' + FONT;
    if (NIGHT) { g.save(); g.shadowColor = '#ff9fd6'; g.shadowBlur = 8 * S; g.fillStyle = '#ff9fd6'; g.fillText('대표실', b.x + 8, b.y); g.restore(); g.fillStyle = '#fff0f8'; }
    else g.fillStyle = '#b06a70';
    g.fillText('대표실', b.x + 8, b.y);
    g.textAlign = 'left';
  }

  /* ================================================================
   * 8. 상태
   * ================================================================ */
  var state = {
    agents: [],
    particles: [],
    prints: [],
    motes: [],
    approvals: [],          // {id, agentId, title, detail, status: 'coming'|'arrived'|'approved'|'rejected', ...}
    approvedPile: 0,
    bossHappyUntil: 0,
    docsToday: 0,
    boss: { look: BOSS_LOOK, dir: 'SW', seed: 1.3 },
    elev: { open: 0, target: 0, holdUntil: 0 },
    wind: 0,
    nextGust: performance.now() + 6000,
    birds: null, nextBirds: performance.now() + 9000,
    shooting: null, nextShooting: performance.now() + 7000,
    listeners: [],
    resolvers: {},
    selected: null,
    bossScreen: { x: 0, y: 0 }
  };

  /* 햇살 속 먼지 */
  for (var mi = 0; mi < 50; mi++) state.motes.push({ u: Math.random(), v: Math.random(), w: Math.floor(Math.random() * 8), ph: Math.random() * TAU, sp: 0.01 + Math.random() * 0.02 });

  /* ================================================================
   * 9. 에이전트 동작 (위치는 월드 좌표)
   * ================================================================ */
  function say(agent, text, ms, delay) {
    var start = performance.now() + (delay || 0);
    agent.bubble = { text: text, from: start, until: start + (ms || 2600) };
  }

  function setPath(agent, points, onArrive) {
    agent.path = points.map(function (p) { return { x: p.x, y: p.y }; });
    agent.onArrive = onArrive || null;
    agent.moving = true;
  }

  function openElevator(ms) {
    state.elev.holdUntil = Math.max(state.elev.holdUntil, performance.now() + (ms || 2400));
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
    var pts = [{ x: seat.x, y: seat.aisleY }];
    if (Math.abs(seat.aisleY - BOSS_Y) > 0.05) { pts.push({ x: CROSS_X, y: seat.aisleY }); pts.push({ x: CROSS_X, y: BOSS_Y }); }
    pts.push({ x: BOSS_SLOTS[slot], y: BOSS_Y });
    return pts;
  }
  function pathBossToSeat(agent, seat) {
    var pts = [{ x: agent.x, y: BOSS_Y }];
    if (Math.abs(seat.aisleY - BOSS_Y) > 0.05) { pts.push({ x: CROSS_X, y: BOSS_Y }); pts.push({ x: CROSS_X, y: seat.aisleY }); }
    pts.push({ x: seat.x, y: seat.aisleY }); pts.push({ x: seat.x, y: seat.y });
    return pts;
  }

  function findAgent(id) {
    for (var i = 0; i < state.agents.length; i++) if (state.agents[i].id === id) return state.agents[i];
    return null;
  }

  function shorten(s, n) { s = String(s || ''); return s.length > n ? s.slice(0, n - 1) + '…' : s; }

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
      log('⚠️', (opts.name || id) + ' — ' + DEPTS[dept].name + '에 빈 자리가 없어요.');
      return null;
    }
    var now = performance.now();
    var agent = {
      id: id,
      name: opts.name || id,
      dept: dept,
      look: opts.look || dogLook(id, opts.breed),
      seed: (hashStr(id) % 100) / 10,
      x: DOOR.x, y: DOOR.y, dir: 'SE',
      phase: 0,
      path: [], moving: false,
      startAt: now + 700,              // 엘리베이터 문이 열린 뒤 걸어 나옴
      state: 'entering',
      seat: seat,
      slot: null,
      task: opts.task || '문서 작업',
      percent: 0,
      docs: 0,
      queue: [],
      bubble: null,
      nextDoc: now + 5000 + Math.random() * 3000,
      nextChat: now + 10000 + Math.random() * 6000,
      focusUntil: 0,
      happyUntil: 0,
      alpha: 0,
      carrying: false,
      approval: null,
      since: new Date()
    };
    seat.agent = agent;
    state.agents.push(agent);
    openElevator(2600);
    say(agent, '출근했습니다~', 3000, 650);
    log('🐾', agent.name + '(' + (BREEDS[agent.look.breed] || {}).name + ') 출근 · ' + DEPTS[dept].name + ' (' + agent.task + ')');
    setPath(agent, pathDoorToSeat(seat), function () {
      agent.state = 'working';
      agent.dir = seat.face;
      say(agent, '업무 시작! ' + shorten(agent.task, 12), 2800);
    });
    renderPanels();
    return agent;
  }

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
      a.percent = clamp(opts.percent, 0, 100);
      if (jump >= 25 && a.state === 'working') focus(a);
    }
    if (typeof opts.docs === 'number' && opts.docs > a.docs) {
      for (var i = a.docs; i < opts.docs; i++) produceDoc(a);
    }
    renderPanels();
  }

  function focus(a) {
    a.focusUntil = performance.now() + 4000;
    say(a, '집중 모드!', 1800);
  }

  function produceDoc(a) {
    a.docs += 1;
    a.seat.docs += 1;
    state.docsToday += 1;
    var sp = iso(a.x, a.y, 0);
    state.particles.push({ kind: 'doc', x: sp.x + 8, y: sp.y - 30, vy: -9, life: 1.8, max: 1.8, text: '+1' });
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
      say(a, '결재 받으러 다녀올게요!', 2200);
      setPath(a, pathSeatToBoss(a.seat, a.slot), function () {
        a.state = 'waiting';
        a.dir = 'NE';
        job.approval.status = 'arrived';
        say(a, '대표님, 결재 부탁드려요!', 4000);
        log('🔔', a.name + ' 도착 — "' + job.approval.title + '" 결재 대기');
        renderPanels();
        notify(job.approval);
        maybeOpenDialog(job.approval);
      });
      renderPanels();
    } else if (job.type === 'checkout') {
      a.state = 'leaving';
      say(a, '퇴근할게요~ 오늘도 수고하셨어요!', 3200);
      log('🌙', a.name + ' 퇴근 (오늘 ' + DEPTS[a.dept].unit + ' ' + a.docs + '건)');
      setPath(a, pathSeatToDoor(a.seat), function () {
        a.state = 'gone';
        openElevator(1400);
      });
      renderPanels();
    }
  }

  function decide(approvalId, decision, comment) {
    var ap = null;
    state.approvals.forEach(function (x) { if (x.id === approvalId) ap = x; });
    if (!ap || ap.status !== 'arrived') return false;
    ap.status = decision === 'approved' ? 'approved' : 'rejected';
    ap.comment = comment || '';
    ap.decidedAt = new Date();
    var a = findAgent(ap.agentId);
    if (a) {
      a.approval = null;
      if (ap.status === 'approved') {
        a.carrying = false;
        a.happyUntil = performance.now() + 2600;
        state.approvedPile += 1;
        state.bossHappyUntil = performance.now() + 2200;
        say(a, '멍! 감사합니다, 바로 할게요', 2600);
        var ps = iso(a.x, a.y, 0);
        state.particles.push({ kind: 'stamp', x: ps.x, y: ps.y - 46, vy: -5, life: 1.5, max: 1.5, text: '승인' });
        for (var h = 0; h < 5; h++) state.particles.push({ kind: 'heart', x: ps.x - 6 + Math.random() * 12, y: ps.y - 30, vx: (Math.random() - 0.5) * 8, vy: -10 - Math.random() * 8, life: 1.6, max: 1.6 });
      } else {
        say(a, '넵, 다듬어서 다시 올릴게요', 2800);
        var pr = iso(a.x, a.y, 0);
        state.particles.push({ kind: 'stamp', x: pr.x, y: pr.y - 46, vy: -5, life: 1.5, max: 1.5, text: '반려', bad: true });
      }
      a.state = 'returning';
      setTimeout(function () {
        a.dir = 'NW';
        setPath(a, pathBossToSeat(a, a.seat), function () {
          a.state = 'working';
          a.carrying = false;
          a.dir = a.seat.face;
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
   * 10. 매 프레임 업데이트
   * ================================================================ */
  function gust(now) {
    state.wind = Math.max(state.wind, NIGHT ? 0.45 : 1);
    if (!NIGHT) {
      for (var i = 0; i < 9; i++) {
        var w = WINDOWS[Math.floor(Math.random() * WINDOWS.length)];
        var wp = wallPt(w.wall, w.h0 + Math.random() * (w.h1 - w.h0), WIN_TOP + 15 + Math.random() * 30);
        var dirx = w.wall === 'right' ? -1 : 1;
        state.particles.push({
          kind: 'petal', x: wp.x, y: wp.y,
          vx: dirx * (14 + Math.random() * 18), vy: 8 + Math.random() * 10, rot: Math.random() * TAU, vr: (Math.random() - 0.5) * 6,
          life: 5 + Math.random() * 3, max: 8, col: ['#f7b6c6', '#ffd9e2', '#fbe7a5', '#cfe8c9'][i % 4]
        });
      }
    }
    state.nextGust = now + (NIGHT ? 26000 : 11000) + Math.random() * 14000;
  }

  function update(dt, now) {
    var t = now / 1000;
    // 바람
    state.wind *= Math.exp(-dt / 2.4);
    if (now > state.nextGust) gust(now);
    // 새 / 별똥별
    if (!NIGHT) {
      if (!state.birds && now > state.nextBirds) state.birds = { u: -20, v: 19 + Math.random() * 12, n: 3 + Math.floor(Math.random() * 3) };
      if (state.birds) { state.birds.u += 22 * dt; if (state.birds.u > 500) { state.birds = null; state.nextBirds = now + 25000 + Math.random() * 25000; } }
    } else {
      if (!state.shooting && now > state.nextShooting) {
        var w = WINDOWS[Math.floor(Math.random() * WINDOWS.length)];
        state.shooting = { wall: w.wall, h: w.h0 + 4 + Math.random() * 20, v: WIN_TOP + 1, vh: 55, vv: 24, life: 0.9, max: 0.9 };
      }
      if (state.shooting) {
        var s = state.shooting; s.h += s.vh * dt; s.v += s.vv * dt; s.life -= dt;
        if (s.life <= 0) { state.shooting = null; state.nextShooting = now + 14000 + Math.random() * 20000; }
      }
    }
    // 엘리베이터
    state.elev.target = now < state.elev.holdUntil ? 1 : 0;
    state.elev.open += (state.elev.target - state.elev.open) * Math.min(1, dt * 4.5);

    state.agents.forEach(function (a) {
      if (now < a.startAt) { a.alpha = 0; return; }
      if (a.state === 'gone') a.alpha -= dt * 2.5; else a.alpha = Math.min(1, a.alpha + dt * 2.5);

      if (a.moving && a.path.length) {
        var tp = a.path[0];
        var dx = tp.x - a.x, dy = tp.y - a.y;
        var dist = Math.sqrt(dx * dx + dy * dy);
        var step = SPEED * dt;
        if (Math.abs(dx) > Math.abs(dy)) a.dir = dx < 0 ? 'NW' : 'SE';
        else if (dist > 0.01) a.dir = dy < 0 ? 'NE' : 'SW';
        if (a.state === 'leaving' && a.path.length <= 2) openElevator(1800);
        if (dist <= step) {
          a.x = tp.x; a.y = tp.y; a.path.shift();
          if (!a.path.length) {
            a.moving = false;
            var cb = a.onArrive; a.onArrive = null;
            if (cb) cb();
            renderPanels();
          }
        } else {
          a.x += dx / dist * step; a.y += dy / dist * step;
        }
        a.phase += dt * 9.5;
        a.printAcc = (a.printAcc || 0) + step;
        if (a.printAcc > 0.42 && a.alpha > 0.5) {
          a.printAcc = 0; a.printSide = -(a.printSide || 1);
          var nx = dx / (dist || 1), ny = dy / (dist || 1);
          state.prints.push({ x: a.x - ny * 0.12 * a.printSide, y: a.y + nx * 0.12 * a.printSide, life: 2.6, max: 2.6 });
          if (state.prints.length > 80) state.prints.shift();
        }
      } else if (a.state === 'working') {
        if (now > a.nextDoc) {
          a.nextDoc = now + 6000 + Math.random() * 6000;
          produceDoc(a);
          if (a.percent < 95) a.percent = Math.min(95, a.percent + 4 + Math.round(Math.random() * 6));
          renderPanels();
        }
        if (now > a.nextChat) {
          a.nextChat = now + 13000 + Math.random() * 12000;
          var lines = DEPTS[a.dept].lines;
          if (Math.random() < 0.25) focus(a); else say(a, lines[Math.floor(Math.random() * lines.length)], 2400);
        }
        processQueue(a);
      }

      var bs = hashStr(a.id) % 2900;
      a.blink = ((now + bs) % 3400) < 130;

      if (a.focusUntil > now && Math.random() < dt * 9) {
        state.particles.push({
          kind: 'spark', x: iso(a.x, a.y).x - 10 + Math.random() * 20, y: iso(a.x, a.y).y - 12 - Math.random() * 18,
          vy: -8 - Math.random() * 8, life: 1.2, max: 1.2, size: 1 + Math.random() * 1.3
        });
      }
    });

    state.agents = state.agents.filter(function (a) {
      if (a.state === 'gone' && a.alpha <= 0) {
        if (a.seat.agent === a) a.seat.agent = null;
        if (state.selected === a.id) state.selected = null;
        renderPanels();
        return false;
      }
      return true;
    });

    state.particles.forEach(function (p) {
      p.life -= dt;
      p.y += p.vy * dt;
      if (p.vx) p.x += p.vx * dt;
      if (p.kind === 'petal') { p.vy += Math.sin(t * 3 + p.rot) * 6 * dt; p.vx *= 0.995; p.rot += p.vr * dt; }
    });
    state.particles = state.particles.filter(function (p) { return p.life > 0; });

    state.boss.blink = (now % 4100) < 130;
    state.boss.typing = !state.approvals.some(function (x) { return x.status === 'arrived'; });
  }


  /* ================================================================
   * 11. 빛 (낮: 햇살과 먼지 / 밤: 어둠 + 조명 + 네온)
   * ================================================================ */
  /* 창문 칸마다 바닥에 떨어지는 햇빛 (월드 좌표 사각형) */
  var SUN = [];
  WINDOWS.forEach(function (w) {
    var mid = (w.h0 + w.h1) / 2;
    [[w.h0 + 2, mid - 1], [mid + 1, w.h1 - 2]].forEach(function (p) {
      var a = p[0] / HX, b = p[1] / HX;
      if (w.wall === 'right') SUN.push({ wall: 'right', p: [[a + 0.3, 0.35], [b + 0.3, 0.35], [b + 1.9, 4.6], [a + 1.9, 4.6]], hv: p });
      else SUN.push({ wall: 'left', p: [[0.35, RY - b + 0.3], [0.35, RY - a + 0.3], [4.2, RY - a + 1.6], [4.2, RY - b + 1.6]], hv: p, dim: true });
    });
  });

  function drawDaylight(t) {
    var g = ctx;
    g.save();
    g.globalCompositeOperation = 'screen';
    var breathe = 0.85 + 0.15 * Math.sin(t * 0.35);
    SUN.forEach(function (s, i) {
      var curtain = 1 - state.wind * 0.25 * (0.5 + 0.5 * Math.sin(t * 5 + i));
      var a = (s.dim ? 0.3 : 0.5) * breathe * curtain;
      var q = s.p.map(function (p) { return iso(p[0], p[1], 0); });
      // 바닥의 햇빛
      var near = { x: (q[0].x + q[1].x) / 2, y: (q[0].y + q[1].y) / 2 }, far = { x: (q[2].x + q[3].x) / 2, y: (q[2].y + q[3].y) / 2 };
      g.fillStyle = lin(g, near.x, near.y, far.x, far.y, [[0, 'rgba(255,214,160,' + a.toFixed(3) + ')'], [0.7, 'rgba(255,220,170,' + (a * 0.35).toFixed(3) + ')'], [1, 'rgba(255,226,180,0)']]);
      poly(g, q); g.fill();
      // 창문에서 바닥까지 공기 중의 빛줄기
      var t0 = wallPt(s.wall, s.hv[0], WIN_TOP + 4), t1 = wallPt(s.wall, s.hv[1], WIN_TOP + 4);
      g.fillStyle = lin(g, (t0.x + t1.x) / 2, t0.y, far.x, far.y, [[0, 'rgba(255,240,210,0)'], [0.35, 'rgba(255,240,210,' + (a * 0.3).toFixed(3) + ')'], [1, 'rgba(255,240,210,0)']]);
      poly(g, [t0, t1, q[2], q[3]]); g.fill();
    });
    // 먼지
    state.motes.forEach(function (m) {
      m.v = (m.v + m.sp * 0.016 + state.wind * 0.004) % 1;
      var s = SUN[m.w % SUN.length];
      var t0 = wallPt(s.wall, s.hv[0] + (s.hv[1] - s.hv[0]) * m.u, WIN_TOP + 8);
      var f0 = iso(s.p[3][0] + (s.p[2][0] - s.p[3][0]) * m.u, s.p[3][1] + (s.p[2][1] - s.p[3][1]) * m.u, 0);
      var xx = t0.x + (f0.x - t0.x) * m.v + Math.sin(t * 0.8 + m.ph) * 2;
      var yy = t0.y + (f0.y - t0.y) * m.v;
      var a = 0.6 * Math.max(0, Math.sin(t * 1.2 + m.ph)) * (1 - Math.abs(m.v - 0.5) * 1.4);
      if (a > 0.02) circ(g, xx, yy, 0.5, 'rgba(255,248,225,' + a.toFixed(2) + ')');
    });
    g.restore();
  }

  function lightsList(now) {
    var L = [];
    Object.keys(DEPTS).forEach(function (k) {
      DEPTS[k].seats.forEach(function (seat) {
        var a = seat.agent;
        if (!a || a.state === 'gone') return;
        var dk = seat.desk;
        var c = iso(dk.x + dk.w / 2, dk.y + dk.d / 2, DESK_H);
        L.push({ x: c.x, y: c.y, r: 24, a: 0.7, col: '255,200,130', ga: 0.09 });
        if (a.state === 'working' && !a.moving) {
          var m = iso(seat.x, seat.row === 0 ? dk.y + dk.d : dk.y, DESK_H + 10);
          L.push({ x: m.x, y: m.y, r: 16, a: 0.6, col: '150,215,255', ga: 0.08 });
        }
      });
      var d = DEPTS[k], sp = iso(d.sign.x, d.sign.y, 26);
      L.push({ x: sp.x, y: sp.y, r: 18, a: 0.45, col: '124,231,200', ga: 0.14 });
    });
    var bd = iso(BOSS_DESK.x + 0.7, BOSS_DESK.y + 0.4, 20);
    L.push({ x: bd.x, y: bd.y, r: 24, a: 0.8, col: '150,215,255', ga: 0.15 });
    var bs = iso(BOSS_SEAT.x, BOSS_SEAT.y, 18);
    L.push({ x: bs.x, y: bs.y, r: 36, a: 0.6, col: '255,200,150', ga: 0.12 });
    var fl = iso(0.5, 7.1, 36);
    L.push({ x: fl.x, y: fl.y + 14, r: 46, a: 0.9, col: '255,205,140', ga: 0.22 });
    var nm = wallPt('right', PILLAR_H, 46);
    L.push({ x: nm.x, y: nm.y, r: 34, a: 0.6, col: '255,150,210', ga: 0.16 });
    var pr = iso(6.55, 0.53, 17);
    L.push({ x: pr.x, y: pr.y, r: 8, a: 0.4, col: '124,231,200', ga: 0.12 });
    var cf = iso(5.2, 9.45, 16);
    L.push({ x: cf.x, y: cf.y, r: 14, a: 0.5, col: '170,215,255', ga: 0.1 });
    var e = state.elev.open;
    if (e > 0.05) { var ep = iso(0.6, (ELEV_Y0 + ELEV_Y1) / 2, 4); L.push({ x: ep.x, y: ep.y, r: 40 * e, a: 0.85 * e, col: '255,225,170', ga: 0.2 * e }); }
    var b = iso(BOSS_RUG[0] + 0.9, BOSS_RUG[3] - 0.45, 2);
    L.push({ x: b.x, y: b.y, r: 18, a: 0.4, col: '255,159,214', ga: 0.12 });
    state.agents.forEach(function (a) {
      if (a.focusUntil > now) { var p = iso(a.x, a.y, 18); L.push({ x: p.x, y: p.y, r: 20, a: 0.45, col: '255,220,140', ga: 0.12 }); }
    });
    return L;
  }
  var ELEV_Y0 = RY - ELEV.h1 / HX, ELEV_Y1 = RY - ELEV.h0 / HX;

  function drawNightLight(now, t) {
    var s = shade.getContext('2d');
    s.setTransform(S, 0, 0, S, 0, 0);
    s.globalCompositeOperation = 'source-over';
    s.clearRect(0, 0, W, H);
    // 방 안만 어둡게 (방 모양: 바닥 + 벽)
    s.fillStyle = 'rgba(16,13,46,.55)';
    poly(s, [iso(-WT, RY, -12), iso(-WT, RY, WH), iso(-WT, -WT, WH), iso(RX, -WT, WH), iso(RX, -WT, -12), iso(RX, RY, -12)]); s.fill();
    s.globalCompositeOperation = 'destination-out';
    // 창문은 밝게
    WINDOWS.forEach(function (w) {
      s.save(); wallTransform(s, w.wall);
      s.fillStyle = 'rgba(0,0,0,.9)'; s.fillRect(w.h0, WIN_TOP, w.h1 - w.h0, WIN_BOTTOM - WIN_TOP);
      s.restore();
    });
    // 창밖 도시 불빛이 바닥에 살짝
    SUN.forEach(function (sn) {
      var q = sn.p.map(function (p) { return iso(p[0] * 0.8, p[1] * 0.8, 0); });
      s.fillStyle = 'rgba(0,0,0,.12)'; poly(s, q); s.fill();
    });
    var lights = lightsList(now);
    lights.forEach(function (l) {
      s.fillStyle = rad(s, l.x, l.y, l.r, [[0, 'rgba(0,0,0,' + l.a + ')'], [1, 'rgba(0,0,0,0)']]);
      s.beginPath(); s.arc(l.x, l.y, l.r, 0, TAU); s.fill();
    });
    ctx.drawImage(shade, 0, 0, W, H);

    ctx.save();
    ctx.globalCompositeOperation = 'lighter';
    lights.forEach(function (l, i) {
      var fl = 0.92 + 0.08 * Math.sin(t * 1.7 + i);
      ctx.fillStyle = rad(ctx, l.x, l.y, l.r, [[0, 'rgba(' + l.col + ',' + (l.ga * fl).toFixed(3) + ')'], [1, 'rgba(' + l.col + ',0)']]);
      ctx.beginPath(); ctx.arc(l.x, l.y, l.r, 0, TAU); ctx.fill();
    });
    // 천장 네온 띠가 벽을 물들임
    ['left', 'right'].forEach(function (wl) {
      ctx.save(); wallTransform(ctx, wl);
      ctx.fillStyle = lin(ctx, 0, 0, 0, 34, [[0, 'rgba(190,140,255,.22)'], [1, 'rgba(190,140,255,0)']]);
      ctx.fillRect(0, 0, WALLS[wl].len, 34);
      ctx.restore();
    });
    ctx.restore();
  }

  /* ================================================================
   * 12. 그리기
   * ================================================================ */
  function render(now, dt) {
    var t = now / 1000;
    var g = ctx;
    g.setTransform(S, 0, 0, S, 0, 0);
    g.imageSmoothingEnabled = true;
    g.globalCompositeOperation = 'source-over';
    g.drawImage(backdrop, 0, 0, W, H);
    drawOuterClouds(t, dt);
    g.drawImage(room, 0, 0, W, H);

    drawWindows(t, dt);
    drawCurtains(t);
    drawElevator();
    drawClock();
    drawNeon(t, false);
    drawFloorText();
    // 발자국
    state.prints.forEach(function (pr) {
      pr.life -= dt;
      if (pr.life <= 0) return;
      var c = iso(pr.x, pr.y, 0), a = Math.min(1, pr.life / pr.max * 1.6) * (NIGHT ? 0.25 : 0.32);
      var col = 'rgba(150,100,80,' + a.toFixed(3) + ')';
      ell(g, c.x, c.y + 0.4, 1.5, 1, col);
      circ(g, c.x - 1.5, c.y - 1, 0.55, col); circ(g, c.x, c.y - 1.5, 0.55, col); circ(g, c.x + 1.5, c.y - 1, 0.55, col);
    });
    state.prints = state.prints.filter(function (pr) { return pr.life > 0; });

    // 깊이 정렬 (월드 x + y 가 작을수록 뒤쪽)
    var list = [];
    function add(key, fn) { list.push({ k: key, draw: fn }); }
    add(0.6, function () { drawCabinet(g, t); });
    add(11.2, function () { drawBookshelf(g); });
    Object.keys(DEPTS).forEach(function (k) {
      var d = DEPTS[k];
      add(d.sign.x + d.sign.y, function () { drawDeptSign(g, d, t); });
      d.seats.forEach(function (seat) {
        var dk = seat.desk;
        var sa = seat.agent;
        var seated = !!(sa && !sa.moving && sa.state === 'working');
        var chairKey = seat.x + (seat.row === 0 ? seat.y - 0.3 : seat.y + 0.3);
        add(chairKey, function () { drawChair(g, seat, seated); });
        add(dk.x + dk.w / 2 + dk.y + dk.d / 2, function () { drawDesk(g, seat, t); });
      });
    });
    PLANTS.forEach(function (p) { add(p.x + p.y, function () { drawPlant(g, p, t); }); });
    add(0.7 + 8.7, function () { drawLounge(g, t); });
    add(FUTURE_ZONE[0] + FUTURE_ZONE[1] + 3, function () { drawBoxes(g); });
    add(BOSS_SEAT.x + BOSS_SEAT.y - 0.45, function () { drawBossChair(g); });
    add(BOSS_SEAT.x + BOSS_SEAT.y, function () {
      var waiting = state.approvals.some(function (x) { return x.status === 'arrived'; });
      var p = iso(BOSS_SEAT.x, BOSS_SEAT.y, 0);
      state.bossScreen = p;
      drawPerson(g, { look: BOSS_LOOK, dir: 'SW', typing: !waiting, blink: state.boss.blink, seed: 1.3, happy: state.bossHappyUntil > now, sitting: true }, p.x, p.y, t);
    });
    add(BOSS_DESK.x + BOSS_DESK.w / 2 + BOSS_DESK.y + BOSS_DESK.d / 2, function () { drawBossDesk(g, t); });

    state.agents.forEach(function (a) {
      if (a.alpha <= 0.01) return;
      var sitting = a.state === 'working' && !a.moving;
      add(a.x + a.y + 0.01, function () {
        var p = iso(a.x, a.y, 0);
        a.sx = p.x; a.sy = p.y;
        g.globalAlpha = clamp(a.alpha, 0, 1);
        if (state.selected === a.id) {
          g.save(); g.translate(p.x, p.y); g.scale(1, 0.5);
          g.beginPath(); g.arc(0, 0, 10, 0, TAU); stroke(g, 1.2, NIGHT ? 'rgba(255,143,200,.8)' : 'rgba(232,132,140,.8)');
          g.restore();
        }
        (a.look.breed ? drawDog : drawPerson)(g, {
          look: a.look, dir: a.dir, moving: a.moving, phase: a.phase, seed: a.seed,
          carrying: a.carrying, typing: sitting, sitting: sitting,
          talking: !!(a.bubble && a.bubble.from < now && a.bubble.until > now), blink: a.blink, happy: a.happyUntil > now
        }, p.x, p.y, t);
        g.globalAlpha = 1;
      });
    });
    list.sort(function (a, b) { return a.k - b.k; });
    list.forEach(function (item) { item.draw(); });

    HANGING.forEach(function (h) { drawHanging(g, h, t); });

    if (NIGHT) { drawNightLight(now, t); drawNeon(t, true); }
    else drawDaylight(t);

    // 이펙트
    state.particles.forEach(function (p) {
      var k = p.life / p.max;
      g.globalAlpha = clamp(k * 1.8, 0, 1);
      if (p.kind === 'spark') {
        var sz = p.size * (0.6 + 0.4 * Math.sin(t * 12 + p.x));
        g.save(); g.translate(p.x, p.y);
        g.beginPath(); g.moveTo(0, -sz * 2); g.quadraticCurveTo(0, 0, sz * 2, 0); g.quadraticCurveTo(0, 0, 0, sz * 2); g.quadraticCurveTo(0, 0, -sz * 2, 0); g.quadraticCurveTo(0, 0, 0, -sz * 2);
        g.fillStyle = NIGHT ? '#fff2b0' : '#ffd36b'; g.fill(); g.restore();
      } else if (p.kind === 'heart') {
        g.save(); g.translate(p.x, p.y); g.scale(0.9, 0.9);
        g.beginPath(); g.moveTo(0, 1.8); g.bezierCurveTo(-3.2, -0.6, -1.8, -3.4, 0, -1.6); g.bezierCurveTo(1.8, -3.4, 3.2, -0.6, 0, 1.8);
        g.fillStyle = '#ff9fb4'; g.fill(); g.restore();
      } else if (p.kind === 'petal') {
        g.save(); g.translate(p.x, p.y); g.rotate(p.rot);
        ell(g, 0, 0, 1.8, 1, p.col); g.restore();
      } else if (p.kind === 'doc') {
        rr(g, p.x - 3, p.y - 4, 7, 8.5, 1.4, NIGHT ? '#f3eeff' : '#fffdf8');
        g.fillStyle = '#9cc3e6'; g.fillRect(p.x - 1.5, p.y - 1.8, 4, 0.7); g.fillRect(p.x - 1.5, p.y + 0.2, 4, 0.7); g.fillRect(p.x - 1.5, p.y + 2.2, 2.5, 0.7);
        g.font = (6.5 * TS) + 'px ' + FONT; g.textBaseline = 'middle';
        g.fillStyle = NIGHT ? '#9ff3d6' : '#6fae8f'; g.fillText(p.text, p.x + 6, p.y + 0.5);
      } else if (p.kind === 'stamp') {
        var col = p.bad ? '#6f8fd6' : '#e0605a';
        g.save(); g.translate(p.x, p.y); g.rotate(-0.2);
        g.beginPath(); g.arc(0, 0, 8.5, 0, TAU); g.fillStyle = 'rgba(255,250,246,.92)'; g.fill(); stroke(g, 1.3, col);
        g.font = (6.4 * TS) + 'px ' + FONT; g.textBaseline = 'middle'; g.textAlign = 'center';
        g.fillStyle = col; g.fillText(p.text, 0, 0.5); g.textAlign = 'left';
        g.restore();
      }
      g.globalAlpha = 1;
    });

    // 이름표, 진행률, 말풍선
    state.agents.forEach(function (a) {
      if (a.alpha <= 0.05) return;
      var p = iso(a.x, a.y, 0);
      g.globalAlpha = clamp(a.alpha, 0, 1);
      var sitting = a.state === 'working' && !a.moving;
      var headTop = p.y - 36 + (sitting ? 4 : 0);
      drawName(p.x, headTop - 9 * TS, a.name, a.state === 'waiting');
      var top = headTop - 9 * TS;
      if (sitting) { drawProgress(p.x, top - 4, a.percent); top -= 4; }
      if (a.state === 'waiting') {
        var bounce = Math.abs(Math.sin(t * 3)) * 2;
        circ(g, p.x + 10, headTop + 6 - bounce, 4, '#ff9f8f');
        g.fillStyle = '#fff'; g.font = '6px ' + FONT; g.textAlign = 'center'; g.textBaseline = 'middle';
        g.fillText('!', p.x + 10, headTop + 6.4 - bounce); g.textAlign = 'left';
      }
      if (a.bubble && a.bubble.from < now && a.bubble.until > now) drawBubble(p.x, top, a.bubble.text, a.state === 'waiting' ? { fill: NIGHT ? 'rgba(90,40,80,.95)' : '#fff3ec' } : null);
      g.globalAlpha = 1;
    });
    var waitingCount = state.approvals.filter(function (x) { return x.status === 'arrived'; }).length;
    var bp = iso(BOSS_SEAT.x, BOSS_SEAT.y, 0);
    drawName(bp.x, bp.y - 42, '나 (대표)', true);
    if (waitingCount) drawBubble(bp.x, bp.y - 43, '결재 ' + waitingCount + '건, 볼게요', { fill: NIGHT ? 'rgba(90,40,80,.95)' : '#fff3ec' });

    // 테마 전환 크로스페이드
    if (fadeAlpha > 0) {
      g.globalAlpha = fadeAlpha;
      g.drawImage(fade, 0, 0, W, H);
      g.globalAlpha = 1;
      fadeAlpha = Math.max(0, fadeAlpha - dt * 0.9);
    }
  }

  var last = performance.now();
  function loop(now) {
    var dt = Math.min(0.05, (now - last) / 1000);
    last = now;
    update(dt, now);
    render(now, dt);
    requestAnimationFrame(loop);
  }

  function setTheme(next) {
    next = next === 'night' ? 'night' : 'day';
    if (next === theme) return;
    if (canvas.width) {
      var f = fade.getContext('2d');
      f.setTransform(1, 0, 0, 1, 0, 0);
      f.clearRect(0, 0, fade.width, fade.height);
      f.drawImage(canvas, 0, 0);
      fadeAlpha = 1;
    }
    theme = next; NIGHT = theme === 'night';
    state.wind = 0; state.birds = null; state.shooting = null;
    CITY.stars = [];
    buildStatic();
  }
  window.addEventListener('themechange', function (e) { setTheme(e.detail && e.detail.theme); });

  /* ================================================================
   * 13. 화면 패널 (결재함, 근무 현황, 업무 일지, 결재 창)
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
          (a.look.breed ? '<img class="avatar" alt="" src="' + dogAvatar(a.look) + '">' : '<span class="avatar"></span>') +
          '<span class="ro-body"><strong>' + esc(a.name) + ' <small>' + esc((BREEDS[a.look.breed] || {}).name || '') + '</small></strong><span>' + esc(DEPTS[a.dept].name) + ' · ' + esc(shorten(a.task, 18)) + '</span></span>' +
          '<span class="chip ' + st[1] + '">' + st[0] + '</span>' +
          '<span class="ro-stat">' + a.docs + '건 · ' + a.percent + '%</span></li>';
      }).join('');
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
      if (a.sx != null && mx > a.sx - 9 && mx < a.sx + 9 && my > a.sy - 36 && my < a.sy + 2) hit = a;
    });
    if (hit) {
      if (hit.state === 'waiting' && hit.approval) { openDialog(hit.approval); return; }
      state.selected = hit.id;
      say(hit, hit.state === 'working' ? shorten(hit.task, 14) + ' (' + hit.percent + '%)' : (STATE_LABEL[hit.state] || [''])[0], 2000);
    } else if (Math.abs(mx - state.bossScreen.x) < 12 && my > state.bossScreen.y - 34 && my < state.bossScreen.y) {
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
    var over = state.agents.some(function (a) { return a.sx != null && mx > a.sx - 9 && mx < a.sx + 9 && my > a.sy - 36 && my < a.sy + 2; });
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
   * 14. 외부 연결 (이벤트 피드 / 결정 전송)
   * ================================================================ */
  var CONFIG = window.OFFICE_CONFIG || {};
  var seenEvents = {};
  var feedStart = 0;
  var pendingTimed = [];

  function dispatch(ev) {
    if (!ev || typeof ev !== 'object') return;
    var type = ev.type;
    var id = ev.agentId || ev.agent || ev.id;
    if (type === 'checkin') return checkIn({ id: ev.agentId || ev.agent, name: ev.name, dept: ev.dept, task: ev.task, breed: ev.breed });
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
   * 15. 데모
   * ================================================================ */
  var DEMO_NAMES = ['포미', '사모', '시바', '코기', '말랑', '푸딩', '비비', '허니', '골디'];
  var DEMO_BREEDS = ['pomeranian', 'samoyed', 'shiba', 'corgi', 'maltese', 'poodle', 'beagle', 'husky', 'golden'];
  var DEMO_TASKS = ['주간 회의록 정리', '프로젝트 기획서 초안', '블로그 글 교정', '보고서 요약본 작성', '매뉴얼 목차 구성', '독서 노트 정리', '회의 안건 정리', '자료 목록 만들기', '메일 초안 쓰기'];
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
    return checkIn({ id: id, name: DEMO_NAMES[i], breed: DEMO_BREEDS[i], dept: 'records', task: DEMO_TASKS[i] });
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
      if (k === 'gust') gust(performance.now());
      if (k === 'notify' && 'Notification' in window) Notification.requestPermission().then(function (p) { toast(p === 'granted' ? '알림을 켰어요.' : '알림이 허용되지 않았어요.'); });
    });
  });


  /* ================================================================
   * 16. 시작
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
        return { id: a.id, name: a.name, breed: a.look.breed, dept: a.dept, state: a.state, task: a.task, percent: a.percent, docs: a.docs };
      });
    },
    gust: function () { gust(performance.now()); },
    departments: DEPTS
  };

  resize();
  if (window.ResizeObserver) {
    var lastW = 0;
    new ResizeObserver(function () {
      var w = canvas.clientWidth;
      if (Math.abs(w - lastW) > 2) { lastW = w; resize(); }
    }).observe(canvas);
  } else {
    window.addEventListener('resize', resize);
  }
  drawPanels();
  log('🏢', '42층 사무실 문을 열었어요.');
  requestAnimationFrame(loop);
  // 글꼴이 늦게 도착하면 다시 그리기
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(function () { buildStatic(); renderPanels(); });

  if (CONFIG.feedUrl) {
    pollFeed();
    setInterval(pollFeed, Math.max(2000, CONFIG.pollMs || 5000));
    log('🔌', '이벤트 피드 연결: ' + CONFIG.feedUrl);
  }
  if (CONFIG.autoDemo && !CONFIG.feedUrl) setTimeout(demoScenario, 900);
})();
