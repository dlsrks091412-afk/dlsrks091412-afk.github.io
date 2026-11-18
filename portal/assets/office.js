/*
 * AI 에이전트 사무실 — 구름 위 고층 사무실 일러스트
 * ---------------------------------------------------------------
 * 모든 그림은 이 파일에서 캔버스 코드로 그립니다(이미지 파일 없음).
 * 논리 좌표 480 x 300 을 화면 크기에 맞게 부드럽게 확대해서 그립니다.
 *
 * 낮 테마: 창밖 파스텔 하늘·구름·새, 창으로 들어오는 햇살과 먼지, 가끔 바람이 불어 커튼과 식물이 흔들림
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
   * 1. 기본 상수와 레이아웃
   * ================================================================ */
  var W = 480, H = 300, WALL = 96;
  var FONT = '"Gowun Dodum", "Pretendard Variable", Pretendard, "Apple SD Gothic Neo", sans-serif';
  var SPEED = 44;                     // 걷는 속도 (논리 단위 / 초)
  var TAU = Math.PI * 2;

  var ELEV = { x: 14, y: 26, w: 48, h: 70 };   // 엘리베이터 (고층 사무실 출입구)
  var DOOR = { x: 38, y: 97 };
  var HALL = { x: 38, y: 116 };
  var GLASS_TOP = 14, GLASS_BOTTOM = 86;
  var WINDOWS = [{ x: 74, w: 96 }, { x: 182, w: 94 }, { x: 324, w: 88 }];
  var PILLAR_X = 300;                          // 창문 사이 기둥 (시계, 네온)

  var BOSS_LANE_X = 281;                       // 부서 ↔ 대표실을 잇는 세로 통로
  var BOSS_LANE_Y = 198;                       // 대표 책상 앞 대기 줄
  var BOSS_DESK = { x: 344, y: 148, w: 82 };
  var BOSS_SEAT = { x: 385, y: 154 };
  var BOSS_RUG = { x: 296, y: 110, w: 174, h: 98 };
  var BOSS_SLOTS = [385, 405, 365, 345, 325];

  /* 부서 정의 — 새 부서는 여기에 추가합니다. (README 참고) */
  var DEPTS = {
    records: {
      name: '기록부',
      desc: '문서 작업 담당',
      rug: { x: 66, y: 116, w: 200, h: 146, color: '#d7e6d2', edge: '#bcd5b6', dot: '#c7dcc1' },
      neon: '#7ce7c8',
      chair: '#a9cdb5', chairIn: '#c1dfca',
      desks: [
        { x: 82, y: 148 }, { x: 142, y: 148 }, { x: 202, y: 148 },
        { x: 82, y: 214 }, { x: 142, y: 214 }, { x: 202, y: 214 }
      ],
      deskW: 46,
      lines: ['문서 정리 중…', '오탈자 꼼꼼히 확인 중', '목차를 다듬고 있어요', '표 모양 맞추는 중', '요약본 쓰는 중', '참고자료 읽는 중'],
      unit: '문서'
    }
  };

  /* 아직 없는 부서 자리 */
  var FUTURE_ZONE = { x: 296, y: 222, w: 174, h: 70 };

  Object.keys(DEPTS).forEach(function (key) {
    var d = DEPTS[key];
    d.seats = d.desks.map(function (desk, i) {
      return {
        dept: key, index: i, desk: desk,
        x: desk.x + Math.round(d.deskW / 2),
        y: desk.y + 6,
        aisleY: desk.y - 24,
        agent: null, docs: 0
      };
    });
  });

  /* 바닥 식물 (흔들림) */
  var PLANTS = [
    { x: 68, y: 104, kind: 'monstera', s: 1.0 },
    { x: 290, y: 104, kind: 'snake', s: 1.0 },
    { x: 458, y: 164, kind: 'monstera', s: 1.1 },
    { x: 22, y: 288, kind: 'bush', s: 1.2 },
    { x: 268, y: 292, kind: 'monstera', s: 0.9 },
    { x: 172, y: 292, kind: 'snake', s: 0.85 }
  ];
  PLANTS.forEach(function (p, i) { p.ph = i * 1.7 + 0.3; });
  var HANGING = [{ x: 176, len: 30, ph: 0.4 }, { x: 417, len: 24, ph: 2.1 }];

  /* ================================================================
   * 2. 캔버스 준비
   * ================================================================ */
  var canvas = document.getElementById('office');
  var ctx = canvas.getContext('2d');
  var S = 2.5;                 // 논리 → 실제 픽셀 배율
  var TS = 1;                  // 작은 화면에서 글씨 키우기
  var bg = document.createElement('canvas');
  var sky = document.createElement('canvas');
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
    [canvas, bg, shade, fade].forEach(function (c) { c.width = Math.round(W * S); c.height = Math.round(H * S); });
    sky.width = Math.round(W * S); sky.height = Math.round(WALL * S);
    buildStatic();
  }

  /* ---------- 그리기 도우미 ---------- */
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
  function softShadow(g, x, y, rx, ry, a) {
    g.save(); g.translate(x, y); g.scale(1, ry / rx);
    g.fillStyle = rad(g, 0, 0, rx, [[0, 'rgba(110,70,60,' + a + ')'], [1, 'rgba(110,70,60,0)']]);
    g.beginPath(); g.arc(0, 0, rx, 0, TAU); g.fill(); g.restore();
  }
  function stroke(g, w, c) { g.lineWidth = w; g.strokeStyle = c; g.lineCap = 'round'; g.lineJoin = 'round'; g.stroke(); }
  var seedN = 1;
  function srand(s) { seedN = s; }
  function rnd() { seedN = (seedN * 16807) % 2147483647; return (seedN - 1) / 2147483646; }
  function hashStr(s) { var h = 0; for (var i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0; return Math.abs(h); }
  function clamp(v, a, b) { return Math.max(a, Math.min(b, v)); }

  /* ================================================================
   * 3. 정적인 배경 (벽, 바닥, 러그, 휴게 공간) — 테마/크기 바뀔 때만 다시 그림
   * ================================================================ */
  var CITY = { aviation: [], twinkles: [], stars: [] };

  function buildStatic() {
    drawBackground();
    drawSkyLayer();
  }

  function drawBackground() {
    var g = bg.getContext('2d');
    g.setTransform(S, 0, 0, S, 0, 0);
    g.clearRect(0, 0, W, H);

    // 바닥: 밝은 원목
    g.fillStyle = lin(g, 0, WALL, 0, H, [[0, '#efd9bf'], [1, '#e2c3a3']]);
    g.fillRect(0, WALL, W, H - WALL);
    srand(11);
    for (var y = WALL + 9; y < H; y += 9) {
      g.fillStyle = 'rgba(160,110,70,.10)'; g.fillRect(0, y, W, 0.6);
      var off = rnd() * 60;
      for (var x = off; x < W; x += 70 + rnd() * 40) { g.fillStyle = 'rgba(160,110,70,.08)'; g.fillRect(x, y - 9, 0.6, 9); }
    }
    // 바닥 빛 번짐
    g.fillStyle = rad(g, W * 0.45, WALL + 40, 260, [[0, 'rgba(255,245,230,.35)'], [1, 'rgba(255,245,230,0)']]);
    g.fillRect(0, WALL, W, H - WALL);

    // 벽
    g.fillStyle = lin(g, 0, 0, 0, WALL, [[0, '#f7e4da'], [1, '#f6ebe0']]);
    g.fillRect(0, 0, W, WALL);
    g.fillStyle = '#efd8cb'; g.fillRect(0, 0, W, 7);                 // 천장 몰딩
    g.fillStyle = 'rgba(255,255,255,.6)'; g.fillRect(0, 7, W, 0.8);
    g.fillStyle = '#ecd9c8'; g.fillRect(0, 88, W, 6);                 // 걸레받이
    g.fillStyle = '#dcc0a8'; g.fillRect(0, 93, W, 3);
    g.fillStyle = lin(g, 0, WALL, 0, WALL + 8, [[0, 'rgba(120,80,60,.16)'], [1, 'rgba(120,80,60,0)']]);
    g.fillRect(0, WALL, W, 8);
    // 옆 기둥
    g.fillStyle = '#ead3c4'; g.fillRect(0, 0, 6, H); g.fillRect(W - 6, 0, 6, H);
    g.fillStyle = 'rgba(120,80,60,.08)'; g.fillRect(6, WALL, 3, H - WALL); g.fillRect(W - 9, WALL, 3, H - WALL);

    // 기둥 장식 판 (시계/네온 자리)
    rr(g, PILLAR_X - 18, 12, 36, 74, 8, 'rgba(255,255,255,.35)');

    // 책장 (대표실, 벽 앞)
    softShadow(g, 446, 101, 30, 4, 0.22);
    rr(g, 422, 38, 48, 62, 5, '#d8b694');
    rr(g, 425, 41, 42, 56, 3, '#f0dfcc');
    var bookCols = ['#f0a7b2', '#9cc3e6', '#a6d1b0', '#f3cf86', '#c3b2ea', '#f6c1a4'];
    srand(5);
    for (var sh = 0; sh < 3; sh++) {
      var by = 41 + sh * 19;
      g.fillStyle = '#d8b694'; g.fillRect(425, by + 17, 42, 2);
      var bx = 427;
      while (bx < 463) {
        var bw = 3 + Math.floor(rnd() * 3), bh = 10 + rnd() * 6;
        if (rnd() < 0.15) { bx += 4; continue; }
        rr(g, bx, by + 17 - bh, bw, bh, 1, bookCols[Math.floor(rnd() * bookCols.length)]);
        bx += bw + 0.6;
      }
    }
    circ(g, 459, 49, 3.5, '#a6d1b0'); rr(g, 456, 52, 6, 4, 1.5, '#f6c1a4');

    // 부서 러그
    Object.keys(DEPTS).forEach(function (k) {
      var z = DEPTS[k].rug;
      softShadow(g, z.x + z.w / 2, z.y + z.h, z.w / 2, 5, 0.1);
      rr(g, z.x, z.y, z.w, z.h, 14, z.edge);
      rr(g, z.x + 4, z.y + 4, z.w - 8, z.h - 8, 11, z.color);
      for (var yy = z.y + 12; yy < z.y + z.h - 8; yy += 12) {
        for (var xx = z.x + 12 + ((yy / 12) % 2) * 6; xx < z.x + z.w - 8; xx += 12) circ(g, xx, yy, 1.1, z.dot);
      }
    });
    // 대표실 러그
    var b = BOSS_RUG;
    rr(g, b.x, b.y, b.w, b.h, 16, '#efc9bd');
    rr(g, b.x + 4, b.y + 4, b.w - 8, b.h - 8, 13, '#f7dcd2');
    g.strokeStyle = 'rgba(255,255,255,.7)'; g.lineWidth = 1; g.setLineDash([3, 3]);
    rrPath(g, b.x + 9, b.y + 9, b.w - 18, b.h - 18, 10); g.stroke(); g.setLineDash([]);

    // 비어 있는 부서 자리
    var fz = FUTURE_ZONE;
    rr(g, fz.x, fz.y, fz.w, fz.h, 14, 'rgba(255,255,255,.22)');
    g.strokeStyle = 'rgba(170,120,100,.35)'; g.lineWidth = 1; g.setLineDash([4, 4]);
    rrPath(g, fz.x + 0.5, fz.y + 0.5, fz.w - 1, fz.h - 1, 14); g.stroke(); g.setLineDash([]);
    for (var e = 0; e < 3; e++) rr(g, fz.x + 16 + e * 52, fz.y + 30, 40, 16, 4, 'rgba(170,120,100,.10)');

    // 휴게 공간: 둥근 러그, 소파, 테이블, 스탠드
    ell(g, 108, 286, 64, 13, '#f3e3d3');
    ell(g, 108, 286, 58, 10, '#f8ece0');
    softShadow(g, 112, 284, 44, 5, 0.25);
    rr(g, 70, 262, 84, 14, 8, '#e8aeb0');          // 등받이
    rr(g, 66, 270, 92, 16, 8, '#f0bdbd');          // 좌석
    rr(g, 62, 266, 12, 22, 6, '#e3a3a6'); rr(g, 150, 266, 12, 22, 6, '#e3a3a6');
    rr(g, 80, 268, 20, 9, 4, '#fbe3c8'); rr(g, 122, 268, 20, 9, 4, '#cfe3f3');
    ell(g, 120, 294, 15, 4, '#d1a37f'); ell(g, 120, 292.5, 15, 4, '#e2b892');
    rr(g, 115, 288, 5, 4, 1.5, '#fffaf2'); circ(g, 125, 290, 2, '#f3cf86');
    // 플로어 스탠드
    g.strokeStyle = '#b89b88'; g.lineWidth = 1.2; g.beginPath(); g.moveTo(48, 290); g.lineTo(48, 252); g.stroke();
    ell(g, 48, 291, 6, 1.6, '#b89b88');
    g.beginPath(); g.moveTo(41, 254); g.lineTo(55, 254); g.lineTo(52, 244); g.lineTo(44, 244); g.closePath(); g.fillStyle = '#fbe7cf'; g.fill();
    // 커피 카트
    softShadow(g, 222, 296, 22, 3, 0.2);
    rr(g, 202, 270, 40, 22, 5, '#f4e6d6'); rr(g, 202, 270, 40, 4, 2, '#d8b694');
    rr(g, 208, 262, 12, 9, 2, '#9cc3e6'); rr(g, 224, 264, 6, 7, 2, '#fffaf2'); rr(g, 232, 264, 6, 7, 2, '#f0a7b2');
    circ(g, 207, 293, 2, '#b89b88'); circ(g, 237, 293, 2, '#b89b88');
  }

  /* 창밖 풍경 (하늘 + 도시) — 창문 안쪽에만 보이도록 나중에 잘라서 그림 */
  function drawSkyLayer() {
    var g = sky.getContext('2d');
    g.setTransform(S, 0, 0, S, 0, 0);
    g.clearRect(0, 0, W, WALL);
    CITY.aviation = []; CITY.twinkles = []; CITY.stars = [];
    srand(42);
    if (!NIGHT) {
      g.fillStyle = lin(g, 0, GLASS_TOP, 0, GLASS_BOTTOM, [[0, '#a9d3f0'], [0.5, '#d4e7f4'], [0.78, '#fbe2e3'], [1, '#fde9df']]);
      g.fillRect(0, 0, W, WALL);
      g.fillStyle = rad(g, 396, 24, 80, [[0, 'rgba(255,250,225,.95)'], [0.18, 'rgba(255,240,200,.6)'], [1, 'rgba(255,240,200,0)']]);
      g.fillRect(0, 0, W, WALL);
      circ(g, 396, 24, 7, 'rgba(255,253,240,.95)');
    } else {
      g.fillStyle = lin(g, 0, GLASS_TOP, 0, GLASS_BOTTOM, [[0, '#0d1131'], [0.55, '#271d57'], [0.85, '#5a3470'], [1, '#7a4277']]);
      g.fillRect(0, 0, W, WALL);
      for (var i = 0; i < 70; i++) {
        var sx = rnd() * W, sy = GLASS_TOP + rnd() * 40;
        circ(g, sx, sy, 0.25 + rnd() * 0.45, 'rgba(255,250,230,' + (0.4 + rnd() * 0.5) + ')');
        if (rnd() < 0.3) CITY.stars.push({ x: sx, y: sy, ph: rnd() * TAU });
      }
      // 달
      g.fillStyle = rad(g, 110, 28, 40, [[0, 'rgba(230,220,255,.45)'], [1, 'rgba(230,220,255,0)']]);
      g.fillRect(0, 0, W, WALL);
      circ(g, 110, 28, 8, '#fff8e4');
      circ(g, 114, 25, 7.2, '#1b1745');
    }
    // 먼 산
    g.fillStyle = NIGHT ? '#2b2457' : 'rgba(214,196,226,.85)';
    g.beginPath(); g.moveTo(0, 62);
    for (var mx = 0; mx <= W; mx += 20) g.lineTo(mx, 58 - Math.sin(mx * 0.021) * 4 - Math.sin(mx * 0.057) * 2.5);
    g.lineTo(W, WALL); g.lineTo(0, WALL); g.fill();
    // 도시 세 겹 (멀리 → 가까이)
    var layers = NIGHT
      ? [{ top: 54, var: 10, col: '#29265a', win: 0.25 }, { top: 62, var: 12, col: '#211e4b', win: 0.4 }, { top: 72, var: 10, col: '#191739', win: 0.5 }]
      : [{ top: 54, var: 10, col: 'rgba(196,192,224,.9)', win: 0 }, { top: 62, var: 12, col: '#b7b9db', win: 0.3 }, { top: 72, var: 10, col: '#a8adcf', win: 0.4 }];
    layers.forEach(function (L, li) {
      var x = -4;
      while (x < W) {
        var bw = 6 + rnd() * 12, top = L.top - rnd() * L.var;
        if (rnd() < 0.08) top -= 8;
        rr(g, x, top, bw, WALL - top, 1, L.col);
        if (li === 0 && top < L.top - 8) CITY.aviation.push({ x: x + bw / 2, y: top - 1, ph: rnd() * TAU });
        for (var wy = top + 2.5; wy < WALL - 2; wy += 3.2) {
          for (var wx = x + 1.5; wx < x + bw - 1.5; wx += 2.6) {
            if (rnd() > L.win) continue;
            var col;
            if (NIGHT) {
              var r = rnd();
              col = r < 0.62 ? 'rgba(255,214,140,.9)' : r < 0.85 ? 'rgba(255,240,205,.85)' : 'rgba(150,225,255,.8)';
              if (rnd() < 0.04) CITY.twinkles.push({ x: wx, y: wy, ph: rnd() * TAU });
            } else col = 'rgba(255,255,255,.55)';
            g.fillStyle = col; g.fillRect(wx, wy, 1.1, 1.4);
          }
        }
        x += bw + 0.8 + rnd() * 2;
      }
    });
    // 아래쪽 안개
    g.fillStyle = lin(g, 0, 70, 0, WALL, NIGHT ? [[0, 'rgba(120,70,140,0)'], [1, 'rgba(120,70,140,.35)']] : [[0, 'rgba(255,240,240,0)'], [1, 'rgba(255,240,240,.55)']]);
    g.fillRect(0, 70, W, WALL - 70);
  }

  /* ================================================================
   * 4. 창문, 커튼, 엘리베이터, 벽 장식 (매 프레임)
   * ================================================================ */
  var clouds = [];
  srand(77);
  for (var ci = 0; ci < 7; ci++) clouds.push({ x: rnd() * (W + 80) - 40, y: 20 + rnd() * 30, s: 0.6 + rnd() * 0.7, v: 1.5 + rnd() * 2.5 });

  function clipGlass(g) {
    g.beginPath();
    WINDOWS.forEach(function (w) { g.rect(w.x, GLASS_TOP, w.w, GLASS_BOTTOM - GLASS_TOP); });
    g.clip();
  }

  function drawCloud(g, x, y, s, a) {
    g.save(); g.translate(x, y); g.scale(s, s);
    g.globalAlpha = a;
    ell(g, 0, 4, 22, 5.5, '#ffffff');
    circ(g, -9, 1, 6.5, '#ffffff'); circ(g, 1, -2, 8.5, '#ffffff'); circ(g, 11, 1, 6, '#ffffff');
    ell(g, 0, 7, 20, 2.5, 'rgba(230,215,235,.7)');
    g.restore(); g.globalAlpha = 1;
  }

  function drawWindows(t, dt) {
    var g = ctx;
    g.save();
    clipGlass(g);
    g.drawImage(sky, 0, 0, W, WALL);
    if (!NIGHT) {
      clouds.forEach(function (c) {
        c.x += c.v * dt * (1 + state.wind * 2);
        if (c.x > W + 40) { c.x = -40; c.y = 18 + Math.random() * 30; }
        drawCloud(g, c.x, c.y, c.s, 0.9);
      });
      if (state.birds) drawBirds(g, t);
    } else {
      CITY.stars.forEach(function (s) {
        var a = 0.4 + 0.6 * Math.max(0, Math.sin(t * 1.3 + s.ph));
        circ(g, s.x, s.y, 0.6, 'rgba(255,250,230,' + a.toFixed(2) + ')');
      });
      CITY.twinkles.forEach(function (w) {
        if (Math.sin(t * 0.7 + w.ph) > 0.6) { g.fillStyle = 'rgba(40,30,70,.9)'; g.fillRect(w.x, w.y, 1.1, 1.4); }
      });
      CITY.aviation.forEach(function (p) {
        var on = Math.sin(t * 3 + p.ph) > 0.4;
        if (on) { circ(g, p.x, p.y, 2.2, 'rgba(255,80,100,.25)'); circ(g, p.x, p.y, 0.7, '#ff6a7c'); }
      });
      if (state.shooting) drawShootingStar(g);
    }
    // 유리 반사
    WINDOWS.forEach(function (w) {
      g.fillStyle = 'rgba(255,255,255,' + (NIGHT ? 0.04 : 0.16) + ')';
      g.beginPath(); g.moveTo(w.x + 10, GLASS_BOTTOM); g.lineTo(w.x + 34, GLASS_TOP); g.lineTo(w.x + 46, GLASS_TOP); g.lineTo(w.x + 22, GLASS_BOTTOM); g.fill();
      g.beginPath(); g.moveTo(w.x + 52, GLASS_BOTTOM); g.lineTo(w.x + 68, GLASS_TOP); g.lineTo(w.x + 72, GLASS_TOP); g.lineTo(w.x + 56, GLASS_BOTTOM); g.fill();
    });
    g.restore();

    // 창틀
    var frame = NIGHT ? '#e9dfe8' : '#fffaf4';
    WINDOWS.forEach(function (w) {
      g.strokeStyle = frame; g.lineWidth = 3;
      g.strokeRect(w.x - 1.5, GLASS_TOP - 1.5, w.w + 3, GLASS_BOTTOM - GLASS_TOP + 3);
      g.fillStyle = frame;
      g.fillRect(w.x + w.w / 2 - 1, GLASS_TOP, 2, GLASS_BOTTOM - GLASS_TOP);
      g.fillRect(w.x, GLASS_TOP + 44, w.w, 1.5);
      rr(g, w.x - 5, GLASS_BOTTOM + 1, w.w + 10, 4, 2, '#f2e2d4');
      g.fillStyle = 'rgba(120,80,60,.12)'; g.fillRect(w.x - 4, GLASS_BOTTOM + 5, w.w + 8, 1.5);
    });
  }

  function drawBirds(g, t) {
    var b = state.birds;
    for (var i = 0; i < b.n; i++) {
      var bx = b.x - Math.abs(i - (b.n - 1) / 2) * 6, by = b.y + i * 2.2 + Math.sin(t * 2 + i) * 0.8;
      var flap = Math.sin(t * 12 + i * 1.3) * 1.6;
      g.beginPath(); g.moveTo(bx - 3, by - flap); g.quadraticCurveTo(bx - 1.5, by - 1, bx, by); g.quadraticCurveTo(bx + 1.5, by - 1, bx + 3, by - flap);
      stroke(g, 0.7, 'rgba(90,80,110,.75)');
    }
  }

  function drawShootingStar(g) {
    var s = state.shooting, k = s.life / s.max;
    var tx = s.x - s.vx * 0.35, ty = s.y - s.vy * 0.35;
    g.strokeStyle = lin(g, s.x, s.y, tx, ty, [[0, 'rgba(255,250,235,' + k + ')'], [1, 'rgba(255,250,235,0)']]);
    g.lineWidth = 1; g.lineCap = 'round';
    g.beginPath(); g.moveTo(s.x, s.y); g.lineTo(tx, ty); g.stroke();
  }

  /* 커튼: 바람(state.wind)이 불면 아래쪽부터 펄럭임 */
  function drawCurtains(t) {
    var g = ctx;
    var fabric = NIGHT ? 'rgba(170,160,220,.42)' : 'rgba(255,249,244,.88)';
    var fold = NIGHT ? 'rgba(90,80,150,.35)' : 'rgba(232,196,190,.55)';
    var hem = NIGHT ? 'rgba(200,190,240,.5)' : 'rgba(245,214,208,.9)';
    WINDOWS.forEach(function (w, wi) {
      // 커튼 봉
      g.fillStyle = '#d2b29a'; g.fillRect(w.x - 8, 9, w.w + 16, 1.6);
      circ(g, w.x - 8, 9.8, 1.8, '#c79f84'); circ(g, w.x + w.w + 8, 9.8, 1.8, '#c79f84');
      [0, 1].forEach(function (side) {
        var cw = 17;
        var x0 = side === 0 ? w.x - 6 : w.x + w.w + 6 - cw;
        var ph = wi * 1.9 + side * 0.7;
        var gust = state.wind;
        function sway(y) {
          var k = Math.pow((y - 10) / 78, 1.6);
          var base = Math.sin(t * 0.9 + ph) * 0.9;
          var flutter = gust * (9 + 4 * Math.sin(t * 4.6 + ph + y * 0.08));
          return (base + flutter) * k;
        }
        var bottom = function (xx) { return 88 - gust * 6 * (0.6 + 0.4 * Math.sin(t * 5 + xx * 0.4 + ph)); };
        g.beginPath();
        g.moveTo(x0, 10);
        for (var y = 10; y <= 88; y += 6) g.lineTo(x0 + sway(y) - (side === 0 ? 0 : -0), y);
        var steps = 6;
        for (var i = 0; i <= steps; i++) {
          var xx = x0 + (cw * i) / steps;
          g.lineTo(xx + sway(bottom(xx)), bottom(xx) + Math.sin(i * 1.7 + t * 2) * 0.8);
        }
        for (var y2 = 88; y2 >= 10; y2 -= 6) g.lineTo(x0 + cw + sway(y2) * 1.05, y2);
        g.closePath();
        g.fillStyle = fabric; g.fill();
        // 주름
        for (var f = 1; f < 4; f++) {
          var fx = x0 + (cw * f) / 4;
          g.beginPath(); g.moveTo(fx, 11);
          for (var y3 = 16; y3 <= 86; y3 += 7) g.lineTo(fx + sway(y3) + Math.sin(y3 * 0.15 + f) * 0.5, y3);
          stroke(g, 0.7, fold);
        }
        // 커튼 끈
        var tieY = 58;
        rr(g, x0 + sway(tieY) + 2, tieY, cw - 4, 2, 1, hem);
      });
    });
  }

  function drawElevator(t) {
    var g = ctx, e = ELEV;
    rr(g, e.x - 3, e.y - 8, e.w + 6, e.h + 8, 5, '#eadbd0');
    rr(g, e.x, e.y, e.w, e.h, 3, NIGHT ? '#fff0d4' : '#fff6e6');
    g.fillStyle = lin(g, 0, e.y, 0, e.y + e.h, [[0, 'rgba(255,230,190,.6)'], [1, 'rgba(255,230,190,0)']]);
    g.fillRect(e.x, e.y, e.w, e.h);
    var half = e.w / 2, o = state.elev.open;
    var lw = half * (1 - o);
    var doorFill = lin(g, e.x, 0, e.x + e.w, 0, [[0, '#d9d3d8'], [0.5, '#e8e3e6'], [1, '#cfc8cf']]);
    if (lw > 0.3) {
      rr(g, e.x, e.y, lw, e.h, 1.5, doorFill);
      rr(g, e.x + e.w - lw, e.y, lw, e.h, 1.5, doorFill);
      g.fillStyle = 'rgba(150,130,140,.35)';
      g.fillRect(e.x + lw - 0.5, e.y, 0.6, e.h); g.fillRect(e.x + e.w - lw, e.y, 0.6, e.h);
    }
    // 층 표시
    rr(g, e.x + 14, e.y - 6.5, 20, 5, 2.5, '#5d5068');
    var lit = o > 0.05 || state.elev.target > 0;
    g.fillStyle = lit ? '#ffd27a' : '#8c7f96';
    g.font = '4px ' + FONT; g.textBaseline = 'middle'; g.textAlign = 'center';
    g.fillText('42F', e.x + 24, e.y - 3.8);
    g.textAlign = 'left';
    // 호출 버튼
    rr(g, e.x + e.w + 4, e.y + 34, 4, 8, 2, '#e3d6cc');
    circ(g, e.x + e.w + 6, e.y + 38, 1, lit ? '#ffcf7a' : '#c9b8ae');
  }

  function drawClock() {
    var g = ctx, cx = PILLAR_X, cy = 30;
    softShadow(g, cx + 1, cy + 2, 11, 11, 0.15);
    circ(g, cx, cy, 10, '#d8b694'); circ(g, cx, cy, 8.6, '#fffaf2');
    for (var i = 0; i < 12; i++) { var a = i / 12 * TAU; circ(g, cx + Math.sin(a) * 7, cy - Math.cos(a) * 7, i % 3 ? 0.35 : 0.6, '#c8a98e'); }
    var now = new Date();
    var hA = ((now.getHours() % 12) + now.getMinutes() / 60) / 12 * TAU, mA = (now.getMinutes() + now.getSeconds() / 60) / 60 * TAU;
    g.beginPath(); g.moveTo(cx, cy); g.lineTo(cx + Math.sin(hA) * 4.2, cy - Math.cos(hA) * 4.2); stroke(g, 1.1, '#6b5560');
    g.beginPath(); g.moveTo(cx, cy); g.lineTo(cx + Math.sin(mA) * 6.2, cy - Math.cos(mA) * 6.2); stroke(g, 0.7, '#e8848c');
    circ(g, cx, cy, 0.9, '#6b5560');
  }

  /* 네온 튜브 (낮엔 꺼진 유리관, 밤엔 은은하게 빛남) */
  function neonPath(g, kind) {
    g.beginPath();
    if (kind === 'moon') {
      g.arc(PILLAR_X - 1, 62, 9, Math.PI * 0.55, Math.PI * 1.85);
      g.arc(PILLAR_X + 3, 59.5, 7, Math.PI * 1.75, Math.PI * 0.62, true);
    } else if (kind === 'star') {
      var cx = PILLAR_X + 9, cy = 52, r1 = 3.2, r2 = 1.3;
      for (var i = 0; i < 10; i++) {
        var a = -Math.PI / 2 + i * Math.PI / 5, r = i % 2 ? r2 : r1;
        g[i ? 'lineTo' : 'moveTo'](cx + Math.cos(a) * r, cy + Math.sin(a) * r);
      }
      g.closePath();
    } else if (kind === 'cove') {
      g.moveTo(10, 4.2); g.lineTo(W - 10, 4.2);
    }
  }

  function drawNeon(t, lit) {
    var g = ctx;
    var tubes = [
      { kind: 'cove', col: null, w: 1.3 },
      { kind: 'moon', col: '#ff9fd6', w: 1.3 },
      { kind: 'star', col: '#ffe08a', w: 1 }
    ];
    tubes.forEach(function (tb, i) {
      neonPath(g, tb.kind);
      var col = tb.col || lin(g, 10, 0, W - 10, 0, [[0, '#ff9fd6'], [0.5, '#b49bff'], [1, '#7ce7e0']]);
      if (!lit) { stroke(g, tb.w, 'rgba(240,210,225,.7)'); return; }
      var flick = 0.85 + 0.15 * Math.sin(t * 2.2 + i * 2) + (Math.random() < 0.004 ? -0.4 : 0);
      g.save();
      g.globalAlpha = clamp(flick, 0.3, 1);
      g.shadowColor = tb.col || '#c7a8ff'; g.shadowBlur = 10 * S;
      stroke(g, tb.w * 1.6, col);
      g.shadowBlur = 0;
      stroke(g, tb.w * 0.6, 'rgba(255,255,255,.85)');
      g.restore();
    });
  }

  /* 매달린 화분 */
  function drawHanging(g, h, t) {
    var a = Math.sin(t * 0.9 + h.ph) * 0.05 + state.wind * 0.18 * Math.sin(t * 4 + h.ph);
    g.save(); g.translate(h.x, 7); g.rotate(a);
    g.beginPath(); g.moveTo(-4, h.len - 6); g.lineTo(0, 0); g.lineTo(4, h.len - 6); stroke(g, 0.4, '#b89b88');
    rr(g, -5, h.len - 6, 10, 6, 3, '#fbf3ea');
    ell(g, 0, h.len - 6, 5, 1.2, '#e9d8c8');
    var vines = [-4, -1.5, 1.5, 4];
    vines.forEach(function (vx, i) {
      var len = 16 + (i % 2) * 8;
      var sw = Math.sin(t * 1.5 + i + h.ph) * 1.5 + state.wind * 5 * Math.sin(t * 6 + i);
      g.beginPath(); g.moveTo(vx, h.len - 3);
      g.quadraticCurveTo(vx + sw * 0.5, h.len + len * 0.5, vx + sw, h.len + len);
      stroke(g, 0.5, '#6fa57d');
      for (var k = 1; k <= 4; k++) {
        var f = k / 4, lx = vx + sw * f * f, ly = h.len - 3 + len * f;
        ell(g, lx + (k % 2 ? 1.2 : -1.2), ly, 1.7, 1.1, k % 2 ? '#8cc39a' : '#a5d4ad', k % 2 ? 0.5 : -0.5);
      }
    });
    g.restore();
  }

  /* ================================================================
   * 5. 식물, 가구
   * ================================================================ */
  function plantSway(p, t) {
    return Math.sin(t * 1.1 + p.ph) * 0.025 + state.wind * 0.11 * Math.sin(t * 5.2 + p.ph) * (p.kind === 'snake' ? 0.5 : 1);
  }

  function drawPlant(g, p, t) {
    var a = plantSway(p, t), s = p.s;
    softShadow(g, p.x, p.y, 9 * s, 2.5 * s, 0.22);
    g.save(); g.translate(p.x, p.y); g.scale(s, s);
    // 화분
    g.beginPath(); g.moveTo(-6.5, -11); g.lineTo(6.5, -11); g.lineTo(5, 0); g.lineTo(-5, 0); g.closePath();
    g.fillStyle = p.kind === 'snake' ? '#f6efe7' : '#e7a684'; g.fill();
    rr(g, -7.5, -13, 15, 3, 1.5, p.kind === 'snake' ? '#e9ddd0' : '#d99375');
    g.translate(0, -12); g.rotate(a);
    if (p.kind === 'monstera') {
      var leaves = [[-0.9, 17, '#7ab38a'], [-0.3, 21, '#8cc39a'], [0.35, 19, '#7ab38a'], [0.95, 15, '#a5d4ad'], [0.05, 25, '#6fa57d']];
      leaves.forEach(function (L, i) {
        var la = L[0] + state.wind * 0.12 * Math.sin(t * 7 + i);
        g.save(); g.rotate(la);
        g.beginPath(); g.moveTo(0, 0); g.lineTo(0, -L[1] + 5); stroke(g, 0.7, '#6a9c76');
        ell(g, 0, -L[1] + 2, 5.5, 6.5, L[2]);
        g.strokeStyle = 'rgba(255,255,255,.35)'; g.lineWidth = 0.4;
        g.beginPath(); g.moveTo(0, -L[1] - 3.5); g.lineTo(0, -L[1] + 7); g.stroke();
        g.restore();
      });
    } else if (p.kind === 'snake') {
      var blades = [[-5, 18, -0.15], [-2, 24, -0.05], [1.5, 27, 0.04], [4.5, 20, 0.14]];
      blades.forEach(function (b, i) {
        g.save(); g.translate(b[0], 0); g.rotate(b[2] + state.wind * 0.05 * Math.sin(t * 6 + i));
        g.beginPath(); g.moveTo(-1.8, 0); g.quadraticCurveTo(-2.2, -b[1] * 0.6, 0, -b[1]); g.quadraticCurveTo(2.2, -b[1] * 0.6, 1.8, 0); g.closePath();
        g.fillStyle = i % 2 ? '#7fb08a' : '#94c49f'; g.fill();
        g.beginPath(); g.moveTo(0, -1); g.lineTo(0, -b[1] + 3); stroke(g, 0.35, 'rgba(240,235,180,.6)');
        g.restore();
      });
    } else {
      var balls = [[-6, -5, 6], [5, -6, 6.5], [0, -11, 7], [-3, -16, 5], [4, -15, 5]];
      balls.forEach(function (b, i) { circ(g, b[0] + state.wind * Math.sin(t * 6 + i) * 0.6, b[1], b[2], i % 2 ? '#8cc39a' : '#a5d4ad'); });
      circ(g, -2, -12, 1.2, '#f7b6c6'); circ(g, 3, -8, 1.1, '#fff1f4'); circ(g, -6, -6, 1, '#f7b6c6');
    }
    g.restore();
  }

  function drawChair(g, seat, occupied) {
    var d = DEPTS[seat.dept], x = seat.x, y = seat.desk.y;
    rr(g, x - 8.5, y - 20, 17, 16, 6, d.chair);
    rr(g, x - 6.5, y - 18, 13, 12, 5, d.chairIn);
    if (!occupied) rr(g, x - 9, y - 6, 18, 5, 2.5, d.chair);
  }

  function drawSteam(g, x, y, t, seed) {
    for (var i = 0; i < 2; i++) {
      var p = ((t * 0.5 + i * 0.5 + seed) % 1);
      g.beginPath();
      for (var k = 0; k <= 6; k++) {
        var yy = y - k * 1.3 - p * 3;
        var xx = x + i * 1.6 - 0.8 + Math.sin(k * 0.9 + t * 3 + i) * 0.8;
        g[k ? 'lineTo' : 'moveTo'](xx, yy);
      }
      stroke(g, 0.5, (NIGHT ? 'rgba(255,255,255,' : 'rgba(190,150,140,') + (0.55 * (1 - p)).toFixed(2) + ')');
    }
  }

  function drawDesk(g, seat, t) {
    var d = DEPTS[seat.dept], x = seat.desk.x, y = seat.desk.y, w = d.deskW;
    var working = seat.agent && seat.agent.state === 'working' && !seat.agent.moving;
    softShadow(g, x + w / 2, y + 26, w / 2 + 3, 3.5, 0.22);
    rr(g, x + 1, y + 11, w - 2, 14, 3, '#dcbc98');
    g.fillStyle = 'rgba(255,255,255,.25)'; g.fillRect(x + 4, y + 17, w - 8, 0.6);
    rr(g, x + w / 2 - 3, y + 19, 6, 1.6, 0.8, '#c39f7c');
    rr(g, x, y, w, 14, 4, '#f3e2cd');
    g.fillStyle = 'rgba(255,255,255,.55)'; g.fillRect(x + 3, y + 0.8, w - 6, 0.8);
    // 서류 더미
    var stack = Math.min(6, Math.ceil(seat.docs / 2));
    for (var i = 0; i < stack; i++) rr(g, x + 4 + (i % 2) * 0.6, y + 7 - i * 1.6, 9, 2.2, 0.8, i % 2 ? '#fffdf8' : '#f2ebdf');
    // 머그컵 + 김
    rr(g, x + w - 9, y + 4, 4.5, 5, 1.4, '#fffaf2');
    g.beginPath(); g.arc(x + w - 4.2, y + 6.5, 1.4, -1.2, 1.2); stroke(g, 0.6, '#fffaf2');
    ell(g, x + w - 6.8, y + 4.2, 2, 0.6, '#c99a74');
    if (working) drawSteam(g, x + w - 6.8, y + 3, t, seat.index * 0.37);
    // 스탠드
    g.beginPath(); g.moveTo(x + 6, y + 6); g.lineTo(x + 4.5, y - 7); g.lineTo(x + 9, y - 11); stroke(g, 0.9, '#c9b2a6');
    ell(g, x + 6, y + 6.3, 3, 0.9, '#c9b2a6');
    g.beginPath(); g.moveTo(x + 7, y - 13); g.lineTo(x + 13, y - 10.5); g.lineTo(x + 11, y - 7.5); g.closePath();
    g.fillStyle = '#a9cdb5'; g.fill();
    if (working || NIGHT && seat.agent) circ(g, x + 11.5, y - 8.6, 1, '#fff2c4');
    // 노트북 (뒷면)
    var cx = seat.x;
    rr(g, cx - 9, y + 3.5, 18, 2.6, 1, '#c7ccd8');
    rr(g, cx - 8, y - 6.5, 16, 10.5, 2, '#dde1ea');
    circ(g, cx, y - 1.3, 1.6, '#f0a7b2');
    if (working) {
      g.fillStyle = 'rgba(170,225,255,' + (0.5 + 0.2 * Math.sin(t * 3)).toFixed(2) + ')';
      g.fillRect(cx - 7, y - 7.4, 14, 0.9);
    }
  }

  function drawBossDesk(g, t) {
    var x = BOSS_DESK.x, y = BOSS_DESK.y, w = BOSS_DESK.w;
    softShadow(g, x + w / 2, y + 31, w / 2 + 4, 4, 0.25);
    rr(g, x + 1, y + 14, w - 2, 16, 4, '#bb8a63');
    rr(g, x + 6, y + 18, 18, 9, 2, '#a97a55'); rr(g, x + w - 24, y + 18, 18, 9, 2, '#a97a55');
    circ(g, x + 15, y + 22.5, 0.9, '#f3cf86'); circ(g, x + w - 15, y + 22.5, 0.9, '#f3cf86');
    rr(g, x + w / 2 - 10, y + 18, 20, 7, 2, '#f3cf86');
    g.fillStyle = '#7a5640'; g.font = '4.2px ' + FONT; g.textAlign = 'center'; g.textBaseline = 'middle';
    g.fillText('대표', x + w / 2, y + 21.7); g.textAlign = 'left';
    rr(g, x, y, w, 16, 5, '#d6a983');
    g.fillStyle = 'rgba(255,255,255,.35)'; g.fillRect(x + 4, y + 1, w - 8, 0.8);
    // 모니터 (뒷면, 왼쪽)
    rr(g, x + 13, y + 1, 6, 6, 1, '#c3c8d3');
    rr(g, x + 5, y - 13, 22, 15, 3, '#dde1ea');
    circ(g, x + 16, y - 5.5, 1.4, '#9cc3e6');
    g.fillStyle = 'rgba(170,225,255,' + (0.5 + 0.2 * Math.sin(t * 2.4)).toFixed(2) + ')';
    g.fillRect(x + 7, y - 14, 18, 0.9);
    // 꽃병
    rr(g, x + 30, y + 3, 4, 7, 2, '#cfe3f3');
    [[-2, -3, '#f7b6c6'], [1.5, -4.5, '#fff1f4'], [3, -1.5, '#f0a7b2']].forEach(function (f) {
      g.beginPath(); g.moveTo(x + 32, y + 4); g.lineTo(x + 32 + f[0], y + 4 + f[1]); stroke(g, 0.4, '#7ab38a');
      circ(g, x + 32 + f[0], y + 4 + f[1], 1.5, f[2]);
    });
    // 찻잔
    ell(g, x + 56, y + 9, 4, 1.4, '#f6efe7');
    rr(g, x + 54, y + 5, 4.5, 4, 1.5, '#fffaf2');
    drawSteam(g, x + 56.2, y + 4.5, t, 0.5);
    // 결재함 (승인한 서류가 쌓임)
    rr(g, x + 63, y + 3, 15, 8, 2, '#c99a74'); rr(g, x + 64, y + 4, 13, 6, 1.5, '#e2bd98');
    var approved = Math.min(4, state.approvedPile);
    for (var i = 0; i < approved; i++) rr(g, x + 65, y + 7.5 - i * 1.4, 11, 1.6, 0.6, i % 2 ? '#f6d38d' : '#fffdf8');
  }

  function drawBossChair(g) {
    var x = BOSS_SEAT.x, y = BOSS_DESK.y;
    rr(g, x - 12, y - 27, 24, 24, 9, '#d58f8d');
    rr(g, x - 9.5, y - 24.5, 19, 19, 7, '#e7a9a5');
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

  function hairBack(g, L, x, hy, dir) {
    if (dir === 'up') return;
    if (L.style === 'long') ell(g, x + (dir === 'left' ? 1.5 : dir === 'right' ? -1.5 : 0), hy + 4, 8.4, 10, L.hair);
    if (L.style === 'bob') ell(g, x + (dir === 'left' ? 1 : dir === 'right' ? -1 : 0), hy + 1.5, 8.8, 8, L.hair);
  }

  function hairFront(g, L, x, hy, dir) {
    var c = L.hair;
    if (dir === 'up') {
      circ(g, x, hy - 0.5, 7.9, c);
      if (L.style === 'long') rr(g, x - 7.6, hy, 15.2, 12, 6, c);
      if (L.style === 'bob') rr(g, x - 8, hy - 1, 16, 8, 5, c);
      if (L.style === 'bun') circ(g, x, hy - 8.5, 3.6, c);
      if (L.style === 'curly') for (var j = 0; j < 6; j++) circ(g, x - 7 + j * 2.8, hy - 5 + Math.abs(j - 2.5) * 0.7, 2.6, c);
      return;
    }
    if (dir === 'down') {
      g.beginPath();
      g.moveTo(x - 7.9, hy + 1.5);
      g.bezierCurveTo(x - 9, hy - 10.5, x + 9, hy - 10.5, x + 7.9, hy + 1.5);
      g.quadraticCurveTo(x + 5.5, hy - 4.2, x + 1.5, hy - 3.6);
      g.quadraticCurveTo(x - 3, hy - 5, x - 7.9, hy + 1.5);
      g.fillStyle = c; g.fill();
      if (L.style === 'long' || L.style === 'bob') { rr(g, x - 8.3, hy - 2, 3, L.style === 'long' ? 13 : 8, 1.5, c); rr(g, x + 5.3, hy - 2, 3, L.style === 'long' ? 13 : 8, 1.5, c); }
      if (L.style === 'bun') circ(g, x, hy - 9, 3.6, c);
      if (L.style === 'curly') for (var i = 0; i < 5; i++) circ(g, x - 6 + i * 3, hy - 6.4 + Math.abs(i - 2) * 0.9, 2.6, c);
      g.fillStyle = 'rgba(255,255,255,.18)'; g.beginPath(); g.ellipse(x - 2.5, hy - 5.5, 2.5, 1, -0.3, 0, TAU); g.fill();
    } else {
      var s = dir === 'left' ? 1 : -1;      // 뒤통수 방향
      g.beginPath();
      g.arc(x, hy - 0.5, 7.9, Math.PI, 0);
      g.lineTo(x + s * 7.9, hy + 3);
      g.quadraticCurveTo(x + s * 3, hy + 3, x + s * 1.5, hy - 1.5);
      g.quadraticCurveTo(x - s * 2, hy - 3.5, x - s * 7.9, hy - 0.5);
      g.fillStyle = c; g.fill();
      if (L.style === 'long') rr(g, s > 0 ? x + 1 : x - 8, hy - 1, 7, 13, 3, c);
      if (L.style === 'bob') rr(g, s > 0 ? x + 1 : x - 8.6, hy - 1, 7.6, 8, 3, c);
      if (L.style === 'bun') circ(g, x + s * 5.5, hy - 6, 3.4, c);
      if (L.style === 'curly') for (var k = 0; k < 4; k++) circ(g, x - 5 + k * 3.3, hy - 6.5, 2.5, c);
    }
  }

  function limb(g, x0, y0, x1, y1, w, col) {
    g.beginPath(); g.moveTo(x0, y0); g.lineTo(x1, y1); stroke(g, w, col);
  }

  /*
   * p = { look, dir, moving, phase, carrying, talking, typing, blink, seed }
   * (x, y) = 발 아래 중앙
   */
  function drawPerson(g, p, x, y, t) {
    var L = p.look, dir = p.dir || 'down';
    var walk = p.moving ? p.phase : 0;
    var sw = p.moving ? Math.sin(walk) : 0;
    var bob = p.moving ? -Math.abs(Math.cos(walk)) * 1.1 : p.typing ? (Math.sin(t * 10 + p.seed) > 0.3 ? -0.35 : 0) : Math.sin(t * 1.7 + p.seed) * 0.35;
    var side = dir === 'left' ? -1 : dir === 'right' ? 1 : 0;
    var hy = y - 26 + bob;

    softShadow(g, x, y, 7.5, 2.4, 0.28);
    hairBack(g, L, x, hy, dir);

    // 다리
    var hip = y - 9.5 + bob * 0.4;
    var shoe = '#6b5560';
    if (!side) {
      var l1 = Math.max(0, sw) * 2.2, l2 = Math.max(0, -sw) * 2.2;
      rr(g, x - 4.2, hip, 3.6, 8.5 - l1, 1.8, L.pants); rr(g, x + 0.6, hip, 3.6, 8.5 - l2, 1.8, L.pants);
      ell(g, x - 2.4, y - 0.8 - l1, 2.3, 1.3, shoe); ell(g, x + 2.4, y - 0.8 - l2, 2.3, 1.3, shoe);
    } else {
      [[sw, 0.75], [-sw, 1]].forEach(function (lg) {
        g.save(); g.translate(x, hip); g.rotate(-lg[0] * 0.5 * side);
        g.globalAlpha = lg[1];
        rr(g, -1.8, 0, 3.6, 8.3, 1.8, L.pants);
        ell(g, side * 0.9, 8.6, 2.5, 1.3, shoe);
        g.restore(); g.globalAlpha = 1;
      });
    }

    // 몸통
    var by = y - 20 + bob;
    var armCol = L.shirt;
    if (side) limb(g, x - side * 1.5, by + 2.5, x - side * 1.5 - sw * 3 * side, by + 8.5, 2.8, armCol); // 뒤쪽 팔
    rr(g, x - 6, by, 12, 11.5, 4.8, L.shirt);
    g.fillStyle = 'rgba(80,50,60,.07)'; rrPath(g, x - 6, by + 7.5, 12, 4, 2); g.fill();
    if (L.scarf && dir !== 'up') rr(g, x - 4.5, by - 0.5, 9, 3, 1.5, L.scarf);
    else if (dir === 'down') { g.beginPath(); g.moveTo(x - 2.2, by); g.lineTo(x, by + 2.4); g.lineTo(x + 2.2, by); g.fillStyle = 'rgba(255,255,255,.7)'; g.fill(); }

    // 팔
    if (!side) {
      if (p.carrying && dir === 'down') {
        limb(g, x - 5.4, by + 2.5, x - 4, by + 7, 2.8, armCol); limb(g, x + 5.4, by + 2.5, x + 4, by + 7, 2.8, armCol);
      } else if (p.typing) {
        var tp = Math.sin(t * 14 + p.seed) * 0.6;
        limb(g, x - 5.4, by + 2.5, x - 3.2, by + 8 + tp, 2.8, armCol); limb(g, x + 5.4, by + 2.5, x + 3.2, by + 8 - tp, 2.8, armCol);
      } else {
        limb(g, x - 5.4, by + 2.5, x - 6.6, by + 8.6 + sw * 1.4, 2.8, armCol); limb(g, x + 5.4, by + 2.5, x + 6.6, by + 8.6 - sw * 1.4, 2.8, armCol);
        circ(g, x - 6.6, by + 9.6 + sw * 1.4, 1.5, L.skin); circ(g, x + 6.6, by + 9.6 - sw * 1.4, 1.5, L.skin);
      }
    }

    // 머리
    rr(g, x - 1.5, by - 2.2, 3, 3, 1, L.skin);
    if (!side && dir === 'down') { circ(g, x - 7.2, hy + 1, 1.5, L.skin); circ(g, x + 7.2, hy + 1, 1.5, L.skin); }
    circ(g, x, hy, 7.5, L.skin);
    hairFront(g, L, x, hy, dir);

    // 얼굴
    var eye = '#4a3438';
    if (dir === 'down') {
      if (p.blink) { limb(g, x - 3.6, hy + 1, x - 1.8, hy + 1, 0.55, eye); limb(g, x + 1.8, hy + 1, x + 3.6, hy + 1, 0.55, eye); }
      else if (p.happy) {
        g.beginPath(); g.arc(x - 2.7, hy + 1.4, 1.1, Math.PI, 0); stroke(g, 0.6, eye);
        g.beginPath(); g.arc(x + 2.7, hy + 1.4, 1.1, Math.PI, 0); stroke(g, 0.6, eye);
      } else {
        ell(g, x - 2.7, hy + 0.9, 0.95, 1.25, eye); ell(g, x + 2.7, hy + 0.9, 0.95, 1.25, eye);
        circ(g, x - 2.4, hy + 0.4, 0.35, '#fff'); circ(g, x + 3, hy + 0.4, 0.35, '#fff');
      }
      ell(g, x - 4.5, hy + 3, 1.6, 1, 'rgba(255,140,150,.45)'); ell(g, x + 4.5, hy + 3, 1.6, 1, 'rgba(255,140,150,.45)');
      if (p.talking && Math.sin(t * 14) > 0) ell(g, x, hy + 3.9, 0.9, 0.7, '#c2646c');
      else { g.beginPath(); g.arc(x, hy + 3.2, 0.9, 0.2, Math.PI - 0.2); stroke(g, 0.5, '#b0606a'); }
    } else if (side) {
      if (!p.blink) ell(g, x + side * 3.6, hy + 0.9, 0.85, 1.2, eye);
      else limb(g, x + side * 2.8, hy + 1, x + side * 4.4, hy + 1, 0.5, eye);
      ell(g, x + side * 4.6, hy + 3, 1.3, 0.9, 'rgba(255,140,150,.45)');
    }

    // 앞쪽 팔 (옆모습)
    if (side) {
      if (p.carrying) limb(g, x + side * 1, by + 2.5, x + side * 5, by + 6.5, 2.8, armCol);
      else { limb(g, x + side * 1, by + 2.5, x + side * 1 + sw * 3 * side, by + 8.5, 2.8, armCol); circ(g, x + side * 1 + sw * 3 * side, by + 9.5, 1.5, L.skin); }
    }

    // 결재 파일
    if (p.carrying) {
      var fx = side ? x + side * 6 : x, fy = by + 1.5;
      if (dir === 'up') { rr(g, x - 7.5, fy + 2, 2, 7, 1, '#f4cf7a'); rr(g, x + 5.5, fy + 2, 2, 7, 1, '#f4cf7a'); }
      else {
        g.save(); g.translate(fx, fy + 4); if (side) g.rotate(side * 0.15);
        rr(g, -5.5, -4.5, 11, 9, 1.8, '#e9b85f');
        rr(g, -4.5, -5.5, 9, 3, 1, '#fffdf8');
        rr(g, -5.5, -3.5, 11, 8, 1.8, '#f6d38d');
        circ(g, 0, 0.5, 1.6, '#e8848c');
        g.restore();
        if (!side) { circ(g, x - 4.8, by + 7.4, 1.5, L.skin); circ(g, x + 4.8, by + 7.4, 1.5, L.skin); }
      }
    }
  }

  /* ================================================================
   * 7. 말풍선, 이름표, 이펙트
   * ================================================================ */
  function textWidth(text, font) { ctx.font = font; return ctx.measureText(text).width; }

  function drawBubble(x, y, text, opts) {
    opts = opts || {};
    var g = ctx;
    var fs = 6.6 * TS;
    var font = fs + 'px ' + FONT;
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
    if (opts.border) { rrPath(g, bx, by, w, h, h / 2); stroke(g, 0.6, opts.border); }
    g.fillStyle = ink; g.font = font; g.textBaseline = 'middle';
    g.fillText(text, bx + 4.5 * TS, by + h / 2 + 0.4);
  }

  function drawName(x, y, text, highlight) {
    var g = ctx, fs = 5.2 * TS, font = fs + 'px ' + FONT;
    var tw = textWidth(text, font), w = tw + 6 * TS, h = fs + 3 * TS;
    var bx = x - w / 2;
    rr(g, bx, y, w, h, h / 2, highlight ? (NIGHT ? 'rgba(255,143,200,.85)' : 'rgba(232,132,140,.9)') : (NIGHT ? 'rgba(30,28,70,.72)' : 'rgba(255,255,255,.82)'));
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

  function drawFloorSigns(t) {
    var g = ctx;
    Object.keys(DEPTS).forEach(function (k) {
      var d = DEPTS[k], z = d.rug;
      var font = (7 * TS) + 'px ' + FONT, tw = textWidth(d.name, font);
      var x = z.x + 8, y = z.y + z.h - 15, w = tw + 12;
      rr(g, x, y, w, 11, 5.5, NIGHT ? 'rgba(20,24,40,.55)' : 'rgba(255,255,255,.8)');
      g.font = font; g.textBaseline = 'middle';
      if (NIGHT) {
        g.save(); g.shadowColor = d.neon; g.shadowBlur = 8 * S; g.fillStyle = d.neon;
        g.fillText(d.name, x + 6, y + 6); g.restore();
        g.fillStyle = '#eafff8'; g.fillText(d.name, x + 6, y + 6);
      } else {
        g.fillStyle = '#5b7a62'; g.fillText(d.name, x + 6, y + 6);
      }
      g.font = (5 * TS) + 'px ' + FONT; g.fillStyle = NIGHT ? 'rgba(200,240,230,.6)' : 'rgba(90,120,98,.75)';
      g.fillText(d.desc, x + w + 4, y + 6);
    });
    var bf = (7 * TS) + 'px ' + FONT, bw = textWidth('대표실', bf) + 12, bx = BOSS_RUG.x + 8, byy = BOSS_RUG.y + 8;
    rr(g, bx, byy, bw, 11, 5.5, NIGHT ? 'rgba(30,20,40,.55)' : 'rgba(255,255,255,.85)');
    g.font = bf; g.textBaseline = 'middle';
    if (NIGHT) { g.save(); g.shadowColor = '#ff9fd6'; g.shadowBlur = 8 * S; g.fillStyle = '#ff9fd6'; g.fillText('대표실', bx + 6, byy + 6); g.restore(); g.fillStyle = '#fff0f8'; }
    else g.fillStyle = '#b06a70';
    g.fillText('대표실', bx + 6, byy + 6);
    g.font = (5.5 * TS) + 'px ' + FONT; g.fillStyle = NIGHT ? 'rgba(220,210,255,.45)' : 'rgba(150,110,100,.6)';
    g.fillText('새 부서가 들어올 자리', FUTURE_ZONE.x + 10, FUTURE_ZONE.y + 12);
  }

  /* ================================================================
   * 8. 상태
   * ================================================================ */
  var state = {
    agents: [],
    particles: [],
    motes: [],
    approvals: [],          // {id, agentId, title, detail, status: 'coming'|'arrived'|'approved'|'rejected', ...}
    approvedPile: 0,
    bossHappyUntil: 0,
    docsToday: 0,
    boss: { look: BOSS_LOOK, dir: 'down', seed: 1.3 },
    elev: { open: 0, target: 0, holdUntil: 0 },
    wind: 0,
    nextGust: performance.now() + 6000,
    birds: null, nextBirds: performance.now() + 9000,
    shooting: null, nextShooting: performance.now() + 7000,
    listeners: [],
    resolvers: {},
    selected: null
  };

  /* 햇살 속 먼지 */
  for (var mi = 0; mi < 46; mi++) state.motes.push({ u: Math.random(), v: Math.random(), w: Math.floor(Math.random() * 6), ph: Math.random() * TAU, sp: 0.01 + Math.random() * 0.02 });

  /* ================================================================
   * 9. 에이전트 동작
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
      look: opts.look || LOOKS[hashStr(id) % LOOKS.length],
      seed: (hashStr(id) % 100) / 10,
      x: DOOR.x, y: DOOR.y - 1, dir: 'down',
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
    log('🌤️', agent.name + ' 출근 · ' + DEPTS[dept].name + ' (' + agent.task + ')');
    setPath(agent, pathDoorToSeat(seat), function () {
      agent.state = 'working';
      agent.dir = 'down';
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
    state.particles.push({ kind: 'doc', x: a.x + 8, y: a.y - 34, vy: -9, life: 1.8, max: 1.8, text: '+1' });
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
        a.dir = 'up';
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
        say(a, '감사합니다! 바로 진행할게요', 2600);
        state.particles.push({ kind: 'stamp', x: a.x, y: a.y - 46, vy: -5, life: 1.5, max: 1.5, text: '승인' });
        for (var h = 0; h < 5; h++) state.particles.push({ kind: 'heart', x: a.x - 6 + Math.random() * 12, y: a.y - 30, vx: (Math.random() - 0.5) * 8, vy: -10 - Math.random() * 8, life: 1.6, max: 1.6 });
      } else {
        say(a, '넵, 다듬어서 다시 올릴게요', 2800);
        state.particles.push({ kind: 'stamp', x: a.x, y: a.y - 46, vy: -5, life: 1.5, max: 1.5, text: '반려', bad: true });
      }
      a.state = 'returning';
      setTimeout(function () {
        a.dir = 'left';
        setPath(a, pathBossToSeat(a, a.seat), function () {
          a.state = 'working';
          a.carrying = false;
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
   * 10. 매 프레임 업데이트
   * ================================================================ */
  function gust(now) {
    state.wind = Math.max(state.wind, NIGHT ? 0.45 : 1);
    if (!NIGHT) {
      for (var i = 0; i < 9; i++) {
        var w = WINDOWS[Math.floor(Math.random() * WINDOWS.length)];
        state.particles.push({
          kind: 'petal', x: w.x + Math.random() * w.w, y: GLASS_TOP + 20 + Math.random() * 50,
          vx: 18 + Math.random() * 22, vy: 6 + Math.random() * 10, rot: Math.random() * TAU, vr: (Math.random() - 0.5) * 6,
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
      if (!state.birds && now > state.nextBirds) state.birds = { x: -20, y: 22 + Math.random() * 18, n: 3 + Math.floor(Math.random() * 3) };
      if (state.birds) { state.birds.x += 20 * dt; if (state.birds.x > W + 30) { state.birds = null; state.nextBirds = now + 25000 + Math.random() * 25000; } }
    } else {
      if (!state.shooting && now > state.nextShooting) {
        var w = WINDOWS[Math.floor(Math.random() * WINDOWS.length)];
        state.shooting = { x: w.x + 10 + Math.random() * 30, y: GLASS_TOP + 2, vx: 70, vy: 32, life: 0.9, max: 0.9 };
      }
      if (state.shooting) {
        var s = state.shooting; s.x += s.vx * dt; s.y += s.vy * dt; s.life -= dt;
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
        if (Math.abs(dx) > Math.abs(dy)) a.dir = dx < 0 ? 'left' : 'right';
        else if (dist > 0.01) a.dir = dy < 0 ? 'up' : 'down';
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
          kind: 'spark', x: a.x - 10 + Math.random() * 20, y: a.y - 16 - Math.random() * 18,
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
  function shaftPanes() {
    var panes = [];
    WINDOWS.forEach(function (w) {
      panes.push([w.x + 3, w.x + w.w / 2 - 2]);
      panes.push([w.x + w.w / 2 + 2, w.x + w.w - 3]);
    });
    return panes;
  }
  var PANES = shaftPanes();
  var SUN_DX = 52, SUN_LEN = 120;

  function drawDaylight(t) {
    var g = ctx;
    g.save();
    g.globalCompositeOperation = 'screen';
    var breathe = 0.85 + 0.15 * Math.sin(t * 0.35);
    PANES.forEach(function (p, i) {
      var curtainShade = 1 - state.wind * 0.25 * (0.5 + 0.5 * Math.sin(t * 5 + i));
      var a = 0.5 * breathe * curtainShade;
      // 바닥에 떨어진 햇빛
      g.fillStyle = lin(g, 0, WALL, 0, WALL + SUN_LEN, [[0, 'rgba(255,214,160,' + a.toFixed(3) + ')'], [0.7, 'rgba(255,220,170,' + (a * 0.35).toFixed(3) + ')'], [1, 'rgba(255,226,180,0)']]);
      g.beginPath();
      g.moveTo(p[0], WALL); g.lineTo(p[1], WALL);
      g.lineTo(p[1] + SUN_DX, WALL + SUN_LEN); g.lineTo(p[0] + SUN_DX, WALL + SUN_LEN);
      g.closePath(); g.fill();
      // 공기 중의 빛줄기
      g.fillStyle = lin(g, 0, GLASS_TOP, 0, WALL + SUN_LEN, [[0, 'rgba(255,240,210,0)'], [0.4, 'rgba(255,240,210,' + (a * 0.32).toFixed(3) + ')'], [1, 'rgba(255,240,210,0)']]);
      g.beginPath();
      g.moveTo(p[0], GLASS_TOP + 10); g.lineTo(p[1], GLASS_TOP + 10);
      g.lineTo(p[1] + SUN_DX, WALL + SUN_LEN); g.lineTo(p[0] + SUN_DX, WALL + SUN_LEN);
      g.closePath(); g.fill();
    });
    // 먼지
    state.motes.forEach(function (m) {
      m.v = (m.v + m.sp * 0.016 + state.wind * 0.004) % 1;
      var p = PANES[m.w % PANES.length];
      var yy = GLASS_TOP + 14 + m.v * (WALL + SUN_LEN - GLASS_TOP - 30);
      var k = (yy - GLASS_TOP - 10) / (WALL + SUN_LEN - GLASS_TOP - 10);
      var xx = p[0] + (p[1] - p[0]) * m.u + SUN_DX * k + Math.sin(t * 0.8 + m.ph) * 2;
      var a = 0.55 * Math.max(0, Math.sin(t * 1.2 + m.ph)) * (1 - Math.abs(k - 0.5) * 1.4);
      if (a > 0.02) circ(g, xx, yy, 0.55, 'rgba(255,248,225,' + a.toFixed(2) + ')');
    });
    g.restore();
  }

  function lightsList(t) {
    var L = [];
    Object.keys(DEPTS).forEach(function (k) {
      DEPTS[k].seats.forEach(function (seat) {
        var a = seat.agent;
        if (!a || a.state === 'gone' || a.state === 'entering' && a.moving) return;
        var x = seat.desk.x;
        L.push({ x: x + 12, y: seat.desk.y - 2, r: 30, a: 0.85, col: '255,200,130', ga: 0.2 });   // 스탠드
        if (a.state === 'working' && !a.moving) L.push({ x: seat.x, y: seat.desk.y - 10, r: 18, a: 0.7, col: '150,215,255', ga: 0.16 });
      });
    });
    L.push({ x: BOSS_DESK.x + 16, y: BOSS_DESK.y - 8, r: 26, a: 0.8, col: '150,215,255', ga: 0.16 });
    L.push({ x: BOSS_SEAT.x, y: BOSS_SEAT.y - 10, r: 42, a: 0.65, col: '255,200,150', ga: 0.14 });
    L.push({ x: 48, y: 252, r: 44, a: 0.9, col: '255,205,140', ga: 0.22 });                          // 플로어 스탠드
    L.push({ x: PILLAR_X, y: 60, r: 40, a: 0.6, col: '255,150,210', ga: 0.16 });                      // 네온 달
    L.push({ x: 222, y: 266, r: 14, a: 0.5, col: '170,215,255', ga: 0.1 });                           // 커피 카트
    var e = state.elev.open;
    if (e > 0.05) L.push({ x: ELEV.x + ELEV.w / 2, y: WALL + 4, r: 40 * e, a: 0.85 * e, col: '255,225,170', ga: 0.2 * e });
    Object.keys(DEPTS).forEach(function (k) {
      var z = DEPTS[k].rug;
      L.push({ x: z.x + 30, y: z.y + z.h - 10, r: 26, a: 0.45, col: '124,231,200', ga: 0.12 });
    });
    L.push({ x: BOSS_RUG.x + 26, y: BOSS_RUG.y + 14, r: 24, a: 0.45, col: '255,159,214', ga: 0.12 });
    state.agents.forEach(function (a) {
      if (a.focusUntil > performance.now()) L.push({ x: a.x, y: a.y - 18, r: 22, a: 0.5, col: '255,220,140', ga: 0.2 });
    });
    return L;
  }

  function drawNightLight(t) {
    var s = shade.getContext('2d');
    s.setTransform(S, 0, 0, S, 0, 0);
    s.globalCompositeOperation = 'source-over';
    s.clearRect(0, 0, W, H);
    s.fillStyle = 'rgba(16,13,46,.55)';
    s.fillRect(0, 0, W, H);
    s.globalCompositeOperation = 'destination-out';
    s.fillStyle = 'rgba(0,0,0,.9)';
    WINDOWS.forEach(function (w) { s.fillRect(w.x, GLASS_TOP, w.w, GLASS_BOTTOM - GLASS_TOP); });
    // 창밖 도시 불빛이 바닥에 살짝
    WINDOWS.forEach(function (w) {
      s.fillStyle = lin(s, 0, WALL, 0, WALL + 60, [[0, 'rgba(0,0,0,.16)'], [1, 'rgba(0,0,0,0)']]);
      s.beginPath(); s.moveTo(w.x, WALL); s.lineTo(w.x + w.w, WALL); s.lineTo(w.x + w.w + 22, WALL + 60); s.lineTo(w.x + 22, WALL + 60); s.fill();
    });
    var lights = lightsList(t);
    lights.forEach(function (l) {
      s.fillStyle = rad(s, l.x, l.y, l.r, [[0, 'rgba(0,0,0,' + l.a + ')'], [1, 'rgba(0,0,0,0)']]);
      s.beginPath(); s.arc(l.x, l.y, l.r, 0, TAU); s.fill();
    });
    ctx.drawImage(shade, 0, 0, W, H);

    // 빛의 색
    ctx.save();
    ctx.globalCompositeOperation = 'lighter';
    lights.forEach(function (l, i) {
      var fl = 0.92 + 0.08 * Math.sin(t * 1.7 + i);
      ctx.fillStyle = rad(ctx, l.x, l.y, l.r, [[0, 'rgba(' + l.col + ',' + (l.ga * fl).toFixed(3) + ')'], [1, 'rgba(' + l.col + ',0)']]);
      ctx.beginPath(); ctx.arc(l.x, l.y, l.r, 0, TAU); ctx.fill();
    });
    // 천장 네온 띠가 벽을 물들임
    ctx.fillStyle = lin(ctx, 0, 0, 0, 40, [[0, 'rgba(190,140,255,.22)'], [1, 'rgba(190,140,255,0)']]);
    ctx.fillRect(0, 0, W, 40);
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
    g.drawImage(bg, 0, 0, W, H);

    drawWindows(t, dt);
    drawCurtains(t);
    drawElevator(t);
    drawClock();
    drawNeon(t, false);
    HANGING.forEach(function (h) { drawHanging(g, h, t); });
    drawFloorSigns(t);

    // 깊이 정렬
    var list = [];
    Object.keys(DEPTS).forEach(function (k) {
      DEPTS[k].seats.forEach(function (seat) {
        var sa = seat.agent;
        var seated = sa && !sa.moving && (sa.state === 'working' || sa.state === 'entering' && !sa.path.length);
        list.push({ y: seat.desk.y + 1, draw: function () { drawChair(g, seat, seated); } });
        list.push({ y: seat.desk.y + 12, draw: function () { drawDesk(g, seat, t); } });
      });
    });
    PLANTS.forEach(function (p) { list.push({ y: p.y, draw: function () { drawPlant(g, p, t); } }); });
    list.push({ y: BOSS_DESK.y - 4, draw: function () { drawBossChair(g); } });
    list.push({ y: BOSS_SEAT.y, draw: function () {
      var waiting = state.approvals.some(function (x) { return x.status === 'arrived'; });
      drawPerson(g, { look: BOSS_LOOK, dir: 'down', typing: !waiting, blink: state.boss.blink, seed: 1.3, happy: state.bossHappyUntil > now }, BOSS_SEAT.x, BOSS_SEAT.y, t);
    } });
    list.push({ y: BOSS_DESK.y + 12, draw: function () { drawBossDesk(g, t); } });

    state.agents.forEach(function (a) {
      if (a.alpha <= 0.01) return;
      list.push({
        y: a.y + 0.5, draw: function () {
          g.globalAlpha = clamp(a.alpha, 0, 1);
          if (state.selected === a.id) {
            g.save(); g.translate(a.x, a.y); g.scale(1, 0.32);
            g.beginPath(); g.arc(0, 0, 10, 0, TAU); stroke(g, 1.2, NIGHT ? 'rgba(255,143,200,.8)' : 'rgba(232,132,140,.8)');
            g.restore();
          }
          drawPerson(g, {
            look: a.look, dir: a.dir, moving: a.moving, phase: a.phase, seed: a.seed,
            carrying: a.carrying, typing: a.state === 'working' && !a.moving,
            talking: !!(a.bubble && a.bubble.from < now && a.bubble.until > now), blink: a.blink, happy: a.happyUntil > now
          }, a.x, a.y, t);
          g.globalAlpha = 1;
        }
      });
    });
    list.sort(function (a, b) { return a.y - b.y; });
    list.forEach(function (item) { item.draw(); });

    if (NIGHT) { drawNightLight(t); drawNeon(t, true); }
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
      g.globalAlpha = clamp(a.alpha, 0, 1);
      var headTop = a.y - 36;
      var nh = drawName(a.x, headTop - 9 * TS, a.name, a.state === 'waiting');
      var top = headTop - 9 * TS;
      if (a.state === 'working' && !a.moving) { drawProgress(a.x, top - 4, a.percent); top -= 4; }
      if (a.state === 'waiting') {
        var bounce = Math.abs(Math.sin(t * 3)) * 2;
        circ(g, a.x + 10, headTop + 6 - bounce, 4, '#ff9f8f');
        g.fillStyle = '#fff'; g.font = '6px ' + FONT; g.textAlign = 'center'; g.textBaseline = 'middle';
        g.fillText('!', a.x + 10, headTop + 6.4 - bounce); g.textAlign = 'left';
      }
      if (a.bubble && a.bubble.from < now && a.bubble.until > now) drawBubble(a.x, top, a.bubble.text, a.state === 'waiting' ? { fill: NIGHT ? 'rgba(90,40,80,.95)' : '#fff3ec' } : null);
      g.globalAlpha = 1;
      void nh;
    });
    var waitingCount = state.approvals.filter(function (x) { return x.status === 'arrived'; }).length;
    drawName(BOSS_SEAT.x, BOSS_SEAT.y - 46, '나 (대표)', true);
    if (waitingCount) drawBubble(BOSS_SEAT.x, BOSS_SEAT.y - 47, '결재 ' + waitingCount + '건, 볼게요', { fill: NIGHT ? 'rgba(90,40,80,.95)' : '#fff3ec' });

    // 테마 전환 크로스페이드
    if (fadeAlpha > 0) {
      g.globalAlpha = fadeAlpha;
      g.drawImage(fade, 0, 0, W, H);
      g.globalAlpha = 1;
      fadeAlpha = Math.max(0, fadeAlpha - dt * 0.9);
    }
  }

  var last = performance.now();
  var lastBg = performance.now();
  function loop(now) {
    var dt = Math.min(0.05, (now - last) / 1000);
    last = now;
    if (now - lastBg > 60000) { lastBg = now; }
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
      if (mx > a.x - 9 && mx < a.x + 9 && my > a.y - 36 && my < a.y + 2) hit = a;
    });
    if (hit) {
      if (hit.state === 'waiting' && hit.approval) { openDialog(hit.approval); return; }
      state.selected = hit.id;
      say(hit, hit.state === 'working' ? shorten(hit.task, 14) + ' (' + hit.percent + '%)' : (STATE_LABEL[hit.state] || [''])[0], 2000);
    } else if (Math.abs(mx - BOSS_SEAT.x) < 12 && my > BOSS_SEAT.y - 36 && my < BOSS_SEAT.y) {
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
    var over = state.agents.some(function (a) { return mx > a.x - 9 && mx < a.x + 9 && my > a.y - 36 && my < a.y + 2; });
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
   * 15. 데모
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
        return { id: a.id, name: a.name, dept: a.dept, state: a.state, task: a.task, percent: a.percent, docs: a.docs };
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
