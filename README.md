# dlsrks091412-afk.github.io

이 저장소에는 두 가지가 들어 있습니다.

| 경로 | 내용 | 주소 (GitHub Pages) |
|---|---|---|
| `index.html` | 기존 영어 수행평가 앱 (**건드리지 않았어요**) | `https://dlsrks091412-afk.github.io/` |
| `portal/` | **새 개인 사이트** — 로그인 → 홈 → AI 에이전트 사무실 | `https://dlsrks091412-afk.github.io/portal/` |

> 수행평가 페이지를 학생들이 쓰고 있을 수 있어서 루트 `index.html`은 그대로 두고, 개인 사이트는 `portal/` 폴더에 만들었습니다.
> 개인 사이트를 루트 주소로 옮기고 싶다면 아래 [개인 사이트를 메인 주소로 옮기기](#개인-사이트를-메인-주소로-옮기기)를 보세요.

---

## 1. 파일 구조

```
portal/
├── index.html              ← 사이트 메인 = 로그인 화면
├── home.html               ← 로그인 후 홈 (여러 공간으로 가는 타일)
├── office.html             ← AI 에이전트 사무실
├── assets/
│   ├── ui.js               ← 낮/밤 테마 + 부드러운 페이지 전환 (모든 페이지 공통)
│   ├── auth-config.js      ← 비밀번호 검증 정보 (비밀번호 자체는 들어 있지 않음)
│   ├── auth.js             ← 로그인/세션/잠금 로직
│   ├── guard.js            ← 보호 페이지 공통: 로그인 확인, 로그아웃, 시계
│   ├── login.js / login.css
│   ├── home.css / home.js
│   ├── office-config.js    ← 사무실 설정 (이벤트 피드 주소, 데모 자동 재생)
│   ├── office.js           ← 사무실 엔진 (일러스트·조명·이동·결재 전부)
│   ├── office.css
│   ├── portal.css          ← 공통 스타일
│   └── logo.svg
├── tools/
│   └── set-password.html   ← 비밀번호 바꾸는 도구 (브라우저 안에서만 계산)
└── data/
    └── sample-feed.json    ← 에이전트 이벤트 예시
```

빌드 과정이 없는 순수 HTML/CSS/JS라서, 파일을 올리기만 하면 바로 동작합니다.

---

## 2. 사이트 올리기 (GitHub Pages)

1. **작업 브랜치를 `main`에 합치기**
   지금 코드는 `claude/brave-feynman-8yde3b` 브랜치에 있습니다. GitHub 저장소 페이지에서
   **Pull requests → New pull request → base: `main` ← compare: `claude/brave-feynman-8yde3b` → Merge** 하면 됩니다.
2. **GitHub Pages 켜기** (이미 켜져 있다면 건너뛰기)
   저장소 **Settings → Pages → Build and deployment**
   - Source: `Deploy from a branch`
   - Branch: `main` / `/ (root)` → **Save**
3. 1~2분 뒤 `https://dlsrks091412-afk.github.io/portal/` 로 접속 → 로그인 화면이 나오면 성공입니다.

### 내 컴퓨터에서 미리 보기

`file://`로 열면 일부 기능(암호화 API)이 막힐 수 있어서, 간단한 로컬 서버로 여는 게 좋아요.

```bash
cd dlsrks091412-afk.github.io
```

```bash
python3 -m http.server 8000
```

브라우저에서 `http://localhost:8000/portal/` 를 엽니다.

---

## 3. 로그인 & 비밀번호

- 첫 비밀번호(임시)는 **이 README가 아니라 채팅으로 따로 전달**했습니다. 공개 저장소라서 비밀번호를 파일에 적지 않았어요.
- **처음 접속하면 꼭 비밀번호를 바꾸세요.**

### 비밀번호 바꾸는 법

1. `https://dlsrks091412-afk.github.io/portal/tools/set-password.html` (또는 로컬 서버에서 같은 경로) 열기
2. 새 비밀번호 입력 → **설정 파일 만들기** → **복사하기**
3. `portal/assets/auth-config.js` 내용을 복사한 내용으로 **통째로 교체** → 커밋 & 푸시
4. 1~2분 뒤 적용되고, 기존 로그인 세션은 모두 자동으로 풀립니다.

### 보안 기능

| 기능 | 설명 |
|---|---|
| 비밀번호 비저장 | 저장소에는 비밀번호도, 단순 해시도 없습니다. PBKDF2(SHA-256, 60만 회)로 만든 키로 암호화한 "검증 문장"만 있어서, 무작위 대입 공격이 매우 느립니다. |
| 실패 잠금 | 5회 틀리면 30초 → 1분 → 2분 … 최대 15분까지 잠깁니다. |
| 세션 만료 | 기본 2시간 + 탭을 닫으면 로그아웃. "로그인 유지"를 체크하면 그 기기에서만 7일. |
| 페이지 보호 | `home.html`, `office.html`은 로그인 안 하면 로그인 화면으로 돌려보냅니다. |
| 오픈 리다이렉트 방지 | 로그인 후 이동 주소는 같은 폴더의 `.html`만 허용합니다. |
| CSP | 모든 페이지에 Content-Security-Policy를 걸어 외부 스크립트·인라인 스크립트 실행을 막습니다. |
| 검색 노출 차단 | `noindex, nofollow`로 검색엔진에 뜨지 않게 했습니다. |

### ⚠ 꼭 알아야 할 한계

GitHub Pages는 **정적 호스팅**이라 서버에서 접근을 막을 수 없습니다.
즉 이 로그인은 "화면에 들어오지 못하게" 막는 것이지, 저장소에 올린 HTML·JS·JSON 파일 자체는 공개 저장소에서 누구나 볼 수 있습니다.

- **절대 저장소에 넣지 말 것**: API 키, 토큰, 개인 문서 내용, 에이전트가 만든 실제 업무 데이터
- 진짜 비공개 데이터가 생기면 아래 중 하나를 추천합니다.
  - **Cloudflare Pages + Cloudflare Access** (무료 플랜 가능): 사이트 앞단에서 이메일 인증을 받아 파일 자체를 보호
  - 데이터는 **로그인이 필요한 백엔드**(Supabase, Firebase, Cloudflare Workers 등)에 두고 사무실 화면은 그걸 읽기만 하기
  - 저장소를 **Private**으로 바꾸기 (GitHub Pages를 비공개로 쓰려면 유료 플랜 필요)

---

## 4. AI 에이전트 사무실 사용법

`office.html` 에서 볼 수 있는 흐름:

에이전트들은 셔츠를 입은 **강아지 직원**으로 나와요. 대표 자리의 "나"는 사람 캐릭터입니다.

| 견종 키 (`breed`) | 견종 | 특징 |
|---|---|---|
| `pomeranian` | 포메라니안 | 주황빛 솜털, 풍성한 목털 |
| `samoyed` | 사모예드 | 새하얀 솜털, 웃는 얼굴 |
| `shiba` | 시바견 | 하얀 볼, 말린 꼬리 |
| `corgi` | 웰시코기 | 큰 귀, 짧은 다리, 동그란 꼬리 |
| `maltese` | 말티즈 | 긴 귀, 분홍 리본 |
| `poodle` | 푸들 | 곱슬곱슬 갈색 털 |
| `beagle` | 비글 | 진갈색 늘어진 귀, 흰 줄무늬 |
| `husky` | 허스키 | 회색 털, 흰 얼굴 무늬 |
| `golden` | 골든 리트리버 | 금빛 털, 늘어진 귀 |

견종을 정하지 않으면 에이전트 id에 따라 자동으로 골라지고, 같은 에이전트는 항상 같은 견종으로 나와요. 걸어 다닐 때는 바닥에 발자국이 남았다가 사라지고, 기분이 좋으면 꼬리를 더 빨리 흔들어요.

1. **출근** — 42층 엘리베이터 문이 열리고 에이전트가 "출근했습니다~" 하며 걸어 나옴
2. **착석 & 업무** — 자기 부서(지금은 **기록부**) 빈 자리에 앉아 타이핑. 스탠드가 켜지고 머그컵에서 김이 남. 문서가 완성될 때마다 `+1`이 떠오르고 책상 위 서류가 쌓임. 진행률 바, 가끔 "집중 모드"(반짝이 효과)
3. **결재 요청** — 결재 파일(노란 폴더)을 들고 **대표실 내 자리 앞**으로 걸어옴 → 도착하면 결재 서류 창이 자동으로 열림
4. **승인/반려** — 도장이 찍히고(승인이면 하트가 퐁퐁), 에이전트가 반응한 뒤 자리로 돌아가 업무 계속 (반려면 서류를 다시 들고 감)
5. **퇴근** — "퇴근할게요~" 하고 엘리베이터로 나감

### 낮 / 밤 테마

오른쪽 위 테마 버튼을 누를 때마다 **자동 → 낮 → 밤** 순서로 바뀌고, 선택은 그 브라우저에 기억됩니다. "자동"은 06~18시는 낮, 그 외에는 밤이에요. 사이트 전체(로그인·홈·사무실)가 함께 바뀝니다.

| | 낮 | 밤 |
|---|---|---|
| 창밖 | 파스텔 하늘, 흘러가는 구름, 가끔 날아가는 새 | 도시 야경, 반짝이는 별·달, 깜빡이는 항공등, 가끔 별똥별 |
| 실내 | 창으로 들어오는 햇살과 떠다니는 먼지 | 어두운 실내에 스탠드·모니터·플로어 스탠드 불빛 |
| 움직임 | 10~25초마다 바람이 불어 커튼이 펄럭이고 꽃잎이 날아들어옴, 식물이 흔들림 | 천장 네온 띠·달 모양 네온·부서 네온 간판이 은은하게 빛남 |

식물과 매달린 화분은 항상 살랑살랑 흔들리고, 바람이 불면 더 크게 흔들립니다. 데모의 **창문 바람** 버튼으로 바로 볼 수 있어요.

### 페이지 전환

크롬·엣지·사파리 최신 버전에서는 브라우저의 View Transitions 기능으로 페이지가 부드럽게 흐려졌다 나타나고, 그 밖의 브라우저에서는 페이드 인/아웃으로 대신합니다. 운영체제에서 "동작 줄이기"를 켜 두면 애니메이션이 꺼집니다.

화면 아래 **데모 버튼**으로 직접 해볼 수 있고, 처음 열면 데모 시나리오가 한 번 자동 재생됩니다.
(`assets/office-config.js`의 `autoDemo: false`로 끌 수 있어요.)
결재 대기 중인 에이전트나 내 캐릭터를 클릭해도 서류가 열리고, **브라우저 알림 켜기**를 누르면 다른 탭을 보고 있을 때 알림이 옵니다.

### 4-1. 실제 에이전트와 연결하기

사무실은 **이벤트**를 받아서 움직입니다. 방법은 두 가지예요.

**① 이벤트 피드(JSON 주소) 읽기** — `assets/office-config.js`

```js
window.OFFICE_CONFIG = {
  feedUrl: 'https://내-서버/events',   // 주기적으로 GET
  decisionUrl: 'https://내-서버/decisions', // 승인/반려 결과를 POST
  pollMs: 5000,
  autoDemo: false
};
```

> 외부 주소를 쓰면 `office.html`의 CSP 줄에서 `connect-src 'self'` 뒤에 그 주소를 추가해야 합니다.
> 예: `connect-src 'self' https://my-office-api.example.workers.dev`

먼저 저장소에 들어 있는 샘플로 테스트해 보세요: `feedUrl: 'data/sample-feed.json'`

**② 자바스크립트 API** — 브라우저 콘솔이나 다른 스크립트에서

```js
Office.checkIn({ id: 'writer-1', name: '포미', breed: 'pomeranian', dept: 'records', task: '회의록 정리' });
Office.progress('writer-1', { percent: 60, docs: 3, task: '요약본 작성' });
Office.requestApproval('writer-1', { title: '보고서 배포', detail: '노션에 올려도 될까요?' })
      .then(r => console.log(r.decision, r.comment));   // 'approved' | 'rejected'
Office.checkOut('writer-1');
Office.onDecision(r => console.log('결재 결과', r));
```

### 4-2. 이벤트 형식

피드는 배열 또는 `{ "events": [...] }` 형태로 돌려주면 됩니다. 같은 `eventId`는 한 번만 처리돼요.

| type | 필드 | 동작 |
|---|---|---|
| `checkin` | `agentId`, `name`, `dept`(`records`), `task`, `breed`(선택, 위 견종 키) | 출근해서 부서 자리로 이동 |
| `progress` | `agentId`, `percent`(0~100), `docs`(누적 문서 수), `task` | 진행률/업무 갱신 (25% 이상 오르면 불꽃) |
| `approval` | `agentId`, `approvalId`, `title`, `detail` | 결재 파일 들고 내 자리로 |
| `checkout` | `agentId` | 하던 일(결재 등) 마치고 퇴근 |
| `say` | `agentId`, `text` | 말풍선 띄우기 |

공통 선택 필드: `eventId`(중복 방지), `delayMs`(피드를 처음 읽은 시점부터 지연 실행 — 테스트용)

`decisionUrl`로 보내는 결과:

```json
{ "approvalId": "ap-001", "agentId": "writer-1", "decision": "approved", "comment": "좋아요", "decidedAt": "2026-11-18T00:38:14.421Z" }
```

### 4-3. 새 부서 추가하기 (다음 단계용)

`assets/office.js` 위쪽의 `DEPTS`에 항목을 추가하면 됩니다. 사무실은 아이소메트릭(비스듬히 내려다보는) 시점이고, 바닥은 가로 12 × 세로 10칸 타일로 되어 있어요. 앞쪽 오른편의 "새 부서가 들어올 자리"(`FUTURE_ZONE`, x 8.3~11.6 / y 5.4~9.4, 지금은 이삿짐 상자가 놓여 있음)가 비어 있습니다.

```js
research: {
  name: '조사부',
  desc: '자료 조사 담당',
  x0: 8.6, y0: 6.4, deskW: 1.4, deskD: 0.9, cols: 2,   // 책상 묶음 시작 위치(타일)와 크기, 한 줄 책상 수
  sign: { x: 8.2, y: 7.3 },                            // 부서 팻말 위치
  neon: '#9fb8ff',                                     // 밤에 빛나는 팻말 색
  chair: '#b4c0ea',
  lines: ['자료 찾는 중…', '출처 확인!'],
  unit: '자료'
}
```

책상은 두 줄로 마주 보게 놓이고(뒷줄은 이쪽을 보고 앉고, 앞줄은 등을 보이며 앉아요), 좌석과 통로는 자동으로 계산됩니다.
추가할 때는 `FUTURE_ZONE` 점선 그리기를 지우거나 위치를 옮기고, 새 부서 자리로 가는 길(통로)이 다른 가구와 겹치지 않는지 확인하세요. 원하면 다음 부서도 같이 만들어 드릴게요.

---

## 개인 사이트를 메인 주소로 옮기기

`https://dlsrks091412-afk.github.io/` 로 바로 로그인 화면이 뜨게 하려면:

1. 수행평가 `index.html`을 다른 폴더로 옮기기 (예: `assessment/index.html`) — **학생들이 쓰는 주소가 바뀌니 먼저 공지하세요.**
2. `portal/` 안의 파일들을 저장소 루트로 옮기기 (`portal/index.html` → `index.html`, `portal/assets` → `assets` …)
3. 커밋 & 푸시

파일 사이 경로는 모두 상대 경로라서 폴더째 옮기면 그대로 동작합니다.

---

## 사용한 외부 리소스

- 글꼴 [Pretendard](https://github.com/orioncactus/pretendard) (jsDelivr), [고운돋움 Gowun Dodum](https://fonts.google.com/specimen/Gowun+Dodum) (Google Fonts) — 둘 다 SIL Open Font License
- 사무실·로그인 배경 그림과 캐릭터는 모두 코드(캔버스/SVG)로 직접 그린 오리지널 일러스트입니다.
