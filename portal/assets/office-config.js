/*
 * AI 사무실 설정
 * ---------------------------------------------------------------
 * feedUrl     : 에이전트 이벤트(JSON)를 주기적으로 읽어올 주소.
 *               비워 두면 외부 연결 없이 화면의 데모 버튼/콘솔 API로만 움직입니다.
 *               예) 'data/sample-feed.json'  (저장소에 포함된 샘플)
 *               예) 'https://my-office-api.example.workers.dev/events'  (직접 만든 서버)
 * decisionUrl : 승인/반려 결과를 POST로 보낼 주소. 비워 두면 브라우저 안에서만 처리.
 * pollMs      : feedUrl 확인 간격(밀리초).
 * autoDemo    : true면 페이지를 열 때 데모 시나리오가 자동으로 재생됩니다.
 *
 * ⚠ 이 파일은 공개 저장소에 올라가므로 API 키·토큰 같은 비밀값은 절대 넣지 마세요.
 * ⚠ 외부 주소를 쓰면 office.html 의 Content-Security-Policy connect-src 에 그 도메인을 추가해야 합니다.
 */
window.OFFICE_CONFIG = {
  feedUrl: '',
  decisionUrl: '',
  pollMs: 5000,
  autoDemo: true
};
