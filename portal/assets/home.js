(function () {
  'use strict';
  var h = new Date().getHours();
  var big = h < 5 ? '늦은 밤이에요, 오늘도 수고 많으셨어요.'
    : h < 11 ? '좋은 아침이에요, 대표님.'
    : h < 17 ? '따뜻한 오후예요, 대표님.'
    : h < 21 ? '노을이 예쁜 저녁이에요.'
    : '고요한 밤이에요, 대표님.';
  document.getElementById('greet').textContent = big;
  document.getElementById('greetSmall').textContent = new Date().toLocaleDateString('ko-KR', { month: 'long', day: 'numeric', weekday: 'long' });
})();
