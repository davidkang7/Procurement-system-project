// [홈 로딩 진단 v3] 실제 웹 호출(google.script.run) 안의 단계별 시간 — 임시 (2026-10-29). 확인 후 파일째 삭제하고
//  getHomeDataForClient의 __tm 줄들도 되돌린다.
//  로그 한 줄: '[홈 타이밍] 함수 안 총 Nms · 단계 ms · … · 응답 N자' — 시간·크기만 남긴다(문서 내용 없음).
//  실행수 화면의 소요 시간과 비교하면 '함수 밖'(실행 시작·응답 변환)이 얼마인지 나온다.

// 이 파일이 로드된 시각 — 함수 시작까지의 간격이 크면 스크립트 로드(다른 파일 포함)가 무겁다는 뜻
var HOME_TIMING_LOADED_AT_ = Date.now();

function _homeTiming_() {
  var t0 = Date.now(), last = t0, parts = [];
  return {
    mark: function (label) {
      var now = Date.now();
      parts.push(label + ' ' + (now - last));
      last = now;
    },
    log: function (res) {
      try {
        var json = '', s = Date.now();
        try { json = JSON.stringify(res); } catch (_) {}
        console.log('[홈 타이밍] 함수 안 총 ' + (Date.now() - t0) + 'ms · 로드→시작 ' + (t0 - HOME_TIMING_LOADED_AT_) + 'ms · '
          + parts.join(' · ') + ' · 응답 ' + json.length + '자(직렬화 ' + (Date.now() - s) + 'ms)');
      } catch (_) { /* 진단 로그 실패가 홈을 막으면 안 된다 */ }
    },
  };
}
