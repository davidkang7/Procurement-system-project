// [홈 로딩 진단] 단계별 서버 시간 측정 — 임시 함수 (2026-10-28). 원인을 확인하면 파일째 삭제한다.
//  사용법: Apps Script 편집기에서 관리자 계정으로 diagHomeTiming() 실행 → 실행 로그 전체를 복사해 전달.
//  - 홈 데이터(getHomeDataForClient)를 그대로 호출해 잰다. 관리자 배너 단계(listPendingQuotePdf)는
//    관리자가 홈을 열 때와 똑같이 견적서 PDF 자동 인식·큐 재등록을 할 수 있다(홈과 같은 동작, 그 이상은 없음).
//  - 단계별 시간은 전역 함수를 이 실행 안에서만 계측 래퍼로 감싸 잰다. 결과는 바꾸지 않고, finally에서 원복한다.
//  - 로그에는 시간·건수·크기만 남긴다(문서 내용 없음).

// 계측 대상 — 홈 경로에서 부르는 전역 함수 (중첩 호출은 각자 합산된다)
var HOME_DIAG_WRAP_ = [
  'getDeptByEmailMap', '_getInspGroupsByPrc', '_getInspMenusForClient', '_isSameDeptEmail',
  '_collectAdminTasks_', 'listPendingQuotePdf', 'reconcileQuotePdf', 'listPoPending', 'listInspAwaitingPdf',
  '_readUserPrefs', 'toDateStr', 'getRowCurrency',
];

function diagHomeTiming() {
  var me = getRequestUserEmail_();
  if (!isAdminUser(me)) throw new Error('관리자만 실행할 수 있습니다. (diagHomeTiming)');

  var L = [];
  var g = globalThis;
  function ms(fn) { var s = Date.now(); fn(); return Date.now() - s; }
  function len(a) { return (a && a.length) || 0; }

  L.push('■ 홈 로딩 진단 ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss') + ' / 실행: ' + me);

  // ── 1) 계측 없이 홈 데이터 전체 3회 — 편집기 실행 기준 순수 서버 시간 ──
  var totals = [], res = null;
  for (var k = 0; k < 3; k++) totals.push(ms(function () { res = getHomeDataForClient(); }));
  L.push('');
  L.push('[1] getHomeDataForClient 전체 (계측 없음) 3회: ' + totals.join(' / ') + ' ms');
  if (!res || !res.ok) {
    L.push('✗ 홈 데이터 실패: ' + (res && res.message));
    console.log(L.join('\n'));
    return;
  }
  var json = '';
  var serMs = ms(function () { json = JSON.stringify(res); });
  var at = res.adminTasks || {};
  L.push('    응답 크기 ' + json.length + '자 (JSON 직렬화 ' + serMs + 'ms)');
  L.push('    건수: myDocs ' + len(res.myDocs) + ' · myPending ' + len(res.myPending) + ' · allPending ' + len(res.allPending)
    + ' · finalDocs ' + len(res.finalDocs) + ' · completedDocs ' + len(res.completedDocs)
    + ' · myInspPending ' + len(res.myInspPending) + ' · allInspPending ' + len(res.allInspPending)
    + ' · approvedInbox ' + len(res.approvedInbox) + ' · myPrc ' + len(res.myPrc)
    + ' · 배너(견적서 ' + len(at.quotePdf) + ' / 주문서 ' + len(at.po) + ' / 검수 ' + len(at.insp) + ')');

  // ── 2) 단계별 — 실제 홈 호출 1회를 계측 래퍼로 감싸서 ──
  var acc = {}, cnt = {}, orig = {}, wrappedTotal = 0;
  try {
    HOME_DIAG_WRAP_.forEach(function (name) {
      if (typeof g[name] !== 'function') return;
      var fn = orig[name] = g[name];
      acc[name] = 0; cnt[name] = 0;
      g[name] = function () {
        var s = Date.now();
        try { return fn.apply(this, arguments); }
        finally { acc[name] += Date.now() - s; cnt[name]++; }
      };
    });
    wrappedTotal = ms(function () { getHomeDataForClient(); });
  } finally {
    Object.keys(orig).forEach(function (name) { g[name] = orig[name]; });
  }
  function line(name, note) {
    return '    ' + name + ': ' + (acc[name] || 0) + 'ms (' + (cnt[name] || 0) + '회)' + (note ? '  ← ' + note : '');
  }
  L.push('');
  L.push('[2] 단계별 (계측 래퍼, 홈 1회 = ' + wrappedTotal + 'ms)');
  if (!cnt.getDeptByEmailMap && !cnt._readUserPrefs) {
    L.push('    ✗ 계측 래퍼가 적용되지 않음 — 이 런타임에서는 전역 함수 가로채기가 안 된다. [3] 개별 측정만 볼 것.');
  } else {
    var top = ['getDeptByEmailMap', '_getInspGroupsByPrc', '_getInspMenusForClient', '_collectAdminTasks_', '_readUserPrefs'];
    var topSum = top.reduce(function (s, n) { return s + (acc[n] || 0); }, 0);
    L.push(line('getDeptByEmailMap', '부서 맵(결재자목록 읽기)'));
    L.push(line('_getInspGroupsByPrc', '검수보고서목록 읽기 #1'));
    L.push(line('_getInspMenusForClient', '검수보고서목록 읽기 #2'));
    L.push(line('_isSameDeptEmail', '위 #2 안 — 일반 사용자 경로'));
    L.push(line('_collectAdminTasks_', '관리자 배너 합계'));
    L.push(line('listPendingQuotePdf', '  └ 견적서 (Y열 JSON 파싱)'));
    L.push(line('reconcileQuotePdf', '      └ PENDING 건 Drive 확인'));
    L.push(line('listPoPending', '  └ 주문서목록 (openById 재호출)'));
    L.push(line('listInspAwaitingPdf', '  └ 검수 PDF 대기 (openById + 검수보고서목록 읽기 #3)'));
    L.push(line('_readUserPrefs', '설정(PropertiesService)'));
    L.push('    그 밖(시트 열기·품의서목록 읽기·메인 루프·정렬): ' + (wrappedTotal - topSum) + 'ms');
    L.push(line('toDateStr', '날짜 변환 — 위 단계들에 걸쳐 합산'));
    L.push(line('getRowCurrency', '통화 추출 — 메인 루프 안'));
  }

  // ── 3) 개별 측정 — 시트 열기·읽기 비용, 같은 시트 반복 읽기가 빨라지는지 ──
  L.push('');
  L.push('[3] 개별 측정');
  var ss = null;
  var open1 = ms(function () { ss = SpreadsheetApp.openById(CONFIG.SHEET_ID); });
  var open2 = ms(function () { SpreadsheetApp.openById(CONFIG.SHEET_ID); });
  L.push('    openById 1차 ' + open1 + 'ms / 2차 ' + open2 + 'ms');

  var mainRows = null;
  [CONFIG.SHEET_NAME, '결재자목록', CONFIG.INSP_SHEET_NAME, CONFIG.PO_SHEET_NAME].forEach(function (name) {
    var sh = null, lr = 0, lc = 0, rows = null;
    var tGet = ms(function () { sh = ss.getSheetByName(name); });
    if (!sh) { L.push('    ' + name + ': 시트 없음'); return; }
    var tDim = ms(function () { lr = sh.getLastRow(); lc = sh.getLastColumn(); });
    var r1 = ms(function () { rows = sh.getDataRange().getValues(); });
    var r2 = ms(function () { sh.getDataRange().getValues(); });
    var r3 = ms(function () { sh.getDataRange().getValues(); });
    if (name === CONFIG.SHEET_NAME) mainRows = rows;
    L.push('    ' + name + ': ' + lr + '행 × ' + lc + '열 = ' + (lr * lc) + '칸 · getSheetByName ' + tGet + 'ms · 크기 확인 ' + tDim
      + 'ms · 전체 읽기 ' + r1 + ' / ' + r2 + ' / ' + r3 + 'ms');
  });

  var deptTimes = [];
  for (var d = 0; d < 3; d++) deptTimes.push(ms(function () { getDeptByEmailMap(ss); }));
  L.push('    getDeptByEmailMap 연속 3회: ' + deptTimes.join(' / ') + 'ms');

  if (mainRows) {
    var dateCells = 0, n = mainRows.length - 1;
    for (var i = 1; i < mainRows.length; i++) {
      [COL.ISSUE_DATE, COL.SUBMIT_AT, COL.CLAIMED_AT].forEach(function (c) { if (mainRows[i][c] instanceof Date) dateCells++; });
    }
    var tDates = ms(function () {
      for (var i = 1; i < mainRows.length; i++) {
        toDateStr(mainRows[i][COL.ISSUE_DATE]);
        toDateTimeStr(mainRows[i][COL.SUBMIT_AT]);
        toDateTimeStr(mainRows[i][COL.CLAIMED_AT]);
      }
    });
    var tCur = ms(function () { for (var i = 1; i < mainRows.length; i++) getRowCurrency(mainRows[i][COL.ITEMS]); });
    var tAttach = ms(function () { for (var i = 1; i < mainRows.length; i++) parseAttachments(mainRows[i][COL.ATTACH_LIST]); });
    L.push('    품의서목록 ' + n + '행 기준: 날짜 변환 ' + tDates + 'ms (Date 셀 ' + dateCells + '개) · 통화 추출 ' + tCur
      + 'ms · 첨부 JSON 파싱 ' + tAttach + 'ms');
  }

  var REP = 200, now = new Date(), tz = '';
  var tTz = ms(function () { for (var i = 0; i < REP; i++) tz = Session.getScriptTimeZone(); });
  var tFmt = ms(function () { for (var i = 0; i < REP; i++) Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm'); });
  L.push('    Session.getScriptTimeZone ' + REP + '회 ' + tTz + 'ms · Utilities.formatDate ' + REP + '회 ' + tFmt + 'ms');

  var tProps = ms(function () { PropertiesService.getScriptProperties().getProperty(_prefsKeyFor(me)); });
  L.push('    ScriptProperties.getProperty 1회 ' + tProps + 'ms');

  var out = L.join('\n');
  console.log(out);
  return out;
}
