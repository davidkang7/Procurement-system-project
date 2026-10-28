// [홈 로딩 진단] 단계별 서버 시간 측정 — 임시 함수 (2026-10-28). 원인을 확인하면 파일째 삭제한다.
//  사용법: Apps Script 편집기에서 관리자 계정으로 diagHomeTiming() 실행 → 실행 로그 전체를 복사해 전달.
//  - 홈 데이터(getHomeDataForClient)를 그대로 호출해 잰다. 관리자 배너 단계(listPendingQuotePdf)는
//    관리자가 홈을 열 때와 똑같이 견적서 PDF 자동 인식·큐 재등록을 할 수 있다(홈과 같은 동작, 그 이상은 없음).
//  - 단계별 시간은 전역 함수를 이 실행 안에서만 계측 래퍼로 감싸 잰다. 결과는 바꾸지 않고, finally에서 원복한다.
//  - 로그에는 시간·건수·크기만 남긴다(문서 내용 없음).
//  v2: [0] 첫 호출(콜드) 단계 분해, [4] 날짜 고속 변환 결과 대조, [5] 스프레드시트 구성 추가.

// 계측 대상 — 홈 경로에서 부르는 전역 함수 (중첩 호출은 각자 합산된다)
var HOME_DIAG_WRAP_ = [
  'getDeptByEmailMap', '_getInspGroupsByPrc', '_getInspMenusForClient', '_isSameDeptEmail',
  '_collectAdminTasks_', 'listPendingQuotePdf', 'reconcileQuotePdf', 'listPoPending', 'listInspAwaitingPdf',
  '_readUserPrefs', 'toDateStr', 'getRowCurrency',
];

// 날짜 고속 변환 후보 — Asia/Seoul(1988년 이후 서머타임 없음, UTC+9 고정)만 순수 계산, 그 밖은 null(기존 방식으로)
function homeDiagFastKst_(d, withTime) {
  var y = d.getUTCFullYear();
  if (isNaN(d.getTime()) || y < 1989 || y > 9998) return null;
  var t = new Date(d.getTime() + 9 * 3600 * 1000);
  function p(n) { return (n < 10 ? '0' : '') + n; }
  var s = t.getUTCFullYear() + '-' + p(t.getUTCMonth() + 1) + '-' + p(t.getUTCDate());
  return withTime ? s + ' ' + p(t.getUTCHours()) + ':' + p(t.getUTCMinutes()) : s;
}

function diagHomeTiming() {
  var c0 = Date.now();
  var me = getRequestUserEmail_();
  var tSession = Date.now() - c0;
  if (!isAdminUser(me)) throw new Error('관리자만 실행할 수 있습니다. (diagHomeTiming)');

  var L = [];
  var g = globalThis;
  function ms(fn) { var s = Date.now(); fn(); return Date.now() - s; }
  function len(a) { return (a && a.length) || 0; }

  // ── 0) 첫 호출(콜드) — 홈과 같은 순서로 시트를 처음 여는 비용. 반드시 맨 먼저 ──
  var ss = null, sh = null, rng = null, rows = null;
  var cold = [];
  cold.push('세션 확인 ' + tSession);
  cold.push('openById ' + ms(function () { ss = SpreadsheetApp.openById(CONFIG.SHEET_ID); }));
  cold.push('품의서목록 getSheetByName ' + ms(function () { sh = ss.getSheetByName(CONFIG.SHEET_NAME); }));
  cold.push('getLastRow ' + ms(function () { sh.getLastRow(); }));
  cold.push('getDataRange ' + ms(function () { rng = sh.getDataRange(); }));
  cold.push('getValues ' + ms(function () { rows = rng.getValues(); }));
  cold.push('| 결재자목록 getSheetByName ' + ms(function () { sh = ss.getSheetByName('결재자목록'); }));
  cold.push('getDataRange ' + ms(function () { rng = sh.getDataRange(); }));
  cold.push('getValues ' + ms(function () { rng.getValues(); }));
  cold.push('| 검수보고서목록 getSheetByName ' + ms(function () { sh = ss.getSheetByName(CONFIG.INSP_SHEET_NAME); }));
  cold.push('getLastRow ' + ms(function () { sh.getLastRow(); }));
  cold.push('getDataRange ' + ms(function () { rng = sh.getDataRange(); }));
  cold.push('getValues ' + ms(function () { rng.getValues(); }));
  cold.push('| 설정 getProperty ' + ms(function () { PropertiesService.getScriptProperties().getProperty(_prefsKeyFor(me)); }));
  cold.push('시간대 조회 ' + ms(function () { Session.getScriptTimeZone(); }));
  var coldTotal = Date.now() - c0;

  L.push('■ 홈 로딩 진단 v2 ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss') + ' / 실행: ' + me);
  L.push('');
  L.push('[0] 첫 호출(콜드) 단계, ms — 합계 ' + coldTotal + 'ms');
  L.push('    ' + cold.join(' · '));
  L.push('    (같은 순서를 한 번 더, 웜) openById ' + ms(function () { SpreadsheetApp.openById(CONFIG.SHEET_ID); })
    + ' · 품의서목록 getDataRange+getValues ' + ms(function () { ss.getSheetByName(CONFIG.SHEET_NAME).getDataRange().getValues(); }));

  // ── 1) 계측 없이 홈 데이터 전체 2회 (웜) ──
  var totals = [], res = null;
  for (var k = 0; k < 2; k++) totals.push(ms(function () { res = getHomeDataForClient(); }));
  L.push('');
  L.push('[1] getHomeDataForClient 전체 (계측 없음, 웜) 2회: ' + totals.join(' / ') + ' ms');
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

  // ── 2) 단계별 — 실제 홈 호출 1회를 계측 래퍼로 감싸서 (웜) ──
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
  L.push('[2] 단계별 (계측 래퍼, 웜, 홈 1회 = ' + wrappedTotal + 'ms)');
  if (!cnt.getDeptByEmailMap && !cnt._readUserPrefs) {
    L.push('    ✗ 계측 래퍼가 적용되지 않음 — 이 런타임에서는 전역 함수 가로채기가 안 된다. [3] 개별 측정만 볼 것.');
  } else {
    var top = ['getDeptByEmailMap', '_getInspGroupsByPrc', '_getInspMenusForClient', '_collectAdminTasks_', '_readUserPrefs'];
    var topSum = top.reduce(function (s, n) { return s + (acc[n] || 0); }, 0);
    L.push(line('getDeptByEmailMap', '부서 맵(결재자목록 읽기)'));
    L.push(line('_getInspGroupsByPrc', '검수보고서목록 읽기 #1'));
    L.push(line('_getInspMenusForClient', '검수보고서목록 읽기 #2'));
    L.push(line('_collectAdminTasks_', '관리자 배너 합계'));
    L.push(line('listPendingQuotePdf', '  └ 견적서'));
    L.push(line('reconcileQuotePdf', '      └ PENDING 건 Drive 확인'));
    L.push(line('listPoPending', '  └ 주문서목록'));
    L.push(line('listInspAwaitingPdf', '  └ 검수 PDF 대기(검수보고서목록 읽기 #3)'));
    L.push(line('_readUserPrefs', '설정'));
    L.push('    그 밖(시트 열기·품의서목록 읽기·메인 루프·정렬): ' + (wrappedTotal - topSum) + 'ms');
    L.push(line('toDateStr', '날짜 변환 — 위 단계들에 걸쳐 합산'));
  }

  // ── 3) 개별 측정 (웜) ──
  L.push('');
  L.push('[3] 개별 측정 (웜)');
  var mainRows = rows, inspRows = null;
  [CONFIG.SHEET_NAME, '결재자목록', CONFIG.INSP_SHEET_NAME, CONFIG.PO_SHEET_NAME].forEach(function (name) {
    var s2 = ss.getSheetByName(name), lr = 0, lc = 0, rv = null;
    if (!s2) { L.push('    ' + name + ': 시트 없음'); return; }
    var tDim = ms(function () { lr = s2.getLastRow(); lc = s2.getLastColumn(); });
    var r1 = ms(function () { rv = s2.getDataRange().getValues(); });
    var r2 = ms(function () { s2.getDataRange().getValues(); });
    if (name === CONFIG.INSP_SHEET_NAME) inspRows = rv;
    L.push('    ' + name + ': ' + lr + '행 × ' + lc + '열 · 크기 확인 ' + tDim + 'ms · 전체 읽기 ' + r1 + ' / ' + r2 + 'ms');
  });
  var REP = 200, now = new Date(), tz = '';
  var tTz = ms(function () { for (var i = 0; i < REP; i++) tz = Session.getScriptTimeZone(); });
  var tFmt = ms(function () { for (var i = 0; i < REP; i++) Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm'); });
  L.push('    Session.getScriptTimeZone ' + REP + '회 ' + tTz + 'ms · Utilities.formatDate ' + REP + '회 ' + tFmt + 'ms · 시간대 "' + tz + '"');

  // ── 4) 날짜 고속 변환 — 운영 데이터 전수 대조 (두 시트의 모든 Date 셀, 두 형식) ──
  var dates = [];
  [mainRows, inspRows].forEach(function (rs) {
    if (!rs) return;
    for (var i = 1; i < rs.length; i++) for (var j = 0; j < rs[i].length; j++) if (rs[i][j] instanceof Date) dates.push(rs[i][j]);
  });
  var mism = [], fallback = 0, slowMs = 0, fastMs = 0, slowOut = [], fastOut = [];
  slowMs = ms(function () {
    var z = Session.getScriptTimeZone();
    dates.forEach(function (d) { slowOut.push(Utilities.formatDate(d, z, 'yyyy-MM-dd'), Utilities.formatDate(d, z, 'yyyy-MM-dd HH:mm')); });
  });
  fastMs = ms(function () {
    dates.forEach(function (d) { fastOut.push(homeDiagFastKst_(d, false), homeDiagFastKst_(d, true)); });
  });
  for (var q = 0; q < slowOut.length; q++) {
    if (fastOut[q] === null) { fallback++; continue; }
    if (fastOut[q] !== slowOut[q] && mism.length < 5) mism.push(slowOut[q] + ' ≠ ' + fastOut[q]);
    else if (fastOut[q] !== slowOut[q]) mism.push('…');
  }
  L.push('');
  L.push('[4] 날짜 고속 변환 대조 (품의서목록+검수보고서목록의 Date 셀 ' + dates.length + '개 × 2형식)');
  L.push('    기존(formatDate) ' + slowMs + 'ms · 순수 계산 ' + fastMs + 'ms · 불일치 ' + mism.length + '건 · 기존 방식 유지(범위 밖) ' + fallback + '건'
    + (mism.length ? '\n    불일치 예: ' + mism.slice(0, 5).join(' / ') : ''));

  // ── 5) 스프레드시트 구성 — 여는 비용이 문서 전체 크기에 좌우되는지 보기 위함 ──
  var sheets = ss.getSheets(), totalCells = 0, totalGrid = 0, parts = [];
  sheets.forEach(function (s3) {
    var lr = s3.getLastRow(), lc = s3.getLastColumn(), mr = s3.getMaxRows(), mc = s3.getMaxColumns();
    totalCells += lr * lc; totalGrid += mr * mc;
    parts.push(s3.getName() + ' ' + lr + '×' + lc + ' (격자 ' + mr + '×' + mc + ')');
  });
  L.push('');
  L.push('[5] 스프레드시트 구성: 시트 ' + sheets.length + '개 · 데이터 ' + totalCells + '칸 · 격자 ' + totalGrid + '칸');
  L.push('    ' + parts.join(' · '));

  var out = L.join('\n');
  console.log(out);
  return out;
}
