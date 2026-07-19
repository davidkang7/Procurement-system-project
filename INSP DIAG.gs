// [결재 메뉴 세분화 진단] 서버 페이로드 검증 — 임시 함수
// Apps Script 편집기에서 이 함수를 실행하고 '실행 로그'를 확인할 것.
// (getHomeDataForClient는 값을 return만 하므로 편집기에서 직접 실행해도 로그가 비어 있음)
function debugPendingMeta() {
  var res = getHomeDataForClient();
  var L = [];

  if (!res || !res.ok) {
    console.log('✗ getHomeDataForClient 실패: ' + (res && res.message ? res.message : '(응답 없음)'));
    return;
  }

  var u = res.user || {};
  L.push('▶ 로그인: ' + u.email + '  / 부서: ' + (u.dept || '(없음)'));
  L.push('▶ isAdmin=' + u.isAdmin + '  isProcurement=' + u.isProcurement);
  L.push('  (isAdmin=false 인데 관리자여야 한다면 CONFIG.ADMIN_EMAILS 확인)');
  L.push('');

  L.push('▶ 배열 건수');
  L.push('   myPending      = ' + (res.myPending || []).length + '   ← 기존 "결재 대기" 숫자와 비교. 줄었으면 문제');
  L.push('   allPending     = ' + (res.allPending || []).length + '   ← 관리자만 채워짐. 비관리자는 0이어야 정상');
  L.push('   myInspPending  = ' + (res.myInspPending || []).length);
  L.push('   allInspPending = ' + (res.allInspPending || []).length);
  L.push('');

  // 새 meta 필드가 실제로 붙어 나오는지 + 메뉴 배지가 어떻게 갈릴지 미리보기
  function dump(title, arr) {
    arr = arr || [];
    L.push('════ ' + title + ' (' + arr.length + '건) ════');
    if (arr.length === 0) { L.push('   (없음)'); L.push(''); return; }

    var missing = 0, buckets = {};
    arr.forEach(function (x) {
      if (typeof x.stageIdx !== 'number') missing++;
      var kind = (typeof x.stageIdx !== 'number') ? '(stageIdx없음)'
               : (x.stageIdx === 0 ? '기안자 결재' : '결재자 결재');
      var key = (x.docType || '?') + ' / ' + kind;
      buckets[key] = (buckets[key] || 0) + 1;
    });

    L.push('   ─ 메뉴 배지 미리보기 (문서종류 / 결재성격) ─');
    Object.keys(buckets).sort().forEach(function (k) { L.push('     ' + k + ' : ' + buckets[k] + '건'); });
    if (missing > 0) {
      L.push('   ⚠ stageIdx 없는 항목 ' + missing + '건 — 서버 배포가 반영되지 않았을 수 있음');
    }

    L.push('   ─ 상위 5건 상세 ─');
    arr.slice(0, 5).forEach(function (x) {
      L.push('     [' + (x.docType || '?') + '] ' + (x.docNo || '') + ' "' + (x.subject || '') + '"');
      L.push('        stageIdx=' + x.stageIdx + ' stageKind=' + x.stageKind
           + ' / 현재결재자: ' + (x.curApprName || '') + ' <' + (x.curApprEmail || '') + '>'
           + ' / apprCount=' + x.apprCount);
    });
    L.push('');
  }

  dump('myPending (내 품의·구매 결재 대기)', res.myPending);
  dump('myInspPending (내 검수보고서 결재 대기)', res.myInspPending);
  if (u.isAdmin) {
    dump('allPending (전사 품의·구매)', res.allPending);
    dump('allInspPending (전사 검수보고서)', res.allInspPending);
    var okSuperset = (res.allPending || []).length >= (res.myPending || []).length;
    L.push(okSuperset
      ? '✓ allPending >= myPending (상위집합 정상)'
      : '✗ allPending < myPending — 상위집합이 아님. 수집 조건 확인 필요');
  } else {
    var leak = (res.allPending || []).length + (res.allInspPending || []).length;
    L.push(leak === 0
      ? '✓ 비관리자에게 전사 데이터 미노출 (권한 백스톱 정상)'
      : '✗ 비관리자인데 전사 데이터가 ' + leak + '건 넘어옴 — 권한 게이트 확인 필요');
  }

  console.log(L.join('\n'));
}

// [INSP 진단 Step5] 검수 결재 대기 미표시 원인 추적 — 임시 함수
function debugInspPending() {
  var actor = (getActiveUserEmail() || '').toLowerCase();
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.INSP_SHEET_NAME);
  var L = [];
  L.push('▶ 로그인(actor): "' + actor + '"');
  L.push('▶ isProcurementUser: ' + isProcurementUser(getActiveUserEmail()));

  if (!sheet) { L.push('✗ 검수보고서목록 시트 없음'); console.log(L.join('\n')); return; }
  L.push('▶ 시트 마지막 행: ' + sheet.getLastRow() + ' / 컬럼: ' + sheet.getLastColumn());
  L.push('▶ INSP_TOTAL_COLS 기대: ' + INSP_TOTAL_COLS + ' / MAX_APPROVERS: ' + INSP_COL.MAX_APPROVERS);

  if (sheet.getLastRow() < 2) { L.push('(검수보고서 데이터 없음)'); console.log(L.join('\n')); return; }
  var rows = sheet.getDataRange().getValues();
  L.push('────────────────────────────────');
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var docNo  = String(r[INSP_COL.DOC_NO] || '');
    var status = String(r[INSP_COL.STATUS] || '');
    var apprCount = parseInt(r[INSP_COL.APPR_COUNT]) || 0;
    var curIdx = parseInt(r[INSP_COL.APPR_IDX]) || 0;
    L.push('[' + docNo + '] status="' + status + '" apprCount=' + apprCount + ' curIdx=' + curIdx);

    // 결재 블록 덤프
    for (var a = 0; a < apprCount; a++) {
      var nm = String(r[inspApprCol(a, 1)] || '');
      var em = String(r[inspApprCol(a, 2)] || '');
      var st = String(r[inspApprCol(a, 3)] || '');
      var mark = (a === curIdx) ? ' ←현재차례' : '';
      var isMe = (em.toLowerCase() === actor) ? ' [=나]' : '';
      L.push('   블록' + a + ': ' + nm + ' <' + em + '> status="' + st + '"' + mark + isMe);
    }

    // 대기 판정 시뮬레이션
    var myIdx = -1;
    for (var b = 0; b < apprCount; b++) {
      if (String(r[inspApprCol(b, 2)] || '').toLowerCase() === actor) { myIdx = b; break; }
    }
    var statusOk = (status === '검토중' || status.indexOf('결재중') >= 0);
    var myStatus = (myIdx >= 0) ? String(r[inspApprCol(myIdx, 3)] || '') : '(나 없음)';
    var pass = (myIdx >= 0 && myIdx === curIdx && statusOk && myStatus === '대기');
    L.push('   → myIdx=' + myIdx + ' / statusOk=' + statusOk + ' / myStatus="' + myStatus + '" / 대기표시=' + (pass ? 'YES ✓' : 'NO ✗'));
    L.push('');
  }
  console.log(L.join('\n'));
}
