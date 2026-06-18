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
