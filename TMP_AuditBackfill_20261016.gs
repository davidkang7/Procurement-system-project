// ================================================================
// [임시] 시스템로그 누락 1건 보정 — TO-PO-26-318 기안자 승인 (2026-10-15 16:55:50)
//  GAS 편집기는 인자를 넘길 수 없어서 무인자 래퍼를 둔다. 실행 후 이 파일째 제거하고 다시 push 한다.
//
//  실행 순서 (GAS 편집기, 관리자 계정):
//    1) _tmp_auditBackfill_20261016_dryRun — 쓰기 없이 검증하고, 넣을 행만 출력
//    2) _tmp_auditBackfill_20261016        — 실제 기록 (다시 돌려도 안전: 이미 있으면 SKIP)
//
//  원인: 결재는 시트에 반영됐는데 withLock의 releaseLock()이 Google 서버 오류를 던져
//        락 밖의 DOC_APPROVE 기록이 실행되지 않았다 (Version 106 실행 16:55:48 Failed, 16:55:56 오류).
//  넣는 행: 그때 남았어야 할 DOC_APPROVE와 같은 형식. 처리자·시각은 품의서목록 결재란 0번 값 그대로,
//           사유 끝에 '(사후 보정)', payload.backfill에 보정자·원인. 보정 자체는 ADMIN_AUDIT_BACKFILL로 따로 남긴다.
//  두 행 모두 시트 맨 아래에 붙는다 — 시스템로그는 timestamp 열 기준으로 정렬해서 볼 것(§3-1-1 ⑤).
// ================================================================

var _TMP_AUDIT_BACKFILL_20261016 = {
  token:      '94cbcfaa-2797-4d64-9637-cb9296bc621b',
  docNo:      'TO-PO-26-318',
  docType:    'PRC',
  stageIdx:   0,
  approver:   'davidkang@inlct.com',
  approvedAt: '2026-10-15T16:55:50.564+09:00',   // 결재란 0번 처리시각 (±1초 안이면 시트 값을 그대로 쓴다)
  cause:      'withLock releaseLock() Google 서버 오류 — 2026-10-15 16:55:56, Version 106 processDecisionFromClient',
};

function _tmp_auditBackfill_20261016_dryRun() { return _tmpAuditBackfill_(true); }
function _tmp_auditBackfill_20261016()        { return _tmpAuditBackfill_(false); }

function _tmpAuditBackfill_(dryRun) {
  // 인가 전용 신원 — getActiveUserEmail()은 신원을 못 잡으면 배포 계정(관리자)으로 폴백하므로 쓰지 않는다
  var actor = getRequestUserEmail_();
  if (!isAdminUser(actor)) throw new Error('관리자만 실행할 수 있습니다: ' + (actor || '(신원 확인 불가)'));
  var J = _TMP_AUDIT_BACKFILL_20261016;
  var tag = dryRun ? '[DRY RUN] ' : '';

  // ── 1) 결재란 대조: 정말 승인됐고, 처리자·시각이 예상과 같은가 ──
  var ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet  = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
  var rowNum = findRowNumByToken(sheet, J.token);
  if (rowNum < 0) return _tmpAuditBackfillDone_(tag + '중단: 토큰에 해당하는 행이 없음 (' + J.token + ')', false);
  var r = readRow(sheet, rowNum);

  var stageLabel = String(r[reqApprCol(J.stageIdx, 0)] || '');
  var stageEmail = String(r[reqApprCol(J.stageIdx, 2)] || '').trim().toLowerCase();
  var stageStat  = String(r[reqApprCol(J.stageIdx, 3)] || '');
  var stageAt    = r[reqApprCol(J.stageIdx, 4)];
  var expectedMs = new Date(J.approvedAt).getTime();
  var mismatch = [];
  if (String(r[COL.DOC_NO] || '') !== J.docNo)        mismatch.push('품의번호 ' + r[COL.DOC_NO]);
  if (String(r[COL.DOC_TYPE] || '') !== J.docType)    mismatch.push('문서유형 ' + r[COL.DOC_TYPE]);
  if (stageEmail !== J.approver)                      mismatch.push('결재자 ' + stageEmail);
  if (stageStat !== '승인')                           mismatch.push('결재 상태 ' + stageStat);
  if (!(stageAt instanceof Date) || Math.abs(stageAt.getTime() - expectedMs) >= 1000) mismatch.push('처리시각 ' + stageAt);
  if (mismatch.length) return _tmpAuditBackfillDone_(tag + '중단: 결재란이 예상과 다름 — ' + mismatch.join(' / '), false);

  // ── 2) 이미 있으면 건너뜀 (같은 토큰·같은 단계의 DOC_APPROVE) ──
  var audit = ensureAuditLogSheet();
  var last  = audit.getLastRow();
  if (last >= 2) {
    var vals = audit.getRange(2, 1, last - 1, 12).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][AUDIT_COL.EVENT_TYPE] !== AUDIT_EVENT.DOC_APPROVE || vals[i][AUDIT_COL.DOC_TOKEN] !== J.token) continue;
      var p = {};
      try { p = JSON.parse(vals[i][AUDIT_COL.PAYLOAD] || '{}'); } catch (_) {}
      if (p.stageIdx === J.stageIdx) {
        return _tmpAuditBackfillDone_(tag + 'SKIP: 이미 기록돼 있음 (시스템로그 ' + (i + 2) + '행)', true);
      }
    }
  }

  // ── 3) 그때 남았어야 할 행 — _decisionCore_의 정상 처리 기록과 같은 형식, 시각만 결재란 값 ──
  var row = _buildAuditRow_({
    eventType:  AUDIT_EVENT.DOC_APPROVE,
    actor:      stageEmail,
    actorRole:  'approver',
    docNo:      J.docNo,
    docToken:   J.token,
    docType:    J.docType,
    targetUser: '',
    reason:     (stageLabel || ('결재자 ' + J.stageIdx)) + ' 단계 승인 (사후 보정)',
    payload: {
      stageIdx:    J.stageIdx,
      onBehalfOf:  '',
      adminReason: '',
      backfill:    { by: actor, at: new Date().toISOString(), cause: J.cause, source: '품의서목록 결재란 ' + J.stageIdx + '번' },
    },
  });
  row[AUDIT_COL.TIMESTAMP] = stageAt;

  var summary = tag + '넣을 행: ' + JSON.stringify(row);
  if (dryRun) return _tmpAuditBackfillDone_(summary + '\n' + tag + '검증 통과 (쓰기 없음)', true);

  audit.appendRow(row);
  writeAuditLog({
    eventType: 'ADMIN_AUDIT_BACKFILL',
    actor:     actor,
    actorRole: 'admin',
    docNo:     J.docNo,
    docToken:  J.token,
    docType:   J.docType,
    reason:    '시스템로그 누락 1건 보정: ' + J.docNo + ' DOC_APPROVE (' + (stageLabel || J.stageIdx) + ', ' + J.approvedAt + ')',
    payload:   { backfilledLogId: row[AUDIT_COL.LOG_ID], stageIdx: J.stageIdx, cause: J.cause },
  });
  return _tmpAuditBackfillDone_(summary + '\n완료: DOC_APPROVE 1행 + ADMIN_AUDIT_BACKFILL 1행 기록', true);
}

function _tmpAuditBackfillDone_(message, ok) {
  console.log(message);
  return { ok: ok, message: message };
}
