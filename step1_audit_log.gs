// ================================================================
// Step 1: 시스템로그 (AuditLog) 인프라
// ================================================================
// 적용 위치: Code.gs 내부
// 추가 위치 권장:
//   - CONFIG 추가 키: line 44 (CONFIG 닫는 괄호 직전)
//   - 상수/헬퍼: line 322 (isAdminUser 다음, "// 4. 스키마 마이그레이션" 직전)
// ================================================================


// ─────────────────────────────────────────────────────────────────
// 【1】 CONFIG에 추가할 키
//   기존 CONFIG 객체 내부, ADMIN_EMAILS 다음 줄에 아래 한 줄 추가
// ─────────────────────────────────────────────────────────────────
/*
  ADMIN_EMAILS: [
    'davidkang@inlct.com',
  ],
  AUDIT_SHEET_NAME: '시스템로그',   // ← 이 줄 추가
*/


// ─────────────────────────────────────────────────────────────────
// 【2】 AUDIT_COL 컬럼 인덱스 상수 (0-based)
//   COL 객체 정의 직후에 추가
// ─────────────────────────────────────────────────────────────────

var AUDIT_COL = {
  LOG_ID:      0,   // UUID
  TIMESTAMP:   1,   // Date 객체
  ACTOR:       2,   // 행위자 이메일
  ACTOR_ROLE:  3,   // 'admin' | 'procurement' | 'approver' | 'requester' | 'unknown'
  EVENT_TYPE:  4,   // AUDIT_EVENT 상수 중 하나
  DOC_NO:      5,   // 대상 문서번호 (없을 수 있음)
  DOC_TOKEN:   6,   // 대상 문서 토큰 (없을 수 있음)
  DOC_TYPE:    7,   // 'REQ' | 'PRC' | ''
  TARGET_USER: 8,   // 영향받는 대상자 이메일 (대리결재의 원 결재자 등)
  REASON:      9,   // 관리자 액션 사유 (관리자 액션 시 필수)
  PAYLOAD:    10,   // JSON string — 변경 전후 값 등 가변 데이터
  IP_ADDRESS: 11,   // 향후 확장용. 현재 GAS에서 취득 불가, 빈값 유지
};


// ─────────────────────────────────────────────────────────────────
// 【3】 이벤트 타입 상수
// ─────────────────────────────────────────────────────────────────

var AUDIT_EVENT = {
  // 관리자 액션 (1차 A 그룹)
  ADMIN_FORCE_DISCARD:        'ADMIN_FORCE_DISCARD',        // A-1 강제 폐기
  ADMIN_PROXY_APPROVE:        'ADMIN_PROXY_APPROVE',        // A-2 대리 승인
  ADMIN_PROXY_REJECT:         'ADMIN_PROXY_REJECT',         // A-2 대리 반려
  ADMIN_CHANGE_APPROVER:      'ADMIN_CHANGE_APPROVER',      // A-3 결재자 변경
  ADMIN_FORCE_RELEASE_LOCK:   'ADMIN_FORCE_RELEASE_LOCK',   // A-4 락 강제 해제
  ADMIN_FORCE_STATUS_CHANGE:  'ADMIN_FORCE_STATUS_CHANGE',  // A-5 상태 강제 변경

  // 권한 거부 (보안 기록)
  ADMIN_ACCESS_DENIED:        'ADMIN_ACCESS_DENIED',

  // 추후 확장 자리 (Step 3 이후 추가):
  // DOC_SUBMIT, DOC_APPROVE, DOC_REJECT, DOC_DISCARD,
  // PRC_CLAIM, PRC_RELEASE, PRC_SUBMIT,
  // EMAIL_SEND_FAIL, DRIVE_SAVE_FAIL,
  // LINK_INVALID_ACCESS
};


// ─────────────────────────────────────────────────────────────────
// 【4】 헬퍼 함수
//   기존 isAdminUser() 다음, runSchemaMigration() 직전에 추가 권장
// ─────────────────────────────────────────────────────────────────

/**
 * 시스템로그 시트 보장
 * - 시트가 없으면 생성
 * - 헤더 행이 비어있으면 헤더 작성 + 1행 고정 + 강조 스타일
 * - 매 writeAuditLog 호출 시 검증 (이미 있으면 즉시 통과)
 * @returns {Sheet} 시스템로그 시트
 */
function ensureAuditLogSheet() {
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.AUDIT_SHEET_NAME);

  if (sheet.getLastRow() > 0) return sheet;

  var headers = [
    'logId', 'timestamp', 'actor', 'actorRole',
    'eventType', 'docNo', 'docToken', 'docType',
    'targetUser', 'reason', 'payload', 'ipAddress',
  ];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#f0f0f0');

  // 컬럼 폭 권장 (선택적 — Apps Script 권한으로 가능)
  sheet.setColumnWidth(AUDIT_COL.LOG_ID + 1, 250);     // UUID
  sheet.setColumnWidth(AUDIT_COL.TIMESTAMP + 1, 150);  // 일시
  sheet.setColumnWidth(AUDIT_COL.EVENT_TYPE + 1, 200); // 이벤트 타입
  sheet.setColumnWidth(AUDIT_COL.PAYLOAD + 1, 400);    // JSON

  return sheet;
}

/**
 * 시스템로그 기록 (append-only, fail-safe)
 *
 * @param {Object} entry
 *   - eventType    {string}  AUDIT_EVENT 상수 중 하나 (필수)
 *   - actor        {string}  행위자 이메일 (생략 시 자동 취득)
 *   - actorRole    {string}  'admin' | 'procurement' | 'approver' | 'requester' (생략 시 자동 판정)
 *   - docNo        {string=} 대상 문서번호
 *   - docToken     {string=} 대상 문서 토큰
 *   - docType      {string=} 'REQ' | 'PRC'
 *   - targetUser   {string=} 영향받는 대상자 이메일
 *   - reason       {string=} 사유 (관리자 액션은 별도로 requireAdminReason로 검증 권장)
 *   - payload      {Object=} 변경 전후 값 등 가변 데이터 (자동 JSON.stringify)
 *
 * @returns {boolean} 기록 성공 여부 (실패해도 throw하지 않음)
 *
 * @example
 *   writeAuditLog({
 *     eventType: AUDIT_EVENT.ADMIN_FORCE_DISCARD,
 *     docNo: 'REQ-2026-0123',
 *     docToken: 'abc-uuid',
 *     docType: 'REQ',
 *     reason: '중복 제출로 인한 관리자 강제 폐기',
 *     payload: { oldStatus: '결재중', newStatus: '폐기' },
 *   });
 */
function writeAuditLog(entry) {
  try {
    if (!entry || !entry.eventType) {
      console.error('[AuditLog] eventType is required');
      return false;
    }

    var sheet = ensureAuditLogSheet();

    var actor = entry.actor || getActiveUserEmail() || '';
    var role  = entry.actorRole || _deriveActorRole(actor);

    var payloadStr = '';
    if (entry.payload !== undefined && entry.payload !== null) {
      try {
        payloadStr = (typeof entry.payload === 'string')
          ? entry.payload
          : JSON.stringify(entry.payload);
      } catch(_) {
        payloadStr = '[unserializable]';
      }
    }

    var row = new Array(12);
    row[AUDIT_COL.LOG_ID]      = Utilities.getUuid();
    row[AUDIT_COL.TIMESTAMP]   = new Date();
    row[AUDIT_COL.ACTOR]       = actor;
    row[AUDIT_COL.ACTOR_ROLE]  = role;
    row[AUDIT_COL.EVENT_TYPE]  = entry.eventType;
    row[AUDIT_COL.DOC_NO]      = entry.docNo      || '';
    row[AUDIT_COL.DOC_TOKEN]   = entry.docToken   || '';
    row[AUDIT_COL.DOC_TYPE]    = entry.docType    || '';
    row[AUDIT_COL.TARGET_USER] = entry.targetUser || '';
    row[AUDIT_COL.REASON]      = entry.reason     || '';
    row[AUDIT_COL.PAYLOAD]     = payloadStr;
    row[AUDIT_COL.IP_ADDRESS]  = '';  // GAS에서 IP 취득 불가, 향후 확장용

    sheet.appendRow(row);
    return true;
  } catch (err) {
    // fail-safe: 로그 기록 실패가 본 업무를 막으면 안 됨
    console.error('[AuditLog] 기록 실패: ' + err.toString(), entry);
    return false;
  }
}

/**
 * 행위자의 역할 자동 판정
 * 우선순위: admin > procurement > (이외는 unknown)
 * (approver/requester는 문서 컨텍스트가 있어야 판정 가능하므로 호출자가 직접 명시)
 */
function _deriveActorRole(email) {
  if (!email) return 'unknown';
  if (isAdminUser(email)) return 'admin';
  if (isProcurementUser(email)) return 'procurement';
  return 'unknown';
}

/**
 * 관리자 액션용 사유 검증
 * - reason이 비어있거나 공백뿐이면 에러 throw
 * - 관리자 액션 함수 진입 직후에 호출
 *
 * @param {string} reason
 * @throws {Error} 사유 미입력 시
 */
function requireAdminReason(reason) {
  if (!reason || !String(reason).trim()) {
    throw new Error('관리자 액션에는 사유 입력이 필수입니다.');
  }
}

/**
 * 관리자 권한 검증 + 거부 시 자동 로그 기록
 * - 관리자가 아니면 ADMIN_ACCESS_DENIED 로그 후 throw
 * - 통과하면 actor 이메일 반환
 *
 * @param {string} eventType  시도한 액션 (로그 기록용)
 * @param {Object=} payload   시도 컨텍스트 (로그 기록용)
 * @returns {string} 검증된 관리자 이메일
 * @throws {Error} 관리자가 아닐 때
 *
 * @example
 *   function forceDiscardForClient(payload) {
 *     return withLock(function() {
 *       var actor = assertAdminWithLog(AUDIT_EVENT.ADMIN_FORCE_DISCARD, payload);
 *       requireAdminReason(payload.reason);
 *       // ... 실제 처리
 *     });
 *   }
 */
function assertAdminWithLog(eventType, payload) {
  var actor = getActiveUserEmail();

  if (!isAdminUser(actor)) {
    // 권한 거부 기록
    writeAuditLog({
      eventType: AUDIT_EVENT.ADMIN_ACCESS_DENIED,
      actor: actor,
      actorRole: _deriveActorRole(actor),
      reason: '관리자 권한 없음 (시도: ' + eventType + ')',
      payload: payload || null,
    });
    throw new Error('관리자 권한이 필요합니다.');
  }

  return actor;
}


// ─────────────────────────────────────────────────────────────────
// 【5】 검증 함수 (Apps Script 에디터에서 수동 실행하여 테스트)
//   실행 후 시스템로그 시트에 1행 들어가는지 확인
// ─────────────────────────────────────────────────────────────────

/**
 * Step 1 인프라 검증용 함수
 * Apps Script 에디터에서 testAuditLogInfra 선택 후 실행
 *
 * 기대 동작:
 *   1) '시스템로그' 시트 자동 생성
 *   2) 헤더 12개 작성 (lOG_ID ~ ipAddress)
 *   3) 테스트 행 1개 append
 *   4) 콘솔에 성공 메시지
 */
function testAuditLogInfra() {
  Logger.log('===== Step 1: AuditLog 인프라 테스트 시작 =====');

  // 1) 시트 보장
  var sheet = ensureAuditLogSheet();
  Logger.log('[1/3] 시스템로그 시트 확보: ' + sheet.getName());
  Logger.log('      현재 행 수: ' + sheet.getLastRow());

  // 2) 테스트 로그 기록
  var ok = writeAuditLog({
    eventType: AUDIT_EVENT.ADMIN_ACCESS_DENIED,  // 테스트용으로 가장 안전한 이벤트
    docNo: 'TEST-DOCNO',
    docToken: 'test-token-' + Date.now(),
    docType: 'REQ',
    targetUser: '',
    reason: '[TEST] testAuditLogInfra 실행 — 실제 권한 거부 아님',
    payload: {
      test: true,
      step: 'Step 1 검증',
      timestamp: new Date().toISOString(),
    },
  });
  Logger.log('[2/3] 테스트 로그 기록: ' + (ok ? '성공' : '실패'));

  // 3) 검증
  var newRowCount = sheet.getLastRow();
  Logger.log('[3/3] 기록 후 행 수: ' + newRowCount);

  if (ok && newRowCount >= 2) {
    Logger.log('===== ✅ Step 1 인프라 정상 동작 =====');
    Logger.log('스프레드시트의 "시스템로그" 시트를 직접 열어 확인하세요.');
  } else {
    Logger.log('===== ❌ Step 1 인프라 이상. 콘솔 에러를 확인하세요 =====');
  }

  return {
    ok: ok,
    sheetName: sheet.getName(),
    rowCount: newRowCount,
  };
}
