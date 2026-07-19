// ================================================================
// 구매품의서 시스템 - Google Apps Script (Code.gs)
// v2.0 — REQ + PRC + 대시보드 홈 통합 버전
// ================================================================
// 사전 준비:
//   1) Apps Script 에디터 → "서비스(Services)" → "Drive API" 추가
//      (식별자: Drive, 버전: v3)
//   2) 아래 CONFIG 값 확인/수정
//   3) 첫 배포 후 한 번 runSchemaMigration() 수동 실행 (시트 컬럼 확장)
// ================================================================

var CONFIG = {
  SHEET_ID:         '1Jg0BfLiMrI9MUDMqwU7DWIAA8loVU5XwfumiWHMqMTA',
  STAGING_ROOT_ID:  '17MouMkuLeAcIkfml9kRwIs6OQVdYKKRc',
  FINAL_ROOT_ID:    '1_dXx1O-jyKHa1mnJ48eLzTEhzGB1t9VE',
  SHEET_NAME:       '품의서목록',
  WEBAPP_URL:       'https://script.google.com/a/macros/inlct.com/s/AKfycbw-iL2l4m59rcZLuiBSgWkumFjBhg80QRm3Y2Kdz1_l9YJYqnWPDw_lhuTurXiCJ8Ip/exec',
  FROM_NAME:        '구매품의 시스템',
  // 모든 발신 메일의 From 주소. 스크립트는 David 계정에서 실행되지만 메일은 이 주소로 나간다.
  //  ⚠ 사전 준비: David Gmail → 설정 → 계정 → '다른 주소에서 메일 보내기'에 이 주소를
  //    별칭(alias)으로 등록·인증해야 한다. 미등록 상태면 별칭 없이(=davidkang) 발송된다.
  //    등록 여부는 GAS 에디터에서 sendTestMailFromAlias() 실행으로 확인.
  MAIL_FROM:        'admin-inlct@inlct.com',
  USE_SHARED_DRIVE: true,

  // 파일 제약
  MAX_FILE_SIZE:         25 * 1024 * 1024,  // 파일당 25MB
  MAX_TOTAL_SIZE:        30 * 1024 * 1024,  // REQ+PRC 합산 30MB (FR-09a)
  ALLOWED_EXTS:          ['pdf','jpg','jpeg','png','xlsx','xls','docx','doc','pptx','hwp','zip'],

  // 락
  LOCK_WAIT_MS:          30000,             // LockService 대기 (NFR/FR-51)
  PRC_CLAIM_TTL_HOURS:   24,                // PRC_DRAFT 락 자동 해제 시간 (예외처리)

  // 이메일 재시도 (NFR-05)
  EMAIL_RETRY_COUNT:     3,
  EMAIL_RETRY_DELAY_MS:  1000,

  // ── 권한 판정용 계정 목록 (실제 로그인 계정만 — 메일 수신처와 별개) ──
  //  ⚠ 여기서 계정을 빼면 해당 사용자의 구매팀/관리자 메뉴 권한이 사라진다.
  //    알림 수신처만 바꾸려면 아래 *_NOTIFY_EMAILS 를 수정할 것.
  PROCUREMENT_TEAM_EMAILS: [
    'davidkang@inlct.com',
    //'mhpark@inlct.com',
    //'bret@inlct.com',
    //'kskim@inlct.com',
    'admin-inlct@inlct.com',   // mujung@inlct.com 대체 (2026-07)
  ],
  ADMIN_EMAILS: [
    'davidkang@inlct.com',
  ],

  // ── 알림 수신처 (권한과 무관 — 순수 메일 수신 대상) ──
  // 구매팀 알림 (REQ 1차 승인 시 픽업 요청)
  PROCUREMENT_NOTIFY_EMAILS: [
    'admin-inlct@inlct.com',
  ],
  // 관리자 알림 (시스템 오류, INSP PDF 작업요청 등)
  ADMIN_NOTIFY_EMAILS: [
    'admin-inlct@inlct.com',
  ],
  AUDIT_SHEET_NAME: '시스템로그',
  INSP_SHEET_NAME:  '검수보고서목록',   // [INSP Step 1] 검수보고서 시트 탭명
  VENDOR_SHEET_NAME: '업체목록',        // 공급업체 마스터 (REQ 폼 조회 자동입력용)
};

// ================================================================
// Sheets 열 인덱스 상수 (0-based)
// v2.0: 기존 26개 컬럼 + 4개 추가 = 30개
// ================================================================
var COL = {
  // 기존 (0~16)
  SUBMIT_AT:     0,
  DOC_NO:        1,
  ISSUE_DATE:    2,
  DRAFTER:       3,
  DEPT:          4,
  SUBJECT:       5,
  PURPOSE:       6,
  VENDOR_NAME:   7,
  VENDOR_EMAIL:  8,
  VENDOR_CTC:    9,
  VENDOR_PHONE:  10,
  DELIVERY_DATE: 11,
  DELIVERY_ADDR: 12,
  ITEMS:         13,
  TOTAL_AMT:     14,
  TOTAL_VAT:     15,
  REMARKS:       16,

  // 결재 관리 (17~25)
  STATUS:        17,
  TOKEN:         18,
  DRIVE_ID:      19,
  APPR_COUNT:    20,
  APPR_IDX:      21,
  RESUB_COUNT:   22,
  REJECT_LOG:    23,
  ATTACH_LIST:   24,
  MOVE_STATUS:   25,

  // v2.0 신규 (26~29)
  DOC_TYPE:      26,  // 'REQ' | 'PRC'
  PARENT_DOC_ID: 27,  // PRC인 경우 원본 REQ의 token
  CLAIMED_BY:    28,  // PRC 락 점유자 이메일
  CLAIMED_AT:    29,  // PRC 락 점유 시각

  // v2.1 신규 (30~31) — 구매 조건 (PRC 폼 입력값)
  PURCHASE_METHOD: 30,  // 구매방법 (계좌이체/법인카드 등)
  PAYMENT_INFO:    31,  // 지급정보/결제조건 (자유 텍스트)

  // 결재자 블록 (32~)
  APPR_START:    32,
  APPR_COLS:     6,   // label, name, email, status, processedAt, comment
};

// 별칭 (오타 방지)
COL._VERSION = 'v2.1';

// ================================================================
// 1. Web App 진입점 (라우팅)
// ================================================================

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'home';

  // 1) 결재 액션 (이메일 링크)
  if (action === 'approve') {
    var _rawIdxA = e.parameter.idx;
    return _renderTemplate('Procurement_Approval', {
      token:         e.parameter.token    || '',
      decision:      e.parameter.decision || '',
      docNo:         e.parameter.docNo    || '',
      // idx 파라미터가 있을 때만 사용, 없으면 빈 문자열 (서버가 본인 단계 자동 판정)
      approverIndex: (_rawIdxA !== undefined && _rawIdxA !== '' && !isNaN(_rawIdxA))
                     ? parseInt(_rawIdxA)
                     : '',
      docKind:       e.parameter.docKind || 'REQ',   // [INSP Step 4] 'REQ' | 'INSP'
      webappUrl:     CONFIG.WEBAPP_URL,
    }, '결재 처리');
  }

  // 2) 반려 후 수정/폐기 화면
  if (action === 'reject_view') {
    return _renderTemplate('Procurement_Reject', {
      token:     e.parameter.token || '',
      docNo:     e.parameter.docNo || '',
      webappUrl: CONFIG.WEBAPP_URL,
    }, '반려 처리');
  }

  // 3) 결재자 뷰어 (이메일 링크 → 본문 확인)
  if (action === 'view') {
    var _rawIdxV = e.parameter.idx;
    return _renderTemplate('Procurement_Viewer', {
      token:         e.parameter.token    || '',
      decision:      e.parameter.decision || '',
      // idx 파라미터가 있을 때만 사용, 없으면 빈 문자열
      // (홈에서 진입 시 idx 없음 → 서버가 대기 중인 본인 단계를 자동 판정)
      approverIndex: (_rawIdxV !== undefined && _rawIdxV !== '' && !isNaN(_rawIdxV))
                     ? parseInt(_rawIdxV)
                     : '',
      webappUrl:     CONFIG.WEBAPP_URL,
    }, '구매품의서 확인');
  }

  // 4) 신규 REQ 작성 (FR/AC-13: 명시적 진입)
  if (action === 'new') {
    return _renderTemplate('Procurement_Form', {
      webappUrl: CONFIG.WEBAPP_URL,
    }, '구매품의서 작성');
  }

  // 5) PRC 작성 진입 (FR-51: 락 점유 후 진입)
  if (action === 'prc') {
    return _renderTemplate('Procurement_PRC_Form', {
      parentToken: e.parameter.parentId || e.parameter.parentToken || '',
      webappUrl:   CONFIG.WEBAPP_URL,
    }, '구매팀 구매품의서 작성');
  }

  // [INSP Step 3] 검수보고서 작성 화면
  if (action === 'insp_form') {
    return _renderTemplate('Procurement_INSP_Form', {
      prcToken:   e.parameter.token || '',
      resubToken: e.parameter.resub || '',
      webappUrl:  CONFIG.WEBAPP_URL,
    }, '검수보고서 작성');
  }

  // [INSP Step 5] 검수보고서 전용 뷰어 (열람 + 결재)
  if (action === 'insp_view') {
    var _rawIdxI = e.parameter.idx;
    return _renderTemplate('Procurement_INSP_Viewer', {
      token:    e.parameter.token || '',
      decision: e.parameter.decision || '',
      idxHint:  (_rawIdxI !== undefined && _rawIdxI !== '' && !isNaN(_rawIdxI)) ? parseInt(_rawIdxI) : '',
      webappUrl: CONFIG.WEBAPP_URL,
    }, '검수보고서');
  }

  // 5-1) 관리자 페이지 진입 (A 그룹 — 강제 폐기/대리 결재/결재자 변경/락 해제/상태 변경)
  if (action === 'admin') {
    var fn = e.parameter.fn || '';
    var actor = getActiveUserEmail();

    // 권한 거부 페이지 (모든 관리자 기능 공통)
    if (!isAdminUser(actor)) {
      writeAuditLog({
        eventType: AUDIT_EVENT.ADMIN_ACCESS_DENIED,
        actor: actor,
        reason: '관리자 페이지 접근 시도 (fn=' + fn + ')',
      });
      return HtmlService.createHtmlOutput(
        '<div style="font-family:Malgun Gothic,sans-serif;text-align:center;padding:60px 20px;">' +
        '<div style="font-size:48px;opacity:0.4;margin-bottom:16px;">🔒</div>' +
        '<h2 style="color:#4a5260;">관리자 권한이 필요합니다</h2>' +
        '<p style="color:#8b94a3;">이 페이지는 관리자로 등록된 계정만 접근할 수 있습니다.</p>' +
        '<p style="color:#8b94a3;font-size:12px;margin-top:20px;">접근 시도자: ' + (actor || '(미인증)') + '</p>' +
        '<a href="' + CONFIG.WEBAPP_URL + '" style="display:inline-block;margin-top:20px;padding:8px 16px;background:#1a6fc4;color:#fff;text-decoration:none;border-radius:5px;">홈으로 돌아가기</a>' +
        '</div>'
      ).setTitle('접근 거부').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // A-1: 강제 폐기
    if (fn === 'discard') {
      return _renderTemplate('Procurement_Admin_Discard', {
        webappUrl: CONFIG.WEBAPP_URL,
      }, '관리자 — 강제 폐기');
    }

    // A-4: 강제 락 해제
    if (fn === 'force_release_lock') {
      return _renderTemplate('Procurement_Admin_ForceReleaseLock', {
        webappUrl: CONFIG.WEBAPP_URL,
      }, '관리자 — 강제 락 해제');
    }

    // 향후 A-2 ~ A-5 추가 자리:
    // if (fn === 'proxy_approve') { ... }
    // if (fn === 'change_approver') { ... }
    // if (fn === 'change_status') { ... }

    // 알 수 없는 fn
    return HtmlService.createHtmlOutput(
      '<p style="font-family:Malgun Gothic,sans-serif;padding:40px;">알 수 없는 관리자 기능입니다: ' + fn + '</p>'
    ).setTitle('오류');
  }

  // 6) 단순 GET 콜백 (호환성용)
  if (action === 'getApprovers')  return getApproverList();
  if (action === 'getRejectData') return getRejectData(e.parameter.token);

  // 7) 기본 진입 = 대시보드 홈 (§12.2, AC-13)
  //    menu 파라미터: 홈 진입 시 기본 선택할 메뉴 키 (예: 메일 링크의 menu=completed)
  return _renderTemplate('Procurement_Home', {
    webappUrl:   CONFIG.WEBAPP_URL,
    initialMenu: e.parameter.menu || '',
  }, '구매결재 시스템');
}

function _renderTemplate(filename, vars, title) {
  var tmpl = HtmlService.createTemplateFromFile(filename);
  Object.keys(vars).forEach(function(k) { tmpl[k] = vars[k]; });
  return tmpl.evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action || '';

    if (action === 'submit')        return jsonResponse(_submitCore(payload));
    if (action === 'decide')        return jsonResponse(_decisionCore(payload));
    if (action === 'resubmit')      return jsonResponse(_resubmitCore(payload));
    if (action === 'discard')       return jsonResponse(_discardCore(payload));
    if (action === 'get_doc')       return jsonResponse(_getDocumentDataCore(payload));
    if (action === 'claim_req')     return jsonResponse(_claimRequestCore(payload));
    if (action === 'release_claim') return jsonResponse(_releaseClaimCore(payload));
    if (action === 'submit_prc')    return jsonResponse(_submitPrcCore(payload));

    // 관리자 액션 (A 그룹)
    if (action === 'admin_list_discardable') return jsonResponse(_adminListDiscardableCore(payload));
    if (action === 'admin_force_discard')    return jsonResponse(_adminForceDiscardCore(payload));

    return jsonResponse({ ok: false, message: '알 수 없는 액션입니다.' });
  } catch(err) {
    return jsonResponse({ ok: false, message: err.toString() });
  }
}

// ================================================================
// 2. LockService 헬퍼 (FR-51)
// ================================================================

function withLock(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_WAIT_MS);
  } catch(e) {
    throw new Error('시스템이 바쁩니다. 잠시 후 다시 시도해주세요.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
// 3. 유틸리티
// ================================================================

/**
 * 셀 값(Date 객체 또는 문자열)을 안전하게 날짜 문자열로 변환
 * google.script.run은 Date 객체 직렬화에 실패할 수 있어 문자열로 강제 변환
 */
function toDateStr(v, fmt) {
  if (v === null || v === undefined || v === '') return '';
  fmt = fmt || 'yyyy-MM-dd';
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return Utilities.formatDate(v, Session.getScriptTimeZone(), fmt);
  }
  var s = String(v).trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  var d = new Date(s);
  if (isNaN(d.getTime())) return s;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), fmt);
}

function toDateTimeStr(v) {
  return toDateStr(v, 'yyyy-MM-dd HH:mm');
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, name) {
  var s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  return s;
}

function normalizeDocNo(docNo) {
  return String(docNo || '').trim().toUpperCase();
}

/**
 * TextFinder 기반 행 검색 (성능 개선)
 * getDataRange 전체 로딩 대신 사용
 * @returns 1-based row number, 없으면 -1
 */
function findRowNumByToken(sheet, token) {
  if (!token) return -1;
  var finder = sheet.createTextFinder(String(token))
    .matchEntireCell(true)
    .matchCase(true);
  var found = finder.findNext();
  if (!found) return -1;
  // 토큰 컬럼(S, 19번째 컬럼)이 맞는지 확인
  if (found.getColumn() !== COL.TOKEN + 1) return -1;
  return found.getRow();
}

/**
 * 단일 행을 읽어 반환 (성능 개선)
 */
function readRow(sheet, rowNum) {
  var lastCol = sheet.getLastColumn();
  return sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
}

/**
 * 일괄 업데이트 헬퍼 (성능 개선)
 * updates: [[col(1-based), value], ...]
 */
function batchUpdate(sheet, rowNum, updates) {
  if (!updates || !updates.length) return;
  // 연속된 컬럼은 묶어서 setValues, 비연속은 개별 setValue
  updates.sort(function(a, b) { return a[0] - b[0]; });
  var i = 0;
  while (i < updates.length) {
    var startCol = updates[i][0];
    var values = [updates[i][1]];
    var j = i + 1;
    while (j < updates.length && updates[j][0] === updates[j-1][0] + 1) {
      values.push(updates[j][1]);
      j++;
    }
    sheet.getRange(rowNum, startCol, 1, values.length).setValues([values]);
    i = j;
  }
}

function getActiveUserEmail() {
  try {
    var e = Session.getActiveUser().getEmail();
    if (e) return e;
  } catch(_) {}
  try {
    return Session.getEffectiveUser().getEmail() || '';
  } catch(_) { return ''; }
}

// 기안자 본인 검증: 결재란 1번(기안자)의 이메일이 현재 로그인 사용자와 일치해야 함.
// 클라이언트 잠금이 우회/변조된 경우의 서버측 안전망. 불일치 시 에러객체, 정상 시 null 반환.
function assertDrafterIsSelf(data) {
  var me = (getActiveUserEmail() || '').toLowerCase();
  var approvers = data.approvers || [];
  var drafterEmail = (approvers[0] && approvers[0].email ? approvers[0].email : '').toLowerCase();
  if (!me) {
    return { ok: false, message: '로그인 사용자를 확인할 수 없습니다. 다시 로그인 후 시도해 주세요.' };
  }
  if (!drafterEmail) {
    return { ok: false, message: '기안자(본인) 정보가 비어 있습니다. 페이지를 새로고침 후 다시 시도해 주세요.' };
  }
  if (drafterEmail !== me) {
    return { ok: false, message: '기안자는 본인만 지정할 수 있습니다. (로그인: ' + me + ' / 기안자: ' + drafterEmail + ')' };
  }
  return null;
}

function isProcurementUser(email) {
  if (!email) return false;
  var list = CONFIG.PROCUREMENT_TEAM_EMAILS || [];
  if (list.length === 0) return true; // 미설정 시 전원 허용 (개발 편의)
  return list.indexOf(email.toLowerCase()) >= 0
      || list.indexOf(email) >= 0;
}

function isAdminUser(email) {
  if (!email) return false;
  return (CONFIG.ADMIN_EMAILS || []).indexOf(email) >= 0;
}


// ================================================================
// 3-1. 시스템로그 (AuditLog) 인프라 [v2.0 신규]
// ================================================================
// - 관리자 액션 및 보안 이벤트를 append-only로 기록
// - fail-safe: 로그 기록 실패가 본 업무를 중단시키지 않음
// - ensureAuditLogSheet()로 시트 자동 보장
// ================================================================

// AUDIT_COL 컬럼 인덱스 상수 (0-based)
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

// 이벤트 타입 상수
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

  // 검수보고서(INSP) 라이프사이클 — [INSP Step 1]
  INSP_SUBMIT:    'INSP_SUBMIT',     // 검수보고서 제출
  INSP_APPROVE:   'INSP_APPROVE',    // 검수 결재 승인 (단계별)
  INSP_REJECT:    'INSP_REJECT',     // 검수 결재 반려
  INSP_RESUBMIT:  'INSP_RESUBMIT',   // 반려 후 수정 재상신
  INSP_FINALIZED: 'INSP_FINALIZED',  // 최종 검수 승인 → 품의 종결

  // 추후 확장 자리 (Step 3 이후 추가):
  // DOC_SUBMIT, DOC_APPROVE, DOC_REJECT, DOC_DISCARD,
  // PRC_CLAIM, PRC_RELEASE, PRC_SUBMIT,
  // EMAIL_SEND_FAIL, DRIVE_SAVE_FAIL,
  // LINK_INVALID_ACCESS
};

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

  // 컬럼 폭 권장
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
 *   - actor        {string=} 행위자 이메일 (생략 시 자동 취득)
 *   - actorRole    {string=} 'admin' | 'procurement' | 'approver' | 'requester' (생략 시 자동 판정)
 *   - docNo        {string=} 대상 문서번호
 *   - docToken     {string=} 대상 문서 토큰
 *   - docType      {string=} 'REQ' | 'PRC'
 *   - targetUser   {string=} 영향받는 대상자 이메일
 *   - reason       {string=} 사유 (관리자 액션은 requireAdminReason으로 별도 검증 권장)
 *   - payload      {Object=} 변경 전후 값 등 가변 데이터 (자동 JSON.stringify)
 *
 * @returns {boolean} 기록 성공 여부 (실패해도 throw하지 않음)
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
 * 우선순위: admin > procurement > unknown
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
 */
function assertAdminWithLog(eventType, payload) {
  var actor = getActiveUserEmail();

  if (!isAdminUser(actor)) {
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

/**
 * Step 1 인프라 검증용 함수
 * Apps Script 에디터에서 testAuditLogInfra 선택 후 실행
 *
 * 기대 동작:
 *   1) '시스템로그' 시트 자동 생성
 *   2) 헤더 12개 작성
 *   3) 테스트 행 1개 append
 *   4) 로그에 성공 메시지
 */
function testAuditLogInfra() {
  Logger.log('===== Step 1: AuditLog 인프라 테스트 시작 =====');

  var sheet = ensureAuditLogSheet();
  Logger.log('[1/3] 시스템로그 시트 확보: ' + sheet.getName());
  Logger.log('      현재 행 수: ' + sheet.getLastRow());

  var ok = writeAuditLog({
    eventType: AUDIT_EVENT.ADMIN_ACCESS_DENIED,
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


// ================================================================
// 4. 스키마 마이그레이션 (최초 1회 수동 실행)
// ================================================================

/**
 * v1.x → v2.0 시트 스키마 마이그레이션
 * - 헤더 30컬럼으로 확장 (docType, parentDocId, claimedBy, claimedAt 추가)
 * - 기존 행에 docType='REQ' 채움
 * Apps Script 에디터에서 한 번 실행하면 완료
 */
function runSchemaMigration() {
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);

  var v2Headers = [
    '제출일시','품의번호','발행일자','기안자','부서','품의제목','용도/공사내역',
    '업체명','담당자이메일','업체담당자','연락처','납기일','납품장소',
    '품목내역','합계금액','부가세포함합계','비고',
    '결재상태','토큰','Drive폴더ID','결재자수','현재결재순번',
    '재상신횟수','반려이력누적','첨부파일ID목록','이동상태',
    'DocType','ParentDocId','ClaimedBy','ClaimedAt',  // v2.0 신규
    '구매방법','구매조건',                              // v2.1 신규
  ];

  if (sheet.getLastRow() === 0) {
    // 빈 시트 → 새 헤더 작성
    sheet.appendRow(v2Headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, v2Headers.length).setFontWeight('bold').setBackground('#f0f0f0');
    Logger.log('[Migration] 빈 시트에 v2.0 헤더 작성 완료');
    return { ok: true, message: '빈 시트에 v2.0 헤더 작성 완료' };
  }

  var currentLastCol = sheet.getLastColumn();
  var existingHeaders = sheet.getRange(1, 1, 1, currentLastCol).getValues()[0];

  // 26번째 컬럼(Z, MOVE_STATUS)까지만 있는 v1.x 시트인지 확인
  var hasV2Cols = existingHeaders.indexOf('DocType') >= 0;

  if (hasV2Cols) {
    Logger.log('[Migration] 이미 v2.0 스키마. 스킵.');
    return { ok: true, message: '이미 v2.0 스키마입니다.', skipped: true };
  }

  // 1) v2.0 신규 컬럼(AA~AD)을 결재자 블록 앞에 삽입
  // 기존 26컬럼 = A~Z. 결재자 블록은 AA(27)부터.
  // 새 컬럼 4개를 AA~AD에 삽입 → 결재자 블록은 AE(31)부터.
  sheet.insertColumnsAfter(26, 4);

  // 헤더 채우기
  sheet.getRange(1, 27, 1, 4).setValues([['DocType','ParentDocId','ClaimedBy','ClaimedAt']]);
  sheet.getRange(1, 27, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');

  // 2) 기존 행에 DocType='REQ' 채움
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var fill = [];
    for (var i = 2; i <= lastRow; i++) fill.push(['REQ', '', '', '']);
    sheet.getRange(2, 27, fill.length, 4).setValues(fill);
  }

  Logger.log('[Migration] v2.0 스키마 마이그레이션 완료 (행 ' + (lastRow - 1) + '건 업데이트)');
  return { ok: true, message: 'v2.0 스키마 마이그레이션 완료', rows: lastRow - 1 };
}

/**
 * v2.1 마이그레이션 — 구매방법/구매조건 컬럼 2개를 결재자 블록 앞에 삽입.
 * 기존 v2.0 시트(ClaimedAt=30번째 컬럼까지, 결재자 블록은 31번째부터)에서
 * 31번째 위치에 컬럼 2개를 삽입 → 결재자 블록은 33번째(AG)부터로 자동 우측 이동.
 * 한 번만 실행. 이미 적용된 시트는 스킵.
 */
function runV21PurchaseConditionMigration() {
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    // 빈 시트 → ensureHeaders가 v2.1 헤더로 초기화하므로 마이그레이션 불필요
    return { ok: true, message: '빈 시트 — ensureHeaders가 v2.1 헤더로 처리합니다.', skipped: true };
  }

  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  if (headerRow.indexOf('구매조건') >= 0) {
    Logger.log('[Migration v2.1] 이미 적용됨. 스킵.');
    return { ok: true, message: '이미 v2.1(구매조건) 스키마입니다.', skipped: true };
  }

  // ClaimedAt 컬럼(30번째, 1-based)을 기준으로 그 뒤에 2개 삽입
  var claimedAtCol = headerRow.indexOf('ClaimedAt') + 1; // 1-based
  if (claimedAtCol <= 0) {
    return { ok: false, message: 'ClaimedAt 컬럼을 찾지 못했습니다. v2.0 마이그레이션을 먼저 실행하세요.' };
  }

  sheet.insertColumnsAfter(claimedAtCol, 2);
  sheet.getRange(1, claimedAtCol + 1, 1, 2).setValues([['구매방법', '구매조건']]);
  sheet.getRange(1, claimedAtCol + 1, 1, 2).setFontWeight('bold').setBackground('#f0f0f0');

  // 기존 행은 빈 값 유지(REQ/이전 PRC는 구매조건 없음). 별도 채움 불필요.
  var rows = Math.max(0, sheet.getLastRow() - 1);
  Logger.log('[Migration v2.1] 구매방법/구매조건 컬럼 삽입 완료 (기존 행 ' + rows + '건은 빈 값 유지, 결재자 블록 우측 이동됨)');
  return { ok: true, message: 'v2.1(구매조건) 마이그레이션 완료. 결재자 블록이 2칸 우측 이동되었습니다.', rows: rows };
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() > 0) return;
  var headers = [
    '제출일시','품의번호','발행일자','기안자','부서','품의제목','용도/공사내역',
    '업체명','담당자이메일','업체담당자','연락처','납기일','납품장소',
    '품목내역','합계금액','부가세포함합계','비고',
    '결재상태','토큰','Drive폴더ID','결재자수','현재결재순번',
    '재상신횟수','반려이력누적','첨부파일ID목록','이동상태',
    'DocType','ParentDocId','ClaimedBy','ClaimedAt',
    '구매방법','구매조건',
  ];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
}

// ================================================================
// 5. 품의번호 중복 검사
// ================================================================

function isDocNoDuplicated(docNo, excludeToken) {
  var target = normalizeDocNo(docNo);
  if (!target) return false;

  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
  if (sheet.getLastRow() < 2) return false;

  // 품의번호 컬럼(B)만 읽어서 비교 — 전체 시트 스캔 회피
  var docNoCol = sheet.getRange(2, COL.DOC_NO + 1, sheet.getLastRow() - 1, 1).getValues();
  var statusCol = sheet.getRange(2, COL.STATUS + 1, sheet.getLastRow() - 1, 1).getValues();
  var tokenCol = sheet.getRange(2, COL.TOKEN + 1, sheet.getLastRow() - 1, 1).getValues();

  for (var i = 0; i < docNoCol.length; i++) {
    if (statusCol[i][0] === '폐기') continue;
    if (excludeToken && tokenCol[i][0] === excludeToken) continue;
    if (normalizeDocNo(docNoCol[i][0]) === target) return true;
  }
  return false;
}

function checkDocNoForClient(docNo) {
  try {
    return { ok: true, duplicated: isDocNoDuplicated(docNo, null) };
  } catch(err) {
    return { ok: false, message: err.toString() };
  }
}


// ================================================================
// 6. 첨부파일 합산 용량 검증 (FR-09a)
// ================================================================

/**
 * 첨부 배열의 합산 용량을 계산하고 30MB 초과 시 throw
 */
function validateTotalAttachmentSize(attachments, additionalSize) {
  additionalSize = additionalSize || 0;
  var total = additionalSize;
  (attachments || []).forEach(function(att) {
    total += Number(att.size) || 0;
  });
  if (total > CONFIG.MAX_TOTAL_SIZE) {
    var mb = Math.round(total / 1024 / 1024 * 10) / 10;
    var maxMb = Math.round(CONFIG.MAX_TOTAL_SIZE / 1024 / 1024);
    throw new Error('첨부파일 합산 용량 초과: ' + mb + 'MB (최대 ' + maxMb + 'MB)');
  }
  return total;
}

// ================================================================
// 7. 품의서 최초 제출 (REQ)
// ================================================================

function submitRequisitionForClient(payload) {
  return _submitCore(payload);
}

function _submitCore(data) {
  // 0단계: 기안자 본인 검증 (본인이 아니면 제출 차단)
  var drafterChk = assertDrafterIsSelf(data);
  if (drafterChk) return drafterChk;

  // 1단계: 합산 용량 사전 검증 (락 밖에서 미리 검사 — 성능 개선)
  try {
    validateTotalAttachmentSize(data.attachments || []);
  } catch(e) {
    return { ok: false, message: e.message };
  }

  // 2단계: 최소 락 — 행 추가/토큰/폴더만 처리, 첨부와 이메일은 락 밖
  var lockResult = withLock(function() {
    try {
      if (isDocNoDuplicated(data.docNo, null)) {
        return { ok: false, message: '이미 사용 중인 품의번호입니다. 다른 번호를 사용해 주세요.' };
      }

      var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
      ensureHeaders(sheet);

      var token        = Utilities.getUuid();
      var folder       = getOrCreateFolder(data.docNo, data.issueDate, 'staging');
      var itemsSummary = buildItemsSummary(data.items);
      var totalAmt     = calcTotal(data.items);
      var totalVat     = totalAmt + computeVat(totalAmt, data.items);
      var approvers    = data.approvers || [];

      var newRow = [
        new Date(),           // A  제출일시
        data.docNo,           // B  품의번호
        data.issueDate,       // C  발행일자
        data.drafter,         // D  기안자
        data.dept,            // E  부서
        data.subject,         // F  품의제목
        data.purpose,         // G  용도/공사내역
        data.vendorName,      // H  업체명
        data.vendorEmail,       // I  담당자이메일
        data.vendorContact,   // J  업체담당자
        data.vendorPhone,     // K  연락처
        data.deliveryDate,    // L  납기일
        data.deliveryAddr,    // M  납품장소
        itemsSummary,         // N  품목내역
        totalAmt,             // O  합계
        totalVat,             // P  부가세포함
        data.remarks,         // Q  비고
        '검토중',             // R  결재상태
        token,                // S  토큰
        folder.getId(),       // T  Drive폴더ID
        approvers.length,     // U  결재자수
        0,                    // V  현재결재순번
        0,                    // W  재상신횟수
        '',                   // X  반려이력누적
        '[]',                 // Y  첨부파일ID목록 (락 밖에서 채움)
        'STAGING',            // Z  이동상태
        'REQ',                // AA DocType
        '',                   // AB ParentDocId
        '',                   // AC ClaimedBy
        '',                   // AD ClaimedAt
        '',                   // AE 구매방법 (REQ는 없음)
        '',                   // AF 구매조건 (REQ는 없음)
      ];

      // AE~ 결재자 블록
      approvers.forEach(function(ap) {
        newRow.push(ap.label, ap.name, ap.email, '대기', '', '');
      });

      sheet.appendRow(newRow);
      var rowNum = sheet.getLastRow();

      return {
        ok:       true,
        rowNum:   rowNum,
        token:    token,
        folderId: folder.getId(),
      };
    } catch(err) {
      return { ok: false, message: err.toString() };
    }
  });

  if (!lockResult.ok) return lockResult;

  // 3단계: 첨부 업로드 (락 밖 — 시간이 가장 오래 걸리는 작업)
  var folder;
  try {
    folder = DriveApp.getFolderById(lockResult.folderId);
  } catch(e) {
    return { ok: false, message: '폴더 접근 실패: ' + e.toString() };
  }

  var savedFiles = [];
  try {
    savedFiles = saveAttachmentsToFolder(folder, data.attachments || []);
  } catch(e) {
    // 업로드 실패 시 행은 남기고 사용자에게 안내
    rollbackFiles(savedFiles);
    // 시트 상태를 ERROR로 표시
    try {
      var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
      sheet.getRange(lockResult.rowNum, COL.STATUS + 1).setValue('업로드오류');
    } catch(_) {}
    return { ok: false, message: '첨부파일 업로드 실패: ' + e.toString() };
  }

  // 4단계: 첨부 메타 시트에 기록 (단일 setValue)
  try {
    var ss2    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet2 = getOrCreateSheet(ss2, CONFIG.SHEET_NAME);
    sheet2.getRange(lockResult.rowNum, COL.ATTACH_LIST + 1).setValue(JSON.stringify(savedFiles));

    // 요약 파일 저장 (오류 시에도 진행)
    try {
      var totalAmt = calcTotal(data.items);
      var totalVat = totalAmt + computeVat(totalAmt, data.items);
      var itemsSummary = buildItemsSummary(data.items);
      saveSummaryToDrive(folder, data, itemsSummary, totalAmt, totalVat, data.approvers || []);
    } catch(_) {}
  } catch(e) {
    return { ok: false, message: '시트 업데이트 실패: ' + e.toString() };
  }

  // 5단계: 이메일 발송 (락 밖, 재시도 포함)
  var approvers = data.approvers || [];
  if (approvers.length > 0 && approvers[0].email) {
    try {
      sendApprovalEmailWithRetry(approvers[0], data, lockResult.token, lockResult.rowNum, 0);
    } catch(e) {
      notifyAdminError('이메일 발송 실패 (제출): ' + data.docNo + ' / ' + e.toString());
      return {
        ok:       true,
        rowNum:   lockResult.rowNum,
        uploaded: savedFiles.length,
        message:  '제출은 완료되었으나 이메일 발송에 실패했습니다. 관리자에게 알림이 전송되었습니다.',
        warning:  true,
      };
    }
  }

  return {
    ok:       true,
    rowNum:   lockResult.rowNum,
    uploaded: savedFiles.length,
    message:  '제출 완료. 결재자에게 이메일이 발송되었습니다.',
  };
}


// ================================================================
// 8. 결재 승인/반려
// ================================================================

function processDecisionFromClient(payload) {
  return _decisionCore(payload);
}
function processDecisionForViewer(payload) {
  return _decisionCore(payload);
}

function _decisionCore(payload) {
  // 락 안: 시트 업데이트만 빠르게
  var lockResult = withLock(function() {
    try {
      var token    = payload.token;
      var decision = payload.decision;
      var comment  = payload.comment || '';
      var idx      = parseInt(payload.idx || '0');

      var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
      var rowNum = findRowNumByToken(sheet, token);

      if (rowNum < 0) return { ok: false, message: '유효하지 않은 토큰입니다.' };

      var r = readRow(sheet, rowNum);
      if (r[COL.STATUS] === '폐기') return { ok: false, message: '폐기된 품의서입니다.' };

      var apprCount = parseInt(r[COL.APPR_COUNT]);
      var curIdx    = parseInt(r[COL.APPR_IDX]);

      if (curIdx !== idx) return { ok: false, message: '이미 처리되었거나 결재 순서가 맞지 않습니다.' };

      var baseCol = COL.APPR_START + idx * COL.APPR_COLS;
      var now     = new Date();

      // 결재자 블록 일괄 업데이트 (3개 연속 컬럼)
      sheet.getRange(rowNum, baseCol + 4, 1, 3).setValues([[
        decision === 'approve' ? '승인' : '반려',
        now,
        comment,
      ]]);

      var docNo     = String(r[COL.DOC_NO] || '');
      var drafter   = String(r[COL.DRAFTER] || '');
      var subject   = String(r[COL.SUBJECT] || '');
      var docType   = String(r[COL.DOC_TYPE] || 'REQ');
      var issueDate = r[COL.ISSUE_DATE];

      if (decision === 'reject') {
        // 반려 처리
        sheet.getRange(rowNum, COL.STATUS + 1).setValue('반려');

        // 첨부 삭제 시도
        try {
          deleteUserAttachments(r[COL.ATTACH_LIST] || '[]');
        } catch(_) {}
        sheet.getRange(rowNum, COL.ATTACH_LIST + 1).setValue('[]');

        var resubCount = parseInt(r[COL.RESUB_COUNT] || 0);
        var existLog   = r[COL.REJECT_LOG] || '';
        var approverNm = r[COL.APPR_START + idx * COL.APPR_COLS + 1] || '-';
        var entry      = '[' + (resubCount + 1) + '차 반려] ' + approverNm + ' / ' +
                         toDateTimeStr(now) + ' / ' + (comment || '사유 없음');
        sheet.getRange(rowNum, COL.REJECT_LOG + 1)
          .setValue(existLog ? existLog + '\n' + entry : entry);

        var drafterEmail = String(r[COL.APPR_START + 2] || '');

        return {
          ok: true,
          message: '반려 처리되었습니다.',
          notify: {
            type: 'reject',
            toEmail: drafterEmail,
            drafter: drafter,
            docNo: docNo,
            subject: subject,
            approverName: String(approverNm),
            comment: comment,
            token: token,
          },
        };
      }

      // 승인 처리
      var nextIdx = idx + 1;
      sheet.getRange(rowNum, COL.APPR_IDX + 1).setValue(nextIdx);

      if (nextIdx < apprCount) {
        // 다음 결재자에게 이메일
        var nb = COL.APPR_START + nextIdx * COL.APPR_COLS;
        sheet.getRange(rowNum, COL.STATUS + 1).setValue('결재중 (' + r[nb] + ')');

        return {
          ok: true,
          message: '승인 처리되었습니다.',
          notify: {
            type: 'next_approval',
            approver: { label: String(r[nb]), name: String(r[nb+1]), email: String(r[nb+2]) },
            data: {
              docNo: docNo,
              subject: subject,
              drafter: drafter,
              issueDate: toDateStr(issueDate),
            },
            token: token,
            rowNum: rowNum,
            idx: nextIdx,
          },
        };
      }

      // 최종 승인
      var finalStatus = (docType === 'PRC') ? '최종승인(PRC)' : '최종승인';
      sheet.getRange(rowNum, COL.STATUS + 1).setValue(finalStatus);
      var drafterEmail2 = String(r[COL.APPR_START + 2] || '');

      var completionNotify = {
        type: 'completion',
        toEmail: drafterEmail2,
        drafter: drafter,
        docNo: docNo,
        subject: subject,
        issueDate: toDateStr(issueDate),
        token: token,
        docType: docType,
      };

      // PRC 최종 승인 → 원 REQ 기안자에게 '검수보고서 작성 가능' 안내를 위해
      // 원본 REQ 행에서 기안자/이메일/품의번호/토큰을 락 안에서 미리 조회 (락 밖 발송용)
      if (docType === 'PRC') {
        var reqToken = String(r[COL.PARENT_DOC_ID] || '');
        if (reqToken) {
          var reqRowNum = findRowNumByToken(sheet, reqToken);
          if (reqRowNum > 0) {
            var reqRow = readRow(sheet, reqRowNum);
            completionNotify.reqToken       = reqToken;
            completionNotify.reqDocNo       = String(reqRow[COL.DOC_NO] || '');
            completionNotify.reqDrafter     = String(reqRow[COL.DRAFTER] || '');
            completionNotify.reqDrafterEmail = String(reqRow[COL.APPR_START + 2] || '');
          }
        }
      }

      return {
        ok: true,
        message: '승인 처리되었습니다.',
        notify: completionNotify,
      };
    } catch(err) {
      return { ok: false, message: err.toString() };
    }
  });

  if (!lockResult.ok || !lockResult.notify) return lockResult;

  // 락 밖: 이메일 발송 (재시도)
  var n = lockResult.notify;
  try {
    if (n.type === 'reject' && n.toEmail) {
      sendRejectionNoticeWithRetry(n.toEmail, n.drafter, n.docNo, n.subject, n.approverName, n.comment, n.token);
    } else if (n.type === 'next_approval') {
      sendApprovalEmailWithRetry(n.approver, n.data, n.token, n.rowNum, n.idx);
    } else if (n.type === 'completion') {
      // 완료 알림 메일은 수신자 이메일이 있을 때만 발송 (없어도 PDF 생성·큐에는 영향 없음)
      if (n.toEmail) {
        sendCompletionNoticeWithRetry(n.toEmail, n.drafter, n.docNo, n.subject, n.issueDate, n.token);
      } else {
        // 수신자 누락은 조용히 넘기지 않고 관리자에게 알림 — PDF 생성은 계속 진행
        notifyAdminError('완료 알림 수신자 이메일 누락: ' + n.docNo + ' (' + n.docType + ') — 메일은 건너뛰고 PDF 생성은 진행');
      }

      // REQ 최종 승인 → 구매팀 전체 알림 (FR-50, REQ-07)
      if (n.docType === 'REQ') {
        try {
          notifyProcurementTeamOfApproved1st(n.docNo, n.subject, n.drafter, n.token);
        } catch(e) {
          notifyAdminError('구매팀 알림 실패: ' + n.docNo + ' / ' + e.toString());
        }
      }

      // PRC 최종 승인 → 원 REQ 기안자에게 '구매팀 결재 완료 & 검수보고서 작성 가능' 안내
      // (검수보고서 제출이 실제로 열리는 시점 = PRC 최종 승인. 원 기안자는 이 시점 알림이 없었음)
      if (n.docType === 'PRC') {
        if (n.reqDrafterEmail) {
          try {
            sendInspReadyNoticeWithRetry(n.reqDrafterEmail, n.reqDrafter, n.reqDocNo, n.subject, n.docNo);
          } catch(e) {
            notifyAdminError('검수보고서 안내 메일 실패: ' + n.docNo + ' / ' + e.toString());
          }
        } else {
          notifyAdminError('검수보고서 안내 수신자(원 REQ 기안자) 이메일 누락: ' + n.docNo + ' — 안내 메일 건너뜀');
        }
      }

      // PDF 합성 + FINAL 이동은 비동기 큐로 위임 (NFR-07 응답성)
      // REQ: pdf-only job (PO 번호 확정 전이라 이동 안 함)
      // PRC: pdf + 통합 이동 job
      // ※ 이메일 수신자 유무와 무관하게 항상 등록 (필수 산출물이므로 메일 조건에서 분리)
      try {
        var jobType = (n.docType === 'PRC') ? 'pdf_and_consolidate' : 'pdf_only';
        enqueueJob({
          type:    jobType,
          token:   n.token,
          docNo:   n.docNo,
          docType: n.docType,
        });
      } catch(e) {
        // 큐 등록 실패는 치명적 — 관리자에게 즉시 알림
        notifyAdminError('큐 등록 실패: ' + n.docNo + ' (' + n.docType + ') / ' + e.toString());
      }
    }
  } catch(e) {
    notifyAdminError('이메일 발송 실패: ' + n.docNo + ' / ' + e.toString());
    return Object.assign({}, lockResult, {
      warning: true,
      message: lockResult.message + ' (이메일 발송 실패 — 관리자 알림 전송)',
    });
  }

  return { ok: lockResult.ok, message: lockResult.message };
}


// ================================================================
// 9. 재상신
// ================================================================

function resubmitRequisitionForClient(payload) {
  return _resubmitCore(payload);
}

function _resubmitCore(payload) {
  // 사전 검증
  try {
    validateTotalAttachmentSize((payload.data && payload.data.attachments) || []);
  } catch(e) {
    return { ok: false, message: e.message };
  }

  var lockResult = withLock(function() {
    try {
      var token = payload.token;
      var data  = payload.data;

      var ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet  = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
      var rowNum = findRowNumByToken(sheet, token);

      if (rowNum < 0) return { ok: false, message: '유효하지 않은 토큰입니다.' };
      var r = readRow(sheet, rowNum);
      if (r[COL.STATUS] === '폐기') return { ok: false, message: '폐기된 품의서는 재상신할 수 없습니다.' };

      // 품의번호 강제 유지
      data.docNo = String(r[COL.DOC_NO] || '');

      var resubCount   = parseInt(r[COL.RESUB_COUNT] || 0) + 1;
      var newToken     = Utilities.getUuid();
      var itemsSummary = buildItemsSummary(data.items);
      var totalAmt     = calcTotal(data.items);
      var totalVat     = totalAmt + computeVat(totalAmt, data.items);
      var approvers    = data.approvers || [];

      // 폴더
      var folderId = r[COL.DRIVE_ID];
      var folder;
      try {
        folder = DriveApp.getFolderById(folderId);
      } catch(e) {
        folder = getOrCreateFolder(data.docNo, data.issueDate, 'staging');
        sheet.getRange(rowNum, COL.DRIVE_ID + 1).setValue(folder.getId());
      }

      // 기본정보 일괄 업데이트 (연속 컬럼 묶음)
      // C(3) ~ Q(17): 발행일자~비고
      sheet.getRange(rowNum, 3, 1, 15).setValues([[
        data.issueDate, data.drafter, data.dept, data.subject, data.purpose,
        data.vendorName, data.vendorEmail, data.vendorContact, data.vendorPhone,
        data.deliveryDate, data.deliveryAddr,
        itemsSummary, totalAmt, totalVat, data.remarks,
      ]]);

      // 결재 관리 컬럼 일괄
      // R(18) ~ Z(26): 상태~이동상태
      sheet.getRange(rowNum, COL.STATUS + 1, 1, 9).setValues([[
        '재상신',          // R
        newToken,          // S
        folderId,          // T (유지)
        approvers.length,  // U
        0,                 // V
        resubCount,        // W
        r[COL.REJECT_LOG] || '', // X (유지)
        '[]',              // Y (락 밖에서 채움)
        'STAGING',         // Z
      ]]);

      // 구매 조건 갱신 (PRC 재상신 시 폼 입력값 반영, REQ는 빈 값)
      // 기존 값을 폼이 보내지 않으면 보존하도록 fallback 처리
      sheet.getRange(rowNum, COL.PURCHASE_METHOD + 1, 1, 2).setValues([[
        String(data.purchaseMethod != null ? data.purchaseMethod : (r[COL.PURCHASE_METHOD] || '')),
        String(data.paymentInfo    != null ? data.paymentInfo    : (r[COL.PAYMENT_INFO]   || '')),
      ]]);

      // 결재자 블록 리셋 (한 번에 setValues)
      var apprRow = [];
      approvers.forEach(function(ap) {
        apprRow.push(ap.label, ap.name, ap.email, '대기', '', '');
      });
      if (apprRow.length > 0) {
        sheet.getRange(rowNum, COL.APPR_START + 1, 1, apprRow.length).setValues([apprRow]);
      }

      return {
        ok: true,
        rowNum: rowNum,
        newToken: newToken,
        folderId: folder.getId(),
        resubCount: resubCount,
      };
    } catch(err) {
      return { ok: false, message: err.toString() };
    }
  });

  if (!lockResult.ok) return lockResult;

  // 락 밖: 첨부 업로드
  var folder;
  try { folder = DriveApp.getFolderById(lockResult.folderId); }
  catch(e) { return { ok: false, message: '폴더 접근 실패: ' + e.toString() }; }

  var savedFiles = [];
  try {
    savedFiles = saveAttachmentsToFolder(folder, payload.data.attachments || []);
  } catch(e) {
    rollbackFiles(savedFiles);
    return { ok: false, message: '첨부파일 업로드 실패: ' + e.toString() };
  }

  // 시트에 첨부 메타 기록
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
  sheet.getRange(lockResult.rowNum, COL.ATTACH_LIST + 1).setValue(JSON.stringify(savedFiles));

  // 재상신 요약 파일
  try {
    var label = lockResult.resubCount + '차_재상신_' +
                Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    var itemsSummary = buildItemsSummary(payload.data.items);
    var totalAmt = calcTotal(payload.data.items);
    var totalVat = totalAmt + computeVat(totalAmt, payload.data.items);
    createFileInFolder(folder, payload.data.docNo + '_' + label + '.txt',
      buildSummaryText(payload.data, itemsSummary, totalAmt, totalVat, payload.data.approvers || []),
      MimeType.PLAIN_TEXT);
  } catch(_) {}

  // 이메일
  var approvers = payload.data.approvers || [];
  if (approvers.length > 0 && approvers[0].email) {
    try {
      sendApprovalEmailWithRetry(approvers[0], payload.data, lockResult.newToken, lockResult.rowNum, 0);
    } catch(e) {
      notifyAdminError('이메일 발송 실패 (재상신): ' + payload.data.docNo + ' / ' + e.toString());
    }
  }

  return {
    ok: true,
    message: '재상신 완료. 결재자에게 이메일이 발송되었습니다.',
    uploaded: savedFiles.length,
  };
}

// ================================================================
// 10. 폐기
// ================================================================

function discardRequisitionForClient(payload) {
  return _discardCore(payload);
}

function _discardCore(payload) {
  return withLock(function() {
    try {
      var token  = payload.token;
      var reason = payload.reason || '기안자 폐기';

      var ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet  = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
      var rowNum = findRowNumByToken(sheet, token);

      if (rowNum < 0) return { ok: false, message: '유효하지 않은 토큰입니다.' };
      var r = readRow(sheet, rowNum);
      if (r[COL.STATUS] === '폐기') return { ok: false, message: '이미 폐기된 품의서입니다.' };

      // 상태/토큰 일괄
      sheet.getRange(rowNum, COL.STATUS + 1, 1, 2).setValues([[
        '폐기',
        'DISCARDED_' + Utilities.getUuid(),
      ]]);

      // Drive 폴더 휴지통 이동
      var folderId = r[COL.DRIVE_ID];
      if (folderId) {
        try {
          DriveApp.getFolderById(folderId).setTrashed(true);
        } catch(_) {}
      }

      var existLog    = r[COL.REJECT_LOG] || '';
      var discardLine = '[폐기] ' + toDateTimeStr(new Date()) + ' / ' + reason;
      sheet.getRange(rowNum, COL.REJECT_LOG + 1)
        .setValue(existLog ? existLog + '\n' + discardLine : discardLine);

      return { ok: true, message: '품의서가 폐기 처리되었습니다.' };
    } catch(err) {
      return { ok: false, message: err.toString() };
    }
  });
}


// ================================================================
// 10-1. 관리자 A-1: 강제 폐기 [v2.0 신규]
// ================================================================
// - 진행 중 REQ/PRC를 관리자 권한으로 강제 폐기
// - 폐기 대상: 모든 진행 상태 (단, '최종승인(PRC)' = 보호 대상은 거부)
// - PRC 폐기 시 부모 REQ 락 자동 해제 (구매팀 기안 문서로 환원)
// - 알림 이메일: 기안자 + 점유자 (관리자 본인 제외, 중복 통합)
// - 시스템로그: ADMIN_FORCE_DISCARD 기록
// ================================================================

/**
 * 폐기 가능 문서 목록 조회 (관리자 페이지 로딩 시 호출)
 * @param {Object} payload
 *   - dateFrom    {string=} 'YYYY-MM-DD'  필터: 제출일 시작
 *   - dateTo      {string=} 'YYYY-MM-DD'  필터: 제출일 종료
 *   - requester   {string=}               필터: 기안자 이메일 또는 이름 부분일치
 *   - status      {string=}               필터: 상태 정확일치 (예: '결재중')
 *   - docNoLike   {string=}               필터: 문서번호 부분일치 (예: 'REQ-2026-01')
 *   - hasFilter   {boolean}               필터 적용 여부 (true면 최대 500건, false면 100건)
 * @returns {Object} { ok, docs: [...], truncated, totalScanned }
 */
function adminListDiscardableForClient(payload) {
  return _adminListDiscardableCore(payload || {});
}

function _adminListDiscardableCore(payload) {
  try {
    // 권한 체크 + 거부 시 자동 로그
    assertAdminWithLog('admin_list_discardable', payload);

    var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
    ensureHeaders(sheet);

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { ok: true, docs: [], truncated: false, totalScanned: 0 };
    }

    // 전체 데이터 한 번에 읽기 (헤더 제외)
    var allRows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    // 필터 파싱
    var dateFrom  = String(payload.dateFrom  || '').trim();
    var dateTo    = String(payload.dateTo    || '').trim();
    var requester = String(payload.requester || '').trim().toLowerCase();
    var status    = String(payload.status    || '').trim();
    var docNoLike = String(payload.docNoLike || '').trim().toLowerCase();
    var hasFilter = !!(dateFrom || dateTo || requester || status || docNoLike);
    var maxRows   = hasFilter ? 500 : 100;

    // dateTo는 해당일 23:59:59까지 포함
    var dateFromObj = dateFrom ? new Date(dateFrom + 'T00:00:00') : null;
    var dateToObj   = dateTo   ? new Date(dateTo   + 'T23:59:59') : null;

    // 필터링 + 폐기 가능 여부 판정
    var matched = [];
    for (var i = 0; i < allRows.length; i++) {
      var r = allRows[i];
      var rowStatus = String(r[COL.STATUS] || '');
      var docType   = String(r[COL.DOC_TYPE] || 'REQ');

      // 이미 폐기된 문서는 제외 (목록에서 숨김)
      if (rowStatus === '폐기') continue;

      // 보호 대상은 목록에는 포함하되 클라이언트에서 비활성화 표시
      var isProtected = (docType === 'PRC' && rowStatus === '최종승인(PRC)');

      // 필터 적용
      var ts = r[COL.TIMESTAMP];
      var tsDate = (ts instanceof Date) ? ts : (ts ? new Date(ts) : null);
      if (dateFromObj && (!tsDate || tsDate < dateFromObj)) continue;
      if (dateToObj   && (!tsDate || tsDate > dateToObj))   continue;

      if (status && rowStatus !== status) continue;

      if (docNoLike) {
        var docNoStr = String(r[COL.DOC_NO] || '').toLowerCase();
        if (docNoStr.indexOf(docNoLike) < 0) continue;
      }

      if (requester) {
        var drafterStr = String(r[COL.DRAFTER] || '').toLowerCase();
        var emailStr   = String(r[COL.APPR_START + 2] || '').toLowerCase();
        if (drafterStr.indexOf(requester) < 0 && emailStr.indexOf(requester) < 0) continue;
      }

      matched.push({
        rowIndex: i + 2,  // 시트 행 번호 (1-based, 헤더 제외)
        token:           String(r[COL.TOKEN] || ''),
        docNo:           String(r[COL.DOC_NO] || ''),
        docType:         docType,
        timestamp:       tsDate ? toDateTimeStr(tsDate) : '',
        timestampMs:     tsDate ? tsDate.getTime() : 0,
        drafter:         String(r[COL.DRAFTER] || ''),
        requesterEmail:  String(r[COL.APPR_START + 2] || ''),
        vendor:          String(r[COL.VENDOR_NAME] || ''),
        amount:          Number(r[COL.TOTAL_AMT] || 0),
        status:          rowStatus,
        claimedBy:       String(r[COL.CLAIMED_BY] || ''),
        claimedAt:       r[COL.CLAIMED_AT] ? toDateTimeStr(r[COL.CLAIMED_AT]) : '',
        parentToken:     String(r[COL.PARENT_DOC_ID] || ''),
        isProtected:     isProtected,
      });
    }

    // 정렬: TIMESTAMP 내림차순
    matched.sort(function(a, b) { return b.timestampMs - a.timestampMs; });

    // 잘림 표시 (결과 더 있음)
    var truncated = matched.length > maxRows;
    var result = truncated ? matched.slice(0, maxRows) : matched;

    // PRC의 경우 부모 REQ 정보 보강 (PRC 폐기 시 부모 락 해제 안내용)
    var parentTokens = {};
    result.forEach(function(d) {
      if (d.docType === 'PRC' && d.parentToken) parentTokens[d.parentToken] = true;
    });
    if (Object.keys(parentTokens).length > 0) {
      var parentMap = {};
      for (var j = 0; j < allRows.length; j++) {
        var pToken = String(allRows[j][COL.TOKEN] || '');
        if (parentTokens[pToken]) {
          parentMap[pToken] = {
            parentDocNo:    String(allRows[j][COL.DOC_NO] || ''),
            parentClaimedBy: String(allRows[j][COL.CLAIMED_BY] || ''),
          };
        }
      }
      result.forEach(function(d) {
        if (d.docType === 'PRC' && d.parentToken && parentMap[d.parentToken]) {
          d.parentDocNo     = parentMap[d.parentToken].parentDocNo;
          d.parentClaimedBy = parentMap[d.parentToken].parentClaimedBy;
        }
      });
    }

    return {
      ok: true,
      docs: result,
      truncated: truncated,
      totalScanned: allRows.length,
      maxRows: maxRows,
    };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

/**
 * 강제 폐기 실행
 * @param {Object} payload
 *   - token   {string} 폐기 대상 문서 토큰 (필수)
 *   - reason  {string} 폐기 사유 (필수, 시스템로그에 기록)
 * @returns {Object} { ok, message, restoredParent }
 */
function adminForceDiscardForClient(payload) {
  return _adminForceDiscardCore(payload || {});
}

function _adminForceDiscardCore(payload) {
  // 락 진입 전 권한 1차 체크 (실패 시 락 점유 낭비 방지)
  try {
    assertAdminWithLog(AUDIT_EVENT.ADMIN_FORCE_DISCARD, { token: payload.token });
    requireAdminReason(payload.reason);
  } catch (e) {
    return { ok: false, message: e.message };
  }

  // 락 안: 권한 재확인 + 시트 변경
  var lockResult = withLock(function() {
    try {
      var actor = getActiveUserEmail();
      if (!isAdminUser(actor)) {
        return { ok: false, message: '관리자 권한이 필요합니다.' };
      }

      var token  = String(payload.token  || '').trim();
      var reason = String(payload.reason || '').trim();
      if (!token)  return { ok: false, message: '폐기 대상 토큰이 필요합니다.' };
      if (!reason) return { ok: false, message: '폐기 사유가 필요합니다.' };

      var ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet  = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
      ensureHeaders(sheet);

      var rowNum = findRowNumByToken(sheet, token);
      if (rowNum < 0) return { ok: false, message: '대상 문서를 찾을 수 없습니다.' };

      var r = readRow(sheet, rowNum);
      var oldStatus = String(r[COL.STATUS] || '');
      var docType   = String(r[COL.DOC_TYPE] || 'REQ');

      // 이미 폐기된 문서 거부
      if (oldStatus === '폐기') {
        return { ok: false, message: '이미 폐기된 문서입니다.' };
      }

      // 보호 대상 거부: 최종승인(PRC)
      if (docType === 'PRC' && oldStatus === '최종승인(PRC)') {
        return { ok: false, message: '최종승인(PRC) 상태의 문서는 강제 폐기할 수 없습니다. (보호 대상)' };
      }

      // 폐기 처리 준비
      var docNo          = String(r[COL.DOC_NO] || '');
      // 기안자 이메일: 결재자 블록의 첫 번째 슬롯(기안자)의 email 위치
      //   APPR_START + 0: 'label' (예: '기안자')
      //   APPR_START + 1: 'name'
      //   APPR_START + 2: 'email'  ← 여기
      var requesterEmail = String(r[COL.APPR_START + 2] || '');
      var drafter        = String(r[COL.DRAFTER] || '');
      var claimedBy      = String(r[COL.CLAIMED_BY] || '');
      var folderId       = String(r[COL.DRIVE_ID] || '');
      var parentToken    = String(r[COL.PARENT_DOC_ID] || '');

      // 부모 REQ 정보 (PRC 폐기 시 환원 처리용)
      var parentInfo = null;
      if (docType === 'PRC' && parentToken) {
        var parentRowNum = findRowNumByToken(sheet, parentToken);
        if (parentRowNum > 0) {
          var pr = readRow(sheet, parentRowNum);
          parentInfo = {
            rowNum:         parentRowNum,
            docNo:          String(pr[COL.DOC_NO] || ''),
            status:         String(pr[COL.STATUS] || ''),
            claimedBy:      String(pr[COL.CLAIMED_BY] || ''),
            requesterEmail: String(pr[COL.APPR_START + 2] || ''),
          };
        }
      }

      // 1) 대상 행: 상태 '폐기' + 토큰 무효화
      sheet.getRange(rowNum, COL.STATUS + 1, 1, 2).setValues([[
        '폐기',
        'DISCARDED_' + Utilities.getUuid(),
      ]]);

      // 2) 락 해제 (REQ가 점유 중이었다면)
      if (claimedBy) {
        sheet.getRange(rowNum, COL.CLAIMED_BY + 1, 1, 2).setValues([['', '']]);
      }

      // 3) REJECT_LOG 추가 (관리자 강제폐기로 구분)
      var existLog    = r[COL.REJECT_LOG] || '';
      var discardLine = '[관리자 강제폐기] ' + toDateTimeStr(new Date()) + ' / ' + actor + ' / ' + reason;
      sheet.getRange(rowNum, COL.REJECT_LOG + 1)
        .setValue(existLog ? existLog + '\n' + discardLine : discardLine);

      // 4) PRC 폐기 시: 부모 REQ 환원 (구매팀 기안 문서로 복귀)
      //    PRC 라이프사이클별로 부모 REQ가 처한 상태가 다름:
      //    - PRC_DRAFT (제출 전): 부모 REQ는 status='최종승인', claimedBy=점유자
      //      → claimedBy만 비우면 인박스 복귀
      //    - 결재중(PRC) / 결재중 (결재N): 부모 REQ는 status='PRC생성됨', claimedBy 비어있음
      //      (PRC 제출 시점에 _submitPrcCore가 두 컬럼 모두 처리)
      //      → status를 '최종승인'으로 환원해야 인박스 필터(status==='최종승인')에 걸림
      if (parentInfo) {
        if (parentInfo.status === 'PRC생성됨') {
          sheet.getRange(parentInfo.rowNum, COL.STATUS + 1).setValue('최종승인');
        }
        if (parentInfo.claimedBy) {
          sheet.getRange(parentInfo.rowNum, COL.CLAIMED_BY + 1, 1, 2).setValues([['', '']]);
        }
      }

      // 락 안에서는 시트 변경만 수행하고, 결과 객체 반환
      // (Drive/이메일/시스템로그는 락 밖에서 처리)
      return {
        ok: true,
        message: '문서가 강제 폐기되었습니다.',
        _ctx: {
          actor:           actor,
          docNo:           docNo,
          docType:         docType,
          token:           token,
          oldStatus:       oldStatus,
          reason:          reason,
          requesterEmail:  requesterEmail,
          drafter:         drafter,
          claimedBy:       claimedBy,
          folderId:        folderId,
          parentInfo:      parentInfo,
        },
      };
    } catch (err) {
      return { ok: false, message: err.toString() };
    }
  });

  if (!lockResult || !lockResult.ok) return lockResult;

  // ─── 락 밖 처리 ───────────────────────────────
  var ctx = lockResult._ctx;
  var restoredParent = null;

  try {
    // 5) Drive 폴더 휴지통 이동 (실패해도 폐기 처리는 계속)
    if (ctx.folderId) {
      try {
        DriveApp.getFolderById(ctx.folderId).setTrashed(true);
      } catch(_) {}
    }

    // 6) 알림 이메일 발송
    _sendAdminDiscardNotifications(ctx);

    // 7) 시스템로그 기록
    var payloadLog = {
      oldStatus: ctx.oldStatus,
      newStatus: '폐기',
      docType:   ctx.docType,
    };
    if (ctx.claimedBy) {
      payloadLog.claimedByCleared = ctx.claimedBy;
    }
    if (ctx.parentInfo) {
      payloadLog.parentReqRestored = {
        parentDocNo:        ctx.parentInfo.docNo,
        oldParentStatus:    ctx.parentInfo.status,
        newParentStatus:    ctx.parentInfo.status === 'PRC생성됨' ? '최종승인' : ctx.parentInfo.status,
        parentClaimedBy:    ctx.parentInfo.claimedBy,
      };
      restoredParent = {
        docNo:     ctx.parentInfo.docNo,
        claimedBy: ctx.parentInfo.claimedBy,
      };
    }

    writeAuditLog({
      eventType:   AUDIT_EVENT.ADMIN_FORCE_DISCARD,
      actor:       ctx.actor,
      actorRole:   'admin',
      docNo:       ctx.docNo,
      docToken:    ctx.token,
      docType:     ctx.docType,
      targetUser:  ctx.requesterEmail,
      reason:      ctx.reason,
      payload:     payloadLog,
    });
  } catch (err) {
    // 락 밖 처리 실패는 폐기 자체를 실패로 만들지 않음 (시트 변경은 이미 커밋됨)
    console.error('[A-1] 락 밖 처리 실패: ' + err.toString());
  }

  // 결과 반환 (클라이언트에서 환원된 부모 REQ 시각화에 사용)
  return {
    ok:             true,
    message:        '문서가 강제 폐기되었습니다.',
    restoredParent: restoredParent,
  };
}

/**
 * 강제 폐기 알림 이메일 발송 (락 밖에서 호출)
 * - 기안자 + 점유자 (모두 다른 사람일 때 각각 발송)
 * - 관리자 본인은 발송 대상에서 제외
 * - 동일 이메일이 여러 사유로 대상이면 1통에 사유 통합
 *
 * @param {Object} ctx _adminForceDiscardCore에서 만든 컨텍스트
 */
function _sendAdminDiscardNotifications(ctx) {
  var notifyMap = {};  // email -> { to, reasons: [] }

  function addNotify(email, reasonText) {
    if (!email) return;
    if (email === ctx.actor) return;  // 관리자 본인은 제외
    if (!notifyMap[email]) {
      notifyMap[email] = { to: email, reasons: [] };
    }
    notifyMap[email].reasons.push(reasonText);
  }

  // 기안자 (항상 후보)
  addNotify(ctx.requesterEmail, '당신이 기안한 ' + (ctx.docType === 'PRC' ? '구매' : '품의') + ' 문서가 강제 폐기됨');

  // REQ 락 점유자 (락 점유 중이었을 때)
  if (ctx.claimedBy) {
    addNotify(ctx.claimedBy, '구매 문서 작업을 위해 점유 중이던 품의가 폐기되어 락이 자동 해제됨');
  }

  // PRC 폐기 시 부모 REQ 점유자 (=PRC 작성자와 시스템상 동일)
  if (ctx.docType === 'PRC' && ctx.parentInfo && ctx.parentInfo.claimedBy) {
    addNotify(ctx.parentInfo.claimedBy,
      '진행 중이던 구매 문서가 폐기되어 부모 품의(' + ctx.parentInfo.docNo + ') 락이 자동 해제되고 인박스로 복귀됨');
  }

  var tsStr = toDateTimeStr(new Date());

  Object.keys(notifyMap).forEach(function(email) {
    var n = notifyMap[email];
    try {
      var reasonsBlock = n.reasons.map(function(r, i) { return '  ' + (i + 1) + ') ' + r; }).join('\n');
      var subject = '[구매결재] ' + ctx.docNo + ' 문서가 관리자에 의해 강제 폐기되었습니다';
      var body =
        '안녕하세요.\n\n' +
        '아래 문서가 관리자에 의해 강제 폐기되었음을 알려드립니다.\n\n' +
        '▶ 문서번호: ' + ctx.docNo + ' (' + ctx.docType + ')\n' +
        '▶ 처리자: ' + ctx.actor + '\n' +
        '▶ 처리 일시: ' + tsStr + '\n' +
        '▶ 폐기 사유: ' + ctx.reason + '\n\n' +
        '▶ 귀하와 관련된 영향:\n' + reasonsBlock + '\n\n' +
        '문의사항은 관리자에게 연락 바랍니다.\n\n' +
        'InLC 구매결재 시스템';

      // GmailApp 사용 — MailApp은 From 별칭(CONFIG.MAIL_FROM)을 지원하지 않는다
      GmailApp.sendEmail(n.to, subject, body, _mailOptions());
    } catch (err) {
      console.error('[A-1] 알림 이메일 발송 실패 (' + n.to + '): ' + err.toString());
      // 실패해도 다른 수신자는 계속 발송
    }
  });
}


// ================================================================
// A-4. 관리자 강제 락 해제 (Admin Force Release Lock)
// ================================================================
//
// 【설계 결정사항】
//   Q1 (함수 구조): 독립 함수 신설 — _releaseClaimCore와 분리하여
//                   트랜잭션 무결성 확보 (점유자 정보 캡처와 락 해제를 동일 락 안에서)
//   Q2 (손실 처리): 관리자 화면 경고 다이얼로그(HTML) + 점유자 이메일 통지
//   Q3 (알림 대상): 원 점유자에게만 (관리자 본인은 자동 제외)
//   Q4 (허용 범위): DOC_TYPE='REQ'로만 제한 (방어적 — PRC에 락이 있으면 이상 신호)
//
// 【공통 패턴 (A-1과 동일)】
//   1. 권한 검증: assertAdminWithLog + 락 안 getActiveUserEmail() 재확인
//   2. 사유 입력: requireAdminReason(reason) 필수
//   3. 시스템로그: writeAuditLog({ADMIN_FORCE_RELEASE_LOCK}) 락 밖 기록
//   4. 알림 메일: 점유자 (관리자 본인 제외)
//   5. UI 진입: ?action=admin&fn=force_release_lock
//   6. 사이드바: Procurement_Home.html admin 섹션
//   7. OAuth scope: 추가 불필요 (기존 scope로 충분)
// ================================================================


/**
 * 강제 해제 가능한 락 목록 조회
 * - 현재 CLAIMED_BY가 채워진 REQ만 반환 (DOC_TYPE='REQ' + 상태 '최종승인')
 * - TTL 경과 여부를 함께 계산하여 클라이언트에 제공 (만료 항목 시각 구분용)
 *
 * @param {Object} payload
 *   - docNoLike   {string=} 문서번호 부분 일치 필터
 *   - claimerLike {string=} 점유자 이메일 부분 일치 필터
 *   - staleOnly   {bool=}   TTL 만료 항목만 표시
 * @returns {Object} { ok, locks: [...], totalScanned, ttlHours }
 */
function adminListClaimedLocksForClient(payload) {
  return _adminListClaimedLocksCore(payload || {});
}

function _adminListClaimedLocksCore(payload) {
  try {
    // 권한 체크 + 거부 시 자동 로그
    assertAdminWithLog('admin_list_claimed_locks', payload);

    var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
    ensureHeaders(sheet);

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { ok: true, locks: [], totalScanned: 0, ttlHours: CONFIG.PRC_CLAIM_TTL_HOURS };
    }

    var allRows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    // 필터 파싱
    var docNoLike   = String(payload.docNoLike   || '').trim().toLowerCase();
    var claimerLike = String(payload.claimerLike || '').trim().toLowerCase();
    var staleOnly   = !!payload.staleOnly;

    var ttlMs   = CONFIG.PRC_CLAIM_TTL_HOURS * 3600 * 1000;
    var nowMs   = Date.now();

    var locks = [];
    for (var i = 0; i < allRows.length; i++) {
      var r = allRows[i];

      // 락이 걸려있지 않으면 스킵
      var claimedBy = String(r[COL.CLAIMED_BY] || '').trim();
      if (!claimedBy) continue;

      // Q4 정책: REQ만 대상 (정상 시스템에서 PRC에는 락이 없음)
      var docType = String(r[COL.DOC_TYPE] || 'REQ');
      if (docType !== 'REQ') {
        // PRC에 락이 걸려있다면 이상 상태 — 목록에는 표시하되 별도 플래그
        // (관리자가 인지할 수 있도록. 단, 강제 해제는 백엔드에서 거부됨)
      }

      // 상태 확인
      var rowStatus = String(r[COL.STATUS] || '');

      // TTL 계산
      var claimedAt   = r[COL.CLAIMED_AT];
      var claimedMs   = (claimedAt instanceof Date) ? claimedAt.getTime()
                        : (claimedAt ? new Date(claimedAt).getTime() : 0);
      var elapsedMs   = (claimedMs && !isNaN(claimedMs)) ? (nowMs - claimedMs) : 0;
      var isStale     = claimedMs && !isNaN(claimedMs) && elapsedMs >= ttlMs;

      // staleOnly 필터
      if (staleOnly && !isStale) continue;

      // docNo 필터
      if (docNoLike) {
        var docNoStr = String(r[COL.DOC_NO] || '').toLowerCase();
        if (docNoStr.indexOf(docNoLike) < 0) continue;
      }

      // claimer 필터
      if (claimerLike && claimedBy.toLowerCase().indexOf(claimerLike) < 0) continue;

      locks.push({
        token:        String(r[COL.TOKEN] || ''),
        docNo:        String(r[COL.DOC_NO] || ''),
        docType:      docType,
        status:       rowStatus,
        drafter:      String(r[COL.DRAFTER] || ''),
        vendor:       String(r[COL.VENDOR_NAME] || ''),
        amount:       Number(r[COL.TOTAL_AMT] || 0),
        claimedBy:    claimedBy,
        claimedAt:    claimedAt ? toDateTimeStr(claimedAt) : '',
        claimedAtMs:  claimedMs || 0,
        elapsedHours: claimedMs ? Math.floor(elapsedMs / 3600000) : 0,
        isStale:      isStale,
        isAbnormal:   (docType !== 'REQ'),  // PRC에 락 = 이상 상태
      });
    }

    // 정렬: 점유 시각 오래된 순 (가장 먼저 풀어야 할 것 위로)
    locks.sort(function(a, b) { return a.claimedAtMs - b.claimedAtMs; });

    return {
      ok: true,
      locks: locks,
      totalScanned: allRows.length,
      ttlHours: CONFIG.PRC_CLAIM_TTL_HOURS,
    };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}


/**
 * 강제 락 해제 실행
 * @param {Object} payload
 *   - token   {string} 대상 REQ 토큰 (필수)
 *   - reason  {string} 해제 사유 (필수, 시스템로그에 기록)
 * @returns {Object} { ok, message, previousClaimer }
 */
function adminForceReleaseLockForClient(payload) {
  return _adminForceReleaseLockCore(payload || {});
}

function _adminForceReleaseLockCore(payload) {
  // 락 진입 전 권한 1차 체크 + 사유 검증 (실패 시 락 점유 낭비 방지)
  try {
    assertAdminWithLog(AUDIT_EVENT.ADMIN_FORCE_RELEASE_LOCK, { token: payload.token });
    requireAdminReason(payload.reason);
  } catch (e) {
    return { ok: false, message: e.message };
  }

  // 락 안: 권한 재확인 + 점유자 정보 캡처 + 락 해제 (단일 트랜잭션)
  var lockResult = withLock(function() {
    try {
      var actor = getActiveUserEmail();
      if (!isAdminUser(actor)) {
        return { ok: false, message: '관리자 권한이 필요합니다.' };
      }

      var token  = String(payload.token  || '').trim();
      var reason = String(payload.reason || '').trim();
      if (!token)  return { ok: false, message: '대상 토큰이 필요합니다.' };
      if (!reason) return { ok: false, message: '해제 사유가 필요합니다.' };

      var ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet  = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
      ensureHeaders(sheet);

      var rowNum = findRowNumByToken(sheet, token);
      if (rowNum < 0) return { ok: false, message: '대상 문서를 찾을 수 없습니다.' };

      var r = readRow(sheet, rowNum);
      var docType   = String(r[COL.DOC_TYPE] || 'REQ');
      var status    = String(r[COL.STATUS] || '');

      // Q4 정책: REQ만 허용 (방어적 검증 — PRC에 락이 있으면 이상 신호)
      if (docType !== 'REQ') {
        return {
          ok: false,
          message: '품의 문서의 락만 해제할 수 있습니다. (현재 문서 유형: ' + docType + ')\n'
                 + '구매 문서에 락이 걸려있다면 시스템 이상일 수 있으니 관리자에게 문의하세요.',
        };
      }

      // 폐기된 문서 거부
      if (status === '폐기') {
        return { ok: false, message: '폐기된 문서입니다.' };
      }

      // 점유자 정보 캡처 (락 안에서 — 무결성 보장)
      var claimedBy = String(r[COL.CLAIMED_BY] || '').trim();
      var claimedAt = r[COL.CLAIMED_AT];
      var docNo     = String(r[COL.DOC_NO] || '');

      // 멱등성: 이미 비어있으면 성공 처리 (실패 아님)
      if (!claimedBy) {
        return {
          ok: true,
          skipped: true,
          message: '이미 락이 해제되어 있습니다.',
          previousClaimer: '',
        };
      }

      // 락 비우기 (CLAIMED_BY, CLAIMED_AT 2셀)
      sheet.getRange(rowNum, COL.CLAIMED_BY + 1, 1, 2).setValues([['', '']]);

      // 컨텍스트 반환 (락 밖에서 로그/메일용)
      return {
        ok: true,
        message: '락이 강제 해제되었습니다.',
        _ctx: {
          actor:           actor,
          docNo:           docNo,
          token:           token,
          previousClaimer: claimedBy,
          claimedAt:       claimedAt,
          reason:          reason,
        },
      };
    } catch (err) {
      return { ok: false, message: err.toString() };
    }
  });

  if (!lockResult || !lockResult.ok) return lockResult;

  // 멱등 스킵된 경우 (이미 해제됨) — 로그/메일 없이 그대로 반환
  if (lockResult.skipped) return lockResult;

  // ─── 락 밖 처리 (시스템로그 + 알림 메일) ───────────────────
  var ctx = lockResult._ctx;
  try {
    // 1) 시스템로그 기록
    writeAuditLog({
      eventType:   AUDIT_EVENT.ADMIN_FORCE_RELEASE_LOCK,
      actor:       ctx.actor,
      actorRole:   'admin',
      docNo:       ctx.docNo,
      docToken:    ctx.token,
      docType:     'REQ',
      targetUser:  ctx.previousClaimer,
      reason:      ctx.reason,
      payload: {
        previousClaimer: ctx.previousClaimer,
        claimedAt: (ctx.claimedAt instanceof Date)
                   ? toDateTimeStr(ctx.claimedAt)
                   : String(ctx.claimedAt || ''),
      },
    });

    // 2) 원 점유자에게 통지 메일 (Q3: 점유자만, 관리자 본인 제외)
    _sendForceReleaseLockNotification(ctx);
  } catch (err) {
    // 락 밖 처리 실패는 해제 자체를 실패로 만들지 않음 (시트 변경은 이미 커밋됨)
    console.error('[A-4] 락 밖 처리 실패: ' + err.toString());
  }

  return {
    ok: true,
    message: '락이 강제 해제되었습니다.',
    previousClaimer: ctx.previousClaimer,
  };
}


/**
 * 강제 락 해제 알림 이메일 (락 밖에서 호출)
 * - 원 점유자에게만 발송 (Q3)
 * - 관리자 본인이 점유자였다면 발송 안 함 (자기 자신 제외)
 * - 발송 실패해도 해제 자체는 성공으로 처리 (메일은 부수 효과)
 *
 * @param {Object} ctx _adminForceReleaseLockCore에서 만든 컨텍스트
 */
function _sendForceReleaseLockNotification(ctx) {
  var target = ctx.previousClaimer;
  if (!target) return;
  if (target === ctx.actor) return;  // 관리자 본인은 제외

  var tsStr = toDateTimeStr(new Date());

  try {
    var subject = '[구매결재] 작업 중이던 ' + ctx.docNo + '의 구매 문서 작성 락이 해제되었습니다';
    var body =
      '안녕하세요.\n\n' +
      '귀하가 구매 문서 작성을 위해 점유 중이던 아래 품의 문서의 락이\n' +
      '관리자에 의해 해제되었음을 알려드립니다.\n\n' +
      '▶ 문서번호: ' + ctx.docNo + '\n' +
      '▶ 처리자: ' + ctx.actor + '\n' +
      '▶ 처리 일시: ' + tsStr + '\n' +
      '▶ 해제 사유: ' + ctx.reason + '\n\n' +
      '⚠️ 주의사항\n' +
      '  - 현재 작성 중이던 구매 문서 내용이 있다면 저장되지 않을 수 있습니다.\n' +
      '  - 이 품의는 다시 구매팀 인박스에 노출되어 다른 담당자가\n' +
      '    픽업할 수 있는 상태가 되었습니다.\n' +
      '  - 계속 작업이 필요하시면 인박스에서 다시 픽업해 주세요.\n\n' +
      '문의사항은 관리자에게 연락 바랍니다.\n\n' +
      'InLC 구매결재 시스템';

    // GmailApp 사용 — MailApp은 From 별칭(CONFIG.MAIL_FROM)을 지원하지 않는다
    GmailApp.sendEmail(target, subject, body, _mailOptions());
  } catch (err) {
    console.error('[A-4] 점유자 알림 이메일 발송 실패 (' + target + '): ' + err.toString());
  }
}


// ─────────────────────────────────────────────────────────────────
// A-4 검증 함수 (Apps Script 에디터에서 수동 실행)
// ─────────────────────────────────────────────────────────────────

/**
 * A-4 목록 조회 검증
 * Apps Script 에디터에서 testA4ListLocks 선택 후 실행
 * 기대: 현재 락이 걸린 REQ 목록을 콘솔에 출력
 */
function testA4ListLocks() {
  Logger.log('===== A-4 락 목록 조회 테스트 =====');
  var result = _adminListClaimedLocksCore({});
  Logger.log('ok: ' + result.ok);
  if (result.ok) {
    Logger.log('총 스캔: ' + result.totalScanned + '행');
    Logger.log('락 걸린 항목: ' + result.locks.length + '개');
    Logger.log('TTL: ' + result.ttlHours + '시간');
    result.locks.forEach(function(lk, i) {
      Logger.log('  [' + (i+1) + '] ' + lk.docNo
        + ' | 점유자: ' + lk.claimedBy
        + ' | ' + lk.elapsedHours + 'h 경과'
        + (lk.isStale ? ' (TTL만료)' : '')
        + (lk.isAbnormal ? ' ⚠️PRC이상' : ''));
    });
  } else {
    Logger.log('에러: ' + result.message);
  }
  return result;
}

/**
 * A-4 강제 해제 드라이런 (실제 해제하지 않고 권한/사유 검증만 테스트)
 * 주의: 실제 토큰을 넣으면 실제로 해제됩니다. 테스트 시 주의.
 */
function testA4ForceReleaseDryRun() {
  Logger.log('===== A-4 강제 해제 검증 (사유 누락 케이스) =====');
  // 사유 없이 호출 → requireAdminReason에서 거부되어야 정상
  var result = _adminForceReleaseLockCore({ token: 'NONEXISTENT_TOKEN', reason: '' });
  Logger.log('ok: ' + result.ok + ' (false여야 정상 — 사유 누락)');
  Logger.log('message: ' + result.message);
  return result;
}


// ================================================================
// 11. PRC 락 / 프리필 / 해제 (FR-50, FR-51, FR-52, FR-53)
// ================================================================

/**
 * PRC 생성 락 점유 (FR-51)
 * 동일 REQ에 대해 1명만 진입 가능
 * payload: { parentToken: string }
 */
function claimRequestForClient(payload) { return _claimRequestCore(payload); }

function _claimRequestCore(payload) {
  return withLock(function() {
    try {
      var parentToken = payload.parentToken || payload.parentId;
      if (!parentToken) return { ok: false, message: '원본 토큰이 필요합니다.' };

      var actor = getActiveUserEmail();
      if (!actor) return { ok: false, message: '사용자 인증 실패. Google 계정으로 로그인되어 있어야 합니다.' };
      if (!isProcurementUser(actor)) {
        return { ok: false, message: '구매팀 권한이 없습니다.' };
      }

      var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
      var rowNum = findRowNumByToken(sheet, parentToken);
      if (rowNum < 0) return { ok: false, message: '원본 품의를 찾을 수 없습니다.' };

      var r = readRow(sheet, rowNum);
      if (r[COL.STATUS] !== '최종승인') {
        return { ok: false, message: '1차 결재가 완료된 품의만 픽업할 수 있습니다. (현재 상태: ' + r[COL.STATUS] + ')' };
      }
      if (String(r[COL.DOC_TYPE] || 'REQ') !== 'REQ') {
        return { ok: false, message: '품의 문서만 픽업할 수 있습니다.' };
      }

      var currentClaim = String(r[COL.CLAIMED_BY] || '').trim();
      var claimedAt   = r[COL.CLAIMED_AT];

      // 이미 다른 사람이 점유 중인지 확인 (TTL 체크)
      if (currentClaim && currentClaim !== actor) {
        var staleMs = CONFIG.PRC_CLAIM_TTL_HOURS * 3600 * 1000;
        var claimedTime = (claimedAt instanceof Date) ? claimedAt.getTime() : new Date(claimedAt).getTime();
        if (!isNaN(claimedTime) && (Date.now() - claimedTime) < staleMs) {
          return {
            ok: false,
            message: '다른 담당자가 작업 중입니다.',
            claimedBy: currentClaim,
            claimedAt: toDateTimeStr(claimedAt),
          };
        }
        // TTL 만료 → 락 강제 해제하고 새로 점유
      }

      // 자기 자신이 이미 점유 중인 경우 → 그대로 유지
      var now = new Date();
      sheet.getRange(rowNum, COL.CLAIMED_BY + 1, 1, 2).setValues([[actor, now]]);

      return {
        ok: true,
        message: '품의를 픽업했습니다.',
        parentToken: parentToken,
        claimedBy: actor,
        claimedAt: toDateTimeStr(now),
      };
    } catch(err) {
      return { ok: false, message: err.toString() };
    }
  });
}

/**
 * PRC 락 해제 (FR-PRC-04a, 회수/폐기)
 * payload: { parentToken: string, silent: bool }
 */
function releaseClaimForClient(payload) { return _releaseClaimCore(payload); }

function _releaseClaimCore(payload) {
  return withLock(function() {
    try {
      var parentToken = payload.parentToken;
      var actor = getActiveUserEmail();
      var force = !!payload.force;  // 관리자 강제 해제

      var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
      var rowNum = findRowNumByToken(sheet, parentToken);
      if (rowNum < 0) return { ok: false, message: '원본 품의를 찾을 수 없습니다.' };

      var r = readRow(sheet, rowNum);
      var claimedBy = String(r[COL.CLAIMED_BY] || '').trim();

      if (!claimedBy) {
        return { ok: true, message: '이미 락이 해제되어 있습니다.', skipped: true };
      }

      if (claimedBy !== actor && !force && !isAdminUser(actor)) {
        return { ok: false, message: '본인이 점유한 락만 해제할 수 있습니다.' };
      }

      sheet.getRange(rowNum, COL.CLAIMED_BY + 1, 1, 2).setValues([['', '']]);

      return { ok: true, message: '락이 해제되었습니다.' };
    } catch(err) {
      return { ok: false, message: err.toString() };
    }
  });
}

/**
 * PRC 프리필용 원본 REQ 데이터 조회 (FR-52)
 * @param parentToken — 원본 REQ의 토큰
 */
function getPrefillDataForClient(parentToken) {
  try {
    if (!parentToken) return { ok: false, message: '원본 토큰이 필요합니다.' };

    var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
    var rowNum = findRowNumByToken(sheet, parentToken);
    if (rowNum < 0) return { ok: false, message: '원본 품의를 찾을 수 없습니다.' };

    var r = readRow(sheet, rowNum);
    if (r[COL.STATUS] !== '최종승인') {
      return { ok: false, message: '1차 결재 완료된 품의만 구매 문서 생성이 가능합니다.' };
    }

    var actor = getActiveUserEmail();
    var claimedBy = String(r[COL.CLAIMED_BY] || '').trim();
    if (claimedBy && claimedBy !== actor) {
      return { ok: false, message: '다른 담당자가 점유 중입니다. (' + claimedBy + ')' };
    }

    // 품목 파싱
    var items = parseItemsSummary(String(r[COL.ITEMS] || ''));

    // 첨부 메타
    var attachmentsRaw = parseAttachments(r[COL.ATTACH_LIST]);
    var attachments = attachmentsRaw.map(function(f) {
      return {
        name: String(f.name || ''),
        id:   String(f.id   || ''),
        size: Number(f.size) || 0,
        type: String(f.type || ''),
        viewUrl:     'https://drive.google.com/file/d/' + f.id + '/view',
      };
    });

    var totalSize = 0;
    attachments.forEach(function(a) { totalSize += a.size; });

    // 부모 REQ의 결재자 내역 (PRC 폼에 표시용)
    var parentApprCount = parseInt(r[COL.APPR_COUNT]) || 0;
    var parentApprovers = [];
    for (var i = 0; i < parentApprCount; i++) {
      var b = COL.APPR_START + i * COL.APPR_COLS;
      parentApprovers.push({
        label:       String(r[b]   || ''),
        name:        String(r[b+1] || ''),
        email:       String(r[b+2] || ''),
        status:      String(r[b+3] || '대기'),
        processedAt: toDateTimeStr(r[b+4]) || '',
        comment:     String(r[b+5] || ''),
      });
    }

    return {
      ok: true,
      data: {
        parentToken:   parentToken,
        parentDocNo:   String(r[COL.DOC_NO] || ''),
        parentStatus:  String(r[COL.STATUS] || ''),
        parentSubmitAt: toDateTimeStr(r[COL.SUBMIT_AT]),
        issueDate:     toDateStr(r[COL.ISSUE_DATE]),
        drafter:       String(r[COL.DRAFTER] || ''),
        dept:          String(r[COL.DEPT] || ''),
        subject:       String(r[COL.SUBJECT] || ''),
        purpose:       String(r[COL.PURPOSE] || ''),
        vendorName:    String(r[COL.VENDOR_NAME] || ''),
        vendorEmail:     String(r[COL.VENDOR_EMAIL] || ''),
        vendorContact: String(r[COL.VENDOR_CTC] || ''),
        vendorPhone:   String(r[COL.VENDOR_PHONE] || ''),
        deliveryDate:  toDateStr(r[COL.DELIVERY_DATE]),
        deliveryAddr:  String(r[COL.DELIVERY_ADDR] || ''),
        remarks:       String(r[COL.REMARKS] || ''),
        items:         items,
        attachments:   attachments,
        attachmentsSize: totalSize,
        maxTotalSize:    CONFIG.MAX_TOTAL_SIZE,
        parentApprovers: parentApprovers,
      },
    };
  } catch(err) {
    return { ok: false, message: err.toString() };
  }
}

/**
 * 품목 요약 문자열 → 배열 파싱
 */
function parseItemsSummary(s) {
  var lines = String(s || '').split('\n');
  var items = [];
  lines.forEach(function(line) {
    // 신규 포맷(품목명·규격 '/' 구분): "1. 품목 명 / 규격 / 100개 / 1,000 KRW"
    // qty(\d+개)를 우측 앵커로 삼아 품목명에 공백/슬래시가 있어도 규격과 정확히 분리
    var mNew = line.match(/^\d+\.\s*(.+?)\s*\/\s*(.+?)\s*\/\s*(\d+)개\s*\/\s*([\d,]+)\s*([A-Z]{3})$/);
    if (mNew) { items.push({ name: mNew[1].trim(), spec: mNew[2].trim(), qty: mNew[3], price: mNew[4].replace(/,/g,''), currency: mNew[5] }); return; }
    // 구 포맷(품목명·규격 공백 구분): "1. 품목 규격 / 100개 / 1,000 KRW"
    var m = line.match(/^\d+\.\s*(.+?)\s+(.+?)\s*\/\s*(\d+)개\s*\/\s*([\d,]+)\s*([A-Z]{3})$/);
    if (m) { items.push({ name: m[1], spec: m[2], qty: m[3], price: m[4].replace(/,/g,''), currency: m[5] }); return; }
    // 하위호환 포맷: "... / 1,000원" → KRW (마이그레이션 없이 기존 행 호환)
    var m2 = line.match(/^\d+\.\s*(.+?)\s+(.+?)\s*\/\s*(\d+)개\s*\/\s*([\d,]+)원$/);
    if (m2) { items.push({ name: m2[1], spec: m2[2], qty: m2[3], price: m2[4].replace(/,/g,''), currency: 'KRW' }); }
  });
  return items;
}

/**
 * PRC 제출 (FR-53)
 * payload: { parentToken, docNo, ..., items, approvers, attachments, prcFields: {...}, fieldChanges: [...] }
 */
function submitPrcForClient(payload) { return _submitPrcCore(payload); }

function _submitPrcCore(data) {
  // 0단계: 구매팀 기안자 본인 검증 (본인이 아니면 제출 차단)
  var drafterChk = assertDrafterIsSelf(data);
  if (drafterChk) return drafterChk;

  // 사전 검증 (원본 REQ의 기존 첨부 + 신규 첨부 합산)
  try {
    var existingSize = 0;
    (data.existingAttachments || []).forEach(function(a) { existingSize += Number(a.size) || 0; });
    validateTotalAttachmentSize(data.attachments || [], existingSize);
  } catch(e) {
    return { ok: false, message: e.message };
  }

  var lockResult = withLock(function() {
    try {
      var parentToken = data.parentToken;
      if (!parentToken) return { ok: false, message: '원본 품의 토큰이 필요합니다.' };

      var actor = getActiveUserEmail();
      if (!isProcurementUser(actor)) return { ok: false, message: '구매팀 권한이 없습니다.' };

      var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
      ensureHeaders(sheet);

      var parentRowNum = findRowNumByToken(sheet, parentToken);
      if (parentRowNum < 0) return { ok: false, message: '원본 품의를 찾을 수 없습니다.' };

      var parentRow = readRow(sheet, parentRowNum);
      var claimedBy = String(parentRow[COL.CLAIMED_BY] || '').trim();
      if (claimedBy !== actor) {
        return { ok: false, message: '본인이 점유한 품의만 구매 문서를 생성할 수 있습니다. (현재 점유: ' + (claimedBy || '없음') + ')' };
      }

      // PO No. 중복 검사 (PRC의 docNo로 사용)
      if (isDocNoDuplicated(data.docNo, null)) {
        return { ok: false, message: '이미 사용 중인 PO 번호입니다.' };
      }

      var token        = Utilities.getUuid();
      var folder       = getOrCreateFolder(data.docNo, data.issueDate, 'staging');
      var itemsSummary = buildItemsSummary(data.items);
      var totalAmt     = calcTotal(data.items);
      var totalVat     = totalAmt + computeVat(totalAmt, data.items);
      var approvers    = data.approvers || [];

      var newRow = [
        new Date(), data.docNo, data.issueDate, data.drafter, data.dept,
        data.subject, data.purpose, data.vendorName, data.vendorEmail,
        data.vendorContact, data.vendorPhone, data.deliveryDate, data.deliveryAddr,
        itemsSummary, totalAmt, totalVat, data.remarks,
        '검토중',
        token,
        folder.getId(),
        approvers.length,
        0, 0, '', '[]', 'STAGING',
        'PRC',         // AA
        parentToken,   // AB ParentDocId
        '',            // AC ClaimedBy (PRC는 비움)
        '',            // AD ClaimedAt
        String(data.purchaseMethod || ''),  // AE 구매방법
        String(data.paymentInfo || ''),     // AF 구매조건
      ];
      approvers.forEach(function(ap) {
        newRow.push(ap.label, ap.name, ap.email, '대기', '', '');
      });
      sheet.appendRow(newRow);
      var rowNum = sheet.getLastRow();

      // 부모 REQ의 락 해제 + 상태 갱신
      sheet.getRange(parentRowNum, COL.CLAIMED_BY + 1, 1, 2).setValues([['', '']]);
      sheet.getRange(parentRowNum, COL.STATUS + 1).setValue('PRC생성됨');

      // 변경 이력 기록 (FR-46)
      if (data.fieldChanges && data.fieldChanges.length > 0) {
        var changeLog = '[PRC 생성 시 변경 필드] ' + toDateTimeStr(new Date()) + ' by ' + actor + '\n' +
          data.fieldChanges.map(function(c) {
            return ' - ' + c.field + ': "' + c.from + '" → "' + c.to + '"';
          }).join('\n');
        // 부모 REQ의 반려이력 컬럼에 변경 이력도 함께 누적
        var existLog = parentRow[COL.REJECT_LOG] || '';
        sheet.getRange(parentRowNum, COL.REJECT_LOG + 1)
          .setValue(existLog ? existLog + '\n' + changeLog : changeLog);
      }

      return {
        ok: true,
        rowNum: rowNum,
        token: token,
        folderId: folder.getId(),
      };
    } catch(err) {
      return { ok: false, message: err.toString() };
    }
  });

  if (!lockResult.ok) return lockResult;

  // 락 밖: 첨부 업로드
  var folder;
  try { folder = DriveApp.getFolderById(lockResult.folderId); }
  catch(e) { return { ok: false, message: '폴더 접근 실패: ' + e.toString() }; }

  var savedFiles = [];
  try {
    savedFiles = saveAttachmentsToFolder(folder, data.attachments || []);
  } catch(e) {
    rollbackFiles(savedFiles);
    return { ok: false, message: '구매 문서 첨부파일 업로드 실패: ' + e.toString() };
  }

  // 첨부 메타 + 요약 파일
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
  sheet.getRange(lockResult.rowNum, COL.ATTACH_LIST + 1).setValue(JSON.stringify(savedFiles));

  try {
    var totalAmt = calcTotal(data.items);
    var totalVat = totalAmt + computeVat(totalAmt, data.items);
    var itemsSummary = buildItemsSummary(data.items);
    saveSummaryToDrive(folder, data, itemsSummary, totalAmt, totalVat, data.approvers || []);
  } catch(_) {}

  // 이메일
  var approvers = data.approvers || [];
  if (approvers.length > 0 && approvers[0].email) {
    try {
      sendApprovalEmailWithRetry(approvers[0], data, lockResult.token, lockResult.rowNum, 0);
    } catch(e) {
      notifyAdminError('PRC 이메일 발송 실패: ' + data.docNo + ' / ' + e.toString());
    }
  }

  return {
    ok: true,
    message: '구매 문서 제출 완료. 구매팀 결재자에게 이메일이 발송되었습니다.',
    uploaded: savedFiles.length,
    rowNum: lockResult.rowNum,
  };
}


// ================================================================
// 12. 대시보드 홈 데이터 API (FR §12.2, FR-50)
// ================================================================

/**
 * 홈 화면용 데이터: 사용자 역할별 필요한 행만 골라 반환
 * (전체 시트 1회 로드 → 메모리에서 필터링)
 */
function getHomeDataForClient() {
  try {
    var actor = getActiveUserEmail();
    var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);

    if (sheet.getLastRow() < 2) {
      return {
        ok: true,
        user: { email: actor, isProcurement: isProcurementUser(actor), isAdmin: isAdminUser(actor) },
        myDrafts: [], myPending: [], myDocs: [],
        approvedInbox: [], myPrc: [],
        completedDocs: [],   // [INSP Step 2]
        myInspPending: [], finalDocs: [],   // [INSP Step 5]
      };
    }

    var rows = sheet.getDataRange().getValues();
    var myDrafts = [];       // 작성 중 (DRAFT는 현재 시스템에 없음 → 검토중)
    var myPending = [];      // 내가 결재해야 할 문서
    var myDocs = [];         // 내가 작성한 문서
    var approvedInbox = [];  // 1차 승인된 REQ (구매팀용 인박스)
    var myPrc = [];          // 내가 점유 중인 REQ / 내가 만든 PRC
    var prcFinals = [];      // [INSP Step 2] 최종승인(PRC) 행 (결재 완료 메뉴 소스)
    var reqMap = {};         // [INSP Step 2] REQ token → 메타 (PRC와 조인용)
    var reqPrcFinals = [];   // [INSP Step 5] 최종완료 통합용: 최종승인 REQ + 최종승인(PRC)

    var isProc = isProcurementUser(actor);

    // 부서 단위 가시성: 결재자목록 부서 열 기준으로 같은 부서 기안 문서를 공유
    // - deptMap: 이메일(소문자) → 부서
    // - myDept 가 비어있으면(등록 안 됨/부서 미기재) 부서 공유 없음(=기존 본인 문서만)
    var deptMap = getDeptByEmailMap(ss);
    var myDept  = deptMap[String(actor || '').toLowerCase()] || '';
    function sameDeptAsMe(email) {
      if (!myDept) return false;
      var d = deptMap[String(email || '').toLowerCase()];
      return !!d && d === myDept;
    }

    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      var status   = String(r[COL.STATUS] || '');
      var docType  = String(r[COL.DOC_TYPE] || 'REQ');
      if (status === '폐기') continue;

      // 기본 메타
      var meta = {
        docNo:      String(r[COL.DOC_NO] || ''),
        token:      String(r[COL.TOKEN] || ''),
        subject:    String(r[COL.SUBJECT] || ''),
        drafter:    String(r[COL.DRAFTER] || ''),
        dept:       String(r[COL.DEPT] || ''),
        vendorName: String(r[COL.VENDOR_NAME] || ''),
        amount:     Number(r[COL.TOTAL_AMT]) || 0,
        status:     status,
        docType:    docType,
        issueDate:  toDateStr(r[COL.ISSUE_DATE]),
        submitAt:   toDateTimeStr(r[COL.SUBMIT_AT]),
        parentToken: String(r[COL.PARENT_DOC_ID] || ''),
        claimedBy:  String(r[COL.CLAIMED_BY] || ''),
        claimedAt:  toDateTimeStr(r[COL.CLAIMED_AT]),
      };

      var drafterEmail = String(r[COL.APPR_START + 2] || '');
      var apprCount = parseInt(r[COL.APPR_COUNT]) || 0;
      var curIdx    = parseInt(r[COL.APPR_IDX]) || 0;

      // [INSP Step 2] 결재 완료 메뉴용 수집
      // - REQ는 PRC 조인용 메타로 적재
      // - 주의: 이 시스템에서 COL.DRAFTER = 기안자 '성명', 기안자 '이메일' = 기안자 블록(APPR_START+2)
      //   (가시성 매칭은 반드시 이메일끼리 비교)
      if (docType === 'REQ') {
        reqMap[meta.token] = {
          docNo:       meta.docNo,
          drafter:     drafterEmail,   // 기안자 이메일 (기안자 블록)
          drafterName: meta.drafter,   // 기안자 성명 (COL.DRAFTER)
          dept:        meta.dept,
          issueDate:   meta.issueDate,
        };
      }
      if (docType === 'PRC' && status === '최종승인(PRC)') {
        prcFinals.push(meta);
      }

      // 1) 내가 작성한 문서 + 같은 부서 동료가 기안한 문서 (부서 단위 공유)
      //    mine 플래그로 클라이언트에서 '내 문서 / 부서 문서' 구분 표시
      if (drafterEmail === actor) {
        myDocs.push(Object.assign({}, meta, { mine: true }));
      } else if (sameDeptAsMe(drafterEmail)) {
        myDocs.push(Object.assign({}, meta, { mine: false }));
      }

      // 2) 내가 결재 대기인 문서
      if (curIdx < apprCount &&
          (status.indexOf('결재') >= 0 || status === '검토중' || status === '재상신')) {
        var curBase = COL.APPR_START + curIdx * COL.APPR_COLS;
        var curApprEmail = String(r[curBase + 2] || '');
        var curApprStatus = String(r[curBase + 3] || '');
        if (curApprEmail === actor && curApprStatus === '대기') {
          myPending.push(meta);
        }
      }

      // 3) 구매팀: 구매팀 기안 문서 (1차 최종 승인 + 아직 PRC 미생성)
      if (isProc && docType === 'REQ' && status === '최종승인') {
        approvedInbox.push(meta);
      }

      // [INSP Step 5] 최종 완료(통합) 후보 수집: 최종승인 REQ + 최종승인(PRC)
      //  가시성: 구매팀 전체 / 일반은 본인 관련(기안자 또는 결재 참여자) + 같은 부서 기안 문서
      //  REQ는 PRC가 생성되면 status가 'PRC생성됨'으로 바뀌므로(결재 자체는 최종승인 완료)
      //  두 상태를 모두 최종 완료로 취급해야 구매까지 진행된 품의가 목록에서 사라지지 않음
      if ((docType === 'REQ' && (status === '최종승인' || status === 'PRC생성됨')) ||
          (docType === 'PRC' && status === '최종승인(PRC)')) {
        var amMine = (drafterEmail === actor);
        var amParticipant = amMine;
        if (!amParticipant) {
          for (var pa = 0; pa < apprCount; pa++) {
            if (String(r[COL.APPR_START + pa * COL.APPR_COLS + 2] || '').toLowerCase()
                === String(actor || '').toLowerCase()) { amParticipant = true; break; }
          }
        }
        if (isProc || amParticipant || sameDeptAsMe(drafterEmail)) {
          reqPrcFinals.push(Object.assign({}, meta, { drafterEmail: drafterEmail, mine: amMine }));
        }
      }

      // 4) 구매팀: 내가 점유한 REQ 또는 내가 작성한 PRC
      if (isProc) {
        if (docType === 'REQ' && meta.claimedBy === actor) {
          myPrc.push(Object.assign({}, meta, { _type: 'claimed_req' }));
        }
        if (docType === 'PRC' && drafterEmail === actor && status !== '최종승인(PRC)') {
          myPrc.push(Object.assign({}, meta, { _type: 'my_prc' }));
        }
      }
    }

    // APPROVED_1ST는 FIFO (오래된 순) 정렬 (Q-13)
    approvedInbox.sort(function(a, b) {
      return (a.submitAt || '').localeCompare(b.submitAt || '');
    });

    // 내 기안 문서 — 최신순
    myDocs.sort(function(a, b) {
      return (b.submitAt || '').localeCompare(a.submitAt || '');
    });

    // [INSP Step 2] 결재 완료 목록 조립
    // - 가시성: 구매팀은 전체, 일반 사용자는 원본 REQ 기안자 본인 건만 (Q-INSP-05)
    // - 검수보고서 그룹은 INSP.gs의 _getInspGroupsByPrc()가 1회 read로 제공
    //   (INSP 파일 미배포 환경에서도 홈이 죽지 않도록 typeof 가드)
    var completedDocs = [];
    if (prcFinals.length > 0) {
      var inspGroups = (typeof _getInspGroupsByPrc === 'function')
        ? _getInspGroupsByPrc(ss)
        : {};
      for (var c = 0; c < prcFinals.length; c++) {
        var pf  = prcFinals[c];
        var req = reqMap[pf.parentToken] || null;
        var reqDrafter = req ? req.drafter : '';
        var completedMine = (String(reqDrafter).toLowerCase() === String(actor || '').toLowerCase());
        // 가시성: 구매팀 전체 / 본인 기안 / 같은 부서 기안 (검수보고서 제출도 같은 부서 허용)
        if (!isProc && !completedMine && !sameDeptAsMe(reqDrafter)) continue;

        var insps = inspGroups[pf.token] || [];
        // 종결(최종 승인 완료) / 잠금대기(최종 검수가 반려 외 상태로 진행중) 구분
        var inspClosed = false;        // 최종 검수 '최종승인(INSP)' → 완전 종결
        var inspLockedPending = false; // 최종 검수 제출됐으나 결재 진행중 → 추가 제출만 차단
        for (var ic = 0; ic < insps.length; ic++) {
          if (insps[ic].isFinal && insps[ic].status !== '반려') {
            inspLockedPending = true;
            if (insps[ic].status === '최종승인(INSP)') inspClosed = true;
          }
        }

        completedDocs.push({
          reqNo:       req ? req.docNo : '',
          reqToken:    pf.parentToken,
          poNo:        pf.docNo,
          prcToken:    pf.token,
          subject:     pf.subject,
          vendorName:  pf.vendorName,
          drafter:     reqDrafter,
          drafterName: req ? req.drafterName : '',
          dept:        req ? req.dept : pf.dept,
          issueDate:   req ? req.issueDate : pf.issueDate,
          prcFinalAt:  pf.submitAt,
          insps:       insps,
          inspClosed:  inspClosed,
          inspLocked:  inspLockedPending,   // [INSP Step 4 보강] 최종 검수 진행중 추가제출 차단
          mine:        completedMine,        // 내 기안 여부(부서 공유 목록에서 구분 표시용)
          // 검수보고서 제출 권한: 본인 기안 또는 같은 부서(서버 INSP 게이트와 동일 기준)
          canSubmit:   (completedMine || sameDeptAsMe(reqDrafter)),
        });
      }
      // 최신 PRC 순
      completedDocs.sort(function(a, b) {
        return (b.prcFinalAt || '').localeCompare(a.prcFinalAt || '');
      });
    }

    // [INSP Step 5] 내 검수 결재 대기 + 최종 완료(통합)
    var inspMenus = (typeof _getInspMenusForClient === 'function')
      ? _getInspMenusForClient(ss, actor, isProc)
      : { myInspPending: [], inspFinals: [] };

    // 최종 완료 통합: REQ/PRC 최종 + INSP 최종(IS_FINAL) — 단일 리스트, docType으로 구분
    var finalDocs = [];
    reqPrcFinals.forEach(function(m) {
      // 'PRC생성됨' REQ는 결재상 최종승인 완료 건이므로 목록에서는 '최종승인'으로 표기하고
      // 후속 구매 진행 여부는 prcCreated 플래그로 구분
      var prcCreated = (m.docType === 'REQ' && m.status === 'PRC생성됨');
      finalDocs.push({
        docNo: m.docNo, token: m.token, subject: m.subject,
        vendorName: m.vendorName, drafterName: m.drafter, dept: m.dept,
        status: prcCreated ? '최종승인' : m.status,
        docType: m.docType, issueDate: m.issueDate, submitAt: m.submitAt,
        mine: m.mine, prcCreated: prcCreated,
      });
    });
    (inspMenus.inspFinals || []).forEach(function(m) {
      finalDocs.push({
        docNo: m.docNo, token: m.token, subject: m.subject,
        vendorName: m.vendorName, drafterName: m.drafterName, dept: '',
        status: m.status, docType: 'INSP', issueDate: m.receivedDate, submitAt: m.submitAt,
        poNo: m.poNo, reqNo: m.reqNo, seq: m.seq,
        mine: (String(m.drafter || '').toLowerCase() === String(actor || '').toLowerCase()),
      });
    });
    finalDocs.sort(function(a, b) { return (b.submitAt || '').localeCompare(a.submitAt || ''); });

    return {
      ok: true,
      user: {
        email: actor,
        isProcurement: isProc,
        isAdmin: isAdminUser(actor),
        dept: myDept,
      },
      myDrafts:      myDrafts,
      myPending:     myPending,
      myDocs:        myDocs,
      approvedInbox: approvedInbox,
      myPrc:         myPrc,
      completedDocs: completedDocs,   // [INSP Step 2]
      myInspPending: inspMenus.myInspPending,   // [INSP Step 5]
      finalDocs:     finalDocs,                 // [INSP Step 5]
    };
  } catch(err) {
    return { ok: false, message: err.toString() + ' / ' + (err.stack || '') };
  }
}

// ================================================================
// 13. 결재자 뷰어 / 반려 페이지 데이터
// ================================================================

function getRequisitionForViewer(token, urlIdxHint) {
  try {
    var ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet  = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
    var rowNum = findRowNumByToken(sheet, token);
    if (rowNum < 0) return { ok: false, message: '유효하지 않은 토큰입니다.' };

    var r         = readRow(sheet, rowNum);
    var apprCount = parseInt(r[COL.APPR_COUNT]) || 0;
    var items     = parseItemsSummary(String(r[COL.ITEMS] || ''));

    var approvers = [];
    for (var i = 0; i < apprCount; i++) {
      var b = COL.APPR_START + i * COL.APPR_COLS;
      approvers.push({
        label:       String(r[b]   || ''),
        name:        String(r[b+1] || ''),
        email:       String(r[b+2] || ''),
        status:      String(r[b+3] || '대기'),
        processedAt: toDateTimeStr(r[b+4]) || '—',
        comment:     String(r[b+5] || '—'),
      });
    }

    var attachmentsRaw = parseAttachments(r[COL.ATTACH_LIST]);
    var attachments = attachmentsRaw.map(function(f) {
      return {
        name: String(f.name || ''),
        id:   String(f.id   || ''),
        size: Number(f.size) || 0,
        type: String(f.type || ''),
        viewUrl:     'https://drive.google.com/file/d/' + f.id + '/view',
        downloadUrl: 'https://drive.google.com/uc?export=download&id=' + f.id,
      };
    });

    // ── PRC 문서이면 부모 REQ의 결재 내역 + 원본 첨부도 함께 조회 ──
    // (PRC viewer에서 1차 결재 현황 및 원본 REQ 견적서 등을 보여주기 위함)
    var parentApprovers = [];
    var parentDocNo = '';
    var parentAttachments = [];
    var docType = String(r[COL.DOC_TYPE] || 'REQ');
    var parentToken = String(r[COL.PARENT_DOC_ID] || '');
    if (docType === 'PRC' && parentToken) {
      try {
        var parentRowNum = findRowNumByToken(sheet, parentToken);
        if (parentRowNum > 0) {
          var pr = readRow(sheet, parentRowNum);
          var pCount = parseInt(pr[COL.APPR_COUNT]) || 0;
          parentDocNo = String(pr[COL.DOC_NO] || '');
          for (var pi = 0; pi < pCount; pi++) {
            var pb = COL.APPR_START + pi * COL.APPR_COLS;
            parentApprovers.push({
              label:       String(pr[pb]   || ''),
              name:        String(pr[pb+1] || ''),
              status:      String(pr[pb+3] || '대기'),
              processedAt: toDateTimeStr(pr[pb+4]) || '',
              comment:     String(pr[pb+5] || ''),
            });
          }
          // 원본 REQ의 첨부(견적서 등) 메타 조회
          parentAttachments = parseAttachments(pr[COL.ATTACH_LIST]).map(function(f) {
            return {
              name: String(f.name || ''),
              id:   String(f.id   || ''),
              size: Number(f.size) || 0,
              type: String(f.type || ''),
              viewUrl:     'https://drive.google.com/file/d/' + f.id + '/view',
              downloadUrl: 'https://drive.google.com/uc?export=download&id=' + f.id,
            };
          });
        }
      } catch(_) { /* 부모 조회 실패는 viewer 본 기능에 영향 없음 */ }
    }

    // ── myApproverIdx 계산 (b안 + URL idx 우선) ──
    // 진입 경로별 의도:
    //  1) 이메일 링크 (URL에 idx 있음) → 그 단계로 진입 (단, 본인 단계인지 검증)
    //  2) 홈/내 결재 대기 등 (URL idx 없음) → 대기 중인 가장 빠른 본인 단계 자동 선택
    //  3) 모든 단계가 처리된 경우 → 본인이 처리한 마지막 단계로 fallback
    var myEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();
    var myApproverIdx = -1;

    if (urlIdxHint !== null && urlIdxHint !== undefined && urlIdxHint !== '' && !isNaN(urlIdxHint)) {
      var hintIdx = parseInt(urlIdxHint);
      if (hintIdx >= 0 && hintIdx < approvers.length &&
          String(approvers[hintIdx].email || '').toLowerCase() === myEmail) {
        myApproverIdx = hintIdx;
      }
    }

    if (myApproverIdx === -1) {
      for (var k = 0; k < approvers.length; k++) {
        if (String(approvers[k].email || '').toLowerCase() === myEmail &&
            approvers[k].status === '대기') {
          myApproverIdx = k; break;
        }
      }
    }

    if (myApproverIdx === -1) {
      for (var k2 = approvers.length - 1; k2 >= 0; k2--) {
        if (String(approvers[k2].email || '').toLowerCase() === myEmail) {
          myApproverIdx = k2; break;
        }
      }
    }

    return {
      ok: true,
      data: {
        docNo:         String(r[COL.DOC_NO] || ''),
        issueDate:     toDateStr(r[COL.ISSUE_DATE]),
        drafter:       String(r[COL.DRAFTER] || ''),
        dept:          String(r[COL.DEPT] || ''),
        subject:       String(r[COL.SUBJECT] || ''),
        purpose:       String(r[COL.PURPOSE] || ''),
        vendorName:    String(r[COL.VENDOR_NAME] || ''),
        vendorEmail:     String(r[COL.VENDOR_EMAIL] || ''),
        vendorContact: String(r[COL.VENDOR_CTC] || ''),
        vendorPhone:   String(r[COL.VENDOR_PHONE] || ''),
        deliveryDate:  toDateStr(r[COL.DELIVERY_DATE]),
        deliveryAddr:  String(r[COL.DELIVERY_ADDR] || ''),
        remarks:       String(r[COL.REMARKS] || ''),
        status:        String(r[COL.STATUS] || ''),
        // 반려 상태에서 기안자 본인에게만 재상신/폐기 진입 버튼을 노출하기 위한 플래그
        isDrafter:     !!(approvers[0] && String(approvers[0].email || '').toLowerCase() === myEmail),
        rejectHistory: String(r[COL.REJECT_LOG] || ''),
        moveStatus:    String(r[COL.MOVE_STATUS] || ''),
        docType:       String(r[COL.DOC_TYPE] || 'REQ'),
        parentToken:   String(r[COL.PARENT_DOC_ID] || ''),
        items:         items,
        approvers:     approvers,
        attachments:   attachments,
        myApproverIdx: myApproverIdx,
        parentDocNo:   parentDocNo,
        parentApprovers: parentApprovers,
        parentAttachments: parentAttachments,
      }
    };
  } catch (err) {
    return { ok: false, message: 'getRequisitionForViewer 오류: ' + err.toString() + ' / stack: ' + (err.stack || '') };
  }
}

function getRejectData(token) {
  var result = _getRejectDataCore(token);
  return jsonResponse(result);
}
function getRejectDataForClient(token) { return _getRejectDataCore(token); }

function _getRejectDataCore(token) {
  try {
    var ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet  = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
    var rowNum = findRowNumByToken(sheet, token);
    if (rowNum < 0) return { ok: false, message: '유효하지 않은 토큰입니다.' };

    var r = readRow(sheet, rowNum);
    var hist = String(r[COL.REJECT_LOG] || '');
    var lastEntry = hist ? hist.split('\n').filter(function(l){return l.indexOf('[폐기]') < 0 && l.indexOf('차 반려]') > 0;}).pop() : '';

    var rejectApprover = '', rejectTime = '', rejectReason = '';
    if (lastEntry) {
      var parts = lastEntry.replace(/^\[.*?\]\s*/, '').split(' / ');
      rejectApprover = parts[0] || '';
      rejectTime     = parts[1] || '';
      rejectReason   = parts[2] || '';
    }

    var items = parseItemsSummary(String(r[COL.ITEMS] || ''));

    return {
      ok: true,
      data: {
        docNo:         String(r[COL.DOC_NO] || ''),
        issueDate:     toDateStr(r[COL.ISSUE_DATE]),
        drafter:       String(r[COL.DRAFTER] || ''),
        dept:          String(r[COL.DEPT] || ''),
        subject:       String(r[COL.SUBJECT] || ''),
        purpose:       String(r[COL.PURPOSE] || ''),
        vendorName:    String(r[COL.VENDOR_NAME] || ''),
        vendorEmail:     String(r[COL.VENDOR_EMAIL] || ''),
        vendorContact: String(r[COL.VENDOR_CTC] || ''),
        vendorPhone:   String(r[COL.VENDOR_PHONE] || ''),
        deliveryDate:  toDateStr(r[COL.DELIVERY_DATE]),
        deliveryAddr:  String(r[COL.DELIVERY_ADDR] || ''),
        remarks:       String(r[COL.REMARKS] || ''),
        rejectHistory: hist,
        rejectApprover: rejectApprover,
        rejectTime:    rejectTime,
        rejectReason:  rejectReason,
        items:         items,
      }
    };
  } catch(err) {
    return { ok: false, message: err.toString() };
  }
}

function _getDocumentDataCore(payload) {
  try {
    var token = payload.token;
    var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
    var rowNum = findRowNumByToken(sheet, token);
    if (rowNum < 0) return { ok: false, message: '유효하지 않은 토큰입니다.' };

    var r = readRow(sheet, rowNum);
    var apprCount = parseInt(r[COL.APPR_COUNT]) || 0;
    var approvers = [];
    for (var i = 0; i < apprCount; i++) {
      var b = COL.APPR_START + i * COL.APPR_COLS;
      approvers.push({
        label:   String(r[b] || ''),
        name:    String(r[b+1] || ''),
        email:   String(r[b+2] || ''),
        status:  String(r[b+3] || ''),
        date:    toDateStr(r[b+4]),
        comment: String(r[b+5] || ''),
      });
    }

    return {
      ok: true,
      doc: {
        docNo:         String(r[COL.DOC_NO] || ''),
        issueDate:     toDateStr(r[COL.ISSUE_DATE]),
        drafter:       String(r[COL.DRAFTER] || ''),
        dept:          String(r[COL.DEPT] || ''),
        subject:       String(r[COL.SUBJECT] || ''),
        purpose:       String(r[COL.PURPOSE] || ''),
        vendorName:    String(r[COL.VENDOR_NAME] || ''),
        vendorEmail:     String(r[COL.VENDOR_EMAIL] || ''),
        vendorContact: String(r[COL.VENDOR_CTC] || ''),
        vendorPhone:   String(r[COL.VENDOR_PHONE] || ''),
        deliveryDate:  toDateStr(r[COL.DELIVERY_DATE]),
        deliveryAddr:  String(r[COL.DELIVERY_ADDR] || ''),
        itemsSummary:  String(r[COL.ITEMS] || ''),
        totalAmt:      Number(r[COL.TOTAL_AMT]) || 0,
        totalVat:      Number(r[COL.TOTAL_VAT]) || 0,
        remarks:       String(r[COL.REMARKS] || ''),
        status:        String(r[COL.STATUS] || ''),
        resubCount:    Number(r[COL.RESUB_COUNT]) || 0,
        rejectLog:     String(r[COL.REJECT_LOG] || ''),
        approvers:     approvers,
        folderId:      String(r[COL.DRIVE_ID] || ''),
        attachments:   parseAttachments(r[COL.ATTACH_LIST]),
        moveStatus:    String(r[COL.MOVE_STATUS] || ''),
        docType:       String(r[COL.DOC_TYPE] || 'REQ'),
      }
    };
  } catch(err) {
    return { ok: false, message: err.toString() };
  }
}


// ================================================================
// 14. 이메일 (재시도 포함 NFR-05)
// ================================================================

function escapeHtml(str) {
  return String(str == null ? '' : str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/**
 * 발신 옵션 조립 — CONFIG.MAIL_FROM 이 있으면 From 별칭을 적용.
 *  스크립트 실행 계정(David)과 무관하게 메일이 대표 주소로 나가게 한다.
 * @param {Object=} extra 추가 옵션 (htmlBody 등)
 * @returns {Object} GmailApp.sendEmail options
 */
function _mailOptions(extra) {
  var opts = extra || {};
  opts.name = CONFIG.FROM_NAME;
  if (CONFIG.MAIL_FROM) opts.from = CONFIG.MAIL_FROM;
  return opts;
}

/**
 * 별칭(from) 없이 보내는 폴백 발송.
 *  별칭 지정이 실패하는 원인은 여러 가지다 — 별칭 미등록/미인증, OAuth 스코프 부족
 *  (gmail.settings.basic 누락 시 "Specified permissions are not sufficient") 등.
 *  예외 메시지로 원인을 구분하려 들면 문구가 바뀔 때마다 폴백이 조용히 죽으므로,
 *  from 을 쓴 발송이 실패하면 이유를 따지지 않고 한 번은 별칭 없이 재발송한다.
 *  → 최악의 경우에도 발신자만 David로 표시될 뿐, 결재 알림이 유실되지 않는다.
 * @returns {boolean} 폴백 발송 성공 여부
 */
function _sendWithoutAlias(toEmail, subject, plainBody, htmlBody, cause) {
  if (!CONFIG.MAIL_FROM) return false;   // 별칭을 안 썼으면 폴백할 것도 없음
  var opts = { name: CONFIG.FROM_NAME };
  if (htmlBody) opts.htmlBody = htmlBody;
  try {
    GmailApp.sendEmail(toEmail, subject, plainBody, opts);
    Logger.log('[메일] From 별칭(' + CONFIG.MAIL_FROM + ') 실패 → 기본 계정으로 발송함. 원인: ' + cause);
    return true;
  } catch (_) {
    return false;
  }
}

/**
 * Gmail 발송 + 재시도 (3회). From 별칭(CONFIG.MAIL_FROM) 적용.
 *  별칭 발송이 실패하면 별칭 없이 폴백 발송한다(_sendWithoutAlias 참고).
 */
function sendEmailWithRetry(toEmail, subject, plainBody, htmlBody) {
  var lastErr;
  for (var i = 0; i < CONFIG.EMAIL_RETRY_COUNT; i++) {
    try {
      GmailApp.sendEmail(toEmail, subject, plainBody, _mailOptions({ htmlBody: htmlBody }));
      return true;
    } catch(e) {
      lastErr = e;
      if (_sendWithoutAlias(toEmail, subject, plainBody, htmlBody, e)) return true;
      if (i < CONFIG.EMAIL_RETRY_COUNT - 1) {
        Utilities.sleep(CONFIG.EMAIL_RETRY_DELAY_MS);
      }
    }
  }
  throw new Error('이메일 발송 ' + CONFIG.EMAIL_RETRY_COUNT + '회 재시도 실패: ' + (lastErr && lastErr.toString()));
}

function notifyAdminError(message) {
  var admins = CONFIG.ADMIN_NOTIFY_EMAILS || [];
  if (admins.length === 0) {
    try { console.error('[ADMIN ALERT] ' + message); } catch(_) { Logger.log('[ADMIN ALERT] ' + message); }
    return;
  }
  admins.forEach(function(adminEmail) {
    var subj = '[구매시스템 오류] ' + message.substring(0, 80);
    var body = message + '\n\n발생 시각: ' + toDateTimeStr(new Date());
    try {
      GmailApp.sendEmail(adminEmail, subj, body, _mailOptions());
    } catch(e) {
      // 오류 알림 경로는 재시도/throw 없음(재귀 방지) — 별칭 없이 1회만 폴백
      _sendWithoutAlias(adminEmail, subj, body, null, e);
    }
  });
}

/**
 * [점검용] From 별칭이 실제로 동작하는지 확인. GAS 에디터에서 수동 실행.
 *  수신자는 실행자 본인 — 어느 계정으로 실행하든 자기 수신함에서 발신자를 확인할 수 있다.
 */
function sendTestMailFromAlias() {
  var to = getActiveUserEmail() || (CONFIG.ADMIN_NOTIFY_EMAILS || [])[0];
  try {
    GmailApp.sendEmail(to, '[구매시스템] 발신 별칭 테스트',
      'CONFIG.MAIL_FROM = ' + CONFIG.MAIL_FROM + '\n이 메일의 발신자를 확인하세요.', _mailOptions());
    console.log('✅ 별칭 발송 성공 → ' + to + ' 수신함에서 발신자가 ' + CONFIG.MAIL_FROM + ' 인지 확인하세요.');
  } catch (e) {
    var msg = String(e);
    console.log('❌ 별칭 발송 실패: ' + msg);
    if (/permission/i.test(msg)) {
      console.log('→ OAuth 스코프 부족입니다. appsscript.json에 gmail.settings.basic 이 있는지 확인하고,'
        + ' 스코프를 추가했다면 에디터에서 아무 함수나 한 번 실행해 재승인(권한 검토) 절차를 마치세요.');
    } else {
      console.log('→ David Gmail 설정 → 계정 → "다른 주소에서 메일 보내기"에 '
        + CONFIG.MAIL_FROM + ' 등록·인증 여부를 확인하세요.');
    }
    // 실제 운영 경로가 유실 없이 폴백되는지도 함께 검증
    if (_sendWithoutAlias(to, '[구매시스템] 별칭 폴백 테스트',
        '별칭 실패 시 폴백 경로로 발송된 메일입니다.', null, e)) {
      console.log('ℹ️ 폴백 발송은 정상 동작 — 별칭이 안 되어도 결재 메일은 유실되지 않습니다(발신자만 David).');
    }
  }
}

function sendApprovalEmailWithRetry(approver, data, token, rowNum, idx) {
  var approverName  = escapeHtml(approver.name);
  var approverLabel = escapeHtml(approver.label);
  var approverEmail = approver.email;

  var docNo     = escapeHtml(data.docNo || '-');
  var docTitle  = escapeHtml(data.subject || '구매품의서');
  var drafter   = escapeHtml(data.drafter || '-');
  var issueDate = escapeHtml(toDateStr(data.issueDate) || '-');

  var base = CONFIG.WEBAPP_URL +
    '?token=' + encodeURIComponent(token) +
    '&docNo=' + encodeURIComponent(data.docNo || '');

  var viewUrl    = base + '&action=view&idx=' + encodeURIComponent(idx);
  var approveUrl = base + '&action=approve&decision=approve&idx=' + encodeURIComponent(idx);
  var rejectUrl  = base + '&action=approve&decision=reject&idx=' + encodeURIComponent(idx);

  var subject = '[결재요청] ' + (data.docNo || '-') + ' — ' + (data.subject || '구매품의서');

  var htmlBody =
    '<div style="font-family:\'Malgun Gothic\',Arial,sans-serif;max-width:600px;margin:0 auto;background:#ffffff;">'
    + '<div style="background:#1a6fc4;padding:24px 32px;">'
    + '<div style="color:#ffffff;font-size:18px;font-weight:700;letter-spacing:2px;">구매품의 시스템</div>'
    + '</div>'
    + '<div style="padding:28px 32px;">'
    + '<p style="font-size:15px;color:#111111;margin:0 0 8px;"><b>' + approverName + ' ' + approverLabel + '</b> 님,</p>'
    + '<p style="font-size:14px;color:#444444;margin:0 0 24px;line-height:1.7;">'
    + '아래 구매품의서에 대한 결재를 요청합니다.<br>내용을 확인하신 후 승인 또는 반려해 주세요.</p>'
    + '<div style="background:#f8f9fa;border:1px solid #e0e0e0;border-radius:8px;padding:16px 20px;margin-bottom:24px;">'
    + '<table role="presentation" style="width:100%;border-collapse:collapse;font-size:13px;">'
    + '<tr><td style="color:#888;padding:4px 0;width:80px;">품의번호</td><td style="color:#111;font-weight:600;">' + docNo + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">품의제목</td><td style="color:#111;font-weight:600;">' + docTitle + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">기안자</td><td style="color:#111;">' + drafter + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">발행일자</td><td style="color:#111;">' + issueDate + '</td></tr>'
    + '</table></div>'
    + '<table role="presentation" align="center" style="margin:0 auto 20px auto;border-collapse:collapse;">'
    + '<tr><td align="center" bgcolor="#1a6fc4" style="border-radius:6px;">'
    + '<a href="' + viewUrl + '" style="display:inline-block;background:#1a6fc4;color:#fff;font-size:14px;font-weight:700;padding:12px 36px;border-radius:6px;text-decoration:none;letter-spacing:1px;">구매품의서 확인</a>'
    + '</td></tr></table>'
    + '<div style="border-top:1px solid #e0e0e0;margin:20px 0;"></div>'
    + '<p style="font-size:12px;color:#888;text-align:center;margin:0 0 14px;">품의서 확인 후 아래에서 결재해 주세요.</p>'
    + '<table role="presentation" align="center" style="margin:0 auto;border-collapse:collapse;"><tr>'
    + '<td align="center" style="padding-right:10px;"><a href="' + rejectUrl + '" style="display:inline-block;background:#d64541;color:#fff;font-size:14px;font-weight:700;padding:11px 32px;border-radius:6px;text-decoration:none;border:2px solid #d64541;">반 려</a></td>'
    + '<td align="center" style="padding-left:10px;"><a href="' + approveUrl + '" style="display:inline-block;background:#27ae60;color:#fff;font-size:14px;font-weight:700;padding:11px 32px;border-radius:6px;text-decoration:none;border:2px solid #27ae60;">승 인</a></td>'
    + '</tr></table>'
    + '</div>'
    + '<div style="background:#f0f0f0;padding:16px 32px;text-align:center;">'
    + '<p style="font-size:11px;color:#888;margin:0;">— ' + escapeHtml(CONFIG.FROM_NAME) + ' · 이 메일은 자동 발송되었습니다.</p>'
    + '</div></div>';

  var plainBody = [
    approver.name + ' ' + approver.label + ' 님,', '',
    '아래 구매품의서에 대한 결재를 요청합니다.', '',
    '■ 품의번호: ' + (data.docNo || '-'),
    '■ 품의제목: ' + (data.subject || '-'),
    '■ 기안자:   ' + (data.drafter || '-'),
    '■ 발행일자: ' + (toDateStr(data.issueDate) || '-'), '',
    '▶ 구매품의서 확인: ' + viewUrl,
    '▶ 승인하기: ' + approveUrl,
    '▶ 반려하기: ' + rejectUrl, '',
    '— ' + CONFIG.FROM_NAME,
  ].join('\n');

  return sendEmailWithRetry(approverEmail, subject, plainBody, htmlBody);
}

function sendRejectionNoticeWithRetry(toEmail, drafter, docNo, subject, approverName, comment, token) {
  var drafterName = escapeHtml(drafter || '-');
  var docNoEsc    = escapeHtml(docNo || '-');
  var docTitle    = escapeHtml(subject || '구매품의서');
  var approverEsc = escapeHtml(approverName || '-');
  var commentEsc  = escapeHtml(comment || '-').replace(/\n/g, '<br>');

  var resubUrl = CONFIG.WEBAPP_URL + '?action=reject_view&token=' + token +
                 '&docNo=' + encodeURIComponent(docNo);

  var emailSubject = '[반려] ' + (docNo || '-') + ' — ' + (subject || '구매품의서') + ' 가 반려되었습니다.';

  var htmlBody =
    '<div style="font-family:\'Malgun Gothic\',Arial,sans-serif;max-width:600px;margin:0 auto;background:#ffffff;">'
    + '<div style="background:#1a6fc4;padding:24px 32px;"><div style="color:#fff;font-size:18px;font-weight:700;letter-spacing:2px;">구매품의 시스템</div></div>'
    + '<div style="padding:28px 32px;">'
    + '<p style="font-size:15px;color:#111;margin:0 0 8px;"><b>' + drafterName + '</b> 님,</p>'
    + '<p style="font-size:14px;color:#444;margin:0 0 24px;line-height:1.7;">'
    + '제출하신 구매품의서가 <b style="color:#d64541;">반려</b>되었습니다.<br>'
    + '아래 내용을 확인하신 후 수정하여 재상신하거나 기안을 폐기해 주세요.</p>'
    + '<div style="background:#f8f9fa;border:1px solid #e0e0e0;border-radius:8px;padding:16px 20px;margin-bottom:20px;">'
    + '<table role="presentation" style="width:100%;border-collapse:collapse;font-size:13px;">'
    + '<tr><td style="color:#888;padding:4px 0;width:80px;">품의번호</td><td style="color:#111;font-weight:600;">' + docNoEsc + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">품의제목</td><td style="color:#111;font-weight:600;">' + docTitle + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">반려자</td><td style="color:#111;">' + approverEsc + '</td></tr></table></div>'
    + '<div style="background:#fdecea;border:1px solid #f5c2c2;border-radius:8px;padding:14px 20px;margin-bottom:24px;">'
    + '<div style="font-size:12px;color:#c62828;font-weight:700;margin-bottom:6px;letter-spacing:0.5px;">반려 사유</div>'
    + '<div style="font-size:13px;color:#333;line-height:1.7;">' + commentEsc + '</div></div>'
    + '<table role="presentation" align="center" style="margin:0 auto 20px auto;border-collapse:collapse;">'
    + '<tr><td align="center" bgcolor="#1a6fc4" style="border-radius:6px;">'
    + '<a href="' + resubUrl + '" style="display:inline-block;background:#1a6fc4;color:#fff;font-size:14px;font-weight:700;padding:12px 28px;border-radius:6px;text-decoration:none;letter-spacing:1px;">수정 후 재상신 / 기안 폐기</a>'
    + '</td></tr></table>'
    + '<div style="border-top:1px solid #e0e0e0;margin:20px 0;"></div>'
    + '<p style="font-size:12px;color:#888;text-align:center;margin:0;line-height:1.7;">'
    + '※ 첨부파일은 반려와 함께 삭제되었으니, 재상신 시 다시 업로드해 주세요.</p>'
    + '</div>'
    + '<div style="background:#f0f0f0;padding:16px 32px;text-align:center;">'
    + '<p style="font-size:11px;color:#888;margin:0;">— ' + escapeHtml(CONFIG.FROM_NAME) + ' · 이 메일은 자동 발송되었습니다.</p>'
    + '</div></div>';

  var plainBody = [
    drafter + ' 님,', '',
    '제출하신 품의서가 반려되었습니다.', '',
    '■ 품의번호: ' + docNo,
    '■ 품의제목: ' + (subject || '구매품의서'),
    '■ 반려자:   ' + (approverName || '-'),
    '■ 반려사유: ' + (comment || '-'), '',
    '아래 링크에서 내용을 수정하거나 기안을 폐기할 수 있습니다.',
    '※ 첨부파일은 반려와 함께 삭제되었으니, 재상신 시 다시 업로드해 주세요.', '',
    '▶ 수정 후 재상신 / 기안 폐기: ' + resubUrl, '',
    '— ' + CONFIG.FROM_NAME,
  ].join('\n');

  return sendEmailWithRetry(toEmail, emailSubject, plainBody, htmlBody);
}

function sendCompletionNoticeWithRetry(toEmail, drafter, docNo, subject, issueDate, token) {
  var drafterName  = escapeHtml(drafter || '-');
  var docNoEsc     = escapeHtml(docNo || '-');
  var docTitle     = escapeHtml(subject || '구매품의서');
  var issueDateEsc = escapeHtml(issueDate || '-');

  var viewUrl = token ?
    (CONFIG.WEBAPP_URL +
     '?token=' + encodeURIComponent(token) +
     '&docNo=' + encodeURIComponent(docNo || '') +
     '&action=view') : '';

  var emailSubject = '[승인완료] ' + (docNo || '-') + ' 구매품의서가 최종 승인되었습니다.';

  var htmlBody =
    '<div style="font-family:\'Malgun Gothic\',Arial,sans-serif;max-width:600px;margin:0 auto;background:#ffffff;">'
    + '<div style="background:#1a6fc4;padding:24px 32px;"><div style="color:#fff;font-size:18px;font-weight:700;letter-spacing:2px;">구매품의 시스템</div></div>'
    + '<div style="padding:28px 32px;">'
    + '<p style="font-size:15px;color:#111;margin:0 0 8px;"><b>' + drafterName + '</b> 님,</p>'
    + '<p style="font-size:14px;color:#444;margin:0 0 24px;line-height:1.7;">'
    + '제출하신 구매품의서가 <b style="color:#27ae60;">최종 승인</b>되었습니다.<br>'
    + '아래 버튼을 통해 승인 완료된 구매품의서를 확인하실 수 있습니다.</p>'
    + '<div style="background:#f8f9fa;border:1px solid #e0e0e0;border-radius:8px;padding:16px 20px;margin-bottom:24px;">'
    + '<table role="presentation" style="width:100%;border-collapse:collapse;font-size:13px;">'
    + '<tr><td style="color:#888;padding:4px 0;width:80px;">품의번호</td><td style="color:#111;font-weight:600;">' + docNoEsc + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">품의제목</td><td style="color:#111;font-weight:600;">' + docTitle + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">기안자</td><td style="color:#111;">' + drafterName + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">발행일자</td><td style="color:#111;">' + issueDateEsc + '</td></tr>'
    + '</table></div>';

  if (viewUrl) {
    htmlBody += '<table role="presentation" align="center" style="margin:0 auto 20px auto;border-collapse:collapse;">'
      + '<tr><td align="center" bgcolor="#1a6fc4" style="border-radius:6px;">'
      + '<a href="' + viewUrl + '" style="display:inline-block;background:#1a6fc4;color:#fff;font-size:14px;font-weight:700;padding:12px 36px;border-radius:6px;text-decoration:none;letter-spacing:1px;">구매품의서 확인</a>'
      + '</td></tr></table>';
  }

  htmlBody += '<div style="border-top:1px solid #e0e0e0;margin:20px 0;"></div>'
    + '<p style="font-size:12px;color:#888;text-align:center;margin:0;line-height:1.7;">'
    + '※ 본 품의서는 결재가 완료되어 별도의 추가 조치가 필요하지 않습니다.</p>'
    + '</div>'
    + '<div style="background:#f0f0f0;padding:16px 32px;text-align:center;">'
    + '<p style="font-size:11px;color:#888;margin:0;">— ' + escapeHtml(CONFIG.FROM_NAME) + ' · 이 메일은 자동 발송되었습니다.</p>'
    + '</div></div>';

  var plainBodyArr = [
    drafter + ' 님,', '',
    '제출하신 품의서가 최종 승인되었습니다.', '',
    '■ 품의번호: ' + (docNo || '-'),
    '■ 품의제목: ' + (subject || '구매품의서'),
    '■ 기안자:   ' + (drafter || '-'),
    '■ 발행일자: ' + (issueDate || '-'),
  ];
  if (viewUrl) plainBodyArr.push('', '▶ 구매품의서 확인: ' + viewUrl);
  plainBodyArr.push('', '— ' + CONFIG.FROM_NAME);

  return sendEmailWithRetry(toEmail, emailSubject, plainBodyArr.join('\n'), htmlBody);
}

/**
 * PRC(구매팀 구매품의서) 최종 승인 완료 → 원 REQ 기안자에게
 * '구매팀 결재 완료 & 검수보고서 작성 가능' 안내 메일.
 * 버튼('내 기안 확인')은 홈의 '결재 완료 + 검수보고서 제출' 메뉴가 기본 선택되도록 연결.
 */
function sendInspReadyNoticeWithRetry(toEmail, reqDrafter, reqDocNo, subject, poNo) {
  var drafterName = escapeHtml(reqDrafter || '-');
  var docNoEsc    = escapeHtml(reqDocNo || '-');
  var docTitle    = escapeHtml(subject || '구매품의서');
  var poNoEsc     = escapeHtml(poNo || '-');

  // 홈 화면 진입 시 '결재 완료 + 검수보고서 제출'(menu=completed) 메뉴 기본 선택
  var homeUrl = CONFIG.WEBAPP_URL + '?action=home&menu=completed';

  var emailSubject = '[구매팀 결재 완료] ' + (reqDocNo || '-') + ' — 검수보고서를 작성할 수 있습니다.';

  var htmlBody =
    '<div style="font-family:\'Malgun Gothic\',Arial,sans-serif;max-width:600px;margin:0 auto;background:#ffffff;">'
    + '<div style="background:#1a6fc4;padding:24px 32px;"><div style="color:#fff;font-size:18px;font-weight:700;letter-spacing:2px;">구매품의 시스템</div></div>'
    + '<div style="padding:28px 32px;">'
    + '<p style="font-size:15px;color:#111;margin:0 0 8px;"><b>' + drafterName + '</b> 님,</p>'
    + '<p style="font-size:14px;color:#444;margin:0 0 24px;line-height:1.7;">'
    + '기안하신 구매품의서에 대한 <b style="color:#27ae60;">구매팀 결재가 모두 완료</b>되었습니다.<br>'
    + '이제 아래 버튼을 통해 <b>검수보고서를 작성·제출</b>하실 수 있습니다.</p>'
    + '<div style="background:#f8f9fa;border:1px solid #e0e0e0;border-radius:8px;padding:16px 20px;margin-bottom:24px;">'
    + '<table role="presentation" style="width:100%;border-collapse:collapse;font-size:13px;">'
    + '<tr><td style="color:#888;padding:4px 0;width:80px;">품의번호</td><td style="color:#111;font-weight:600;">' + docNoEsc + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">품의제목</td><td style="color:#111;font-weight:600;">' + docTitle + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">구매문서번호</td><td style="color:#111;">' + poNoEsc + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">기안자</td><td style="color:#111;">' + drafterName + '</td></tr>'
    + '</table></div>'
    + '<table role="presentation" align="center" style="margin:0 auto 20px auto;border-collapse:collapse;">'
    + '<tr><td align="center" bgcolor="#1a6fc4" style="border-radius:6px;">'
    + '<a href="' + homeUrl + '" style="display:inline-block;background:#1a6fc4;color:#fff;font-size:14px;font-weight:700;padding:12px 36px;border-radius:6px;text-decoration:none;letter-spacing:1px;">내 기안 확인</a>'
    + '</td></tr></table>'
    + '<div style="border-top:1px solid #e0e0e0;margin:20px 0;"></div>'
    + '<p style="font-size:12px;color:#888;text-align:center;margin:0;line-height:1.7;">'
    + '※ \'내 기안 확인\'을 누르면 홈의 [결재 완료 + 검수보고서 제출] 메뉴로 이동합니다.</p>'
    + '</div>'
    + '<div style="background:#f0f0f0;padding:16px 32px;text-align:center;">'
    + '<p style="font-size:11px;color:#888;margin:0;">— ' + escapeHtml(CONFIG.FROM_NAME) + ' · 이 메일은 자동 발송되었습니다.</p>'
    + '</div></div>';

  var plainBody = [
    (reqDrafter || '-') + ' 님,', '',
    '기안하신 구매품의서에 대한 구매팀 결재가 모두 완료되었습니다.',
    '이제 검수보고서를 작성·제출하실 수 있습니다.', '',
    '■ 품의번호:     ' + (reqDocNo || '-'),
    '■ 품의제목:     ' + (subject || '구매품의서'),
    '■ 구매문서번호: ' + (poNo || '-'),
    '■ 기안자:       ' + (reqDrafter || '-'), '',
    '▶ 내 기안 확인 (결재 완료 + 검수보고서 제출): ' + homeUrl, '',
    '— ' + CONFIG.FROM_NAME,
  ].join('\n');

  return sendEmailWithRetry(toEmail, emailSubject, plainBody, htmlBody);
}

/**
 * REQ 1차 승인 완료 → 구매팀 전체 알림 (FR-50, REQ-07)
 */
function notifyProcurementTeamOfApproved1st(docNo, subject, drafter, parentToken) {
  var emails = CONFIG.PROCUREMENT_NOTIFY_EMAILS || [];
  if (emails.length === 0) {
    Logger.log('[INFO] 구매팀 알림 대상이 설정되지 않음. CONFIG.PROCUREMENT_NOTIFY_EMAILS를 채워주세요.');
    return;
  }

  var docNoEsc   = escapeHtml(docNo);
  var subjectEsc = escapeHtml(subject);
  var drafterEsc = escapeHtml(drafter);

  var homeUrl = CONFIG.WEBAPP_URL + '?action=home';
  var prcUrl  = CONFIG.WEBAPP_URL + '?action=prc&parentId=' + encodeURIComponent(parentToken);

  var emailSubject = '[구매팀 인박스] ' + docNo + ' — 1차 결재 완료 (' + subject + ')';

  var htmlBody =
    '<div style="font-family:\'Malgun Gothic\',Arial,sans-serif;max-width:600px;margin:0 auto;background:#fff;">'
    + '<div style="background:#6a3eb5;padding:24px 32px;"><div style="color:#fff;font-size:18px;font-weight:700;letter-spacing:2px;">구매팀 작업 대기열</div></div>'
    + '<div style="padding:28px 32px;">'
    + '<p style="font-size:14px;color:#444;margin:0 0 20px;line-height:1.7;">'
    + '아래 구매품의서의 1차 결재가 완료되어 구매팀 작업 대기열(구매팀 기안 문서)에 등록되었습니다.<br>'
    + '구매팀 기안자 누구나 선착순으로 픽업하여 구매 문서를 생성할 수 있습니다.</p>'
    + '<div style="background:#f8f9fa;border:1px solid #e0e0e0;border-radius:8px;padding:16px 20px;margin-bottom:24px;">'
    + '<table role="presentation" style="width:100%;border-collapse:collapse;font-size:13px;">'
    + '<tr><td style="color:#888;padding:4px 0;width:80px;">품의번호</td><td style="color:#111;font-weight:600;">' + docNoEsc + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">품의제목</td><td style="color:#111;font-weight:600;">' + subjectEsc + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">기안자</td><td style="color:#111;">' + drafterEsc + '</td></tr></table></div>'
    + '<table role="presentation" align="center" style="margin:0 auto;border-collapse:collapse;"><tr>'
    + '<td style="padding-right:10px;"><a href="' + homeUrl + '" style="display:inline-block;background:#fff;color:#6a3eb5;font-size:14px;font-weight:700;padding:11px 22px;border-radius:6px;text-decoration:none;border:2px solid #6a3eb5;">대시보드에서 보기</a></td>'
    + '<td style="padding-left:10px;"><a href="' + prcUrl + '" style="display:inline-block;background:#6a3eb5;color:#fff;font-size:14px;font-weight:700;padding:11px 22px;border-radius:6px;text-decoration:none;">구매 문서 생성하기</a></td>'
    + '</tr></table>'
    + '<div style="border-top:1px solid #e0e0e0;margin:24px 0 12px;"></div>'
    + '<p style="font-size:11px;color:#888;text-align:center;margin:0;">동일 품의에 대해 한 명만 구매 문서 생성이 가능합니다 (선착순).</p>'
    + '</div>'
    + '<div style="background:#f0f0f0;padding:16px 32px;text-align:center;">'
    + '<p style="font-size:11px;color:#888;margin:0;">— ' + escapeHtml(CONFIG.FROM_NAME) + '</p>'
    + '</div></div>';

  var plainBody = [
    '[구매팀 작업 대기열]', '',
    '아래 REQ가 1차 결재 완료되어 인박스에 등록되었습니다.', '',
    '■ 품의번호: ' + docNo,
    '■ 품의제목: ' + subject,
    '■ 기안자:   ' + drafter, '',
    '▶ 대시보드: ' + homeUrl,
    '▶ PRC 생성: ' + prcUrl, '',
    '동일 품의에 대해 한 명만 구매 문서 생성이 가능합니다 (선착순).',
    '— ' + CONFIG.FROM_NAME,
  ].join('\n');

  emails.forEach(function(email) {
    try {
      sendEmailWithRetry(email, emailSubject, plainBody, htmlBody);
    } catch(e) {
      Logger.log('[notifyProcurementTeam] ' + email + ' 발송 실패: ' + e);
    }
  });
}


// ================================================================
// 15. Drive 폴더 / 파일 관리
// ================================================================

function getOrCreateFolder(docNo, issueDate, type) {
  var rootId = (type === 'final') ? CONFIG.FINAL_ROOT_ID : CONFIG.STAGING_ROOT_ID;
  var root   = DriveApp.getFolderById(rootId);
  var date   = issueDate ? new Date(issueDate) : new Date();
  if (isNaN(date.getTime())) date = new Date();
  var year   = date.getFullYear().toString();
  var month  = ('0' + (date.getMonth() + 1)).slice(-2) + '월';

  var yearFolder  = getOrCreateSubFolder(root, year);
  var monthFolder = getOrCreateSubFolder(yearFolder, month);
  var docFolder   = getOrCreateSubFolder(monthFolder, docNo);
  return docFolder;
}

function getOrCreateSubFolder(parent, name) {
  var iter = parent.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return parent.createFolder(name);
}

function createFileInFolder(folder, fileName, content, mimeType) {
  return folder.createFile(fileName, content, mimeType);
}

function uniqueFileName(folder, name) {
  if (!folder.getFilesByName(name).hasNext()) return name;
  var dot = name.lastIndexOf('.');
  var base = dot > 0 ? name.substring(0, dot) : name;
  var ext  = dot > 0 ? name.substring(dot)    : '';
  for (var i = 1; i < 100; i++) {
    var candidate = base + '(' + i + ')' + ext;
    if (!folder.getFilesByName(candidate).hasNext()) return candidate;
  }
  return base + '_' + Date.now() + ext;
}

function saveAttachmentsToFolder(folder, attachments) {
  var saved = [];
  (attachments || []).forEach(function(att) {
    if (!att || !att.data) return;

    if (att.size && att.size > CONFIG.MAX_FILE_SIZE) {
      throw new Error('파일 크기 초과 (' + att.name + '): 최대 ' + (CONFIG.MAX_FILE_SIZE / 1024 / 1024) + 'MB');
    }
    var ext = (att.name || '').split('.').pop().toLowerCase();
    if (CONFIG.ALLOWED_EXTS.indexOf(ext) < 0) {
      throw new Error('허용되지 않는 파일 형식: ' + att.name);
    }

    var bytes    = Utilities.base64Decode(att.data);
    var finalName = uniqueFileName(folder, att.name);
    var blob     = Utilities.newBlob(bytes, att.type || 'application/octet-stream', finalName);
    var file     = folder.createFile(blob);

    // 폴더(구매팀 전용) 권한과 무관하게, 결재 라인의 타 부서 기안자/결재자가
    // 링크로 열람할 수 있도록 파일 단위로 도메인 링크 공유(읽기)를 부여.
    // 도메인 정책으로 막히면 조용히 무시(업로드 자체는 계속 진행).
    try {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(e) {
      Logger.log('[첨부 공유 실패] ' + finalName + ' (id: ' + file.getId() + '): ' + e);
    }

    saved.push({
      name: finalName,
      id:   file.getId(),
      size: att.size || bytes.length,
      type: att.type || '',
    });
  });
  return saved;
}

function rollbackFiles(savedFiles) {
  (savedFiles || []).forEach(function(f) {
    try { DriveApp.getFileById(f.id).setTrashed(true); } catch(_) {}
  });
}

function deleteUserAttachments(attachJson) {
  var list = parseAttachments(attachJson);
  list.forEach(function(f) {
    try { DriveApp.getFileById(f.id).setTrashed(true); } catch(_) {}
  });
}

function parseAttachments(json) {
  try {
    var parsed = JSON.parse(json || '[]');
    return Array.isArray(parsed) ? parsed : [];
  } catch(_) { return []; }
}

// ================================================================
// 16. Drive 요약 파일
// ================================================================

function saveSummaryToDrive(folder, data, itemsSummary, totalAmt, totalVat, approvers) {
  createFileInFolder(folder, data.docNo + '_품의서요약.txt',
    buildSummaryText(data, itemsSummary, totalAmt, totalVat, approvers), MimeType.PLAIN_TEXT);
}

function buildSummaryText(data, itemsSummary, totalAmt, totalVat, approvers) {
  var _cur = getItemsCurrency((data && data.items) ? data.items : parseItemsSummary(itemsSummary));
  var lines = [
    '========================================',
    '구 매 품 의 서',
    '========================================',
    '품의번호: ' + (data.docNo || ''),     '발행일자: ' + (toDateStr(data.issueDate) || ''),
    '기안자:   ' + (data.drafter || ''),   '소속부서: ' + (data.dept || ''),
    '품의제목: ' + (data.subject || ''),   '용도:     ' + (data.purpose || ''),
    '', '[공급업체]',
    '업체명:       ' + (data.vendorName || ''),
    '담당자:       ' + (data.vendorContact || ''),
    '담당자이메일: ' + (data.vendorEmail || ''),
    '연락처:       ' + (data.vendorPhone || ''),
    '납기일:       ' + (toDateStr(data.deliveryDate) || ''),
    '납품장소:     ' + (data.deliveryAddr || ''),
    '', '[품목내역]', itemsSummary, '',
    '공급가액:     ' + Number(totalAmt).toLocaleString() + ' ' + _cur,
    '부가세:       ' + Number(totalVat - totalAmt).toLocaleString() + ' ' + _cur,
    '합계:         ' + Number(totalVat).toLocaleString() + ' ' + _cur,
    '', '[결재자]',
  ];
  approvers.forEach(function(ap, i) {
    lines.push((i + 1) + '. ' + ap.label + ' / ' + ap.name + ' <' + ap.email + '>');
  });
  lines.push('', '비고: ' + (data.remarks || '-'), '========================================');
  return lines.join('\n');
}

// ================================================================
// 17. 결재자 목록
// ================================================================

function getApproverList() {
  return jsonResponse(getApproverListForClient());
}

// 현재 로그인 사용자 정보 조회 (기안자 본인 자동 고정용)
// 결재자목록에서 세션 이메일로 매칭 → 이름/부서/직급 반환. 미등록 시 이메일 앞부분을 이름으로.
function getCurrentUserForClient() {
  try {
    var email = getActiveUserEmail() || '';
    var name = '', dept = '', rank = '', found = false;
    if (email) {
      var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet = ss.getSheetByName('결재자목록');
      if (sheet) {
        var rows = sheet.getDataRange().getValues();
        for (var i = 1; i < rows.length; i++) {
          if (String(rows[i][3] || '').toLowerCase() === email.toLowerCase()) {
            name = String(rows[i][0] || '');
            dept = String(rows[i][1] || '');
            rank = String(rows[i][2] || '');
            found = true;
            break;
          }
        }
      }
    }
    if (!name) name = email ? email.split('@')[0] : '';
    return { ok: true, user: { name: name, dept: dept, rank: rank, email: email, found: found } };
  } catch(err) {
    return { ok: false, message: err.toString(), user: { name: '', dept: '', rank: '', email: '', found: false } };
  }
}

function getApproverListForClient() {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet = ss.getSheetByName('결재자목록');
    if (!sheet) return { ok: true, approvers: [] };
    var rows = sheet.getDataRange().getValues();
    var approvers = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      approvers.push({
        name:   String(r[0] || ''),
        dept:   String(r[1] || ''),
        rank:   String(r[2] || ''),
        email:  String(r[3] || ''),
        active: r[4] !== '비활성',
      });
    }
    return { ok: true, approvers: approvers };
  } catch(err) {
    return { ok: false, message: err.toString(), approvers: [] };
  }
}

/**
 * 공급업체 마스터 조회 (REQ 폼 '조회' 모달용)
 * - 폼 로드 시 1회만 호출해 전량 내려주고, 검색/필터는 클라이언트 메모리에서 처리
 * - 업체목록 열: A:업체명 B:담당자 C:담당자이메일 D:연락처 E:사업자번호 F:활성여부
 * - 시트가 없거나 읽기에 실패해도 throw 하지 않는다 (조회 실패가 폼 자체를 막으면 안 됨)
 */
function getVendorListForClient() {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.VENDOR_SHEET_NAME);
    if (!sheet) return { ok: true, vendors: [] };
    var rows = sheet.getDataRange().getValues();
    var vendors = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      vendors.push({
        name:    String(r[0] || '').trim(),
        contact: String(r[1] || '').trim(),
        email:   String(r[2] || '').trim(),
        phone:   String(r[3] || '').trim(),
        bizNo:   String(r[4] || '').trim(),
        active:  r[5] !== '비활성',
      });
    }
    return { ok: true, vendors: vendors };
  } catch(err) {
    return { ok: false, message: err.toString(), vendors: [] };
  }
}

/**
 * 업체목록 탭이 없으면 헤더만 만들어 둔다 (데이터는 사용자가 직접 입력)
 * - GAS 편집기에서 1회 수동 실행용
 */
function ensureVendorSheet() {
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var exists = !!ss.getSheetByName(CONFIG.VENDOR_SHEET_NAME);
  var sheet = getOrCreateSheet(ss, CONFIG.VENDOR_SHEET_NAME);
  if (exists) return '이미 존재: ' + CONFIG.VENDOR_SHEET_NAME;

  var headers = ['업체명', '담당자', '담당자이메일', '연락처', '사업자번호', '활성여부'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
       .setFontWeight('bold')
       .setBackground('#f0f0f0');
  sheet.setFrozenRows(1);
  return '생성 완료: ' + CONFIG.VENDOR_SHEET_NAME;
}

/**
 * 결재자목록 탭 기반 이메일→부서 맵 (부서 단위 문서 가시성 판정용)
 * - key: 이메일(소문자), value: 부서 문자열(trim)
 * - 부서/이메일 미기재 행은 스킵
 * - ss 인자를 재사용해 추가 시트 오픈 없이 1회 read
 */
function getDeptByEmailMap(ss) {
  var map = {};
  try {
    ss = ss || SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet = ss.getSheetByName('결재자목록');
    if (!sheet) return map;
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var email = String(rows[i][3] || '').trim().toLowerCase();
      var dept  = String(rows[i][1] || '').trim();
      if (!email || !dept) continue;
      map[email] = dept;
    }
  } catch(_) { /* 부서 맵 실패는 가시성 축소로만 이어짐(본인 문서는 항상 표시) */ }
  return map;
}

// ================================================================
// 18. 유틸 (계산)
// ================================================================

// 폼의 단가 입력란은 콤마 표시(1,430,000)를 쓴다. 클라이언트가 콤마를 떼고 보내지만,
// Number('1,430,000')은 NaN → 0이 되어 금액을 조용히 0으로 만들므로 서버에서도 한 번 더 막는다.
function toNum(v) {
  return Number(String(v == null ? '' : v).replace(/,/g, '')) || 0;
}

function buildItemsSummary(items) {
  return (items || []).map(function(it, i) {
    var cur = String(it.currency || 'KRW');
    // 품목명·규격을 ' / '로 구분 — 품목명에 공백이 있어도 규격 셀로 넘어가지 않도록 명시적 구분자 사용
    return (i + 1) + '. ' + (it.name || '') + ' / ' + (it.spec || '') +
           ' / ' + (it.qty || 0) + '개 / ' + toNum(it.price).toLocaleString() + ' ' + cur;
  }).join('\n');
}

// 통화 판정: 첫 품목의 통화 (단일통화 폼 전제, 미지정 시 KRW)
function getItemsCurrency(items) {
  if (items && items.length && items[0] && items[0].currency) return String(items[0].currency);
  return 'KRW';
}

// 부가세 계산: 원화(KRW)만 10%, 외화는 0 (B안)
function computeVat(totalAmt, items) {
  return getItemsCurrency(items) === 'KRW' ? Math.round(totalAmt * 0.1) : 0;
}

function calcTotal(items) {
  return (items || []).reduce(function(s, it) {
    return s + toNum(it.qty) * toNum(it.price);
  }, 0);
}

// ================================================================
// 19. Staging → Final 이동 (최종 트리거, FR-32~36 토대)
// ================================================================

function moveAttachmentsToFinal(token) {
  return withLock(function() {
    try {
      var ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet  = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
      var rowNum = findRowNumByToken(sheet, token);
      if (rowNum < 0) return { ok: false, message: '유효하지 않은 토큰입니다.' };

      var r = readRow(sheet, rowNum);
      var moveStatus = String(r[COL.MOVE_STATUS] || '');
      if (moveStatus === 'FINAL') {
        return { ok: true, message: '이미 FINAL로 이동된 품의서입니다.', skipped: true };
      }
      var status = String(r[COL.STATUS] || '');
      if (status.indexOf('최종승인') < 0) {
        return { ok: false, message: '최종승인 상태의 품의서만 이동할 수 있습니다.' };
      }

      var docNo     = r[COL.DOC_NO];
      var issueDate = r[COL.ISSUE_DATE];
      var stagingFolderId = String(r[COL.DRIVE_ID] || '');

      if (!stagingFolderId) {
        return { ok: false, message: 'Drive 폴더 ID가 없습니다.' };
      }

      var finalFolder   = getOrCreateFolder(docNo, issueDate, 'final');
      var finalFolderId = finalFolder.getId();

      // 폴더 안의 모든 파일을 FINAL로 이동 (첨부 + 요약 .txt + 생성된 PDF 모두 포함)
      var stagingFolder = DriveApp.getFolderById(stagingFolderId);
      var fileIter = stagingFolder.getFiles();
      var movedCount = 0;
      var failedFiles = [];

      while (fileIter.hasNext()) {
        var file = fileIter.next();
        var fileId = file.getId();
        try {
          var meta = Drive.Files.get(fileId, { supportsAllDrives: true, fields: 'id,name,parents' });
          var currentParents = (meta.parents || []).join(',');
          Drive.Files.update({}, fileId, null, {
            supportsAllDrives: true,
            addParents:        finalFolderId,
            removeParents:     currentParents,
            fields:            'id,parents',
          });
          movedCount++;
        } catch(e) {
          failedFiles.push(file.getName() + ': ' + e.toString());
        }
      }

      // 빈 STAGING 폴더는 휴지통으로 (선택적 — 폴더 자체는 비워둠)
      // 실수 복구를 위해 즉시 삭제하지 않고 유지

      sheet.getRange(rowNum, COL.DRIVE_ID + 1, 1, 1).setValue(finalFolderId);
      sheet.getRange(rowNum, COL.MOVE_STATUS + 1, 1, 1).setValue('FINAL');

      if (failedFiles.length > 0) {
        return {
          ok: true,
          warning: true,
          message: 'Final 폴더 이동 완료 (일부 실패: ' + failedFiles.length + '건)',
          moved: movedCount,
          failed: failedFiles,
        };
      }
      return { ok: true, message: 'Final 폴더로 이동 완료.', moved: movedCount };
    } catch(err) {
      return { ok: false, message: err.toString() };
    }
  });
}

/**
 * PRC 최종승인 시점에 REQ + PRC를 한 FINAL 폴더(PO 번호 기준)로 통합 이동
 *
 * 시점: PRC가 최종승인된 직후 호출
 * 결과: FINAL/{YYYY}/{MM월}/{PO번호}/ 폴더 하나에 REQ와 PRC의 모든 파일이 모임
 *
 * @param prcToken — PRC 행의 token
 * @returns {Object} ok, message, movedFromReq, movedFromPrc, folderId
 */
function consolidateToFinalByPrc(prcToken) {
  return withLock(function() {
    try {
      var ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet  = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
      var prcRowNum = findRowNumByToken(sheet, prcToken);
      if (prcRowNum < 0) return { ok: false, message: 'PRC 행을 찾을 수 없습니다.' };

      var prcRow = readRow(sheet, prcRowNum);
      var docType = String(prcRow[COL.DOC_TYPE] || '');
      if (docType !== 'PRC') {
        return { ok: false, message: 'PRC 문서만 통합 이동 가능. 현재 docType: ' + docType };
      }

      var prcStatus = String(prcRow[COL.STATUS] || '');
      if (prcStatus.indexOf('최종승인') < 0) {
        return { ok: false, message: 'PRC가 최종승인 상태가 아닙니다. 현재: ' + prcStatus };
      }

      var prcMoveStatus = String(prcRow[COL.MOVE_STATUS] || '');
      if (prcMoveStatus === 'FINAL') {
        return { ok: true, message: '이미 FINAL로 통합 이동되었습니다.', skipped: true };
      }

      // 부모 REQ 행 찾기
      var parentToken = String(prcRow[COL.PARENT_DOC_ID] || '');
      if (!parentToken) {
        return { ok: false, message: 'PRC에 부모 REQ 토큰이 없습니다.' };
      }
      var reqRowNum = findRowNumByToken(sheet, parentToken);
      if (reqRowNum < 0) {
        return { ok: false, message: '부모 REQ 행을 찾을 수 없습니다.' };
      }
      var reqRow = readRow(sheet, reqRowNum);

      var prcDocNo     = String(prcRow[COL.DOC_NO] || '');         // PO 번호
      var prcIssueDate = prcRow[COL.ISSUE_DATE];
      var reqDocNo     = String(reqRow[COL.DOC_NO] || '');         // 품의번호
      var prcStagingId = String(prcRow[COL.DRIVE_ID] || '');
      var reqStagingId = String(reqRow[COL.DRIVE_ID] || '');

      if (!prcStagingId) return { ok: false, message: 'PRC 폴더 ID가 없습니다.' };
      if (!reqStagingId) return { ok: false, message: 'REQ 폴더 ID가 없습니다.' };

      // 1) FINAL 폴더 생성 — PO 번호 기준
      var finalFolder = getOrCreateFolder(prcDocNo, prcIssueDate, 'final');
      var finalFolderId = finalFolder.getId();

      // 2) PRC STAGING 폴더의 모든 파일 → FINAL 이동
      var prcResult = _moveAllFilesToFolder(prcStagingId, finalFolderId);

      // 3) REQ STAGING 폴더의 모든 파일 → 같은 FINAL 폴더로 이동
      var reqResult = _moveAllFilesToFolder(reqStagingId, finalFolderId);

      // 4) 두 행 모두 시트 업데이트
      sheet.getRange(prcRowNum, COL.DRIVE_ID + 1, 1, 1).setValue(finalFolderId);
      sheet.getRange(prcRowNum, COL.MOVE_STATUS + 1, 1, 1).setValue('FINAL');
      sheet.getRange(reqRowNum, COL.DRIVE_ID + 1, 1, 1).setValue(finalFolderId);
      sheet.getRange(reqRowNum, COL.MOVE_STATUS + 1, 1, 1).setValue('FINAL');

      var totalMoved = prcResult.moved + reqResult.moved;
      var totalFailed = (prcResult.failed || []).concat(reqResult.failed || []);

      if (totalFailed.length > 0) {
        return {
          ok: true,
          warning: true,
          message: '통합 이동 완료 (일부 실패: ' + totalFailed.length + '건)',
          movedFromReq: reqResult.moved,
          movedFromPrc: prcResult.moved,
          failed: totalFailed,
          folderId: finalFolderId,
        };
      }
      return {
        ok: true,
        message: 'PO 번호 ' + prcDocNo + ' 기준 FINAL 폴더로 통합 이동 완료.',
        movedFromReq: reqResult.moved,
        movedFromPrc: prcResult.moved,
        folderId: finalFolderId,
      };
    } catch(err) {
      return { ok: false, message: err.toString() };
    }
  });
}

/**
 * 한 폴더의 모든 파일을 다른 폴더로 이동 (부모 교체 방식)
 * 폴더 자체는 비워진 채 유지됨 (실수 복구 대비)
 */
function _moveAllFilesToFolder(srcFolderId, dstFolderId) {
  var moved = 0;
  var failed = [];
  try {
    var srcFolder = DriveApp.getFolderById(srcFolderId);
    var iter = srcFolder.getFiles();
    while (iter.hasNext()) {
      var file = iter.next();
      var fileId = file.getId();
      try {
        var meta = Drive.Files.get(fileId, { supportsAllDrives: true, fields: 'id,name,parents' });
        var currentParents = (meta.parents || []).join(',');
        Drive.Files.update({}, fileId, null, {
          supportsAllDrives: true,
          addParents:        dstFolderId,
          removeParents:     currentParents,
          fields:            'id,parents',
        });
        moved++;
      } catch(e) {
        failed.push(file.getName() + ': ' + e.toString());
      }
    }
  } catch(e) {
    failed.push('폴더 접근 실패 (' + srcFolderId + '): ' + e.toString());
  }
  return { moved: moved, failed: failed };
}

// ================================================================
// 20. PDF 합성 (Sprint 4) — 최종승인 시 자동 호출
// ================================================================

/**
 * PDF용 데이터 페이로드 (화이트리스트)
 * 결재 단계에서만 보이는 내부 메모(remarks, prcMemo, rejectLog 등)는 의도적으로 제외.
 *
 * @param row — 시트의 한 행 (readRow 결과)
 * @returns {Object} PDF 템플릿에 안전하게 넘길 수 있는 필드만 포함
 */
function buildPdfPayload(row) {
  var docType = String(row[COL.DOC_TYPE] || 'REQ');
  var apprCount = parseInt(row[COL.APPR_COUNT]) || 0;

  var approvers = [];
  for (var i = 0; i < apprCount; i++) {
    var b = COL.APPR_START + i * COL.APPR_COLS;
    approvers.push({
      label:       String(row[b]   || ''),
      name:        String(row[b+1] || ''),
      status:      String(row[b+3] || '대기'),
      processedAt: toDateTimeStr(row[b+4]),
    });
  }

  return {
    docType:       docType,
    isPrc:         docType === 'PRC',
    docNo:         String(row[COL.DOC_NO] || ''),
    docToken:      String(row[COL.TOKEN] || ''),
    parentDocNo:   '',
    parentSubmitAt: '',
    issueDate:     toDateStr(row[COL.ISSUE_DATE]),
    drafter:       String(row[COL.DRAFTER] || ''),
    dept:          String(row[COL.DEPT] || ''),
    subject:       String(row[COL.SUBJECT] || ''),
    purpose:       String(row[COL.PURPOSE] || ''),
    vendorName:    String(row[COL.VENDOR_NAME] || ''),
    vendorEmail:     String(row[COL.VENDOR_EMAIL] || ''),
    vendorContact: String(row[COL.VENDOR_CTC] || ''),
    vendorPhone:   String(row[COL.VENDOR_PHONE] || ''),
    deliveryDate:  toDateStr(row[COL.DELIVERY_DATE]),
    deliveryAddr:  String(row[COL.DELIVERY_ADDR] || ''),
    purchaseMethod: String(row[COL.PURCHASE_METHOD] || ''),
    paymentInfo:    String(row[COL.PAYMENT_INFO] || ''),
    items:         parseItemsSummary(String(row[COL.ITEMS] || '')),
    totalAmt:      Number(row[COL.TOTAL_AMT]) || 0,
    totalVat:      Number(row[COL.TOTAL_VAT]) || 0,
    approvers:     approvers,
    parentApprovers: [],
    hasParentApprovers: false,
    generatedAt:   toDateTimeStr(new Date()),
    title:         (docType === 'PRC' ? '구매품의서(PRC)_' : '구매품의서_') + String(row[COL.DOC_NO] || ''),

    // 의도적으로 제외:
    //   remarks, prcMemo, rejectLog, claimedBy, claimedAt, fieldChanges
  };
}

/**
 * 페이로드 화이트리스트 검증 — 금지 필드가 끼어 들어갔는지 자체 점검
 */
function _verifyPdfPayloadWhitelist(payload) {
  var FORBIDDEN = ['remarks', 'prcMemo', 'rejectLog', 'rejectHistory',
                   'fieldChanges', 'discardReason', 'claimedBy', 'claimedAt'];
  for (var i = 0; i < FORBIDDEN.length; i++) {
    if (Object.prototype.hasOwnProperty.call(payload, FORBIDDEN[i])) {
      throw new Error('PDF 페이로드 화이트리스트 위반: 금지 필드 ' + FORBIDDEN[i] + ' 발견');
    }
  }
}

/**
 * 부모 REQ의 결재자 정보를 PRC payload에 머지
 */
function _attachParentApprovers(payload, parentRow) {
  if (!parentRow) return payload;
  payload.parentDocNo    = String(parentRow[COL.DOC_NO] || '');
  payload.parentSubmitAt = toDateTimeStr(parentRow[COL.SUBMIT_AT]);

  var parentApprCount = parseInt(parentRow[COL.APPR_COUNT]) || 0;
  var parentApprovers = [];
  for (var i = 0; i < parentApprCount; i++) {
    var b = COL.APPR_START + i * COL.APPR_COLS;
    parentApprovers.push({
      label:       String(parentRow[b]   || ''),
      name:        String(parentRow[b+1] || ''),
      status:      String(parentRow[b+3] || '대기'),
      processedAt: toDateTimeStr(parentRow[b+4]),
    });
  }
  payload.parentApprovers = parentApprovers;
  payload.hasParentApprovers = parentApprovers.length > 0;
  return payload;
}

/**
 * PDF 파일 생성 — REQ든 PRC든 자기 token 기준으로 자기 PDF 1개를 만듦
 *
 * 흐름:
 *   1. 시트에서 행 조회
 *   2. 최종승인 상태 검증
 *   3. buildPdfPayload로 화이트리스트 페이로드 구성
 *   4. PRC인 경우 부모 REQ 결재자 머지
 *   5. PDF_Template 렌더 → Blob → setName → Drive 폴더에 저장
 *
 * @param token — 문서의 token (REQ 또는 PRC)
 * @returns {{ok: bool, message: string, fileId?: string, fileName?: string}}
 */
function generatePdfForDocument(token) {
  try {
    if (!token) return { ok: false, message: 'token이 필요합니다.' };

    var ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet  = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
    var rowNum = findRowNumByToken(sheet, token);
    if (rowNum < 0) return { ok: false, message: '문서를 찾을 수 없습니다.' };

    var row = readRow(sheet, rowNum);
    var status = String(row[COL.STATUS] || '');
    if (status.indexOf('최종승인') < 0) {
      return { ok: false, message: '최종승인 상태에서만 PDF 생성 가능. 현재: ' + status };
    }

    var payload = buildPdfPayload(row);

    // PRC인 경우 부모 REQ 결재자 머지
    if (payload.isPrc) {
      var parentToken = String(row[COL.PARENT_DOC_ID] || '');
      if (parentToken) {
        var parentRowNum = findRowNumByToken(sheet, parentToken);
        if (parentRowNum > 0) {
          var parentRow = readRow(sheet, parentRowNum);
          _attachParentApprovers(payload, parentRow);
        }
      }
    }

    // 화이트리스트 자체 검증
    _verifyPdfPayloadWhitelist(payload);

    // HTML 템플릿 렌더링
    var tmpl = HtmlService.createTemplateFromFile('PDF_Template');
    Object.keys(payload).forEach(function(k) { tmpl[k] = payload[k]; });
    var html = tmpl.evaluate().getContent();

    // Blob → PDF 변환
    var blob = Utilities.newBlob(html, 'text/html', payload.docNo + '.html')
                        .getAs(MimeType.PDF);

    // 파일명 규칙: {docType}_{docNo}_v{생성일}.pdf
    var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    var fileName = (payload.isPrc ? 'PRC_' : 'REQ_') + payload.docNo + '_' + ts + '.pdf';
    blob.setName(fileName);

    // 현재 Drive 폴더에 저장 (FINAL 이동 전에 호출되면 STAGING 폴더에, 이후라면 FINAL 폴더에 저장됨)
    var folderId = String(row[COL.DRIVE_ID] || '');
    if (!folderId) {
      return { ok: false, message: '폴더 ID가 없습니다.' };
    }
    var folder = DriveApp.getFolderById(folderId);

    // 같은 이름의 이전 PDF가 있으면 휴지통으로 (중복 방지)
    var existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) {
      try { existing.next().setTrashed(true); } catch(_) {}
    }

    // 동일 docType+docNo의 이전 버전들도 정리 (타임스탬프만 다른 것들)
    var pattern = (payload.isPrc ? 'PRC_' : 'REQ_') + payload.docNo + '_';
    var allFiles = folder.getFiles();
    while (allFiles.hasNext()) {
      var f = allFiles.next();
      var fname = f.getName();
      if (fname.indexOf(pattern) === 0 && fname.indexOf('.pdf') > 0 && fname !== fileName) {
        try { f.setTrashed(true); } catch(_) {}
      }
    }

    var file = folder.createFile(blob);

    return {
      ok: true,
      message: 'PDF 생성 완료',
      fileId: file.getId(),
      fileName: fileName,
      docType: payload.docType,
      folderId: folderId,
    };
  } catch(err) {
    return { ok: false, message: 'PDF 생성 오류: ' + err.toString() };
  }
}

/**
 * 수동 재생성용 진입점 (대시보드나 관리자 콘솔에서 호출 가능)
 * @param token — 문서 token
 */
function regeneratePdfForClient(payload) {
  if (!payload || !payload.token) return { ok: false, message: 'token이 필요합니다.' };
  var actor = getActiveUserEmail();
  if (!isAdminUser(actor)) {
    return { ok: false, message: '관리자만 PDF 재생성이 가능합니다.' };
  }
  return generatePdfForDocument(payload.token);
}

// ================================================================
// 21. 비동기 작업 큐 (NFR-07 응답성)
// ================================================================
// 큐 키:        Script Properties의 'PENDING_JOB_QUEUE' (JSON 배열)
// Job 구조:     { id, type, token, docNo, docType, attempts, status, enqueuedAt, lastError? }
// Job 타입:     'pdf_only'              REQ 최종승인 → PDF 생성만
//               'pdf_and_consolidate'   PRC 최종승인 → PDF + FINAL 통합 이동
// 처리 주기:    Time-based Trigger, 1분 간격 — installQueueTrigger() 1회 호출하여 설치
// 한 트리거당:  while 루프로 시간 한도 안에 최대한 많이 처리
//               (예: 4분 한도 - 30초 안전 마진 → 12초/건 시 약 17건 가능)
// 재시도:       3회까지, 그 후엔 큐에서 제거하고 관리자 알림
// ================================================================

var QUEUE_KEY = 'PENDING_JOB_QUEUE';
var QUEUE_MAX_ATTEMPTS = 3;
var QUEUE_MAX_RUNTIME_MS = 4 * 60 * 1000;   // 트리거 자체 안전 한도 (4분)
var QUEUE_SAFETY_MARGIN_MS = 30 * 1000;     // 다음 job 진입 여유 (30초)

/**
 * 작업을 큐에 추가
 * @param job — { type, token, docNo, docType }
 * @returns 추가된 job (id 포함)
 */
function enqueueJob(job) {
  return withLock(function() {
    var props = PropertiesService.getScriptProperties();
    var queueStr = props.getProperty(QUEUE_KEY) || '[]';
    var queue;
    try { queue = JSON.parse(queueStr); }
    catch(_) { queue = []; }

    // 중복 방지: 같은 토큰의 pending job이 이미 있으면 추가하지 않음
    var existing = queue.filter(function(j) {
      return j.token === job.token && j.status !== 'done' && j.status !== 'failed';
    });
    if (existing.length > 0) {
      Logger.log('[Queue] 중복 enqueue 스킵: ' + job.docNo);
      return existing[0];
    }

    var fullJob = {
      id:         'job_' + Utilities.getUuid().substring(0, 8),
      type:       job.type,
      token:      job.token,
      docNo:      job.docNo,
      docType:    job.docType,
      attempts:   0,
      status:     'pending',
      enqueuedAt: new Date().toISOString(),
      lastError:  '',
    };
    queue.push(fullJob);
    props.setProperty(QUEUE_KEY, JSON.stringify(queue));
    Logger.log('[Queue] enqueue: ' + fullJob.id + ' (' + fullJob.type + ' ' + fullJob.docNo + ')');
    return fullJob;
  });
}

/**
 * 큐 전체 조회 (관리자 모니터링용)
 */
function getQueueStatus() {
  try {
    var props = PropertiesService.getScriptProperties();
    var queueStr = props.getProperty(QUEUE_KEY) || '[]';
    var queue = JSON.parse(queueStr);
    return {
      ok: true,
      total: queue.length,
      pending:   queue.filter(function(j){ return j.status === 'pending'; }).length,
      processing: queue.filter(function(j){ return j.status === 'processing'; }).length,
      failed:    queue.filter(function(j){ return j.status === 'failed'; }).length,
      jobs:      queue,
    };
  } catch(err) {
    return { ok: false, message: err.toString() };
  }
}

/**
 * Time-based Trigger의 진입점.
 * 큐에서 다음 pending job을 1건 꺼내 처리.
 *
 * 설치: installQueueTrigger() 1회 호출
 */
function processQueueTrigger() {
  var startTime = Date.now();
  var processedCount = 0;
  var skippedReason = '';

  // 시간 한도 안에서 계속 다음 job을 꺼내 처리
  while (true) {
    var elapsed = Date.now() - startTime;

    // 안전 마진 체크: 다음 job 한 건이 들어올 여유가 있는지
    // 평균 job 소요 시간을 보수적으로 가정하여 남은 시간이 안전 마진보다 적으면 중단
    if (elapsed > QUEUE_MAX_RUNTIME_MS - QUEUE_SAFETY_MARGIN_MS) {
      skippedReason = '시간 한도 임박 — 다음 트리거에서 계속';
      break;
    }

    // 다음 pending job 1건 선점
    var job = _claimNextJob();
    if (!job) {
      skippedReason = processedCount === 0 ? '처리할 job 없음' : '큐 비움';
      break;
    }

    Logger.log('[Queue] processing 시작: ' + job.id + ' (' + job.type + ' ' + job.docNo + ')');

    // 실제 작업 수행 (락 밖)
    var result = { ok: false, message: '' };
    try {
      if (job.type === 'pdf_only') {
        result = _processPdfOnlyJob(job);
      } else if (job.type === 'pdf_and_consolidate') {
        result = _processPdfAndConsolidateJob(job);
      } else if (job.type === 'insp_pdf') {
        // [이관됨] INSP PDF는 로컬 파이썬(코워크)이 생성. 신규 enqueue는 없으며,
        //  큐에 남은 과거 insp_pdf job은 _processInspPdfJob(무력화)이 조용히 배수한다.
        result = (typeof _processInspPdfJob === 'function')
          ? _processInspPdfJob(job)
          : { ok: true, skipped: true, message: 'INSP PDF 이관됨 — 큐 처리 없음' };
      } else {
        result = { ok: false, message: '알 수 없는 job type: ' + job.type };
      }
    } catch(e) {
      result = { ok: false, message: '처리 중 예외: ' + e.toString() };
    }

    // 결과에 따라 큐 업데이트 (락 안)
    _finalizeJob(job.id, result);
    processedCount++;

    var jobElapsed = Date.now() - startTime;
    Logger.log('[Queue] processing 완료: ' + job.id + ' / ' + result.ok +
               ' / 누적 ' + processedCount + '건 / ' + jobElapsed + 'ms');
  }

  var totalElapsed = Date.now() - startTime;
  Logger.log('[Queue] 트리거 종료: ' + processedCount + '건 처리 / ' +
             totalElapsed + 'ms / ' + skippedReason);

  // 전체 시간 임계 알림 (5분 이상)
  if (totalElapsed > 5 * 60 * 1000) {
    notifyAdminError('큐 트리거 시간 초과 경고: ' + totalElapsed + 'ms / 처리 ' +
                     processedCount + '건');
  }
}

/**
 * 다음 pending job을 선점 (status를 processing으로 변경하고 반환)
 * 동시 트리거 실행 시 같은 job이 두 번 처리되는 것을 막음
 */
function _claimNextJob() {
  return withLock(function() {
    var props = PropertiesService.getScriptProperties();
    var queueStr = props.getProperty(QUEUE_KEY) || '[]';
    var queue;
    try { queue = JSON.parse(queueStr); }
    catch(_) { return null; }

    // FIFO 순서로 첫 pending 찾기
    for (var i = 0; i < queue.length; i++) {
      if (queue[i].status === 'pending') {
        queue[i].status = 'processing';
        queue[i].attempts = (queue[i].attempts || 0) + 1;
        queue[i].lastStartedAt = new Date().toISOString();
        props.setProperty(QUEUE_KEY, JSON.stringify(queue));
        return queue[i];
      }
    }
    return null;
  });
}

/**
 * job 완료/실패 처리 후 큐에서 제거 또는 재시도 상태로 되돌림
 */
function _finalizeJob(jobId, result) {
  return withLock(function() {
    var props = PropertiesService.getScriptProperties();
    var queueStr = props.getProperty(QUEUE_KEY) || '[]';
    var queue;
    try { queue = JSON.parse(queueStr); }
    catch(_) { return; }

    var idx = -1;
    for (var i = 0; i < queue.length; i++) {
      if (queue[i].id === jobId) { idx = i; break; }
    }
    if (idx < 0) return;

    var job = queue[idx];

    if (result.ok) {
      // 성공 — 큐에서 제거
      queue.splice(idx, 1);
      props.setProperty(QUEUE_KEY, JSON.stringify(queue));
      Logger.log('[Queue] 완료 및 제거: ' + jobId);
      return;
    }

    // 실패 — 재시도 가능 여부 확인
    job.lastError = result.message || '알 수 없는 오류';

    if (job.attempts >= QUEUE_MAX_ATTEMPTS) {
      // 최대 재시도 초과 → failed로 표시 (큐에는 남겨두어 관리자 확인 가능)
      job.status = 'failed';
      props.setProperty(QUEUE_KEY, JSON.stringify(queue));
      notifyAdminError('큐 작업 최종 실패: ' + job.docNo + ' (' + job.type + ')\n' +
                       'job ID: ' + job.id + '\n' +
                       '시도 횟수: ' + job.attempts + '\n' +
                       '마지막 오류: ' + job.lastError);
    } else {
      // 다시 pending으로 되돌려 다음 트리거에서 재시도
      job.status = 'pending';
      props.setProperty(QUEUE_KEY, JSON.stringify(queue));
      Logger.log('[Queue] 재시도 대기: ' + jobId + ' (' + job.attempts + '/' + QUEUE_MAX_ATTEMPTS + ')');
    }
  });
}

/**
 * REQ pdf-only job 처리
 */
function _processPdfOnlyJob(job) {
  var pdfResult = generatePdfForDocument(job.token);
  if (!pdfResult.ok) {
    return { ok: false, message: 'PDF 생성 실패: ' + pdfResult.message };
  }
  return { ok: true, message: 'PDF 생성 완료', fileId: pdfResult.fileId };
}

/**
 * PRC pdf + 통합 이동 job 처리
 * 순서: PDF 합성 → FINAL 통합 이동
 */
function _processPdfAndConsolidateJob(job) {
  // 1) PDF 합성 먼저
  var pdfResult = generatePdfForDocument(job.token);
  if (!pdfResult.ok) {
    return { ok: false, message: 'PDF 생성 실패 (이동 중단): ' + pdfResult.message };
  }

  // 2) 통합 이동
  var moveResult = consolidateToFinalByPrc(job.token);
  if (!moveResult.ok && !moveResult.skipped) {
    return { ok: false, message: 'FINAL 이동 실패: ' + moveResult.message };
  }

  return {
    ok: true,
    message: 'PDF + FINAL 통합 이동 완료',
    pdfFileId: pdfResult.fileId,
    movedCount: (moveResult.movedFromReq || 0) + (moveResult.movedFromPrc || 0),
  };
}

/**
 * Time-based Trigger 설치 (관리자가 1회 수동 실행)
 * Apps Script 에디터에서 installQueueTrigger 함수 선택 → ▶ 실행
 */
function installQueueTrigger() {
  // 기존 트리거 제거 (중복 설치 방지)
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'processQueueTrigger') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 1분 간격 트리거 생성
  ScriptApp.newTrigger('processQueueTrigger')
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log('[Queue] Time-based Trigger 설치 완료 (1분 간격)');
  return { ok: true, message: '큐 처리 트리거 설치 완료 (1분 간격)' };
}

/**
 * 큐 트리거 제거 (디버깅용)
 */
function uninstallQueueTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'processQueueTrigger') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  return { ok: true, message: '트리거 제거: ' + removed + '개' };
}

/**
 * 실패한 job들을 다시 pending으로 (관리자 콘솔에서 호출)
 */
function retryFailedJobs() {
  return withLock(function() {
    var props = PropertiesService.getScriptProperties();
    var queueStr = props.getProperty(QUEUE_KEY) || '[]';
    var queue;
    try { queue = JSON.parse(queueStr); }
    catch(_) { return { ok: false, message: '큐 파싱 실패' }; }

    var resetCount = 0;
    queue.forEach(function(j) {
      if (j.status === 'failed') {
        j.status = 'pending';
        j.attempts = 0;
        j.lastError = '';
        resetCount++;
      }
    });
    props.setProperty(QUEUE_KEY, JSON.stringify(queue));
    return { ok: true, message: '재시도 등록: ' + resetCount + '건' };
  });
}

/**
 * 큐 전체 초기화 (긴급 상황용 — 매우 신중하게 사용)
 */
function clearQueue() {
  var actor = getActiveUserEmail();
  if (!isAdminUser(actor)) {
    return { ok: false, message: '관리자만 큐 초기화 가능' };
  }
  PropertiesService.getScriptProperties().deleteProperty(QUEUE_KEY);
  Logger.log('[Queue] 전체 초기화됨 (by ' + actor + ')');
  return { ok: true, message: '큐 초기화 완료' };
}

// ================================================================
// 22. 테스트
// ================================================================

function testCheckDocNo() {
  Logger.log(JSON.stringify(checkDocNoForClient('PRQ-2026-001')));
}

function testApproverList() {
  Logger.log(JSON.stringify(getApproverListForClient()));
}

function testHomeData() {
  Logger.log(JSON.stringify(getHomeDataForClient(), null, 2));
}

function dumpApproverEmails() {
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;  // 단순 형식 검사

  for (var i = 1; i < data.length; i++) {
    var row    = data[i];
    var docNo  = String(row[COL.DOC_NO] || '');
    var status = String(row[COL.STATUS] || '');
    if (!docNo) continue;
    if (status.indexOf('최종승인') >= 0 || status === '폐기' || status.indexOf('반려') >= 0) continue;

    var cur = parseInt(row[COL.CURR_IDX]); if (isNaN(cur)) cur = 0;   // 현재 결재 순번(V열)
    var apprCount = parseInt(row[COL.APPR_COUNT]) || 0;
    Logger.log('▼ [' + (i+1) + '행] ' + docNo + ' / ' + status + ' / 현재순번(V)=' + cur);

    for (var k = 0; k < apprCount; k++) {
      var b = COL.APPR_START + k * COL.APPR_COLS;
      var name  = String(row[b+1] || '');
      var email = String(row[b+2] || '');
      var st    = String(row[b+3] || '');
      var valid = re.test(email);
      Logger.log('   ' + (k+1) + '번째 ' + name +
                 ' / 상태=' + st +
                 ' / email=[' + email + ']' +
                 ' / 형식=' + (valid ? 'OK' : '⚠ 깨짐') +
                 (k === cur ? '  ◀ 지금 결재 차례' : ''));
    }
  }
}


function diagnosePdfAuto() {
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
  var data  = sheet.getDataRange().getValues();

  Logger.log('=== 최종승인 상태 문서 탐색 ===');
  var found = 0;
  for (var i = 1; i < data.length; i++) {
    var row    = data[i];
    var docNo  = String(row[COL.DOC_NO] || '');
    var status = String(row[COL.STATUS] || '');
    var token  = String(row[COL.TOKEN] || '');
    if (!docNo || status.indexOf('최종승인') < 0) continue;

    found++;
    var docType = String(row[COL.DOC_TYPE] || '');
    var driveId = String(row[COL.DRIVE_ID] || '');
    Logger.log('--------------------------------');
    Logger.log((i+1) + '행 / ' + docNo + ' / ' + docType +
               ' / status=' + status + ' / DRIVE_ID=[' + driveId + ']');
    Logger.log('generatePdfForDocument 직접 호출 결과:');
    try {
      var r = generatePdfForDocument(token);
      Logger.log(JSON.stringify(r, null, 2));
    } catch(e) {
      Logger.log('함수 밖 예외: ' + e.toString());
    }
    if (found >= 3) break;  // 너무 많지 않게 최대 3건만
  }
  if (found === 0) Logger.log('최종승인 상태인 문서가 없습니다. (그렇다면 문제는 PDF가 아니라 승인 흐름)');
}

function fixConsolidatePO8852() {
  var TARGET = 'PO-2026-8852';
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (String(r[COL.DOC_NO] || '') !== TARGET) continue;
    if (String(r[COL.DOC_TYPE] || '') !== 'PRC') continue;

    var token = String(r[COL.TOKEN] || '');
    Logger.log('대상 PRC 토큰: ' + token);

    // 안전을 위해 PDF가 있는지 먼저 확인 후 통합 이동
    var pdf = generatePdfForDocument(token);   // 이미 있으면 재생성(이전 버전 정리됨)
    Logger.log('PDF: ' + JSON.stringify(pdf));

    var mv = consolidateToFinalByPrc(token);
    Logger.log('통합 이동 결과: ' + JSON.stringify(mv, null, 2));
    return;
  }
  Logger.log('대상을 찾지 못했습니다: ' + TARGET);
}