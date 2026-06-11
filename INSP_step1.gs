// ================================================================
// INSP.gs — 검수보고서(INSP) 모듈 Step 1: 스키마 + 시트 보장
// ================================================================
// 구매결재시스템 Phase 2 / INSP 구현설계서 v1.0 기준
//
// [배포 방법]
// 1. Apps Script 편집기에서 새 스크립트 파일 "INSP" 생성 후 본 파일 전체 붙여넣기
//    (같은 프로젝트의 .gs 파일은 전역 스코프를 공유하므로 Code.gs의
//     CONFIG, getOrCreateSheet, writeAuditLog 등을 그대로 사용 가능)
// 2. Code.gs의 CONFIG 객체에 아래 1줄 추가 (AUDIT_SHEET_NAME 행 바로 아래):
//
//      INSP_SHEET_NAME:  '검수보고서목록',
//
// 3. Code.gs의 AUDIT_EVENT 객체에 아래 블록 추가
//    ("추후 확장 자리" 주석 위치에 붙여넣기):
//
//      // 검수보고서(INSP) 라이프사이클 — Phase 2 Step 1
//      INSP_SUBMIT:    'INSP_SUBMIT',     // 검수보고서 제출
//      INSP_APPROVE:   'INSP_APPROVE',    // 검수 결재 승인 (단계별)
//      INSP_REJECT:    'INSP_REJECT',     // 검수 결재 반려
//      INSP_RESUBMIT:  'INSP_RESUBMIT',   // 반려 후 수정 재상신
//      INSP_FINALIZED: 'INSP_FINALIZED',  // 최종 검수 승인 → 품의 종결
//
// 4. 편집기에서 testInspStep1() 실행 → 로그 'ALL PASS' 확인
//    (클라이언트 호출 함수가 없으므로 이 단계는 웹앱 재배포 불필요)
// ================================================================

// ================================================================
// 검수보고서목록 열 인덱스 상수 (0-based) — insp-v1.0
// ----------------------------------------------------------------
// 주의:
// - 품의서목록의 COL과 절대 혼용 금지. 검수보고서목록 접근은 반드시 INSP_COL.
// - 결재자 블록은 기존과 동일한 6컬럼 패턴 (label/name/email/status/processedAt/comment)
// - 결재자 블록 접근은 INSP_COL.APPR_START + (단계 idx × APPR_COLS) + 오프셋만 사용
// - PO_NO·REQ_NO·SUBJECT·VENDOR_NAME은 제출 시점 스냅샷(비정규화) — 원본과 동기화하지 않음
// - 상태값은 한글 리터럴: '검토중' | '결재중 (이름)' | '반려' | '최종승인(INSP)'
// - VERDICT는 한글 리터럴: '합격' | '불합격'
// ================================================================
var INSP_COL = {
  SUBMIT_AT:     0,   // 제출 시각 (Date)
  DOC_NO:        1,   // 검수보고서 번호: {REQ docNo}-{NN}  ex) TH-AP-26-029-01
  SEQ:           2,   // 회차 (number, 1~99)
  PRC_TOKEN:     3,   // 대상 PRC token (품의서목록 조인 키)
  REQ_TOKEN:     4,   // 원본 REQ token
  PO_NO:         5,   // PO 번호 (PRC docNo 스냅샷)
  REQ_NO:        6,   // 품의 번호 (REQ docNo 스냅샷)
  SUBJECT:       7,   // 품명 (스냅샷)
  VENDOR_NAME:   8,   // 업체명 (스냅샷)
  DRAFTER:       9,   // 작성자 이메일
  DRAFTER_NAME:  10,  // 작성자 성명
  DEPT:          11,  // 부서
  ISSUE_DATE:    12,  // 품의 기안일자 (스냅샷)
  RECEIVED_DATE: 13,  // 입고일자
  RECEIVED_NOTE: 14,  // 입고 내역 (텍스트 1줄)
  VERDICT:       15,  // '합격' | '불합격'
  IS_FINAL:      16,  // 'Y' | ''  — 최종 검수 여부 (Q-INSP-09)
  COMMENT:       17,  // 검수 의견
  STATUS:        18,  // 상태 (한글)
  TOKEN:         19,  // INSP 고유 토큰 (Utilities.getUuid)
  DRIVE_ID:      20,  // 사진 임시 폴더 ID (STAGING) — PDF 생성 성공 후 비움
  APPR_COUNT:    21,  // 결재자 수 (1~3)
  APPR_IDX:      22,  // 현재 결재 단계 (0-based)
  RESUB_COUNT:   23,  // 재상신 횟수
  REJECT_LOG:    24,  // 반려 이력 누적 텍스트
  PHOTO_LIST:    25,  // JSON: [{id, name, size}] — 임시 사진 파일 목록
  MOVE_STATUS:   26,  // 'STAGING' | 'FINAL'
  APPR_START:    27,  // 결재자 블록 시작 (결재자 1~3, 기안자 미포함)
  APPR_COLS:     6,   // 블록당 컬럼 수: label, name, email, status, processedAt, comment
  MAX_APPROVERS: 3,   // Q-INSP-03: 결재자 최대 3인
};
INSP_COL._VERSION = 'insp-v1.0';

// 시트 총 컬럼 수 (헤더 검증 기준): 27 + 3×6 = 45
var INSP_TOTAL_COLS = INSP_COL.APPR_START + (INSP_COL.MAX_APPROVERS * INSP_COL.APPR_COLS);

/**
 * 검수보고서목록 시트 보장
 * - 시트가 없으면 생성, 헤더가 비어있으면 헤더 작성 + 1행 고정 + 스타일
 * - ensureAuditLogSheet()와 동일 패턴: 이미 데이터가 있으면 즉시 통과
 * @returns {Sheet} 검수보고서목록 시트
 */
function ensureInspSheet() {
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.INSP_SHEET_NAME);

  if (sheet.getLastRow() > 0) return sheet;

  var headers = [
    'submitAt', 'docNo', 'seq', 'prcToken', 'reqToken',
    'poNo', 'reqNo', 'subject', 'vendorName',
    'drafter', 'drafterName', 'dept', 'issueDate',
    'receivedDate', 'receivedNote', 'verdict', 'isFinal', 'comment',
    'status', 'token', 'driveId',
    'apprCount', 'apprIdx', 'resubCount', 'rejectLog',
    'photoList', 'moveStatus',
  ];
  // 결재자 블록 헤더 (결재자 1~3 × 6컬럼)
  for (var i = 1; i <= INSP_COL.MAX_APPROVERS; i++) {
    headers.push(
      'appr' + i + '_label',
      'appr' + i + '_name',
      'appr' + i + '_email',
      'appr' + i + '_status',
      'appr' + i + '_processedAt',
      'appr' + i + '_comment'
    );
  }

  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#e4f4f2');   // INSP 테마 (청록) — 품의서목록과 시각적 구분

  // 컬럼 폭 권장
  sheet.setColumnWidth(INSP_COL.SUBMIT_AT + 1, 150);
  sheet.setColumnWidth(INSP_COL.DOC_NO + 1, 160);
  sheet.setColumnWidth(INSP_COL.PRC_TOKEN + 1, 250);
  sheet.setColumnWidth(INSP_COL.REQ_TOKEN + 1, 250);
  sheet.setColumnWidth(INSP_COL.SUBJECT + 1, 220);
  sheet.setColumnWidth(INSP_COL.RECEIVED_NOTE + 1, 220);
  sheet.setColumnWidth(INSP_COL.COMMENT + 1, 250);
  sheet.setColumnWidth(INSP_COL.TOKEN + 1, 250);
  sheet.setColumnWidth(INSP_COL.PHOTO_LIST + 1, 300);

  return sheet;
}

/**
 * INSP 스키마 자가 점검 (PDF payload 화이트리스트의 자체 점검 throw 패턴)
 * - INSP_COL 본체 인덱스(메타 키 제외)가 0부터 빈틈없이 연속하는지 검증
 * - 컬럼 추가/삭제 시 인덱스 누락·중복을 배포 전에 잡아내기 위함
 * @throws {Error} 스키마 불일치 시
 */
function _inspSchemaSelfCheck() {
  var META_KEYS = { APPR_COLS: true, MAX_APPROVERS: true, _VERSION: true };
  var indices = [];
  for (var key in INSP_COL) {
    if (META_KEYS[key]) continue;
    indices.push(INSP_COL[key]);
  }
  indices.sort(function (a, b) { return a - b; });

  for (var i = 0; i < indices.length; i++) {
    if (indices[i] !== i) {
      throw new Error(
        '[INSP] 스키마 자가 점검 실패: 인덱스 ' + i + ' 자리에 ' + indices[i] +
        ' 발견 (누락 또는 중복). INSP_COL 정의를 확인하세요. (' + INSP_COL._VERSION + ')'
      );
    }
  }
  if (indices.length !== INSP_COL.APPR_START + 1) {
    throw new Error(
      '[INSP] 스키마 자가 점검 실패: 본체 키 수(' + indices.length +
      ')가 APPR_START+1(' + (INSP_COL.APPR_START + 1) + ')과 불일치.'
    );
  }
  return true;
}

/**
 * 결재자 블록 셀 인덱스 계산 헬퍼
 * - 이후 모든 Step에서 결재 블록 접근은 반드시 이 함수를 통할 것 (하드코딩 금지)
 * @param {number} stageIdx 결재 단계 (0-based, 0~2)
 * @param {number} offset   0:label 1:name 2:email 3:status 4:processedAt 5:comment
 * @returns {number} 0-based 컬럼 인덱스
 */
function inspApprCol(stageIdx, offset) {
  if (stageIdx < 0 || stageIdx >= INSP_COL.MAX_APPROVERS) {
    throw new Error('[INSP] 잘못된 결재 단계: ' + stageIdx);
  }
  if (offset < 0 || offset >= INSP_COL.APPR_COLS) {
    throw new Error('[INSP] 잘못된 블록 오프셋: ' + offset);
  }
  return INSP_COL.APPR_START + (stageIdx * INSP_COL.APPR_COLS) + offset;
}

// ================================================================
// Step 1 테스트 — 편집기에서 직접 실행
// ================================================================

/**
 * Step 1 통합 테스트
 * 검증: ① 스키마 자가 점검 ② CONFIG 키 존재 ③ AUDIT_EVENT 확장 반영
 *       ④ 시트 생성 + 헤더 45컬럼 ⑤ 헬퍼 함수 경계값
 * 실행 후 로그에서 'ALL PASS' 확인. 실패 항목은 [FAIL]로 표시됨.
 */
function testInspStep1() {
  var results = [];
  function check(name, fn) {
    try {
      var r = fn();
      results.push((r ? '[PASS] ' : '[FAIL] ') + name);
      return !!r;
    } catch (e) {
      results.push('[FAIL] ' + name + ' — ' + e.message);
      return false;
    }
  }

  check('① INSP_COL 스키마 자가 점검', function () {
    return _inspSchemaSelfCheck();
  });

  check('② CONFIG.INSP_SHEET_NAME 정의됨', function () {
    return CONFIG.INSP_SHEET_NAME === '검수보고서목록';
  });

  check('③ AUDIT_EVENT에 INSP 이벤트 5종 추가됨', function () {
    return AUDIT_EVENT.INSP_SUBMIT === 'INSP_SUBMIT'
      && AUDIT_EVENT.INSP_APPROVE === 'INSP_APPROVE'
      && AUDIT_EVENT.INSP_REJECT === 'INSP_REJECT'
      && AUDIT_EVENT.INSP_RESUBMIT === 'INSP_RESUBMIT'
      && AUDIT_EVENT.INSP_FINALIZED === 'INSP_FINALIZED';
  });

  check('④ ensureInspSheet — 시트 생성 + 헤더 ' + INSP_TOTAL_COLS + '컬럼', function () {
    var sheet = ensureInspSheet();
    var lastCol = sheet.getLastColumn();
    if (lastCol !== INSP_TOTAL_COLS) {
      throw new Error('헤더 컬럼 수 ' + lastCol + ' ≠ 기대값 ' + INSP_TOTAL_COLS);
    }
    var h = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (h[INSP_COL.STATUS] !== 'status') throw new Error('STATUS 헤더 위치 불일치: ' + h[INSP_COL.STATUS]);
    if (h[INSP_COL.APPR_START] !== 'appr1_label') throw new Error('APPR_START 헤더 위치 불일치: ' + h[INSP_COL.APPR_START]);
    return true;
  });

  check('⑤ inspApprCol 헬퍼 — 정상/경계값', function () {
    if (inspApprCol(0, 0) !== 27) throw new Error('(0,0) → ' + inspApprCol(0, 0));
    if (inspApprCol(2, 5) !== 44) throw new Error('(2,5) → ' + inspApprCol(2, 5));
    var threw = false;
    try { inspApprCol(3, 0); } catch (e) { threw = true; }
    if (!threw) throw new Error('단계 범위 초과가 차단되지 않음');
    return true;
  });

  var allPass = results.every(function (r) { return r.indexOf('[PASS]') === 0; });
  results.push(allPass ? '=== ALL PASS — Step 2 진행 가능 ===' : '=== 실패 항목 확인 필요 ===');
  console.log(results.join('\n'));
  return allPass;
}
