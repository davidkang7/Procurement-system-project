// ================================================================
// INSP.gs — 검수보고서(INSP) 모듈 (누적본: Step 1~3)
// Step 1: 스키마 + 시트 보장 / Step 2: 홈 "결재 완료" 그룹 조회
// Step 3: 작성 폼 + 제출 + 사진 임시 저장
// Step 4: 결재 승인/반려 + 최종 진행중 차단 + 결재란 기안자 포함(4블록)
// Step 5: 전용 뷰어 + 내 검수 결재 대기 + 최종 완료(INSP)
// Step 6: 최종승인 시 PDF 생성 → FINAL/PO폴더 이동 → 사진 삭제 (insp_pdf 큐 job)
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
  MOVE_STATUS:   26,  // 'STAGING' | 'PENDING_PDF'(결재완료·로컬 PDF 대기) | 'FINAL'(PDF 생성 완료)
  APPR_START:    27,  // 결재 블록 시작 (블록0=기안자, 블록1~3=결재자 1~3 — REQ/PRC 동일 구조)
  APPR_COLS:     6,   // 블록당 컬럼 수: label, name, email, status, processedAt, comment
  MAX_APPROVERS: 4,   // 기안자 1 + 결재자 최대 3 (Q-INSP-03). approvers[0]=기안자
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
  // 결재 블록 헤더 (블록0=기안자, 블록1~3=결재자 1~3 — REQ/PRC 동일 구조)
  // approvers[0]=기안자(자기결재), approvers[1..3]=결재자
  var apprPrefix = ['drafter_appr', 'appr1', 'appr2', 'appr3'];
  for (var i = 0; i < INSP_COL.MAX_APPROVERS; i++) {
    var px = apprPrefix[i];
    headers.push(
      px + '_label', px + '_name', px + '_email',
      px + '_status', px + '_processedAt', px + '_comment'
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
// [INSP Step 2] 홈 대시보드용 — 검수보고서 그룹 조회
// ================================================================

/**
 * 검수보고서목록을 1회 read하여 PRC token별로 그룹핑해 반환
 * - getHomeDataForClient에서 호출 (동일 스프레드시트 — openById 재사용, 추가 read 1회)
 * - 시트가 없거나 데이터가 없으면 빈 객체 반환 (홈 로딩을 중단시키지 않음)
 * - Date 필드는 전부 문자열 변환 (google.script.run 직렬화 시 Date → null 방지)
 * @param {Spreadsheet} ss 이미 열린 스프레드시트 객체
 * @returns {Object} { prcToken: [inspObj, ...] } — 각 그룹은 회차(seq) 오름차순
 */
function _getInspGroupsByPrc(ss) {
  var groups = {};
  try {
    var sheet = ss.getSheetByName(CONFIG.INSP_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return groups;

    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      var prcToken = String(r[INSP_COL.PRC_TOKEN] || '');
      if (!prcToken) continue;

      var status = String(r[INSP_COL.STATUS] || '');
      var curIdx = parseInt(r[INSP_COL.APPR_IDX]) || 0;

      // 결재 진행 중일 때 현재 결재자명 (블록 접근은 inspApprCol 헬퍼만 사용)
      var currentApprName = '';
      if ((status === '검토중' || status.indexOf('결재중') >= 0) &&
          curIdx >= 0 && curIdx < INSP_COL.MAX_APPROVERS) {
        currentApprName = String(r[inspApprCol(curIdx, 1)] || '');
      }

      var obj = {
        docNo:           String(r[INSP_COL.DOC_NO] || ''),
        seq:             parseInt(r[INSP_COL.SEQ]) || 0,
        status:          status,
        receivedDate:    toDateStr(r[INSP_COL.RECEIVED_DATE]),
        receivedNote:    String(r[INSP_COL.RECEIVED_NOTE] || ''),
        verdict:         String(r[INSP_COL.VERDICT] || ''),
        isFinal:         String(r[INSP_COL.IS_FINAL] || '') === 'Y',
        comment:         String(r[INSP_COL.COMMENT] || ''),
        token:           String(r[INSP_COL.TOKEN] || ''),
        submitAt:        toDateTimeStr(r[INSP_COL.SUBMIT_AT]),
        currentApprName: currentApprName,
      };

      if (!groups[prcToken]) groups[prcToken] = [];
      groups[prcToken].push(obj);
    }

    // 각 그룹 회차 오름차순 정렬
    for (var key in groups) {
      groups[key].sort(function (a, b) { return a.seq - b.seq; });
    }
  } catch (e) {
    // fail-safe: 그룹 조회 실패가 홈 로딩 전체를 중단시키지 않음
    console.error('[INSP] _getInspGroupsByPrc 실패: ' + e.toString());
  }
  return groups;
}

// ================================================================
// [INSP 부서 공유] 같은 부서 판정 (검수보고서 제출 권한 확장)
// - 결재자목록 부서 열 기준. Code.gs의 getDeptByEmailMap 재사용.
// - 두 이메일이 모두 등록되어 있고 부서가 같으면 true
// ================================================================
function _isSameDeptEmail(ss, emailA, emailB) {
  if (!emailA || !emailB) return false;
  if (typeof getDeptByEmailMap !== 'function') return false;
  var m  = getDeptByEmailMap(ss);
  var da = m[String(emailA).trim().toLowerCase()];
  var db = m[String(emailB).trim().toLowerCase()];
  return !!da && !!db && da === db;
}

// ================================================================
// [INSP Step 3] 작성 폼 데이터 조회
// ================================================================

/**
 * 검수보고서 작성 폼 초기 데이터
 * - PRC token으로 품의(PRC)+원본(REQ) 스냅샷 필드를 모아 폼에 자동 채움
 * - 권한: 원본 REQ 기안자 본인만 (이메일 소문자 정규화 비교)
 * - 종결 차단: 동일 PRC에 최종 검수(IS_FINAL=Y)가 최종승인(INSP)이면 추가 작성 불가
 * - 재상신: resubToken이 있으면 기존 반려 건의 값을 불러와 prefill
 * @param {string} prcToken 대상 PRC token
 * @param {string=} resubToken 재상신할 기존 INSP token (선택)
 * @returns {Object} ok / form(스냅샷) / nextSeq / resub(기존값) / approverList
 */
function getInspFormDataForClient(prcToken, resubToken) {
  try {
    var actor = (getActiveUserEmail() || '').toLowerCase();
    if (!prcToken) return { ok: false, message: '구매 문서 토큰이 필요합니다.' };

    var ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet  = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
    var rows   = sheet.getDataRange().getValues();

    // 대상 PRC 찾기
    var prc = null, reqByToken = {};
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      var docType = String(r[COL.DOC_TYPE] || 'REQ');
      var token   = String(r[COL.TOKEN] || '');
      if (docType === 'REQ') {
        reqByToken[token] = {
          docNo:       String(r[COL.DOC_NO] || ''),
          drafterEmail:String(r[COL.APPR_START + 2] || ''),
          drafterName: String(r[COL.DRAFTER] || ''),
          dept:        String(r[COL.DEPT] || ''),
          issueDate:   toDateStr(r[COL.ISSUE_DATE]),
        };
      }
      if (docType === 'PRC' && token === prcToken) {
        prc = {
          token:       token,
          docNo:       String(r[COL.DOC_NO] || ''),
          subject:     String(r[COL.SUBJECT] || ''),
          vendorName:  String(r[COL.VENDOR_NAME] || ''),
          parentToken: String(r[COL.PARENT_DOC_ID] || ''),
          status:      String(r[COL.STATUS] || ''),
        };
      }
    }

    if (!prc) return { ok: false, message: '대상 구매품의서(구매)를 찾을 수 없습니다.' };
    if (prc.status !== '최종승인(PRC)') {
      return { ok: false, message: '최종승인(구매) 상태의 품의에만 검수보고서를 제출할 수 있습니다.' };
    }

    var req = reqByToken[prc.parentToken] || null;
    var drafterEmail = req ? String(req.drafterEmail || '') : '';

    // 권한: 원본 REQ 기안자 본인 또는 같은 부서 담당자 (부서 공유)
    if (!drafterEmail ||
        (drafterEmail.toLowerCase() !== actor && !_isSameDeptEmail(ss, drafterEmail, actor))) {
      return { ok: false, message: '검수보고서는 원본 품의 기안자 또는 같은 부서 담당자만 제출할 수 있습니다.' };
    }

    // 기존 검수보고서 그룹 (회차 채번 + 종결 차단 + 재상신 prefill)
    var groups = _getInspGroupsByPrc(ss);
    var existing = groups[prcToken] || [];

    var locked = _isInspLocked(existing);
    if (locked && !resubToken) {
      var finalApproved = existing.some(function(x){ return x.isFinal && x.status === '최종승인(INSP)'; });
      return { ok: false, message: finalApproved
        ? '이미 최종 검수가 승인되어 종결된 품의입니다. 추가 제출이 불가합니다.'
        : '최종 검수 보고서가 이미 제출되어 결재 진행 중입니다. 추가 제출이 불가합니다. (해당 건이 반려되면 재상신할 수 있습니다)' };
    }

    var maxSeq = 0;
    existing.forEach(function(x) { if (x.seq > maxSeq) maxSeq = x.seq; });
    var nextSeq = maxSeq + 1;
    if (nextSeq > 99) return { ok: false, message: '검수보고서 회차가 한도(99)를 초과했습니다.' };

    // 재상신 prefill (반려 건만)
    var resub = null;
    if (resubToken) {
      var rinfo = _readInspRowByToken(ss, resubToken);
      if (!rinfo) return { ok: false, message: '재상신할 검수보고서를 찾을 수 없습니다.' };
      if (rinfo.row[INSP_COL.STATUS] !== '반려') {
        return { ok: false, message: '반려 상태의 검수보고서만 재상신할 수 있습니다.' };
      }
      var rr = rinfo.row;
      resub = {
        token:        resubToken,
        seq:          parseInt(rr[INSP_COL.SEQ]) || nextSeq,
        receivedDate: toDateStr(rr[INSP_COL.RECEIVED_DATE]),
        receivedNote: String(rr[INSP_COL.RECEIVED_NOTE] || ''),
        verdict:      String(rr[INSP_COL.VERDICT] || ''),
        isFinal:      String(rr[INSP_COL.IS_FINAL] || '') === 'Y',
        comment:      String(rr[INSP_COL.COMMENT] || ''),
        rejectLog:    String(rr[INSP_COL.REJECT_LOG] || ''),
      };
    }

    // 현재 로그인 사용자 (담당자 표시용)
    var cur = getCurrentUserForClient();
    var apprList = getApproverListForClient();

    return {
      ok: true,
      form: {
        prcToken:    prc.token,
        reqToken:    prc.parentToken,
        poNo:        prc.docNo,
        reqNo:       req ? req.docNo : '',
        subject:     prc.subject,
        vendorName:  prc.vendorName,
        drafterEmail:drafterEmail,
        drafterName: req ? req.drafterName : (cur.user ? cur.user.name : ''),
        dept:        req ? req.dept : (cur.user ? cur.user.dept : ''),
        issueDate:   req ? req.issueDate : '',
        nextSeq:     resub ? resub.seq : nextSeq,
        docNo:       (req ? req.docNo : prc.docNo) + '-' + _pad2(resub ? resub.seq : nextSeq),
      },
      resub:        resub,
      currentUser:  cur.user || null,
      approverList: apprList.approvers || [],
    };
  } catch (err) {
    return { ok: false, message: err.toString() + ' / ' + (err.stack || '') };
  }
}

// ================================================================
// [INSP Step 3] 검수보고서 제출 / 재상신
// ================================================================

/**
 * 검수보고서 제출 (신규 또는 재상신)
 * - 권한·종결·중복회차 검증은 락 안에서 (동시 제출 경합 방지 — 회차 채번 포함)
 * - 사진 업로드는 락 밖에서 (기존 첨부 처리 원칙과 동일)
 * - 사진은 INSP STAGING 폴더에 임시 저장, PDF 생성(Step 6) 후 삭제 예정
 * @param {Object} data 폼 페이로드
 * @returns {Object} ok / docNo / token
 */
function submitInspForClient(data) {
  return _submitInspCore(data);
}

function _submitInspCore(data) {
  try {
    var actor = (getActiveUserEmail() || '').toLowerCase();
    if (!actor) return { ok: false, message: '로그인 사용자를 확인할 수 없습니다.' };

    // 판정/입고내역 필수
    if (data.verdict !== '합격' && data.verdict !== '불합격') {
      return { ok: false, message: '검수 판정(합격/불합격)을 선택해 주세요.' };
    }
    if (!data.receivedNote || !String(data.receivedNote).trim()) {
      return { ok: false, message: '입고 내역을 입력해 주세요.' };
    }
    // approvers[0]=기안자(본인), [1..]=결재자 (REQ/PRC 동일 구조)
    var approvers = data.approvers || [];
    if (approvers.length < 2) return { ok: false, message: '결재자를 1명 이상 지정해 주세요. (기안자 외)' };
    if (approvers.length > INSP_COL.MAX_APPROVERS) {
      return { ok: false, message: '결재자는 최대 ' + (INSP_COL.MAX_APPROVERS - 1) + '인까지 지정 가능합니다.' };
    }
    // 기안자(0번) 본인 검증
    var drafterAppr = (approvers[0] && approvers[0].email ? approvers[0].email : '').toLowerCase();
    if (!drafterAppr) return { ok: false, message: '기안자(본인) 정보가 비어 있습니다. 새로고침 후 다시 시도해 주세요.' };
    if (drafterAppr !== actor) {
      return { ok: false, message: '기안자는 본인만 지정할 수 있습니다. (로그인: ' + actor + ' / 기안자: ' + drafterAppr + ')' };
    }

    // 사진 사전 검증 (락 밖)
    var photos = data.photos || [];
    if (photos.length > 6) return { ok: false, message: '검수 사진은 최대 6장까지 첨부 가능합니다.' };
    for (var p = 0; p < photos.length; p++) {
      if (photos[p].size && photos[p].size > 5 * 1024 * 1024) {
        return { ok: false, message: '검수 사진은 장당 5MB를 초과할 수 없습니다: ' + (photos[p].name || '') };
      }
    }

    // ── 락 안: 검증 + 행 기록 (회차 채번 경합 방지) ──
    var lockResult = withLock(function() {
      try {
        var ss        = SpreadsheetApp.openById(CONFIG.SHEET_ID);
        var docSheet  = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
        var inspSheet = ensureInspSheet();

        // 대상 PRC + 원본 REQ 재검증 (폼 데이터 신뢰 금지)
        var prc = _findRowByTokenInDoc(docSheet, data.prcToken);
        if (!prc) return { ok: false, message: '대상 구매 문서를 찾을 수 없습니다.' };
        if (String(prc.row[COL.STATUS] || '') !== '최종승인(PRC)') {
          return { ok: false, message: '최종승인(구매) 상태가 아닙니다.' };
        }
        var parentToken = String(prc.row[COL.PARENT_DOC_ID] || '');
        var reqInfo = _findRowByTokenInDoc(docSheet, parentToken);
        var drafterEmail = reqInfo ? String(reqInfo.row[COL.APPR_START + 2] || '') : '';
        if (!drafterEmail ||
            (drafterEmail.toLowerCase() !== actor && !_isSameDeptEmail(ss, drafterEmail, actor))) {
          return { ok: false, message: '원본 품의 기안자 또는 같은 부서 담당자만 제출할 수 있습니다.' };
        }

        // 기존 INSP 그룹 (회차 + 종결)
        var groups = _getInspGroupsByPrc(ss);
        var existing = groups[data.prcToken] || [];
        var locked = _isInspLocked(existing);

        // 재상신 분기
        var isResub = !!data.resubToken;
        var resubRow = null;
        if (isResub) {
          resubRow = _readInspRowByToken(ss, data.resubToken);
          if (!resubRow) return { ok: false, message: '재상신 대상을 찾을 수 없습니다.' };
          if (resubRow.row[INSP_COL.STATUS] !== '반려') {
            return { ok: false, message: '반려 상태만 재상신할 수 있습니다.' };
          }
        } else {
          if (locked) {
            var fa = existing.some(function(x){ return x.isFinal && x.status === '최종승인(INSP)'; });
            return { ok: false, message: fa
              ? '이미 최종 검수가 승인되어 종결된 품의입니다.'
              : '최종 검수 보고서가 결재 진행 중입니다. 추가 제출이 불가합니다.' };
          }
        }

        // 스냅샷 (제출 시점 고정 — 원본 변경과 비동기화)
        var poNo       = String(prc.row[COL.DOC_NO] || '');
        var reqNo      = reqInfo ? String(reqInfo.row[COL.DOC_NO] || '') : '';
        var subject    = String(prc.row[COL.SUBJECT] || '');
        var vendorName = String(prc.row[COL.VENDOR_NAME] || '');
        var drafterName= reqInfo ? String(reqInfo.row[COL.DRAFTER] || '') : '';
        var dept       = reqInfo ? String(reqInfo.row[COL.DEPT] || '') : '';
        var issueDate  = reqInfo ? toDateStr(reqInfo.row[COL.ISSUE_DATE]) : '';

        var seq, token, docNo, targetRowNum, resubCount = 0;

        if (isResub) {
          seq        = parseInt(resubRow.row[INSP_COL.SEQ]) || 1;
          token      = String(resubRow.row[INSP_COL.TOKEN] || Utilities.getUuid());
          docNo      = reqNo + '-' + _pad2(seq);
          resubCount = (parseInt(resubRow.row[INSP_COL.RESUB_COUNT]) || 0) + 1;
          targetRowNum = resubRow.rowNum;
        } else {
          var maxSeq = 0;
          existing.forEach(function(x) { if (x.seq > maxSeq) maxSeq = x.seq; });
          seq   = maxSeq + 1;
          if (seq > 99) return { ok: false, message: '회차 한도(99) 초과.' };
          token = Utilities.getUuid();
          docNo = reqNo + '-' + _pad2(seq);
        }

        // INSP STAGING 폴더 (사진 임시 보관) — 사진 저장은 락 밖에서
        var folder = getOrCreateFolder('INSP_' + docNo, issueDate, 'staging');
        var folderId = folder.getId();

        // 행 구성 (45컬럼)
        var firstApprName = approvers[0].name || '';
        var rowArr = new Array(INSP_TOTAL_COLS);
        rowArr[INSP_COL.SUBMIT_AT]     = new Date();
        rowArr[INSP_COL.DOC_NO]        = docNo;
        rowArr[INSP_COL.SEQ]           = seq;
        rowArr[INSP_COL.PRC_TOKEN]     = data.prcToken;
        rowArr[INSP_COL.REQ_TOKEN]     = parentToken;
        rowArr[INSP_COL.PO_NO]         = poNo;
        rowArr[INSP_COL.REQ_NO]        = reqNo;
        rowArr[INSP_COL.SUBJECT]       = subject;
        rowArr[INSP_COL.VENDOR_NAME]   = vendorName;
        rowArr[INSP_COL.DRAFTER]       = drafterEmail;   // 이메일
        rowArr[INSP_COL.DRAFTER_NAME]  = drafterName;    // 성명
        rowArr[INSP_COL.DEPT]          = dept;
        rowArr[INSP_COL.ISSUE_DATE]    = issueDate;
        rowArr[INSP_COL.RECEIVED_DATE] = data.receivedDate || '';
        rowArr[INSP_COL.RECEIVED_NOTE] = String(data.receivedNote || '');
        rowArr[INSP_COL.VERDICT]       = data.verdict;
        rowArr[INSP_COL.IS_FINAL]      = data.isFinal ? 'Y' : '';
        rowArr[INSP_COL.COMMENT]       = String(data.comment || '');
        rowArr[INSP_COL.STATUS]        = '검토중';
        rowArr[INSP_COL.TOKEN]         = token;
        rowArr[INSP_COL.DRIVE_ID]      = folderId;
        rowArr[INSP_COL.APPR_COUNT]    = approvers.length;
        rowArr[INSP_COL.APPR_IDX]      = 0;
        rowArr[INSP_COL.RESUB_COUNT]   = resubCount;
        rowArr[INSP_COL.REJECT_LOG]    = isResub ? String(resubRow.row[INSP_COL.REJECT_LOG] || '') : '';
        rowArr[INSP_COL.PHOTO_LIST]    = '[]';   // 락 밖에서 채움
        rowArr[INSP_COL.MOVE_STATUS]   = 'STAGING';

        // 결재자 블록 (label/name/email/status/processedAt/comment)
        for (var a = 0; a < INSP_COL.MAX_APPROVERS; a++) {
          var ap = approvers[a];
          rowArr[inspApprCol(a, 0)] = ap ? (ap.label || (a === 0 ? '기안자' : ('결재자 ' + a))) : '';
          rowArr[inspApprCol(a, 1)] = ap ? (ap.name || '') : '';
          rowArr[inspApprCol(a, 2)] = ap ? (ap.email || '') : '';
          rowArr[inspApprCol(a, 3)] = ap ? '대기' : '';
          rowArr[inspApprCol(a, 4)] = '';
          rowArr[inspApprCol(a, 5)] = '';
        }
        for (var z = 0; z < rowArr.length; z++) { if (rowArr[z] === undefined) rowArr[z] = ''; }

        if (isResub) {
          inspSheet.getRange(targetRowNum, 1, 1, INSP_TOTAL_COLS).setValues([rowArr]);
        } else {
          inspSheet.appendRow(rowArr);
          targetRowNum = inspSheet.getLastRow();
        }

        return { ok: true, rowNum: targetRowNum, token: token, docNo: docNo,
                 folderId: folderId, seq: seq, isResub: isResub,
                 firstApprover: approvers[0], snapshot: {
                   reqNo: reqNo, poNo: poNo, subject: subject, vendorName: vendorName,
                   drafterName: drafterName, dept: dept,
                 } };
      } catch (e) {
        return { ok: false, message: '제출 처리 실패: ' + e.toString() };
      }
    });

    if (!lockResult.ok) return lockResult;

    // ── 락 밖: 사진 임시 저장 ──
    var savedPhotos = [];
    try {
      var folder = DriveApp.getFolderById(lockResult.folderId);
      savedPhotos = _saveInspPhotos(folder, photos);
      var ss2    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet2 = ss2.getSheetByName(CONFIG.INSP_SHEET_NAME);
      sheet2.getRange(lockResult.rowNum, INSP_COL.PHOTO_LIST + 1)
        .setValue(JSON.stringify(savedPhotos));
    } catch (e) {
      // 사진 실패는 제출 자체를 막지 않음 (행은 이미 기록됨) — 관리자 알림
      try { notifyAdminError('[INSP] 사진 임시 저장 실패: ' + lockResult.docNo + ' / ' + e.toString()); } catch (_) {}
    }

    // ── 락 밖: 1차 결재자 메일 ──
    try {
      var s = lockResult.snapshot;
      var mailData = {
        docNo:      lockResult.docNo,
        subject:    '[검수] ' + s.subject,
        drafter:    lockResult.snapshot.drafterName,
        vendorName: s.vendorName,
        dept:       s.dept,
      };
      if (lockResult.firstApprover && lockResult.firstApprover.email) {
        _sendInspApprovalEmail(lockResult.firstApprover, mailData, lockResult.token, lockResult.rowNum, 0);
      }
    } catch (e) {
      try { notifyAdminError('[INSP] 결재 요청 메일 실패: ' + lockResult.docNo + ' / ' + e.toString()); } catch (_) {}
    }

    // 감사 로그
    try {
      writeAuditLog({
        eventType: lockResult.isResub ? AUDIT_EVENT.INSP_RESUBMIT : AUDIT_EVENT.INSP_SUBMIT,
        docNo: lockResult.docNo, docToken: lockResult.token, docType: 'INSP',
        actor: actor, reason: '검수보고서 ' + (lockResult.isResub ? '재상신' : '제출') + ' (' + lockResult.seq + '회차)',
      });
    } catch (_) {}

    return {
      ok: true,
      message: '검수보고서 ' + lockResult.docNo + ' 제출 완료. 1차 결재자에게 메일이 발송됩니다.',
      docNo: lockResult.docNo, token: lockResult.token, photos: savedPhotos.length,
    };
  } catch (err) {
    return { ok: false, message: err.toString() + ' / ' + (err.stack || '') };
  }
}

// ================================================================
// [INSP Step 3] 내부 헬퍼
// ================================================================

function _pad2(n) { n = parseInt(n) || 0; return n < 10 ? '0' + n : String(n); }

/**
 * 본인의 결재 단계(idx) 판정 — 동일인이 여러 단계에 배정된 경우 대응
 * - 현재 차례(curIdx)의 결재자가 본인이면 그 단계를 우선 반환 (이미 승인한 앞 단계가 아니라 '지금 처리할' 단계)
 * - 아니면 본인이 배정된 첫 단계 반환 (열람/대기 판정용)
 * @param {Array} row INSP 행 배열
 * @param {number} apprCount 결재자 수
 * @param {number} curIdx 현재 결재 차례
 * @param {string} actor 로그인 이메일(소문자)
 * @returns {number} 본인 결재 단계 idx, 없으면 -1
 */
function _findMyInspIdx(row, apprCount, curIdx, actor) {
  actor = String(actor || '').toLowerCase();
  // 1) 현재 차례가 본인이면 그 단계 우선 (동일인 다단계 핵심)
  if (curIdx >= 0 && curIdx < apprCount &&
      String(row[inspApprCol(curIdx, 2)] || '').toLowerCase() === actor) {
    return curIdx;
  }
  // 2) 본인이 배정된 첫 단계
  for (var a = 0; a < apprCount; a++) {
    if (String(row[inspApprCol(a, 2)] || '').toLowerCase() === actor) return a;
  }
  return -1;
}

/**
 * 품의 종결(추가 제출 차단) 여부 판정
 * - 최종 검수(isFinal) 회차가 '반려'가 아닌 상태로 존재하면 차단
 *   (= 결재 대기/진행중 '검토중'·'결재중', 또는 이미 '최종승인(INSP)')
 * - 반려된 최종 회차는 재상신 대상이므로 차단에서 제외
 * @param {Array} inspList _getInspGroupsByPrc 그룹의 항목 배열
 * @returns {boolean} true면 신규 제출 차단
 */
function _isInspLocked(inspList) {
  return (inspList || []).some(function (x) {
    return x.isFinal && x.status !== '반려';
  });
}

/** 품의서목록에서 token으로 행 1건 조회 → {rowNum, row} | null */
function _findRowByTokenInDoc(docSheet, token) {
  if (!token) return null;
  var rows = docSheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][COL.TOKEN] || '') === token) {
      return { rowNum: i + 1, row: rows[i] };
    }
  }
  return null;
}

/** 검수보고서목록에서 token으로 행 1건 조회 → {rowNum, row} | null */
function _readInspRowByToken(ss, token) {
  if (!token) return null;
  var sheet = ss.getSheetByName(CONFIG.INSP_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][INSP_COL.TOKEN] || '') === token) {
      return { rowNum: i + 1, row: rows[i] };
    }
  }
  return null;
}

/**
 * 검수 사진 임시 저장 (이미지 전용 — ALLOWED_EXTS 검증 우회)
 * 클라이언트에서 canvas 압축(긴 변 1600px, JPEG 0.8) 후 base64 전송 전제
 * @returns {Array} [{id, name, size}]
 */
function _saveInspPhotos(folder, photos) {
  var saved = [];
  (photos || []).forEach(function (ph, idx) {
    if (!ph || !ph.data) return;
    var bytes = Utilities.base64Decode(ph.data);
    var name  = ph.name || ('insp_photo_' + (idx + 1) + '.jpg');
    var blob  = Utilities.newBlob(bytes, ph.type || 'image/jpeg', name);
    var file  = folder.createFile(blob);
    saved.push({ id: file.getId(), name: name, size: ph.size || bytes.length });
  });
  return saved;
}

/**
 * 검수보고서 결재 요청 메일 — 기존 sendApprovalEmailWithRetry 재사용 시도,
 * 시그니처/링크가 INSP와 달라 별도 래퍼로 분리(approve 링크에 docKind=INSP 부여).
 * Step 4(결재 처리)에서 본 함수의 링크를 받는 라우팅을 추가한다.
 */
function _sendInspApprovalEmail(approver, data, token, rowNum, idx) {
  // 구매품의서 결재요청 메일과 동일 구성 (확인 + 반려 + 승인 3버튼)
  // 7번 해결책: 모든 버튼이 전용 뷰어(insp_view)로 가고, 뷰어가 서버 myApproverIdx로 본인 단계 재판정
  var approverName = escapeHtml(approver.name || '');
  var approverLabel = escapeHtml(approver.label || '');
  var docNo    = escapeHtml(data.docNo || '-');
  var docTitle = escapeHtml(data.subject || '검수보고서');
  var vendor   = escapeHtml(data.vendorName || '-');
  var drafter  = escapeHtml(data.drafter || '-');

  // 확인 = 전용 뷰어(본문·사진), 승인/반려 = 경량 결재 화면(Procurement_Approval, docKind=INSP)
  //  — 구매품의서와 동일한 분기. 경량 화면은 본문 없이 의견+버튼만 (가벼운 결재)
  var viewBase = CONFIG.WEBAPP_URL +
    '?action=insp_view&token=' + encodeURIComponent(token) + '&idx=' + encodeURIComponent(idx);
  var apprBase = CONFIG.WEBAPP_URL +
    '?action=approve&docKind=INSP&token=' + encodeURIComponent(token) +
    '&docNo=' + encodeURIComponent(data.docNo || '') + '&idx=' + encodeURIComponent(idx);

  var viewUrl    = viewBase;
  var approveUrl = apprBase + '&decision=approve';
  var rejectUrl  = apprBase + '&decision=reject';

  var subject = '[검수보고서 결재요청] ' + (data.docNo || '-') + ' — ' + (data.subject || '검수보고서');

  var htmlBody =
    '<div style="font-family:\'Malgun Gothic\',Arial,sans-serif;max-width:600px;margin:0 auto;background:#ffffff;">'
    + '<div style="background:#0e7d72;padding:24px 32px;">'
    + '<div style="color:#ffffff;font-size:18px;font-weight:700;letter-spacing:2px;">구매품의 시스템</div>'
    + '</div>'
    + '<div style="padding:28px 32px;">'
    + '<p style="font-size:15px;color:#111111;margin:0 0 8px;"><b>' + approverName + ' ' + approverLabel + '</b> 님,</p>'
    + '<p style="font-size:14px;color:#444444;margin:0 0 24px;line-height:1.7;">'
    + '아래 검수보고서에 대한 결재를 요청합니다.<br>내용을 확인하신 후 승인 또는 반려해 주세요.</p>'
    + '<div style="background:#f8f9fa;border:1px solid #e0e0e0;border-radius:8px;padding:16px 20px;margin-bottom:24px;">'
    + '<table role="presentation" style="width:100%;border-collapse:collapse;font-size:13px;">'
    + '<tr><td style="color:#888;padding:4px 0;width:80px;">문서번호</td><td style="color:#111;font-weight:600;">' + docNo + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">품명</td><td style="color:#111;font-weight:600;">' + docTitle + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">업체명</td><td style="color:#111;">' + vendor + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 0;">기안자</td><td style="color:#111;">' + drafter + '</td></tr>'
    + '</table></div>'
    + '<table role="presentation" align="center" style="margin:0 auto 20px auto;border-collapse:collapse;">'
    + '<tr><td align="center" bgcolor="#0e7d72" style="border-radius:6px;">'
    + '<a href="' + viewUrl + '" style="display:inline-block;background:#0e7d72;color:#fff;font-size:14px;font-weight:700;padding:12px 36px;border-radius:6px;text-decoration:none;letter-spacing:1px;">검수보고서 확인</a>'
    + '</td></tr></table>'
    + '<div style="border-top:1px solid #e0e0e0;margin:20px 0;"></div>'
    + '<p style="font-size:12px;color:#888;text-align:center;margin:0 0 14px;">검수보고서 확인 후 아래에서 결재해 주세요.</p>'
    + '<table role="presentation" align="center" style="margin:0 auto;border-collapse:collapse;"><tr>'
    + '<td align="center" style="padding-right:10px;"><a href="' + rejectUrl + '" style="display:inline-block;background:#d64541;color:#fff;font-size:14px;font-weight:700;padding:11px 32px;border-radius:6px;text-decoration:none;border:2px solid #d64541;">반 려</a></td>'
    + '<td align="center" style="padding-left:10px;"><a href="' + approveUrl + '" style="display:inline-block;background:#27ae60;color:#fff;font-size:14px;font-weight:700;padding:11px 32px;border-radius:6px;text-decoration:none;border:2px solid #27ae60;">승 인</a></td>'
    + '</tr></table>'
    + '</div>'
    + '<div style="background:#f0f0f0;padding:16px 32px;text-align:center;">'
    + '<p style="font-size:11px;color:#888;margin:0;">— ' + escapeHtml(CONFIG.FROM_NAME || '') + ' · 이 메일은 자동 발송되었습니다.</p>'
    + '</div></div>';

  var plainBody = [
    approver.name + ' ' + (approver.label || '') + ' 님,', '',
    '아래 검수보고서에 대한 결재를 요청합니다.', '',
    '■ 문서번호: ' + (data.docNo || '-'),
    '■ 품명:     ' + (data.subject || '-'),
    '■ 업체명:   ' + (data.vendorName || '-'),
    '■ 기안자:   ' + (data.drafter || '-'), '',
    '▶ 검수보고서 확인: ' + viewUrl,
    '▶ 승인하기: ' + approveUrl,
    '▶ 반려하기: ' + rejectUrl, '',
    '— ' + (CONFIG.FROM_NAME || ''),
  ].join('\n');

  return sendEmailWithRetry(approver.email, subject, plainBody, htmlBody);
}

// ================================================================
// [INSP Step 3] 테스트
// ================================================================

/**
 * Step 3 드라이런 — 실제 제출 없이 폼 데이터 조회만 점검
 * 사용법: 운영의 최종승인(PRC) 토큰을 인자로 넣어 실행
 *   testInspStep3('PRC토큰')  → 폼 자동채움 값/회차/권한 결과 로그
 */
function testInspStep3(prcToken) {
  if (!prcToken) {
    console.log('사용법: testInspStep3("최종승인PRC의_token")');
    console.log('힌트: debugInspStep2() 로그의 PARENT_DOC_ID가 아니라 PRC 행의 token이 필요합니다.');
    return;
  }
  var res = getInspFormDataForClient(prcToken, null);
  console.log(JSON.stringify(res, null, 2));
}


// ================================================================
// [INSP Step 4] 검수보고서 결재 승인/반려 처리
// ================================================================

/**
 * 검수보고서 결재 화면용 데이터 (Approval/Viewer 공용 경량 메타)
 * - 토큰으로 INSP 행을 찾아 결재 진행 정보 반환
 * - 현재 결재 단계(curIdx)와 본인 결재 여부를 서버가 판정 (idx 파라미터 신뢰 금지)
 */
function getInspDecisionMetaForClient(token) {
  try {
    var actor = (getActiveUserEmail() || '').toLowerCase();
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var info = _readInspRowByToken(ss, token);
    if (!info) return { ok: false, message: '검수보고서를 찾을 수 없습니다.' };
    var r = info.row;
    var status = String(r[INSP_COL.STATUS] || '');
    var curIdx = parseInt(r[INSP_COL.APPR_IDX]) || 0;
    var apprCount = parseInt(r[INSP_COL.APPR_COUNT]) || 0;

    // 본인이 현재 결재 단계의 결재자인지 자동 판정
    var myIdx = _findMyInspIdx(r, apprCount, curIdx, actor);
    // 현재 결재 차례의 라벨 (기안자/결재자 N) — Approval 경량화면 결재순서 표시용
    var curLabel = (curIdx >= 0 && curIdx < apprCount)
      ? String(r[inspApprCol(curIdx, 0)] || ('결재자 ' + curIdx))
      : '';
    var myLabel = (myIdx >= 0 && myIdx < apprCount)
      ? String(r[inspApprCol(myIdx, 0)] || (myIdx === 0 ? '기안자' : ('결재자 ' + myIdx)))
      : '';

    return {
      ok: true,
      docNo:       String(r[INSP_COL.DOC_NO] || ''),
      curLabel:    curLabel,
      myLabel:     myLabel,
      subject:     String(r[INSP_COL.SUBJECT] || ''),
      vendorName:  String(r[INSP_COL.VENDOR_NAME] || ''),
      drafterName: String(r[INSP_COL.DRAFTER_NAME] || ''),
      receivedDate:toDateStr(r[INSP_COL.RECEIVED_DATE]),
      verdict:     String(r[INSP_COL.VERDICT] || ''),
      status:      status,
      curIdx:      curIdx,
      apprCount:   apprCount,
      myIdx:       myIdx,
      isMyTurn:    (myIdx === curIdx && status !== '반려' && status !== '최종승인(INSP)'),
      alreadyDone: (status === '최종승인(INSP)' || status === '반려'),
    };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

/**
 * 검수보고서 결재 처리 (승인/반려) — _decisionCore의 INSP 버전
 * payload: { token, decision('approve'|'reject'), comment, idx? }
 * - idx는 신뢰하지 않고 서버가 actor로 본인 단계 재판정
 * - 최종 승인 시: 상태 '최종승인(INSP)'. IS_FINAL='Y'면 품의 종결(추가 제출 차단은 제출단에서 검증)
 *   ※ PDF 변환·FINAL 이동·사진 삭제는 Step 6에서 insp_pdf 큐로 처리
 */
function processInspDecisionFromClient(payload) {
  var lockResult = withLock(function () {
    try {
      var actor    = (getActiveUserEmail() || '').toLowerCase();
      var token    = payload.token;
      var decision = payload.decision;
      var comment  = payload.comment || '';

      var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet = ss.getSheetByName(CONFIG.INSP_SHEET_NAME);
      if (!sheet) return { ok: false, message: '검수보고서 시트가 없습니다.' };

      var info = _readInspRowByToken(ss, token);
      if (!info) return { ok: false, message: '유효하지 않은 토큰입니다.' };
      var rowNum = info.rowNum, r = info.row;

      var status = String(r[INSP_COL.STATUS] || '');
      if (status === '반려')          return { ok: false, message: '이미 반려된 검수보고서입니다.' };
      if (status === '최종승인(INSP)') return { ok: false, message: '이미 최종 승인된 검수보고서입니다.' };

      var apprCount = parseInt(r[INSP_COL.APPR_COUNT]) || 0;
      var curIdx    = parseInt(r[INSP_COL.APPR_IDX]) || 0;

      // 본인 단계 자동 판정 (idx 파라미터 의존 금지 — 기존 학습)
      // 동일인 다단계 대응: 현재 차례(curIdx)가 본인이면 그 단계를 처리 단계로 삼음
      var myIdx = _findMyInspIdx(r, apprCount, curIdx, actor);
      if (myIdx < 0)        return { ok: false, message: '결재 권한이 없습니다. (지정된 결재자가 아닙니다)' };
      if (myIdx !== curIdx) return { ok: false, message: '결재 순서가 아닙니다. 앞 단계 결재 완료 후 처리됩니다.' };

      var now = new Date();
      var docNo       = String(r[INSP_COL.DOC_NO] || '');
      var subject     = String(r[INSP_COL.SUBJECT] || '');
      var drafterEmail= String(r[INSP_COL.DRAFTER] || '');
      var drafterName = String(r[INSP_COL.DRAFTER_NAME] || '');
      var isFinal     = String(r[INSP_COL.IS_FINAL] || '') === 'Y';
      var apprName    = String(r[inspApprCol(curIdx, 1)] || '-');

      // 현재 결재자 블록 status/processedAt/comment 기록
      sheet.getRange(rowNum, inspApprCol(curIdx, 3) + 1, 1, 3).setValues([[
        decision === 'approve' ? '승인' : '반려', now, comment,
      ]]);

      if (decision === 'reject') {
        sheet.getRange(rowNum, INSP_COL.STATUS + 1).setValue('반려');
        var existLog = String(r[INSP_COL.REJECT_LOG] || '');
        var resubCount = parseInt(r[INSP_COL.RESUB_COUNT]) || 0;
        var entry = '[' + (resubCount + 1) + '차 반려] ' + apprName + ' / ' + toDateTimeStr(now) + ' / ' + (comment || '사유 없음');
        sheet.getRange(rowNum, INSP_COL.REJECT_LOG + 1).setValue(existLog ? existLog + '\n' + entry : entry);
        return { ok: true, decision: 'reject', message: '반려 처리되었습니다.',
          notify: { type: 'reject', toEmail: drafterEmail, drafter: drafterName,
                    docNo: docNo, subject: subject, approverName: apprName, comment: comment } };
      }

      // 승인
      var nextIdx = curIdx + 1;
      sheet.getRange(rowNum, INSP_COL.APPR_IDX + 1).setValue(nextIdx);

      if (nextIdx < apprCount) {
        var nextName  = String(r[inspApprCol(nextIdx, 1)] || '');
        var nextEmail = String(r[inspApprCol(nextIdx, 2)] || '');
        var nextLabel = String(r[inspApprCol(nextIdx, 0)] || ('결재자 ' + (nextIdx + 1)));
        sheet.getRange(rowNum, INSP_COL.STATUS + 1).setValue('결재중 (' + nextName + ')');
        return { ok: true, decision: 'approve', message: '승인 처리되었습니다.',
          notify: { type: 'next', approver: { label: nextLabel, name: nextName, email: nextEmail },
                    docNo: docNo, subject: subject, drafter: drafterName, token: token, idx: nextIdx, rowNum: rowNum } };
      }

      // 최종 승인 → 종결 처리 (PDF/이동/삭제는 Step 6)
      sheet.getRange(rowNum, INSP_COL.STATUS + 1).setValue('최종승인(INSP)');
      return { ok: true, decision: 'approve', message: '최종 승인 처리되었습니다.', finalized: true,
        notify: { type: 'completion', toEmail: drafterEmail, drafter: drafterName,
                  docNo: docNo, subject: subject, token: token, isFinal: isFinal } };
    } catch (err) {
      return { ok: false, message: err.toString() };
    }
  });

  if (!lockResult.ok || !lockResult.notify) return lockResult;

  // 락 밖: 메일 + 감사 로그
  var n = lockResult.notify;
  try {
    if (n.type === 'reject' && n.toEmail) {
      _sendInspRejectEmail(n.toEmail, n.drafter, n.docNo, n.subject, n.approverName, n.comment, n.token);
    } else if (n.type === 'next' && n.approver.email) {
      _sendInspApprovalEmail(n.approver, { docNo: n.docNo, subject: n.subject, drafter: n.drafter, vendorName: '' }, n.token, n.rowNum, n.idx);
    } else if (n.type === 'completion' && n.toEmail) {
      _sendInspCompletionEmail(n.toEmail, n.drafter, n.docNo, n.subject, n.token, n.isFinal);
    }
  } catch (e) {
    try { notifyAdminError('[INSP] 결재 후 메일 실패: ' + n.docNo + ' / ' + e.toString()); } catch (_) {}
  }

  try {
    var ev = (lockResult.decision === 'reject') ? AUDIT_EVENT.INSP_REJECT
           : (lockResult.finalized ? AUDIT_EVENT.INSP_FINALIZED : AUDIT_EVENT.INSP_APPROVE);
    writeAuditLog({ eventType: ev, docNo: n.docNo, docToken: n.token, docType: 'INSP',
      reason: lockResult.message + (n.isFinal && lockResult.finalized ? ' (최종 검수 — 품의 종결)' : '') });
  } catch (_) {}

  // [INSP] 최종 승인 완료 → PDF 생성을 로컬 파이썬(코워크)으로 이관.
  //  GAS는 구글 변환 서버(불안정)를 쓰지 않는다. 대신 STAGING 폴더에 manifest.json을
  //  기록하고 David에게 작업요청 메일만 보낸다. 사진은 삭제하지 않고 보존한다.
  if (lockResult.finalized) {
    try {
      _prepareInspPdfHandoff(n.token);
    } catch (e) {
      try { notifyAdminError('[INSP] PDF 핸드오프(매니페스트/메일) 실패: ' + n.docNo + ' / ' + e.toString()); } catch (_) {}
    }
  }
  return { ok: true, message: lockResult.message, finalized: !!lockResult.finalized };
}

/** INSP 반려 알림 메일 */
function _sendInspRejectEmail(toEmail, drafter, docNo, subject, approverName, comment, token) {
  var url = CONFIG.WEBAPP_URL + '?action=insp_form&resub=' + encodeURIComponent(token);
  var subj = '[검수보고서 반려] ' + docNo + ' - ' + subject;
  var html = '<div style="font-family:sans-serif;max-width:600px;">'
    + '<h2 style="color:#d64541;">검수보고서가 반려되었습니다</h2>'
    + '<p><b>' + escapeHtml(drafter || '') + '</b> 님, 제출하신 검수보고서가 반려되었습니다.</p>'
    + '<table style="border-collapse:collapse;font-size:14px;">'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">문서번호</td><td>' + escapeHtml(docNo) + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">반려자</td><td>' + escapeHtml(approverName || '') + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">사유</td><td>' + escapeHtml(comment || '사유 없음') + '</td></tr>'
    + '</table>'
    + '<p style="margin-top:16px;"><a href="' + url + '" style="background:#0e7d72;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;">수정 후 재상신</a></p></div>';
  var plain = '검수보고서 반려\n문서번호: ' + docNo + '\n반려자: ' + approverName + '\n사유: ' + (comment || '사유 없음') + '\n재상신: ' + url;
  sendEmailWithRetry(toEmail, subj, plain, html);
}

/** INSP 완료 알림 메일 */
function _sendInspCompletionEmail(toEmail, drafter, docNo, subject, token, isFinal) {
  var subj = '[검수보고서 최종승인] ' + docNo + ' - ' + subject;
  var html = '<div style="font-family:sans-serif;max-width:600px;">'
    + '<h2 style="color:#0e7d72;">검수보고서가 최종 승인되었습니다</h2>'
    + '<p><b>' + escapeHtml(drafter || '') + '</b> 님, 검수보고서 결재가 모두 완료되었습니다.</p>'
    + '<table style="border-collapse:collapse;font-size:14px;">'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">문서번호</td><td>' + escapeHtml(docNo) + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">품명</td><td>' + escapeHtml(subject) + '</td></tr>'
    + (isFinal ? '<tr><td style="color:#888;padding:4px 12px 4px 0;">비고</td><td><b style="color:#0e7d72;">최종 검수 — 본 품의의 검수가 종결되었습니다.</b></td></tr>' : '')
    + '</table>'
    + '<p style="font-size:12px;color:#888;margin-top:14px;">검수보고서 PDF가 곧 PO 폴더에 자동 보관됩니다.</p></div>';
  var plain = '검수보고서 최종 승인\n문서번호: ' + docNo + '\n품명: ' + subject + (isFinal ? '\n(최종 검수 — 품의 종결)' : '');
  sendEmailWithRetry(toEmail, subj, plain, html);
}

/**
 * Step 4 테스트 — 결재 메타 조회 점검
 *   testInspStep4('INSP토큰')
 */
function testInspStep4(token) {
  if (!token) { console.log('사용법: testInspStep4("검수보고서_token")'); return; }
  console.log(JSON.stringify(getInspDecisionMetaForClient(token), null, 2));
}


// ================================================================
// [INSP Step 5] 홈 메뉴용 — 내 검수 결재 대기 / 최종 완료(INSP)
// ================================================================

/**
 * 검수보고서 목록을 1회 read하여 세 가지 목록 생성
 *  - myInspPending:  내가 현재 결재할 차례인 검수보고서 (status가 진행중 + 내 단계 + 대기)
 *  - inspFinals:     최종 선택(IS_FINAL=Y)되어 최종승인(INSP)된 검수보고서
 *  - allInspPending: [결재 메뉴 세분화] 관리자 전용 — 결재자가 누구든 모든 대기 검수보고서
 * 가시성: isAdmin(관리자)·isProc(구매팀)면 inspFinals 전체, 일반은 본인 관련(기안자 또는 결재 참여자)
 * @param {Spreadsheet} ss 열린 스프레드시트
 * @param {string} actor 로그인 이메일(소문자)
 * @param {boolean} isProc 구매팀 여부
 * @param {boolean} isAdmin 관리자 여부 (allInspPending 수집 + 전사 조회 게이트)
 * @returns {Object} { myInspPending: [], inspFinals: [], allInspPending: [] }
 */
function _getInspMenusForClient(ss, actor, isProc, isAdmin, isGView) {
  var out = { myInspPending: [], inspFinals: [], allInspPending: [] };
  try {
    var sheet = ss.getSheetByName(CONFIG.INSP_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return out;
    actor = String(actor || '').toLowerCase();

    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      var token = String(r[INSP_COL.TOKEN] || '');
      if (!token) continue;
      var status    = String(r[INSP_COL.STATUS] || '');
      // ⚠ apprCount는 반드시 MAX_APPROVERS로 클램프할 것.
      //   inspApprCol은 범위 초과 시 throw하고 이 함수 전체가 try/catch로 감싸여 있어,
      //   시트에 apprCount가 깨진 행이 하나만 있어도 아래 amParticipant 루프나
      //   _findMyInspIdx에서 튕겨 모든 사용자의 검수 메뉴가 통째로 비어버린다.
      var apprCount = Math.min(parseInt(r[INSP_COL.APPR_COUNT]) || 0, INSP_COL.MAX_APPROVERS);
      var curIdx    = parseInt(r[INSP_COL.APPR_IDX]) || 0;
      var isFinal   = String(r[INSP_COL.IS_FINAL] || '') === 'Y';

      var meta = {
        docNo:       String(r[INSP_COL.DOC_NO] || ''),
        token:       token,
        seq:         parseInt(r[INSP_COL.SEQ]) || 0,
        poNo:        String(r[INSP_COL.PO_NO] || ''),
        reqNo:       String(r[INSP_COL.REQ_NO] || ''),
        subject:     String(r[INSP_COL.SUBJECT] || ''),
        vendorName:  String(r[INSP_COL.VENDOR_NAME] || ''),
        drafter:     String(r[INSP_COL.DRAFTER] || ''),
        drafterName: String(r[INSP_COL.DRAFTER_NAME] || ''),
        verdict:     String(r[INSP_COL.VERDICT] || ''),
        isFinal:     isFinal,
        status:      status,
        receivedDate:toDateStr(r[INSP_COL.RECEIVED_DATE]),
        submitAt:    toDateTimeStr(r[INSP_COL.SUBMIT_AT]),
        docType:     'INSP',
      };

      // 내가 결재 참여자인지 + 내 단계인지 판정 (curIdx 우선 — 동일인 다단계 대응)
      var myIdx = _findMyInspIdx(r, apprCount, curIdx, actor);
      // 가시성용: 본인이 어느 단계든 배정돼 있는지 (최종완료 노출 판정)
      var amParticipant = false;
      for (var pa = 0; pa < apprCount; pa++) {
        if (String(r[inspApprCol(pa, 2)] || '').toLowerCase() === actor) { amParticipant = true; break; }
      }

      // 결재 대기 판정: 진행중 + 현재 차례 블록 status가 '대기'
      //  - myInspPending  : 현재 차례가 '나' (동일인 다단계 대응은 _findMyInspIdx가 처리)
      //  - allInspPending : 관리자용 — 현재 결재자가 누구든 수집 (반드시 curIdx 기준.
      //                     myIdx는 '접속자의' 단계라 남의 문서에는 의미가 없다)
      // curValid: apprCount는 위에서 이미 클램프됐으므로 curIdx 범위만 확인하면 된다.
      var inProgress = (status === '검토중' || status.indexOf('결재중') >= 0);
      var curValid   = (curIdx >= 0 && curIdx < apprCount);
      if (inProgress && curValid && String(r[inspApprCol(curIdx, 3)] || '') === '대기') {
        var pendMeta = Object.assign({}, meta, {
          stageIdx:     curIdx,
          stageKind:    (curIdx === 0 ? 'drafter' : 'approver'),
          curApprName:  String(r[inspApprCol(curIdx, 1)] || ''),
          curApprEmail: String(r[inspApprCol(curIdx, 2)] || ''),
          apprCount:    apprCount,
        });
        if (myIdx >= 0 && myIdx === curIdx) out.myInspPending.push(pendMeta);
        if (isAdmin) out.allInspPending.push(pendMeta);
      }

      // 최종 완료(INSP): IS_FINAL=Y && 최종승인(INSP)
      // 가시성: 관리자 / 구매팀 / 결재 참여자 / 본인 기안 / 같은 부서(부서 공유)
      if (isFinal && status === '최종승인(INSP)') {
        if (isAdmin || isProc || isGView || amParticipant ||
            meta.drafter.toLowerCase() === actor ||
            _isSameDeptEmail(ss, meta.drafter, actor)) {
          out.inspFinals.push(meta);
        }
      }
    }

    out.myInspPending.sort(function (a, b) { return (a.submitAt || '').localeCompare(b.submitAt || ''); });
    out.allInspPending.sort(function (a, b) { return (a.submitAt || '').localeCompare(b.submitAt || ''); });
    out.inspFinals.sort(function (a, b) { return (b.submitAt || '').localeCompare(a.submitAt || ''); });
  } catch (e) {
    console.error('[INSP] _getInspMenusForClient 실패: ' + e.toString());
  }
  return out;
}

// ================================================================
// [A-3] 검수보고서(INSP) 결재자 변경 — Code.gs 디스패처가 위임
//   REQ/PRC 코어와 동일 정책: 권한검증 · canonical 결재자 · 감사우선+원복 · 메일결과 로그
//   (assertAdminWithLog / findActiveApproverByEmail / _changeApproverSendAndLog /
//    _sendApproverChangedNotice / writeAuditLog 등은 Code.gs의 전역 함수 재사용)
// ================================================================

/** 변경 가능 INSP 목록 (Code.gs _adminListChangeableCore가 호출; 단독 호출 방어 위해 자체 권한검증) */
function inspListChangeableForAdmin(payload) {
  try {
    assertAdminWithLog('insp_list_changeable', payload || {});
    payload = payload || {};
    var drafterLike = String(payload.drafter || '').trim().toLowerCase();
    var docNoLike   = String(payload.docNoLike || '').trim().toLowerCase();
    var statusLike  = String(payload.status || '').trim();

    var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.INSP_SHEET_NAME);
    var docs = [];
    if (sheet && sheet.getLastRow() >= 2) {
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        var token = String(r[INSP_COL.TOKEN] || '');
        if (!token) continue;
        var status = String(r[INSP_COL.STATUS] || '');
        // 진행 중만: 검토중 / 결재중 (반려·최종승인 제외)
        if (!(status === '검토중' || status.indexOf('결재중') >= 0)) continue;

        var docNo        = String(r[INSP_COL.DOC_NO] || '');
        var drafter      = String(r[INSP_COL.DRAFTER_NAME] || '');
        var drafterEmail = String(r[INSP_COL.DRAFTER] || '');
        if (docNoLike && docNo.toLowerCase().indexOf(docNoLike) < 0) continue;
        if (drafterLike && drafter.toLowerCase().indexOf(drafterLike) < 0 &&
            drafterEmail.toLowerCase().indexOf(drafterLike) < 0) continue;
        if (statusLike && status.indexOf(statusLike) < 0) continue;

        var apprCount = Math.min(parseInt(r[INSP_COL.APPR_COUNT]) || 0, INSP_COL.MAX_APPROVERS);
        var curIdx    = parseInt(r[INSP_COL.APPR_IDX]) || 0;
        var chain = [];
        for (var s = 0; s < apprCount; s++) {
          var st = String(r[inspApprCol(s, 3)] || '대기');
          chain.push({
            stageIdx:   s,
            label:      String(r[inspApprCol(s, 0)] || ''),
            name:       String(r[inspApprCol(s, 1)] || ''),
            email:      String(r[inspApprCol(s, 2)] || ''),
            status:     st,
            changeable: (s >= 1 && st === '대기'),
            isCurrent:  (s === curIdx),
          });
        }
        docs.push({
          token:    token,
          docNo:    docNo,
          docType:  'INSP',
          drafter:  drafter,
          vendor:   String(r[INSP_COL.VENDOR_NAME] || ''),
          status:   status,
          curIdx:   curIdx,
          submitAt: toDateTimeStr(r[INSP_COL.SUBMIT_AT]),
          chain:    chain,
        });
      }
    }
    return { ok: true, docs: docs };
  } catch (err) {
    return { ok: false, message: err.toString(), docs: [] };
  }
}

/** INSP 결재자 변경 실행 (Code.gs adminChangeApproverForClient가 docType==='INSP'일 때 위임) */
function adminChangeInspApprover(payload) {
  payload = payload || {};
  var opId = Utilities.getUuid();
  var token, stageIdx, reason, canonical;

  try {
    assertAdminWithLog(AUDIT_EVENT.ADMIN_CHANGE_APPROVER, { token: payload.token, stageIdx: payload.stageIdx, opId: opId, docType: 'INSP' });
    requireAdminReason(payload.reason);

    token    = String(payload.token || '').trim();
    stageIdx = parseInt(payload.stageIdx, 10);
    reason   = String(payload.reason || '').trim();
    if (!token) return { ok: false, message: '문서 토큰이 필요합니다.' };
    if (isNaN(stageIdx) || stageIdx < 1) return { ok: false, message: '변경할 수 없는 단계입니다. (기안자 슬롯은 변경 불가)' };
    if (stageIdx >= INSP_COL.MAX_APPROVERS) return { ok: false, message: '존재하지 않는 결재 단계입니다.' };

    canonical = findActiveApproverByEmail(payload.newEmail);
    if (!canonical) return { ok: false, message: '활성 결재자 명단에 없는 사용자입니다.' };
  } catch (e) {
    return { ok: false, message: e.message };
  }

  var newName  = canonical.name;
  var newEmail = canonical.email;

  var lockResult = withLock(function() {
    try {
      var actor = getActiveUserEmail();
      if (!isAdminUser(actor)) return { ok: false, message: '관리자 권한이 필요합니다.' };

      var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var sheet = ss.getSheetByName(CONFIG.INSP_SHEET_NAME);
      if (!sheet) return { ok: false, message: '검수보고서 시트가 없습니다.' };

      var info = _readInspRowByToken(ss, token);
      if (!info) return { ok: false, message: '대상 검수보고서를 찾을 수 없습니다.' };
      var rowNum = info.rowNum, r = info.row;

      var status    = String(r[INSP_COL.STATUS] || '');
      var docNo     = String(r[INSP_COL.DOC_NO] || '');
      var apprCount = Math.min(parseInt(r[INSP_COL.APPR_COUNT]) || 0, INSP_COL.MAX_APPROVERS);
      var curIdx    = parseInt(r[INSP_COL.APPR_IDX]) || 0;

      if (!(status === '검토중' || status.indexOf('결재중') >= 0))
        return { ok: false, message: '진행 중 문서만 결재자를 변경할 수 있습니다. (현재 상태: ' + status + ')' };
      if (stageIdx >= apprCount) return { ok: false, message: '존재하지 않는 결재 단계입니다.' };

      var stageStatus = String(r[inspApprCol(stageIdx, 3)] || '');
      if (stageStatus !== '대기') return { ok: false, message: '이미 처리된 단계는 변경할 수 없습니다. (단계 상태: ' + stageStatus + ')' };

      var oldName    = String(r[inspApprCol(stageIdx, 1)] || '');
      var oldEmail   = String(r[inspApprCol(stageIdx, 2)] || '');
      var stageLabel = String(r[inspApprCol(stageIdx, 0)] || ('결재자 ' + stageIdx));

      if (oldEmail.toLowerCase() === newEmail.toLowerCase())
        return { ok: false, message: '현재 결재자와 동일한 사람으로는 변경할 수 없습니다.' };

      var isCurrentStage = (stageIdx === curIdx);

      sheet.getRange(rowNum, inspApprCol(stageIdx, 1) + 1, 1, 2).setValues([[newName, newEmail]]);
      if (isCurrentStage && status.indexOf('결재중') >= 0)
        sheet.getRange(rowNum, INSP_COL.STATUS + 1).setValue('결재중 (' + newName + ')');

      var logged = writeAuditLog({
        eventType: AUDIT_EVENT.ADMIN_CHANGE_APPROVER,
        actor: actor, actorRole: 'admin',
        docNo: docNo, docToken: token, docType: 'INSP',
        targetUser: newEmail, reason: reason,
        payload: {
          opId: opId, stageIdx: stageIdx, stageLabel: stageLabel,
          oldName: oldName, oldEmail: oldEmail, newName: newName, newEmail: newEmail,
          wasCurrentStage: isCurrentStage,
          newApproverNotify: isCurrentStage ? 'pending' : 'not_required',
          oldApproverNotify: 'pending',
        },
      });

      if (logged !== true) {
        try {
          sheet.getRange(rowNum, inspApprCol(stageIdx, 1) + 1, 1, 2).setValues([[oldName, oldEmail]]);
          if (isCurrentStage && status.indexOf('결재중') >= 0)
            sheet.getRange(rowNum, INSP_COL.STATUS + 1).setValue(status);
          return { ok: false, message: '감사로그 기록 실패 — 변경을 취소했습니다.' };
        } catch (revErr) {
          return { ok: false, _revertFailed: true, message: '감사로그 기록 실패 + 원복 실패. 관리자에게 통지합니다.',
            _ctx: { opId: opId, docNo: docNo, token: token, stageIdx: stageIdx, oldEmail: oldEmail, newEmail: newEmail } };
        }
      }

      return {
        ok: true, message: '결재자가 변경되었습니다.',
        _ctx: {
          opId: opId, docNo: docNo, token: token, rowNum: rowNum,
          stageIdx: stageIdx, stageLabel: stageLabel, isCurrentStage: isCurrentStage,
          oldName: oldName, oldEmail: oldEmail, newName: newName, newEmail: newEmail, reason: reason,
          subject: String(r[INSP_COL.SUBJECT] || ''), drafter: String(r[INSP_COL.DRAFTER_NAME] || ''),
          vendorName: String(r[INSP_COL.VENDOR_NAME] || ''),
        },
      };
    } catch (err) {
      return { ok: false, message: err.toString() };
    }
  });

  if (!lockResult) return { ok: false, message: '시스템이 바쁩니다. 잠시 후 다시 시도해주세요.' };
  if (lockResult._revertFailed) {
    try { notifyAdminError('[A-3/INSP] 결재자 변경 감사로그 실패+원복 실패 opId=' + opId + ' / ' + JSON.stringify(lockResult._ctx)); } catch (_) {}
    return { ok: false, message: lockResult.message };
  }
  if (!lockResult.ok) return lockResult;

  var ctx = lockResult._ctx;
  var logCtx = { opId: ctx.opId, docNo: ctx.docNo, token: ctx.token, docType: 'INSP' };

  // 1) 새 결재자 결재요청 (현재 차례)
  if (ctx.isCurrentStage) {
    _changeApproverSendAndLog('approve_request', ctx.newEmail, logCtx, function() {
      return _sendInspApprovalEmail(
        { label: ctx.stageLabel, name: ctx.newName, email: ctx.newEmail },
        { docNo: ctx.docNo, subject: ctx.subject, drafter: ctx.drafter, vendorName: ctx.vendorName },
        ctx.token, ctx.rowNum, ctx.stageIdx);
    });
  }

  // 2) 제거자 통지 (항상)
  _changeApproverSendAndLog('removed_notice', ctx.oldEmail, logCtx, function() {
    return _sendApproverChangedNotice(ctx.oldEmail, {
      docNo: ctx.docNo, docType: 'INSP', oldName: ctx.oldName, newName: ctx.newName,
      stageLabel: ctx.stageLabel, reason: ctx.reason, opId: ctx.opId,
    });
  });

  return { ok: true, message: '결재자가 변경되었습니다.',
    changed: { stageIdx: ctx.stageIdx, oldName: ctx.oldName, newName: ctx.newName, isCurrentStage: ctx.isCurrentStage } };
}

// ================================================================
// [INSP Step 5] 전용 뷰어 데이터
// ================================================================

/**
 * 검수보고서 뷰어/결재 화면용 전체 데이터
 *  - 본문(스냅샷·판정·검수의견) + 결재 진행 + 사진 썸네일 URL + myApproverIdx(7번 해결책)
 *  - 사진은 결재 진행중일 때만 thumbnail URL 제공(STAGING에 존재)
 * @param {string} token INSP token
 * @param {string|number=} hintIdx URL idx 힌트(선택) — 서버 판정 우선
 */
function getInspViewerDataForClient(token, hintIdx) {
  try {
    var actor = (getActiveUserEmail() || '').toLowerCase();
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var info = _readInspRowByToken(ss, token);
    if (!info) return { ok: false, message: '검수보고서를 찾을 수 없습니다.' };
    var r = info.row;

    var apprCount = parseInt(r[INSP_COL.APPR_COUNT]) || 0;
    var curIdx    = parseInt(r[INSP_COL.APPR_IDX]) || 0;
    var status    = String(r[INSP_COL.STATUS] || '');

    // 결재자 블록
    var approvers = [];
    for (var a = 0; a < apprCount; a++) {
      approvers.push({
        label:       String(r[inspApprCol(a, 0)] || ''),
        name:        String(r[inspApprCol(a, 1)] || ''),
        email:       String(r[inspApprCol(a, 2)] || ''),
        status:      String(r[inspApprCol(a, 3)] || '대기'),
        processedAt: toDateTimeStr(r[inspApprCol(a, 4)]),
        comment:     String(r[inspApprCol(a, 5)] || ''),
      });
    }

    // myApproverIdx 판정 (7번 해결책 + 동일인 다단계 대응)
    // 우선순위: ①현재 차례가 본인이면 curIdx ②URL 힌트(검증) ③본인 첫 단계
    var myApproverIdx = -1;
    if (curIdx >= 0 && curIdx < apprCount && String(approvers[curIdx].email || '').toLowerCase() === actor) {
      myApproverIdx = curIdx;
    }
    if (myApproverIdx === -1) {
      var h = (hintIdx !== undefined && hintIdx !== '' && !isNaN(hintIdx)) ? parseInt(hintIdx) : -1;
      if (h >= 0 && h < apprCount && String(approvers[h].email || '').toLowerCase() === actor) {
        myApproverIdx = h;
      }
    }
    if (myApproverIdx === -1) {
      for (var k = 0; k < apprCount; k++) {
        if (String(approvers[k].email || '').toLowerCase() === actor) { myApproverIdx = k; break; }
      }
    }

    // 사진 썸네일 (진행중일 때만 — 최종승인 후엔 삭제 예정이라 PDF로 확인)
    var photos = [];
    try {
      var photoList = JSON.parse(String(r[INSP_COL.PHOTO_LIST] || '[]'));
      photoList.forEach(function (p) {
        if (p && p.id) {
          // base64 data URL: Drive 썸네일 URL은 웹앱 iframe에서 미인증으로 깨지므로
          // PDF 생성과 동일하게 STAGING Drive에서 바이트를 읽어 직접 임베드한다.
          var dataUrl = '';
          try {
            var blob = DriveApp.getFileById(p.id).getBlob();
            var mime = blob.getContentType() || 'image/jpeg';
            dataUrl = 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());
          } catch (ePhoto) { /* 개별 사진 실패는 썸네일 URL로 폴백 */ }
          photos.push({
            name: p.name || '',
            dataUrl:  dataUrl,
            thumbUrl: 'https://drive.google.com/thumbnail?id=' + p.id + '&sz=w1000',
            viewUrl:  'https://drive.google.com/file/d/' + p.id + '/view',
          });
        }
      });
    } catch (_) {}

    return {
      ok: true,
      doc: {
        docNo:        String(r[INSP_COL.DOC_NO] || ''),
        seq:          parseInt(r[INSP_COL.SEQ]) || 0,
        poNo:         String(r[INSP_COL.PO_NO] || ''),
        reqNo:        String(r[INSP_COL.REQ_NO] || ''),
        subject:      String(r[INSP_COL.SUBJECT] || ''),
        vendorName:   String(r[INSP_COL.VENDOR_NAME] || ''),
        drafterName:  String(r[INSP_COL.DRAFTER_NAME] || ''),
        dept:         String(r[INSP_COL.DEPT] || ''),
        issueDate:    toDateStr(r[INSP_COL.ISSUE_DATE]),
        receivedDate: toDateStr(r[INSP_COL.RECEIVED_DATE]),
        receivedNote: String(r[INSP_COL.RECEIVED_NOTE] || ''),
        verdict:      String(r[INSP_COL.VERDICT] || ''),
        isFinal:      String(r[INSP_COL.IS_FINAL] || '') === 'Y',
        comment:      String(r[INSP_COL.COMMENT] || ''),
        status:       status,
        rejectLog:    String(r[INSP_COL.REJECT_LOG] || ''),
      },
      approvers:     approvers,
      curIdx:        curIdx,
      myApproverIdx: myApproverIdx,
      photos:        photos,
    };
  } catch (err) {
    return { ok: false, message: err.toString() + ' / ' + (err.stack || '') };
  }
}


// ================================================================
// [INSP Step 6] PDF 생성 → FINAL/PO폴더 이동 → 사진 삭제
// ================================================================

// ================================================================
// [INSP] PDF 생성 로컬 이관 (매니페스트 핸드오프)
//  - 최종 승인 시 GAS는 PDF를 직접 만들지 않는다. STAGING 폴더에 manifest.json
//    (파이썬 렌더러 입력)을 기록하고 David에게 작업요청 메일을 보낸다.
//  - 사진은 삭제하지 않고 보존한다. 실제 PDF 생성은 로컬 파이썬(코워크)이 수행.
// ================================================================

/**
 * 결재 완료된 검수보고서의 PDF 생성 작업을 로컬 파이썬으로 핸드오프.
 *  1) PRC 기안일 기준 FINAL/{PO} 폴더 확보 → finalFolderId (유령 폴더 방지)
 *  2) STAGING 폴더에 manifest.json 기록 (PDF 필드 + 사진 id + finalFolderId)
 *  3) MOVE_STATUS='PENDING_PDF' (워크리스트 상태)
 *  4) David(관리자)에게 작업요청 메일
 * @param {string} token INSP token
 * @returns {Object} { ok, manifestFileId, finalFolderId }
 */
function _prepareInspPdfHandoff(token) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var info = _readInspRowByToken(ss, token);
  if (!info) return { ok: false, message: '검수보고서 행을 찾을 수 없음' };
  var r = info.row, rowNum = info.rowNum;

  var status = String(r[INSP_COL.STATUS] || '');
  if (status !== '최종승인(INSP)') {
    return { ok: false, message: '최종승인(INSP) 상태가 아님: ' + status };
  }

  var manifest = _buildInspManifest(ss, r);

  // STAGING 폴더에 manifest.json 기록 (기존 것 있으면 교체)
  var stagingId = String(r[INSP_COL.DRIVE_ID] || '');
  if (!stagingId) return { ok: false, message: 'STAGING 폴더 ID 없음' };
  var folder = DriveApp.getFolderById(stagingId);
  var existing = folder.getFilesByName('manifest.json');
  while (existing.hasNext()) { try { existing.next().setTrashed(true); } catch (_) {} }
  var mBlob = Utilities.newBlob(JSON.stringify(manifest, null, 2), 'application/json', 'manifest.json');
  var mFile = folder.createFile(mBlob);

  // 워크리스트 상태 마킹
  var sheet = ss.getSheetByName(CONFIG.INSP_SHEET_NAME);
  sheet.getRange(rowNum, INSP_COL.MOVE_STATUS + 1).setValue('PENDING_PDF');

  // David에게 작업요청 메일
  try {
    _sendInspPdfHandoffEmail(manifest, stagingId);
  } catch (e) {
    try { notifyAdminError('[INSP] 핸드오프 메일 실패: ' + manifest.docNo + ' / ' + e.toString()); } catch (_) {}
  }

  Logger.log('[INSP] 핸드오프 완료: ' + manifest.docNo + ' / manifest=' + mFile.getId());
  return { ok: true, manifestFileId: mFile.getId(), finalFolderId: manifest.finalFolderId };
}

/**
 * INSP 행 → 파이썬 렌더러 입력 매니페스트 객체.
 *  - 파이썬이 시트 45컬럼 스키마를 몰라도 되도록 필요한 값을 이름표로 전부 담는다.
 *  - 사진은 파일명이 아니라 Drive 파일 id로 참조 (위치 독립·자족).
 *  - finalFolderId는 REQ/PRC 파일이 들어 있는 FINAL PO 폴더 id.
 *    PRC 행의 DRIVE_ID(통합 완료 시 FINAL 폴더 id) 우선, 없으면 경로로 폴백.
 * @param {Spreadsheet} ss 열린 스프레드시트
 * @param {Array} r INSP 행 배열
 * @returns {Object} manifest
 */
function _buildInspManifest(ss, r) {
  var docNo = String(r[INSP_COL.DOC_NO] || '');
  var poNo  = String(r[INSP_COL.PO_NO] || '');
  var prcToken = String(r[INSP_COL.PRC_TOKEN] || '');

  // 업로드 대상 = REQ/PRC 파일이 실제로 들어 있는 FINAL PO 폴더.
  //  1순위: PRC 행에 저장된 DRIVE_ID를 그대로 사용 (consolidateToFinalByPrc가 FINAL 폴더 ID로
  //         갱신해 둔 값 — Code.gs). 이름·기안일 계산에 의존하지 않아 정확하다.
  //  2순위: 아직 통합 전이면 경로(PO번호 + PRC 기안일)로 확보.
  //  ⚠ 경로 방식은 이름이 안 맞으면 조용히 새 폴더를 만들어 버린다(유령 폴더). 그래서 1순위 우선.
  var prc = null, prcIssueDate = '';
  try {
    var docSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    prc = _findRowByTokenInDoc(docSheet, prcToken);
    if (prc) prcIssueDate = prc.row[COL.ISSUE_DATE];
  } catch (_) {}

  var finalFolderId = '';
  if (prc && String(prc.row[COL.MOVE_STATUS] || '') === 'FINAL') {
    finalFolderId = String(prc.row[COL.DRIVE_ID] || '');
  }
  if (!finalFolderId) {
    finalFolderId = getOrCreateFolder(
      poNo, prcIssueDate || toDateStr(r[INSP_COL.ISSUE_DATE]), 'final'
    ).getId();
  }

  var apprCount = parseInt(r[INSP_COL.APPR_COUNT]) || 0;
  var approvers = [];
  for (var a = 0; a < apprCount; a++) {
    approvers.push({
      label:       String(r[inspApprCol(a, 0)] || ''),
      name:        String(r[inspApprCol(a, 1)] || ''),
      status:      String(r[inspApprCol(a, 3)] || ''),
      processedAt: toDateTimeStr(r[inspApprCol(a, 4)]),
    });
  }

  var photos = [];
  try {
    JSON.parse(String(r[INSP_COL.PHOTO_LIST] || '[]')).forEach(function (p) {
      if (p && p.id) photos.push({ id: p.id, name: p.name || '' });
    });
  } catch (_) {}

  return {
    schemaVersion: 'insp-manifest-v1',
    token:        String(r[INSP_COL.TOKEN] || ''),
    docNo:        docNo,
    seq:          parseInt(r[INSP_COL.SEQ]) || 0,
    isFinal:      String(r[INSP_COL.IS_FINAL] || '') === 'Y',
    reqNo:        String(r[INSP_COL.REQ_NO] || ''),
    poNo:         poNo,
    subject:      String(r[INSP_COL.SUBJECT] || ''),
    vendorName:   String(r[INSP_COL.VENDOR_NAME] || ''),
    drafterName:  String(r[INSP_COL.DRAFTER_NAME] || ''),
    dept:         String(r[INSP_COL.DEPT] || ''),
    issueDate:    toDateStr(r[INSP_COL.ISSUE_DATE]),
    receivedDate: toDateStr(r[INSP_COL.RECEIVED_DATE]),
    receivedNote: String(r[INSP_COL.RECEIVED_NOTE] || ''),
    verdict:      String(r[INSP_COL.VERDICT] || ''),
    comment:      String(r[INSP_COL.COMMENT] || ''),
    approvers:    approvers,
    photos:       photos,
    finalFolderId: finalFolderId,
    generatedAt:  toDateTimeStr(new Date()),
  };
}

/** INSP PDF 작업요청 메일 (David/관리자) — 매니페스트 준비 완료 통지 */
function _sendInspPdfHandoffEmail(m, stagingFolderId) {
  var toList = CONFIG.ADMIN_NOTIFY_EMAILS || [];
  if (!toList.length) return;
  var stagingUrl = 'https://drive.google.com/drive/folders/' + stagingFolderId;
  var finalUrl   = 'https://drive.google.com/drive/folders/' + m.finalFolderId;
  var subj = '[검수보고서 PDF 생성 요청] ' + m.docNo + ' - ' + m.subject;
  var html = '<div style="font-family:sans-serif;max-width:600px;">'
    + '<h2 style="color:#0e7d72;">검수보고서 결재 완료 — PDF 생성 대기</h2>'
    + '<p>아래 검수보고서의 결재가 완료되었습니다. 로컬 렌더러로 PDF를 생성해 PO 폴더에 보관하세요.</p>'
    + '<table style="border-collapse:collapse;font-size:14px;">'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">문서번호</td><td>' + escapeHtml(m.docNo) + ' (' + m.seq + '회차' + (m.isFinal ? ' · 최종' : '') + ')</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">품명</td><td>' + escapeHtml(m.subject) + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">판정</td><td>' + escapeHtml(m.verdict) + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">사진</td><td>' + m.photos.length + '장</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">token</td><td>' + escapeHtml(m.token) + '</td></tr>'
    + '</table>'
    + '<p style="margin-top:12px;">'
    + '· 입력 폴더(manifest.json + 사진): <a href="' + stagingUrl + '">' + stagingUrl + '</a><br>'
    + '· 출력 폴더(PO): <a href="' + finalUrl + '">' + finalUrl + '</a></p>'
    + '<p style="font-size:12px;color:#888;">완료 후 GAS에서 <b>markInspPdfDone("' + escapeHtml(m.token) + '", "생성된PDF파일ID")</b> 실행해 마감하세요.</p>'
    + '</div>';
  var plain = '검수보고서 PDF 생성 요청\n문서번호: ' + m.docNo + '\n품명: ' + m.subject
    + '\n입력 폴더: ' + stagingUrl + '\n출력 폴더(PO): ' + finalUrl + '\ntoken: ' + m.token
    + '\n완료 후: markInspPdfDone("' + m.token + '", "PDF파일ID")';
  sendEmailWithRetry(toList.join(','), subj, plain, html);
}

/**
 * PDF 생성 대기 목록 (David/코워크 워크리스트).
 *  - status=최종승인(INSP) && MOVE_STATUS=PENDING_PDF 인 회차.
 * @returns {Array} [{docNo, token, subject, poNo, stagingFolderId, photoCount}]
 */
function listInspAwaitingPdf() {
  var out = [];
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.INSP_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return out;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[INSP_COL.STATUS] || '') !== '최종승인(INSP)') continue;
    if (String(r[INSP_COL.MOVE_STATUS] || '') !== 'PENDING_PDF') continue;
    var photoCount = 0;
    try { photoCount = JSON.parse(String(r[INSP_COL.PHOTO_LIST] || '[]')).length; } catch (_) {}
    out.push({
      docNo:           String(r[INSP_COL.DOC_NO] || ''),
      token:           String(r[INSP_COL.TOKEN] || ''),
      subject:         String(r[INSP_COL.SUBJECT] || ''),
      poNo:            String(r[INSP_COL.PO_NO] || ''),
      stagingFolderId: String(r[INSP_COL.DRIVE_ID] || ''),
      photoCount:      photoCount,
    });
  }
  Logger.log('[INSP] PDF 대기 ' + out.length + '건: ' + out.map(function(x){return x.docNo;}).join(', '));
  return out;
}

/**
 * PDF 생성·업로드 완료 마감. MOVE_STATUS=FINAL.
 * @param {string} token INSP token
 * @param {string=} pdfFileId 생성된 PDF 파일 id (감사로그 기록용, 선택)
 * @returns {Object} { ok, message }
 */
function markInspPdfDone(token, pdfFileId) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var info = _readInspRowByToken(ss, token);
  if (!info) return { ok: false, message: '행 없음: ' + token };
  var docNo = String(info.row[INSP_COL.DOC_NO] || '');
  var sheet = ss.getSheetByName(CONFIG.INSP_SHEET_NAME);
  sheet.getRange(info.rowNum, INSP_COL.MOVE_STATUS + 1).setValue('FINAL');
  try {
    writeAuditLog({ eventType: AUDIT_EVENT.INSP_FINALIZED, docNo: docNo, docToken: token,
      docType: 'INSP', reason: 'PDF 로컬 생성 완료' + (pdfFileId ? ' / fileId=' + pdfFileId : '') });
  } catch (_) {}
  Logger.log('[INSP] PDF 마감(FINAL): ' + docNo);
  return { ok: true, message: 'FINAL 마감 완료: ' + docNo };
}

/**
 * [임시] 2026-08-28 렌더분 3건 PDF 일괄 마감 — 편집기 Run 전용(무인자).
 *  GAS 편집기는 인자 있는 함수를 직접 실행할 수 없어, 이번 회차 값을 박아둔 래퍼.
 *  실행 후 대기목록(listInspAwaitingPdf)에서 3건이 모두 빠지면 이 함수는 삭제해도 됨.
 *  한 건이 실패해도 나머지는 계속 진행하고, 마지막에 건별 결과를 로그로 남긴다.
 */
function _tmp_markDone_20260828_batch1() {
  var items = [
    // [문서번호, token, 업로드된 PDF 파일 id]
    ['LC-2026-003-01',  'c029277f-24cf-4f8b-af7d-a1d2e64aada5', '1W_Vz65iIkkVdZjxt9iCQBgRtLW1fIMTG'],
    ['LC-2026-004-01',  'cdbf5491-b2fd-4bc7-a2e9-6fd3677332c6', '15Pi3wWvdglxw1avs8dW0qN2gFJoqMy_i'],
    ['TH-AP-26-050-01', 'd03d6b7f-93bb-41bc-a001-252edcf4126a', '1NKCl_quwps8bQQsDxxOOwEKQ0mgK0yzm']
  ];
  var out = [];
  items.forEach(function (it) {
    try {
      var r = markInspPdfDone(it[1], it[2]);
      out.push(it[0] + ': ' + (r && r.ok ? 'OK' : 'FAIL') + ' / ' + (r && r.message));
    } catch (e) {
      out.push(it[0] + ': ERROR / ' + e.toString());
    }
  });
  console.log(out.join('\n'));
  console.log('남은 대기: ' + JSON.stringify(listInspAwaitingPdf().map(function (x) { return x.docNo; })));
  return out;
}

/**
 * 실패/누락 회차의 PDF 핸드오프 재생성 (매니페스트 재작성 + 메일).
 *  - 큐 실패로 PDF가 안 만들어진 기존 회차 복구용. 사진이 STAGING에 보존돼 있어야 함.
 *  - 사용법: rerunInspHandoff("검수보고서_token")
 * @param {string} token INSP token
 */
function rerunInspHandoff(token) {
  if (!token) { console.log('사용법: rerunInspHandoff("검수보고서_token")'); return; }
  var res = _prepareInspPdfHandoff(token);
  console.log(JSON.stringify(res, null, 2));
  return res;
}

/**
 * insp_pdf 큐 job 처리 진입점 — [이관됨/무력화]
 *  INSP PDF는 더 이상 GAS(구글 변환 서버)에서 생성하지 않는다. 결재 완료 시
 *  _prepareInspPdfHandoff가 manifest.json을 기록하고, 실제 PDF는 로컬 파이썬(코워크)이 만든다.
 *  큐에 남은 과거 insp_pdf job(예: 예전 실패분)은 사진/폴더를 건드리지 않고 조용히 배수한다.
 *  ⚠ 아래 generateInspPdf / _cleanupInspPhotos / _fetchInspPhotoBlobForPdf 는
 *    현재 호출되지 않는 사양(dead code) — 참고용으로만 남겨둠(사진 삭제·구글 변환 미실행).
 * @param {Object} job { token, docNo, ... }
 * @returns {Object} { ok, message, skipped }
 */
function _processInspPdfJob(job) {
  return { ok: true, skipped: true,
    message: 'INSP PDF는 로컬 렌더러로 이관됨 — 큐 처리 없음 (' + ((job && job.docNo) || '') + ')' };
}

/**
 * 검수 사진을 PDF 임베드용 "축소 이미지" Blob으로 가져온다.
 *  - HTML→PDF 변환기가 큰 base64 data URI를 렌더링하지 못해 사진이 깨지는 문제를
 *    피하기 위해, 원본 대신 Drive 썸네일(약 1000px 축소본)을 사용한다.
 *  - 썸네일을 못 받으면(생성 전·권한 등) 원본 Blob으로 폴백한다.
 * @param {string} fileId Drive 파일 ID
 * @returns {Blob|null}
 */
function _fetchInspPhotoBlobForPdf(fileId) {
  // 1순위: Drive 썸네일(축소본) — data URI 용량을 줄여 PDF에서 정상 렌더되게 함
  try {
    var meta = Drive.Files.get(fileId, { fields: 'thumbnailLink', supportsAllDrives: true });
    var link = meta && meta.thumbnailLink;
    if (link) {
      // 썸네일 URL 끝의 크기 토큰(=s220 등)을 =s1000으로 올려 화질 확보
      link = link.replace(/=s\d+(-[\w-]+)?$/, '=s1000');
      var resp = UrlFetchApp.fetch(link, { muteHttpExceptions: true });
      if (resp.getResponseCode() === 200) {
        var b = resp.getBlob();
        if (b && /^image\//.test(b.getContentType() || '') && b.getBytes().length > 0) {
          return b.setName(fileId + '.jpg');
        }
      }
    }
  } catch (e) { /* 썸네일 실패 → 원본 폴백 */ }
  // 2순위: 원본 Blob (작은 사진은 그대로도 렌더 가능)
  try {
    return DriveApp.getFileById(fileId).getBlob();
  } catch (e2) {
    return null;
  }
}

/**
 * 검수보고서 PDF 생성 → FINAL/{PO번호} 폴더에 저장
 *  - 사진은 Drive(STAGING)에서 읽어 base64로 템플릿에 삽입
 *  - FINAL 폴더는 getOrCreateFolder(PO번호, ..., 'final') — PRC가 이미 만든 PO폴더와 동일
 * @param {string} token INSP token
 * @returns {Object} { ok, fileId, folderId, fileName }
 */
function generateInspPdf(token) {
  try {
    var ss   = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var info = _readInspRowByToken(ss, token);
    if (!info) return { ok: false, message: '행 없음' };
    var r = info.row;

    var docNo      = String(r[INSP_COL.DOC_NO] || '');
    var poNo       = String(r[INSP_COL.PO_NO] || '');
    var issueDate  = toDateStr(r[INSP_COL.ISSUE_DATE]);
    var verdict    = String(r[INSP_COL.VERDICT] || '');
    var isFinal    = String(r[INSP_COL.IS_FINAL] || '') === 'Y';
    var apprCount  = parseInt(r[INSP_COL.APPR_COUNT]) || 0;

    // 사진 → base64 (STAGING Drive에서 읽음)
    // ⚠ HTML→PDF 변환기(getAs(MimeType.PDF))는 용량이 큰 data URI 이미지를
    //   렌더링하지 못하고 "깨진 이미지" 자리표시자로 표시한다(브라우저 뷰어는 정상).
    //   원본(최대 1600px·수백 KB)을 그대로 임베드하면 PDF에서 사진이 깨지므로,
    //   Drive 썸네일(축소본)을 가져와 작은 data URI로 삽입한다.
    var photoTags = [];
    try {
      var photoList = JSON.parse(String(r[INSP_COL.PHOTO_LIST] || '[]'));
      photoList.forEach(function (p) {
        if (!p || !p.id) return;
        try {
          var blob = _fetchInspPhotoBlobForPdf(p.id);
          if (!blob) return;
          var b64  = Utilities.base64Encode(blob.getBytes());
          var mime = blob.getContentType() || 'image/jpeg';
          photoTags.push({ src: 'data:' + mime + ';base64,' + b64, name: p.name || '' });
        } catch (e) { /* 개별 사진 실패는 건너뜀 */ }
      });
    } catch (_) {}

    // 결재자 블록
    var approvers = [];
    for (var a = 0; a < apprCount; a++) {
      approvers.push({
        label:       String(r[inspApprCol(a, 0)] || ''),
        name:        String(r[inspApprCol(a, 1)] || ''),
        status:      String(r[inspApprCol(a, 3)] || ''),
        processedAt: toDateTimeStr(r[inspApprCol(a, 4)]),
      });
    }

    // 템플릿 렌더
    var tmpl = HtmlService.createTemplateFromFile('INSP_PDF_Template');
    tmpl.docNo        = docNo;
    tmpl.seq          = parseInt(r[INSP_COL.SEQ]) || 0;
    tmpl.isFinalLabel = isFinal ? ' · ★최종 검수' : '';
    tmpl.reqNo        = String(r[INSP_COL.REQ_NO] || '');
    tmpl.poNo         = poNo;
    tmpl.subject      = String(r[INSP_COL.SUBJECT] || '');
    tmpl.vendorName   = String(r[INSP_COL.VENDOR_NAME] || '');
    tmpl.drafterName  = String(r[INSP_COL.DRAFTER_NAME] || '');
    tmpl.dept         = String(r[INSP_COL.DEPT] || '');
    tmpl.issueDate    = issueDate;
    tmpl.receivedDate = toDateStr(r[INSP_COL.RECEIVED_DATE]);
    tmpl.receivedNote = String(r[INSP_COL.RECEIVED_NOTE] || '');
    tmpl.verdict      = verdict || '-';
    tmpl.verdictClass = (verdict === '합격') ? 'pass' : 'fail';
    tmpl.comment      = String(r[INSP_COL.COMMENT] || '');
    tmpl.photoTags    = photoTags;
    tmpl.approvers    = approvers;
    tmpl.generatedAt  = toDateTimeStr(new Date());

    var html = tmpl.evaluate().getContent();

    // HTML → Google 문서(임시) → PDF
    //  ⚠ Utilities.newBlob(html).getAs(MimeType.PDF) 변환기는 <img>(data URI·외부 URL
    //    모두)를 렌더링하지 못해 사진이 깨진다(CSS는 그려짐). 그래서 HTML을 먼저
    //    Google 문서로 변환(이미지가 실제로 임베드됨)한 뒤, 그 문서를 PDF로 내보낸다.
    //    무거운 변환은 Google 서버가 처리하므로 스크립트 부하는 작다. 임시 문서는
    //    변환 직후 삭제한다.
    var blob, tmpDocId = null;
    try {
      var htmlBlob = Utilities.newBlob(html, MimeType.HTML, docNo + '.html');
      var tmpDoc = Drive.Files.create(
        { name: 'INSP_TMP_' + docNo, mimeType: MimeType.GOOGLE_DOCS },
        htmlBlob,
        { supportsAllDrives: true }
      );
      tmpDocId = tmpDoc.id;
      blob = DriveApp.getFileById(tmpDocId).getAs(MimeType.PDF);
    } finally {
      if (tmpDocId) { try { Drive.Files.remove(tmpDocId); } catch (_) {} }
    }

    var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    var fileName = 'INSP_' + docNo + '_' + ts + '.pdf';
    blob.setName(fileName);

    // FINAL/{PO번호} 폴더 — PRC가 이미 만든 PO폴더와 동일 위치
    var folder = getOrCreateFolder(poNo, issueDate, 'final');
    var folderId = folder.getId();

    // 동일 검수보고서의 이전 PDF 정리 (재생성 시 중복 방지)
    var pattern = 'INSP_' + docNo + '_';
    var files = folder.getFiles();
    while (files.hasNext()) {
      var ef = files.next();
      if (ef.getName().indexOf(pattern) === 0) { try { ef.setTrashed(true); } catch (_) {} }
    }

    var file = folder.createFile(blob);
    return { ok: true, fileId: file.getId(), fileName: fileName, folderId: folderId };
  } catch (err) {
    return { ok: false, message: 'generateInspPdf 오류: ' + err.toString() };
  }
}

/**
 * 검수 사진 원본 + INSP STAGING 임시폴더 삭제 (PDF 생성 성공 후에만 호출)
 *  - 설계 원칙(요구사항 9): 사진은 PDF 삽입 후 보관하지 않음
 * @param {Array} row INSP 행
 * @returns {string} 처리 결과 메시지
 */
function _cleanupInspPhotos(row) {
  var deleted = 0;
  try {
    var photoList = JSON.parse(String(row[INSP_COL.PHOTO_LIST] || '[]'));
    photoList.forEach(function (p) {
      if (p && p.id) { try { DriveApp.getFileById(p.id).setTrashed(true); deleted++; } catch (_) {} }
    });
  } catch (_) {}

  // STAGING 임시폴더 삭제 (INSP_{docNo} 폴더)
  try {
    var folderId = String(row[INSP_COL.DRIVE_ID] || '');
    if (folderId) {
      var folder = DriveApp.getFolderById(folderId);
      // STAGING 폴더만 삭제 (FINAL 폴더와 혼동 방지: 폴더명이 INSP_로 시작하는 임시폴더)
      if (folder.getName().indexOf('INSP_') === 0) { folder.setTrashed(true); }
    }
  } catch (_) {}

  return '사진 ' + deleted + '장 삭제';
}

/**
 * Step 6 테스트 — 최종승인된 검수보고서로 PDF 생성만 수동 실행 (사진 삭제는 안 함)
 *   testInspStep6('INSP토큰')  → PDF가 FINAL/PO폴더에 생성되는지 확인
 *   ※ 실제 큐 처리는 사진 삭제까지 수행하므로, 이 테스트는 PDF 생성 단계만 점검
 */
function testInspStep6(token) {
  if (!token) { console.log('사용법: testInspStep6("최종승인된_검수보고서_token")'); return; }
  var res = generateInspPdf(token);
  console.log(JSON.stringify(res, null, 2));
  if (res.ok) console.log('→ FINAL 폴더에서 ' + res.fileName + ' 확인하세요. (사진 삭제는 실제 큐 처리 시 수행)');
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
    if (h[INSP_COL.APPR_START] !== 'drafter_appr_label') throw new Error('APPR_START 헤더 위치 불일치: ' + h[INSP_COL.APPR_START]);
    return true;
  });

  check('⑤ inspApprCol 헬퍼 — 정상/경계값', function () {
    if (inspApprCol(0, 0) !== 27) throw new Error('(0,0)=기안자 label → ' + inspApprCol(0, 0));
    if (inspApprCol(3, 5) !== 50) throw new Error('(3,5)=결재자3 comment → ' + inspApprCol(3, 5));
    var threw = false;
    try { inspApprCol(4, 0); } catch (e) { threw = true; }
    if (!threw) throw new Error('단계 범위 초과(4)가 차단되지 않음');
    return true;
  });

  var allPass = results.every(function (r) { return r.indexOf('[PASS]') === 0; });
  results.push(allPass ? '=== ALL PASS — Step 2 진행 가능 ===' : '=== 실패 항목 확인 필요 ===');
  console.log(results.join('\n'));
  return allPass;
}
