// ================================================================
// INSP.gs — 검수보고서(INSP) 모듈 (누적본: Step 1~3)
// Step 1: 스키마 + 시트 보장 / Step 2: 홈 "결재 완료" 그룹 조회
// Step 3: 작성 폼 데이터 조회 + 제출(submitInspForClient) + 사진 임시 저장
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
    if (!prcToken) return { ok: false, message: 'PRC 토큰이 필요합니다.' };

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

    if (!prc) return { ok: false, message: '대상 구매품의서(PRC)를 찾을 수 없습니다.' };
    if (prc.status !== '최종승인(PRC)') {
      return { ok: false, message: '최종승인(PRC) 상태의 품의에만 검수보고서를 제출할 수 있습니다.' };
    }

    var req = reqByToken[prc.parentToken] || null;
    var drafterEmail = req ? String(req.drafterEmail || '') : '';

    // 권한: 원본 REQ 기안자 본인만
    if (!drafterEmail || drafterEmail.toLowerCase() !== actor) {
      return { ok: false, message: '검수보고서는 원본 품의 기안자 본인만 제출할 수 있습니다.' };
    }

    // 기존 검수보고서 그룹 (회차 채번 + 종결 차단 + 재상신 prefill)
    var groups = _getInspGroupsByPrc(ss);
    var existing = groups[prcToken] || [];

    var closed = existing.some(function(x) { return x.isFinal && x.status === '최종승인(INSP)'; });
    if (closed && !resubToken) {
      return { ok: false, message: '이미 최종 검수가 승인되어 종결된 품의입니다. 추가 제출이 불가합니다.' };
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
    var approvers = data.approvers || [];
    if (approvers.length < 1) return { ok: false, message: '결재자를 1명 이상 지정해 주세요.' };
    if (approvers.length > INSP_COL.MAX_APPROVERS) {
      return { ok: false, message: '결재자는 최대 ' + INSP_COL.MAX_APPROVERS + '인까지 지정 가능합니다.' };
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
        if (!prc) return { ok: false, message: '대상 PRC를 찾을 수 없습니다.' };
        if (String(prc.row[COL.STATUS] || '') !== '최종승인(PRC)') {
          return { ok: false, message: '최종승인(PRC) 상태가 아닙니다.' };
        }
        var parentToken = String(prc.row[COL.PARENT_DOC_ID] || '');
        var reqInfo = _findRowByTokenInDoc(docSheet, parentToken);
        var drafterEmail = reqInfo ? String(reqInfo.row[COL.APPR_START + 2] || '') : '';
        if (!drafterEmail || drafterEmail.toLowerCase() !== actor) {
          return { ok: false, message: '원본 품의 기안자 본인만 제출할 수 있습니다.' };
        }

        // 기존 INSP 그룹 (회차 + 종결)
        var groups = _getInspGroupsByPrc(ss);
        var existing = groups[data.prcToken] || [];
        var closed = existing.some(function(x) { return x.isFinal && x.status === '최종승인(INSP)'; });

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
          if (closed) return { ok: false, message: '이미 종결된 품의입니다. 추가 제출이 불가합니다.' };
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
          rowArr[inspApprCol(a, 0)] = ap ? (ap.label || ('결재자 ' + (a + 1))) : '';
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
  var url = CONFIG.WEBAPP_URL + '?action=approve&docKind=INSP'
          + '&token=' + encodeURIComponent(token)
          + '&idx='   + idx;
  var subject = '[검수보고서 결재요청] ' + data.docNo + ' - ' + data.subject;
  var html =
    '<div style="font-family:sans-serif;max-width:600px;">'
    + '<h2 style="color:#0e7d72;">검수보고서 결재 요청</h2>'
    + '<p><b>' + escapeHtml(approver.name || '') + '</b> 님, 아래 검수보고서의 결재를 요청드립니다.</p>'
    + '<table style="border-collapse:collapse;font-size:14px;">'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">문서번호</td><td>' + escapeHtml(data.docNo) + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">품명</td><td>' + escapeHtml(data.subject) + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">업체명</td><td>' + escapeHtml(data.vendorName || '') + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">기안자</td><td>' + escapeHtml(data.drafter || '') + '</td></tr>'
    + '</table>'
    + '<p style="margin-top:16px;"><a href="' + url + '" style="background:#0e7d72;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;">결재하러 가기</a></p>'
    + '</div>';
  var plain = '검수보고서 결재 요청\n문서번호: ' + data.docNo + '\n품명: ' + data.subject + '\n결재: ' + url;
  // 일반 결재 알림 경로(sendEmailWithRetry, GmailApp 기반)로 발송
  sendEmailWithRetry(approver.email, subject, plain, html);
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
