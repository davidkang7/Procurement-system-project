// ================================================================
// PO.gs — 주문서(PURCHASE ORDER) 생성 핸드오프
//  - PRC 최종승인(최종승인(PRC)) → 큐 job(pdf_and_consolidate) 말미에서
//    _preparePoHandoff_() 호출 (Code.gs _processPdfAndConsolidateJob).
//  - GAS는 xlsx를 만들지 않는다. FINAL/{PO} 폴더에 po_manifest.json을 기록하고
//    David에게 작업요청 메일만 보낸다.
//  - 로컬 파이썬(po-renderer/render_po.py)이 매니페스트를 읽어 2026 기준 양식
//    (내자=태성테크 / 외자=Coherent, 표준거래조건 2페이지 포함)으로 주문서를 만들고
//    주문서 폴더에 {PO번호}_{업체명}.xlsx + .pdf 로 올린다.
//  - 설계는 INSP PDF 핸드오프(INSP.gs)와 대칭. (검수보고서 파이썬 렌더러와 같은 방식)
//
//  ※ 주문서 표기 관례(Shipper 상호·REMARK 라벨·Payment Terms 문구·Destination)는
//     렌더러의 vendor_rules.json 이 담당한다. GAS는 품의서 원본 값만 넘긴다.
// ================================================================

var PO_CONFIG = {
  // 공유 드라이브 SCM_Innovation/02. Purchase/주문서 폴더 ID (2026-07-24 확인)
  ORDER_FOLDER_ID: '1Ot_KcJlCS8oZpu_IKCwAqz2CSPUTd81c',
  MANIFEST_NAME:   'po_manifest.json',
  SCHEMA:          'po-manifest-v2',
};

// ================================================================
// 주문서목록 시트 — 주문서 생성 대장
//  · 품의서목록(PRC 행)에는 주문서 상태 칸이 없다(AF열 다음이 결재자 블록이라
//    중간 컬럼 추가가 불가). 검수보고서목록과 같은 방식으로 별도 시트에 추적한다.
//  · 조인 키는 prcToken(품의서목록 S열 토큰). PO번호는 스냅샷.
//  · 이 시트 쓰기는 전부 비치명적이다 — 실패해도 결재·PDF·메일 경로를 막지 않는다.
// ================================================================
var PO_COL = {
  CREATED_AT:    0,   // 주문서 생성 요청(핸드오프) 시각
  PO_NO:         1,   // PO번호 (= PRC 품의번호 스냅샷)
  PRC_TOKEN:     2,   // 품의서목록 조인 키
  REQ_NO:        3,   // 원 REQ 품의번호 (Reference)
  SUBJECT:       4,   // 품의제목 스냅샷
  VENDOR_NAME:   5,   // 업체명 스냅샷
  CURRENCY:      6,   // KRW / USD / JPY …  (내자·외자 구분)
  TOTAL_AMT:     7,   // 합계금액 (부가세 별도)
  ITEM_COUNT:    8,   // 품목 수
  ISSUE_DATE:    9,   // 발행일자
  DELIVERY_DATE: 10,  // 납기일
  PAYMENT_TERMS: 11,  // 구매조건 원문
  STATUS:        12,  // '생성대기' | '생성완료'
  MANIFEST_ID:   13,  // po_manifest.json 파일 id
  FOLDER_ID:     14,  // FINAL/{PO} 폴더 id (매니페스트 위치)
  XLSX_FILE_ID:  15,  // 생성된 주문서 xlsx 파일 id
  XLSX_URL:      16,  // 주문서 파일 링크
  GENERATED_AT:  17,  // 주문서 생성 완료 시각
  GENERATED_BY:  18,  // 마감 실행 계정
  NOTE:          19,  // 비고 (재핸드오프 이력 등)
};
PO_COL._VERSION = 'po-sheet-v1.0';
var PO_TOTAL_COLS = 20;

var PO_STATUS = {
  PENDING: '생성대기',
  DONE:    '생성완료',   // 주문서를 실제로 만들어 주문서 폴더에 올린 건
  SKIP:    '생성불요',   // 주문서 발행이 필요 없다고 판단한 건 (사유는 비고에)
};
// 상태 드롭다운 목록 (손으로 고칠 때 오타로 '생성대기'에 남는 사고 방지)
var PO_STATUS_LIST = [PO_STATUS.PENDING, PO_STATUS.DONE, PO_STATUS.SKIP];

/**
 * 아직 주문서 작업이 남아 있는 상태인가 (미생성 목록에 띄울 대상).
 *  '생성완료'·'생성불요'가 아니면 전부 열린 건으로 본다 — 오타·빈칸이 조용히 사라지지 않게.
 */
function _isPoOpen_(status) {
  var s = String(status || '').trim();
  return s !== PO_STATUS.DONE && s !== PO_STATUS.SKIP;
}

/**
 * 주문서 운영 함수(관리자 콘솔 전용) 권한 관문.
 *  Apps Script는 이름이 밑줄로 **끝나는** 함수만 google.script.run에서 비공개다.
 *  아래 운영 함수들은 에디터에서 실행해야 해서 공개 이름을 유지하므로, 대신 여기서 막는다.
 *  ⚠ 인가 판정에는 getRequestUserEmail_()만 쓴다 — getActiveUserEmail()은 신원 확인 실패 시
 *    배포 계정(=관리자)으로 폴백해 관문이 열려 버린다(Code.gs 주석 참조).
 * @param {string} action 로그에 남길 액션명
 * @returns {string} 검증된 관리자 이메일
 * @throws {Error} 관리자가 아니면
 */
function _assertPoOperator_(action) {
  var actor = (typeof getRequestUserEmail_ === 'function')
    ? getRequestUserEmail_()
    : String(getActiveUserEmail() || '').toLowerCase();
  if (!isAdminUser(actor)) {
    try {
      writeAuditLog({
        eventType: AUDIT_EVENT.ADMIN_ACCESS_DENIED, actor: actor,
        reason: '주문서 운영 함수 무권한 호출: ' + action,
      });
    } catch (_) {}
    throw new Error('관리자 권한이 필요합니다. (' + action + ')');
  }
  return actor;
}

/**
 * PO_COL 인덱스가 0부터 빈틈없이 연속하는지 자가 점검 (INSP _inspSchemaSelfCheck와 같은 패턴).
 * 컬럼 추가·삭제 시 인덱스 누락/중복을 배포 전에 잡는다.
 * @throws {Error} 스키마 불일치 시
 */
function _poSchemaSelfCheck_() {
  var idx = [];
  for (var k in PO_COL) {
    if (!Object.prototype.hasOwnProperty.call(PO_COL, k)) continue;
    if (k === '_VERSION') continue;
    idx.push(PO_COL[k]);
  }
  idx.sort(function (a, b) { return a - b; });
  for (var i = 0; i < idx.length; i++) {
    if (idx[i] !== i) throw new Error('[PO] PO_COL 인덱스 불연속: ' + i + ' 자리에 ' + idx[i]);
  }
  if (idx.length !== PO_TOTAL_COLS) {
    throw new Error('[PO] PO_TOTAL_COLS 불일치: ' + idx.length + ' vs ' + PO_TOTAL_COLS);
  }
  return true;
}

/**
 * 상태 열에 드롭다운(생성대기/생성완료/생성불요)을 건다.
 *  손으로 고칠 때 오타가 나면 그 행이 '미생성'으로 조용히 남기 때문에 입력을 아예 제한한다.
 *  시트 생성 시 1회 + backfillPoSheet 실행 시(이미 만들어진 시트 보정) 호출.
 */
function _setPoStatusValidation_(sheet) {
  try {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(PO_STATUS_LIST, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, PO_COL.STATUS + 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  } catch (e) {
    Logger.log('[PO] 상태 드롭다운 설정 실패(무시): ' + e.toString());
  }
}

/**
 * 주문서목록 시트 보장 — 없으면 생성, 헤더가 비었으면 헤더 작성.
 *  ensureAuditLogSheet / ensureInspSheet 와 같은 패턴 (데이터가 있으면 즉시 통과).
 * @returns {Sheet}
 */
function ensurePoSheet_() {
  _poSchemaSelfCheck_();          // 컬럼 인덱스 드리프트를 첫 쓰기에서 잡는다
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.PO_SHEET_NAME);

  if (sheet.getLastRow() > 0) {
    // 같은 이름의 시트가 이미 있는데 우리 대장이 아니면(수기 시트 등) 엉뚱한 컬럼에 쓰게 된다.
    var hdr = sheet.getRange(1, 1, 1, 3).getValues()[0];
    if (String(hdr[0]) !== 'createdAt' || String(hdr[2]) !== 'prcToken') {
      throw new Error('[PO] 시트 "' + CONFIG.PO_SHEET_NAME + '" 의 머리글이 주문서 대장 형식이 아닙니다 '
        + '(A1=' + hdr[0] + ', C1=' + hdr[2] + '). 기존 시트를 다른 이름으로 옮기고 다시 실행하세요.');
    }
    return sheet;
  }

  var headers = [
    'createdAt', 'poNo', 'prcToken', 'reqNo', 'subject', 'vendorName',
    'currency', 'totalAmt', 'itemCount', 'issueDate', 'deliveryDate', 'paymentTerms',
    'status', 'manifestId', 'folderId', 'xlsxFileId', 'xlsxUrl',
    'generatedAt', 'generatedBy', 'note',
  ];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#fdf0d5');   // PO 테마(연한 주황) — 품의서목록·검수보고서목록과 시각적 구분

  _setPoStatusValidation_(sheet);

  sheet.setColumnWidth(PO_COL.CREATED_AT + 1, 150);
  sheet.setColumnWidth(PO_COL.PO_NO + 1, 140);
  sheet.setColumnWidth(PO_COL.PRC_TOKEN + 1, 250);
  sheet.setColumnWidth(PO_COL.SUBJECT + 1, 220);
  sheet.setColumnWidth(PO_COL.VENDOR_NAME + 1, 160);
  sheet.setColumnWidth(PO_COL.XLSX_URL + 1, 260);
  sheet.setColumnWidth(PO_COL.NOTE + 1, 220);

  return sheet;
}

/**
 * prcToken 으로 주문서목록 행 번호 조회 (1-based, 없으면 -1).
 *  토큰 컬럼만 대상으로 TextFinder를 돌려 오검출(다른 컬럼에 같은 문자열)을 막는다.
 */
function _findPoRowNum_(sheet, prcToken) {
  if (!prcToken) return -1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var found = sheet.getRange(2, PO_COL.PRC_TOKEN + 1, lastRow - 1, 1)
    .createTextFinder(String(prcToken))
    .matchEntireCell(true)
    .matchCase(true)
    .findNext();
  return found ? found.getRow() : -1;
}

/**
 * 주문서목록에 생성 요청 행을 등록/갱신 (upsert, 키=prcToken).
 *  - 신규: 상태 '생성대기'
 *  - 기존: 스냅샷·매니페스트 정보를 갱신. 이미 '생성완료'였다면 재발행 요청으로 보고
 *          '생성대기'로 되돌리고 비고에 이력을 남긴다(주문서를 다시 만들어야 하므로).
 * @param {Object} m po_manifest 객체
 * @returns {number} 행 번호 (실패 시 -1)
 */
function _upsertPoRow_(m) {
  var sheet = ensurePoSheet_();
  var rowNum = _findPoRowNum_(sheet, m.prcToken);
  var now = new Date();

  var prev = null;
  if (rowNum > 0) prev = sheet.getRange(rowNum, 1, 1, PO_TOTAL_COLS).getValues()[0];

  var note = prev ? String(prev[PO_COL.NOTE] || '') : '';
  var status = PO_STATUS.PENDING;
  if (prev && !_isPoOpen_(prev[PO_COL.STATUS])) {
    // 이미 닫힌 건(생성완료·생성불요)에 새 핸드오프가 왔다 = 재발행 요청.
    // 판단을 덮어쓰지 않고 이력을 남긴 채 다시 대기로 올려 사람이 보게 한다.
    note = ('재핸드오프 ' + toDateTimeStr(now) + ' — 이전 상태 ' + String(prev[PO_COL.STATUS] || '')
            + ', 재검토 필요' + (note ? ' / ' + note : ''));
  }

  var row = new Array(PO_TOTAL_COLS);
  row[PO_COL.CREATED_AT]    = prev ? prev[PO_COL.CREATED_AT] : now;
  row[PO_COL.PO_NO]         = m.poNo || '';
  row[PO_COL.PRC_TOKEN]     = m.prcToken || '';
  row[PO_COL.REQ_NO]        = m.reference || '';
  row[PO_COL.SUBJECT]       = m.subject || '';
  row[PO_COL.VENDOR_NAME]   = m.vendorName || '';
  row[PO_COL.CURRENCY]      = m.currency || '';
  row[PO_COL.TOTAL_AMT]     = Number(m.totalAmt) || 0;
  row[PO_COL.ITEM_COUNT]    = (m.items || []).length;
  row[PO_COL.ISSUE_DATE]    = m.issueDate || '';
  row[PO_COL.DELIVERY_DATE] = m.deliveryDate || '';
  row[PO_COL.PAYMENT_TERMS] = m.paymentTerms || '';
  row[PO_COL.STATUS]        = status;
  row[PO_COL.MANIFEST_ID]   = m.manifestFileId || (prev ? prev[PO_COL.MANIFEST_ID] : '');
  row[PO_COL.FOLDER_ID]     = m.stagingFolderId || (prev ? prev[PO_COL.FOLDER_ID] : '');
  row[PO_COL.XLSX_FILE_ID]  = '';      // 재생성 대상이므로 이전 산출물 정보는 비운다
  row[PO_COL.XLSX_URL]      = '';
  row[PO_COL.GENERATED_AT]  = '';
  row[PO_COL.GENERATED_BY]  = '';
  row[PO_COL.NOTE]          = note;

  if (rowNum > 0) {
    sheet.getRange(rowNum, 1, 1, PO_TOTAL_COLS).setValues([row]);
  } else {
    sheet.appendRow(row);
    rowNum = sheet.getLastRow();
  }
  return rowNum;
}

/**
 * 결재 완료된 PRC의 주문서(PO) xlsx 생성 작업을 로컬 파이썬으로 핸드오프.
 *  1) FINAL/{PO} 폴더 확보 (PRC 통합 이동 후엔 PRC.DRIVE_ID가 FINAL 폴더 id)
 *  2) 그 폴더에 po_manifest.json 기록 (렌더러 입력 — 결정적, 값만 담음)
 *  3) David(관리자)에게 작업요청 메일
 * @param {string} prcToken PRC token
 * @returns {Object} { ok, manifestFileId, stagingFolderId, message }
 */
function _preparePoHandoff_(prcToken) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
  var prcRowNum = findRowNumByToken(sheet, prcToken);
  if (prcRowNum < 0) return { ok: false, message: 'PRC 행 없음: ' + prcToken };
  var prc = readRow(sheet, prcRowNum);

  if (String(prc[COL.DOC_TYPE] || '') !== 'PRC') {
    return { ok: false, message: 'PRC 문서가 아님 (docType=' + prc[COL.DOC_TYPE] + ')' };
  }
  var status = String(prc[COL.STATUS] || '');
  if (status.indexOf('최종승인(PRC)') < 0) {
    return { ok: false, message: '최종승인(PRC) 상태가 아님: ' + status };
  }

  // 부모 REQ 조회 (Reference NO.(품의번호) = 원 REQ 문서번호)
  var reqRow = null;
  var reqToken = String(prc[COL.PARENT_DOC_ID] || '');
  if (reqToken) {
    var reqRowNum = findRowNumByToken(sheet, reqToken);
    if (reqRowNum > 0) reqRow = readRow(sheet, reqRowNum);
  }

  var manifest = _buildPoManifest_(prc, reqRow);

  // 매니페스트 기록 위치 = FINAL/{PO} 폴더 (REQ/PRC PDF와 동거).
  //  1순위: 통합 이동 완료면 PRC.DRIVE_ID가 FINAL 폴더 id.
  //  2순위: 아직이면 경로(PO번호+기안일)로 확보 (유령 폴더 방지 위해 1순위 우선).
  var finalFolderId = String(prc[COL.DRIVE_ID] || '');
  var folder;
  if (finalFolderId && String(prc[COL.MOVE_STATUS] || '') === 'FINAL') {
    folder = DriveApp.getFolderById(finalFolderId);
  } else {
    folder = getOrCreateFolder(manifest.poNo, prc[COL.ISSUE_DATE], 'final');
    finalFolderId = folder.getId();
  }
  manifest.stagingFolderId = finalFolderId;

  // 기존 매니페스트 있으면 교체
  var existing = folder.getFilesByName(PO_CONFIG.MANIFEST_NAME);
  while (existing.hasNext()) { try { existing.next().setTrashed(true); } catch (_) {} }
  var mBlob = Utilities.newBlob(JSON.stringify(manifest, null, 2), 'application/json', PO_CONFIG.MANIFEST_NAME);
  var mFile = folder.createFile(mBlob);
  manifest.manifestFileId = mFile.getId();

  // 주문서목록 대장 등록 (상태=생성대기). 실패해도 핸드오프 자체는 계속한다.
  try {
    _upsertPoRow_(manifest);
    writeAuditLog({ eventType: AUDIT_EVENT.PO_HANDOFF, docNo: manifest.poNo, docToken: manifest.prcToken,
      docType: 'PO', reason: '주문서 생성 요청 등록(생성대기)' });
  } catch (e) {
    try { notifyAdminError('[PO] 주문서목록 등록 실패: ' + manifest.poNo + ' / ' + e.toString()); } catch (_) {}
  }

  // David에게 작업요청 메일
  try {
    _sendPoHandoffEmail_(manifest, finalFolderId);
  } catch (e) {
    try { notifyAdminError('[PO] 핸드오프 메일 실패: ' + manifest.poNo + ' / ' + e.toString()); } catch (_) {}
  }

  Logger.log('[PO] 핸드오프 완료: ' + manifest.poNo + ' / manifest=' + mFile.getId());
  return { ok: true, manifestFileId: mFile.getId(), stagingFolderId: finalFolderId };
}

/**
 * PRC 행 → 주문서 렌더러 입력 매니페스트(po-manifest-v1).
 *  - buildPdfPayload(화이트리스트)를 재사용해 내부 메모 누출 차단.
 * @param {Array} prc PRC 행 배열
 * @param {Array=} reqRow 부모 REQ 행 배열 (Reference용, 없으면 null)
 * @returns {Object} manifest
 */
function _buildPoManifest_(prc, reqRow) {
  var payload = buildPdfPayload(prc);
  try { _verifyPdfPayloadWhitelist(payload); } catch (e) {
    throw new Error('[PO] 페이로드 화이트리스트 위반: ' + e.message);
  }

  var reference = reqRow ? String(reqRow[COL.DOC_NO] || '') : '';
  // Payment Terms 원문: 구매조건(자유텍스트) 우선, 없으면 구매방법.
  //   주문서 표기로의 변환(예: 'NET 30' → 'T/T(NET 30 day)')은 렌더러가 맡는다.
  var paymentTerms = payload.paymentInfo || payload.purchaseMethod || '';
  var currency = (payload.items[0] && payload.items[0].currency) || 'KRW';

  var items = payload.items.map(function (it) {
    return {
      name:     it.name || '',
      spec:     it.spec || '',
      qty:      Number(it.qty) || 0,
      price:    Number(it.price) || 0,
      currency: it.currency || currency,
    };
  });

  return {
    schemaVersion: PO_CONFIG.SCHEMA,
    prcToken:      payload.docToken,
    poNo:          payload.docNo,          // P/O-No = PRC 품의번호(=PO번호)
    reference:     reference,              // Reference NO.(품의번호) = 원 REQ 문서번호
    subject:       payload.subject,        // 파일명 앞부분(품의제목)
    vendorName:    payload.vendorName,     // Shipper
    vendorEmail:   payload.vendorEmail,
    vendorContact: payload.vendorContact,
    vendorPhone:   payload.vendorPhone,
    paymentTerms:  paymentTerms,           // Payment Terms 원문(구매조건/구매방법)
    // 주문서 Destination 은 관례상 내자 'KOREA' / 외자 'INLC Technology, Daejeon, Korea' 이다.
    // 품의서 납품장소(도로명 주소)를 그대로 쓰지 않는다 — 참고용으로만 넘긴다.
    deliveryAddr:  payload.deliveryAddr || '',
    deliveryDate:  payload.deliveryDate,   // DEL'Y DATE (H열)
    issueDate:     payload.issueDate,
    items:         items,                  // [{name,spec,qty,price,currency}]
    totalAmt:      payload.totalAmt,
    currency:      currency,
    orderFolderId: PO_CONFIG.ORDER_FOLDER_ID,   // 업로드 대상 = 주문서 폴더
    generatedAt:   toDateTimeStr(new Date()),
  };
}

/** 주문서 xlsx 작업요청 메일 (David/관리자) — 매니페스트 준비 완료 통지 */
function _sendPoHandoffEmail_(m, stagingFolderId) {
  var toList = CONFIG.ADMIN_NOTIFY_EMAILS || [];
  if (!toList.length) return;
  var stagingUrl = 'https://drive.google.com/drive/folders/' + stagingFolderId;
  var orderUrl   = 'https://drive.google.com/drive/folders/' + m.orderFolderId;
  var subj = '[주문서(PO) 생성 요청] ' + m.poNo + ' - ' + m.subject;
  var html = '<div style="font-family:sans-serif;max-width:600px;">'
    + '<h2 style="color:#0e7d72;">구매 결재 완료 — 주문서(PO) 생성 대기</h2>'
    + '<p>아래 PRC의 결재가 완료되었습니다. 로컬 렌더러로 주문서 Excel을 생성해 주문서 폴더에 보관하세요.</p>'
    + '<table style="border-collapse:collapse;font-size:14px;">'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">P/O-No</td><td>' + escapeHtml(m.poNo) + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">품의제목</td><td>' + escapeHtml(m.subject) + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">업체</td><td>' + escapeHtml(m.vendorName) + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">품목</td><td>' + (m.items ? m.items.length : 0) + '건</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">prcToken</td><td>' + escapeHtml(m.prcToken) + '</td></tr>'
    + '<tr><td style="color:#888;padding:4px 12px 4px 0;">통화</td><td>' + escapeHtml(m.currency || '') + ' ('
    + (String(m.currency || 'KRW') === 'KRW' ? '내자 — 국문 거래조건' : '외자 — 영문 거래조건') + ')</td></tr>'
    + '</table>'
    + '<p style="margin-top:12px;">'
    + '· 입력 폴더(po_manifest.json): <a href="' + stagingUrl + '">' + stagingUrl + '</a><br>'
    + '· 출력 폴더(주문서): <a href="' + orderUrl + '">' + orderUrl + '</a></p>'
    + '<p style="font-size:13px;">실행: <code>po-renderer\\.venv\\Scripts\\python.exe render_po.py --folder ' + stagingUrl + '</code><br>'
    + '→ <b>' + escapeHtml(m.poNo) + '_' + escapeHtml(m.vendorName) + '.xlsx</b> + 같은 이름 PDF 생성·업로드</p>'
    + '<p style="font-size:12px;color:#888;">생성 후 PDF를 열어 1p 주문서 · 2p 표준거래조건과 글자 잘림을 눈으로 확인하세요.<br>'
    + '완료 후 GAS에서 <b>markPoDone("' + escapeHtml(m.prcToken) + '", "생성된파일ID")</b> 실행해 마감(감사로그) 하세요.</p>'
    + '</div>';
  var plain = '주문서(PO) 생성 요청\nP/O-No: ' + m.poNo + '\n품의제목: ' + m.subject + '\n업체: ' + m.vendorName
    + '\n통화: ' + (m.currency || '')
    + '\n입력 폴더: ' + stagingUrl + '\n출력 폴더(주문서): ' + orderUrl
    + '\n실행: render_po.py --folder ' + stagingUrl
    + '\n산출물: ' + m.poNo + '_' + m.vendorName + '.xlsx (+ PDF)'
    + '\n완료 후: markPoDone("' + m.prcToken + '", "파일ID")';
  sendEmailWithRetry(toList.join(','), subj, plain, html);
}

/**
 * 주문서 xlsx 생성·업로드 완료 마감.
 *  - 주문서목록 행을 '생성완료'로 바꾸고(파일 id·링크·시각·실행자 기록) 감사로그를 남긴다.
 *  - 대장에 행이 없으면(과거 건·백필 전) 품의서 정보로 행을 만들어 마감한다.
 * @param {string} prcToken PRC token
 * @param {string=} xlsxFileId 생성된 xlsx 파일 id (선택 — 있으면 링크까지 기록)
 * @returns {Object} { ok, message }
 */
function markPoDone(prcToken, xlsxFileId) {
  _assertPoOperator_('markPoDone');
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
  var rowNum = findRowNumByToken(sheet, prcToken);
  if (rowNum < 0) return { ok: false, message: '행 없음: ' + prcToken };
  var prc = readRow(sheet, rowNum);
  var poNo = String(prc[COL.DOC_NO] || '');

  var sheetMsg = '';
  try {
    var poSheet = ensurePoSheet_();
    var poRow = _findPoRowNum_(poSheet, prcToken);
    if (poRow < 0) {
      // 대장에 없는 과거 건 — 품의서 정보만으로 최소 행을 만든다(핸드오프 정보는 비움).
      poRow = _upsertPoRow_({
        poNo: poNo, prcToken: prcToken, subject: String(prc[COL.SUBJECT] || ''),
        vendorName: String(prc[COL.VENDOR_NAME] || ''), totalAmt: Number(prc[COL.TOTAL_AMT]) || 0,
        items: parseItemsSummary(String(prc[COL.ITEMS] || '')),
        issueDate: toDateStr(prc[COL.ISSUE_DATE]), deliveryDate: toDateStr(prc[COL.DELIVERY_DATE]),
        paymentTerms: String(prc[COL.PAYMENT_INFO] || prc[COL.PURCHASE_METHOD] || ''),
        currency: (parseItemsSummary(String(prc[COL.ITEMS] || ''))[0] || {}).currency || '',
      });
      sheetMsg = ' (대장 행 신규 생성)';
    }
    var updates = [
      [PO_COL.STATUS + 1,       PO_STATUS.DONE],
      [PO_COL.XLSX_FILE_ID + 1, xlsxFileId || ''],
      [PO_COL.XLSX_URL + 1,     xlsxFileId ? ('https://drive.google.com/file/d/' + xlsxFileId + '/view') : ''],
      [PO_COL.GENERATED_AT + 1, new Date()],
      [PO_COL.GENERATED_BY + 1, getActiveUserEmail() || ''],
    ];
    batchUpdate(poSheet, poRow, updates);
  } catch (e) {
    sheetMsg = ' ※ 주문서목록 갱신 실패: ' + e.toString();
    try { notifyAdminError('[PO] 주문서목록 마감 실패: ' + poNo + ' / ' + e.toString()); } catch (_) {}
  }

  try {
    writeAuditLog({ eventType: AUDIT_EVENT.PO_GENERATED, docNo: poNo, docToken: prcToken,
      docType: 'PO', reason: '주문서 xlsx 로컬 생성 완료' + (xlsxFileId ? ' / fileId=' + xlsxFileId : '') });
  } catch (_) {}
  Logger.log('[PO] 주문서 생성 마감: ' + poNo + sheetMsg);
  return { ok: true, message: '주문서 생성 마감: ' + poNo + sheetMsg };
}

/**
 * 주문서 발행이 필요 없는 건 마감 ('생성불요').
 *  - 과거 건 정리·발주 취소·다른 경로로 처리된 건 등. 사유를 반드시 남긴다.
 *  - 시트에서 상태를 직접 '생성불요'로 바꿔도 되지만, 이 함수를 쓰면 사유와 감사로그가 함께 남는다.
 * @param {string} prcToken PRC token (주문서목록 C열)
 * @param {string} reason 사유 (비고에 기록 — 필수)
 * @returns {Object} { ok, message }
 */
function markPoNotRequired(prcToken, reason) {
  _assertPoOperator_('markPoNotRequired');
  if (!prcToken || !String(reason || '').trim()) {
    return { ok: false, message: '사용법: markPoNotRequired("prcToken", "사유")' };
  }
  var sheet = ensurePoSheet_();
  var rowNum = _findPoRowNum_(sheet, prcToken);
  if (rowNum < 0) return { ok: false, message: '주문서목록에 행 없음: ' + prcToken };

  var row = sheet.getRange(rowNum, 1, 1, PO_TOTAL_COLS).getValues()[0];
  var poNo = String(row[PO_COL.PO_NO] || '');
  var prevNote = String(row[PO_COL.NOTE] || '');
  var note = '생성불요 ' + toDateTimeStr(new Date()) + ' — ' + String(reason).trim()
           + (prevNote ? ' / ' + prevNote : '');

  batchUpdate(sheet, rowNum, [
    [PO_COL.STATUS + 1,       PO_STATUS.SKIP],
    [PO_COL.GENERATED_AT + 1, new Date()],
    [PO_COL.GENERATED_BY + 1, getActiveUserEmail() || ''],
    [PO_COL.NOTE + 1,         note],
  ]);
  try {
    writeAuditLog({ eventType: AUDIT_EVENT.PO_SKIPPED, docNo: poNo, docToken: prcToken,
      docType: 'PO', reason: '주문서 생성불요: ' + String(reason).trim() });
  } catch (_) {}
  Logger.log('[PO] 생성불요 처리: ' + poNo);
  return { ok: true, message: '생성불요 처리: ' + poNo };
}

/**
 * 주문서 미생성 목록 (관리자 콘솔).
 *  - 주문서목록에서 상태가 '생성완료'가 아닌 행을 오래된 순으로 보여 준다.
 *  - 대장에 아직 없는 과거 건은 backfillPoSheet()로 먼저 채운다.
 * @returns {Array<Object>}
 */
function listPoPending() {
  _assertPoOperator_('listPoPending');
  var sheet = ensurePoSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { console.log('[PO] 주문서목록이 비어 있습니다. backfillPoSheet() 실행을 검토하세요.'); return []; }

  var rows = sheet.getRange(2, 1, lastRow - 1, PO_TOTAL_COLS).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    if (!_isPoOpen_(rows[i][PO_COL.STATUS])) continue;   // 생성완료·생성불요는 제외
    out.push({
      poNo:       String(rows[i][PO_COL.PO_NO] || ''),
      vendorName: String(rows[i][PO_COL.VENDOR_NAME] || ''),
      currency:   String(rows[i][PO_COL.CURRENCY] || ''),
      totalAmt:   Number(rows[i][PO_COL.TOTAL_AMT]) || 0,
      prcToken:   String(rows[i][PO_COL.PRC_TOKEN] || ''),
      folderId:   String(rows[i][PO_COL.FOLDER_ID] || ''),
      createdAt:  toDateTimeStr(rows[i][PO_COL.CREATED_AT]),
      rowNum:     i + 2,
    });
  }
  console.log('[PO] 주문서 미생성 ' + out.length + '건');
  out.forEach(function (r) {
    console.log('  · ' + r.poNo + ' / ' + r.vendorName + ' / ' + r.currency + ' ' + r.totalAmt + ' / 요청 ' + r.createdAt);
  });
  return out;
}

/**
 * 과거 건 백필 — 결재 완료(PRC)된 품의를 주문서목록에 채워 넣는다.
 *  - 주문서 폴더를 한 번 훑어 `{PO번호}_...xlsx` 가 이미 있으면 '생성완료'로,
 *    없으면 '생성대기'로 기록한다. 이미 대장에 있는 행은 건드리지 않는다(재발행 오인 방지).
 *  - 관리자 수동 실행 전용. 실행 시간이 길어지면 남은 건은 다음 실행에서 이어서 처리한다.
 * @param {Object=} opts { limit: 처리 상한(기본 200) }
 * @returns {Object} { ok, added, done, pending, remaining }
 */
function backfillPoSheet(opts) {
  _assertPoOperator_('backfillPoSheet');
  opts = opts || {};
  var limit = opts.limit || 200;
  var startedAt = Date.now();
  var TIME_BUDGET_MS = 4 * 60 * 1000;

  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
  var poSheet = ensurePoSheet_();
  _setPoStatusValidation_(poSheet);   // 먼저 만들어진 시트에도 드롭다운을 보정해 둔다

  // 주문서 폴더 1회 스캔 → { PO번호: {id, name} }
  var fileMap = {};
  try {
    var files = DriveApp.getFolderById(PO_CONFIG.ORDER_FOLDER_ID).getFiles();
    while (files.hasNext()) {
      var f = files.next();
      var name = f.getName();
      if (name.indexOf('.xlsx') < 0) continue;
      var m = name.match(/^(TO-[A-Z]{2}-\d{2}-\d{3,})_/);
      if (m && !fileMap[m[1]]) fileMap[m[1]] = { id: f.getId(), name: name };
    }
  } catch (e) {
    return { ok: false, message: '주문서 폴더 스캔 실패: ' + e.toString() };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, added: 0, done: 0, pending: 0, remaining: 0, message: '품의서목록이 비어 있음' };
  var rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  var added = 0, doneCnt = 0, pendingCnt = 0, remaining = 0;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[COL.DOC_TYPE] || '') !== 'PRC') continue;
    if (String(r[COL.STATUS] || '').indexOf('최종승인(PRC)') < 0) continue;

    if (added >= limit || (Date.now() - startedAt) > TIME_BUDGET_MS) { remaining++; continue; }

    var token = String(r[COL.TOKEN] || '');
    if (_findPoRowNum_(poSheet, token) > 0) continue;      // 이미 대장에 있음 — 손대지 않는다

    var poNo = String(r[COL.DOC_NO] || '');
    var items = parseItemsSummary(String(r[COL.ITEMS] || ''));
    var hit = fileMap[poNo];
    var now = new Date();

    var row = new Array(PO_TOTAL_COLS);
    row[PO_COL.CREATED_AT]    = r[COL.ISSUE_DATE] || now;
    row[PO_COL.PO_NO]         = poNo;
    row[PO_COL.PRC_TOKEN]     = token;
    row[PO_COL.REQ_NO]        = '';
    row[PO_COL.SUBJECT]       = String(r[COL.SUBJECT] || '');
    row[PO_COL.VENDOR_NAME]   = String(r[COL.VENDOR_NAME] || '');
    row[PO_COL.CURRENCY]      = (items[0] || {}).currency || '';
    row[PO_COL.TOTAL_AMT]     = Number(r[COL.TOTAL_AMT]) || 0;
    row[PO_COL.ITEM_COUNT]    = items.length;
    row[PO_COL.ISSUE_DATE]    = toDateStr(r[COL.ISSUE_DATE]);
    row[PO_COL.DELIVERY_DATE] = toDateStr(r[COL.DELIVERY_DATE]);
    row[PO_COL.PAYMENT_TERMS] = String(r[COL.PAYMENT_INFO] || r[COL.PURCHASE_METHOD] || '');
    row[PO_COL.STATUS]        = hit ? PO_STATUS.DONE : PO_STATUS.PENDING;
    row[PO_COL.MANIFEST_ID]   = '';
    row[PO_COL.FOLDER_ID]     = String(r[COL.DRIVE_ID] || '');
    row[PO_COL.XLSX_FILE_ID]  = hit ? hit.id : '';
    row[PO_COL.XLSX_URL]      = hit ? ('https://drive.google.com/file/d/' + hit.id + '/view') : '';
    row[PO_COL.GENERATED_AT]  = '';      // 과거 건은 생성 시각을 알 수 없다 — 비워 둔다
    row[PO_COL.GENERATED_BY]  = '';
    row[PO_COL.NOTE]          = hit ? ('백필: 주문서 폴더에서 확인 — ' + hit.name) : '백필: 주문서 미확인';

    poSheet.appendRow(row);
    added++;
    if (hit) doneCnt++; else pendingCnt++;
  }

  var msg = '[PO] 백필 완료 — 신규 ' + added + '건 (생성완료 ' + doneCnt + ' / 생성대기 ' + pendingCnt + ')'
          + (remaining ? ', 남은 후보 ' + remaining + '건은 다시 실행하세요' : '');
  Logger.log(msg);
  console.log(msg);
  return { ok: true, added: added, done: doneCnt, pending: pendingCnt, remaining: remaining, message: msg };
}

/**
 * 실패/누락 건의 PO 핸드오프 재생성 (매니페스트 재작성 + 메일).
 *  - 사용법: rerunPoHandoff("PRC_token")
 * @param {string} prcToken PRC token
 */
function rerunPoHandoff(prcToken) {
  _assertPoOperator_('rerunPoHandoff');
  if (!prcToken) { console.log('사용법: rerunPoHandoff("PRC_token")'); return; }
  var res = _preparePoHandoff_(prcToken);
  console.log(JSON.stringify(res));
  return res;
}

/**
 * 주문서 폴더 ID 재확인용 (관리자 콘솔). 폴더명/부모 경로 확인.
 */
function _checkPoOrderFolder() {
  _assertPoOperator_('_checkPoOrderFolder');
  var f = DriveApp.getFolderById(PO_CONFIG.ORDER_FOLDER_ID);
  var parents = f.getParents();
  var parentName = parents.hasNext() ? parents.next().getName() : '(없음)';
  Logger.log('[PO] 주문서 폴더: ' + f.getName() + ' / 상위: ' + parentName + ' / id=' + PO_CONFIG.ORDER_FOLDER_ID);
  return { name: f.getName(), parent: parentName, id: PO_CONFIG.ORDER_FOLDER_ID };
}

// ================================================================
// [임시] 주문서 마감용 무인자 래퍼 — GAS 편집기는 인자를 넘길 수 없어서 둔다.
//  실행 후 제거하고 다시 push 한다. (INSP 마감 래퍼와 같은 관례)
// ================================================================

/**
 * 2026-09-17 생성분 마감 — TO-PO-26-278 에이에스씨 (ESD 자재).
 *  주문서 폴더 업로드 완료: TO-PO-26-278_에이에스씨.xlsx + .pdf
 */
function _tmp_markPoDone_20260917() {
  var jobs = [
    // [PO번호, prcToken, xlsx 파일 ID]
    ['TO-PO-26-278', '903ac2c9-2c18-4ab5-9c4a-84fa4363a9cf', '1pwv2GKiZyxTU7apd_v8TI9ro_wJxK919'],
  ];
  var out = jobs.map(function (j) {
    var r = markPoDone(j[1], j[2]);
    return j[0] + ' → ' + (r.ok ? 'OK' : 'FAIL') + ' / ' + r.message;
  });
  Logger.log(out.join('\n'));
  console.log(out.join('\n'));
  return out;
}
