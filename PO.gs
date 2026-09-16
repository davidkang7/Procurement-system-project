// ================================================================
// PO.gs — 주문서(PURCHASE ORDER) 생성 핸드오프
//  - PRC 최종승인(최종승인(PRC)) → 큐 job(pdf_and_consolidate) 말미에서
//    _preparePoHandoff() 호출 (Code.gs _processPdfAndConsolidateJob).
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

/**
 * 결재 완료된 PRC의 주문서(PO) xlsx 생성 작업을 로컬 파이썬으로 핸드오프.
 *  1) FINAL/{PO} 폴더 확보 (PRC 통합 이동 후엔 PRC.DRIVE_ID가 FINAL 폴더 id)
 *  2) 그 폴더에 po_manifest.json 기록 (렌더러 입력 — 결정적, 값만 담음)
 *  3) David(관리자)에게 작업요청 메일
 * @param {string} prcToken PRC token
 * @returns {Object} { ok, manifestFileId, stagingFolderId, message }
 */
function _preparePoHandoff(prcToken) {
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

  var manifest = _buildPoManifest(prc, reqRow);

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

  // David에게 작업요청 메일
  try {
    _sendPoHandoffEmail(manifest, finalFolderId);
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
function _buildPoManifest(prc, reqRow) {
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
function _sendPoHandoffEmail(m, stagingFolderId) {
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
 * 주문서 xlsx 생성·업로드 완료 마감 (감사로그 기록).
 *  - PO 상태 전용 컬럼이 없어 상태 변경 없이 감사로그만 남긴다.
 * @param {string} prcToken PRC token
 * @param {string=} xlsxFileId 생성된 xlsx 파일 id (감사로그용, 선택)
 * @returns {Object} { ok, message }
 */
function markPoDone(prcToken, xlsxFileId) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
  var rowNum = findRowNumByToken(sheet, prcToken);
  if (rowNum < 0) return { ok: false, message: '행 없음: ' + prcToken };
  var prc = readRow(sheet, rowNum);
  var poNo = String(prc[COL.DOC_NO] || '');
  try {
    writeAuditLog({ eventType: 'PO_GENERATED', docNo: poNo, docToken: prcToken,
      docType: 'PO', reason: '주문서 xlsx 로컬 생성 완료' + (xlsxFileId ? ' / fileId=' + xlsxFileId : '') });
  } catch (_) {}
  Logger.log('[PO] 주문서 생성 마감: ' + poNo);
  return { ok: true, message: '주문서 생성 마감(감사로그): ' + poNo };
}

/**
 * 실패/누락 건의 PO 핸드오프 재생성 (매니페스트 재작성 + 메일).
 *  - 사용법: rerunPoHandoff("PRC_token")
 * @param {string} prcToken PRC token
 */
function rerunPoHandoff(prcToken) {
  if (!prcToken) { console.log('사용법: rerunPoHandoff("PRC_token")'); return; }
  var res = _preparePoHandoff(prcToken);
  console.log(JSON.stringify(res));
  return res;
}

/**
 * 주문서 폴더 ID 재확인용 (관리자 콘솔). 폴더명/부모 경로 확인.
 */
function _checkPoOrderFolder() {
  var f = DriveApp.getFolderById(PO_CONFIG.ORDER_FOLDER_ID);
  var parents = f.getParents();
  var parentName = parents.hasNext() ? parents.next().getName() : '(없음)';
  Logger.log('[PO] 주문서 폴더: ' + f.getName() + ' / 상위: ' + parentName + ' / id=' + PO_CONFIG.ORDER_FOLDER_ID);
  return { name: f.getName(), parent: parentName, id: PO_CONFIG.ORDER_FOLDER_ID };
}
