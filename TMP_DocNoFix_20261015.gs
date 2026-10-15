// ================================================================
// [임시] 품의번호 오기 정정 — 2026-10-15 김민지(mjkim) 기안 3건
//  GAS 편집기는 인자를 넘길 수 없어서 무인자 래퍼를 둔다. 실행 후 이 파일째 제거하고 다시 push 한다.
//
//  실행 순서 (GAS 편집기, 관리자 계정):
//    1) _tmp_docNoFix_20261015_dryRun  — 쓰기 없이 검증·변경 예정 목록만 출력
//    2) _tmp_docNoFix_20261015         — 실제 정정 (다시 돌려도 안전: 끝난 단계는 SKIP)
//
//  3건 모두 결재 완료 + PRC(PO) 생성 후라 재상신(품의번호 강제 유지)·폐기 경로로는 고칠 수 없다.
//  바뀌는 곳
//    시트  품의서목록 B열(REQ 행) · Y열 QUO_ 파일명 / 검수보고서목록 docNo·reqNo / 주문서목록 reqNo
//    Drive FINAL(PO) 폴더: REQ PDF 재생성(옛 번호 PDF 휴지통) · PRC PDF 재생성('참조 원본 품의' 표기)
//          {번호}_품의서요약.txt 이름·본문 · QUO_{번호}.pdf 이름 · po_manifest.json reference
//          STAGING 빈 폴더 이름 · INSP_ 폴더 이름 + manifest.json 재작성(로컬 렌더러 재실행용)
//  안 바꾸는 것: FINAL 폴더(PO번호 이름) · 주문서 xlsx(Reference NO. 관례상 공란) · 과거 시스템로그 · 발송된 메일
//  INSP PDF는 GAS가 만들지 않는다 → 이 래퍼 실행 후 로컬 렌더러로 다시 만들고 옛 PDF를 정리한다.
// ================================================================

var _TMP_DOCNO_FIX_20261015 = [
  // [REQ token, 기존 품의번호, 정정 품의번호, PRC token, PO번호]
  ['003116ab-f981-4a1e-84a9-ce915b042d6f', 'TR-AP-020-055', 'TR-AP-26-055', '190b22a7-77f3-4836-bde8-50f9cbfe94ea', 'TO-PO-26-312'],
  ['fd68747d-4d25-4d67-bb95-cadf49a0558e', 'TR-AP-27-048',  'TR-AP-26-048', '130aa37a-26ba-4865-a286-4610e51974f6', 'TO-PO-26-274'],
  ['ed999a8c-b9c3-4b9c-a970-6c009fe2604d', 'TR-AT-026-50',  'TR-AP-26-050', '4ad096e5-5aff-4c00-8c04-6cd575d3be14', 'TO-PO-26-290'],
];

function _tmp_docNoFix_20261015_dryRun() { return _tmpDocNoFix_(true); }
function _tmp_docNoFix_20261015()        { return _tmpDocNoFix_(false); }

function _tmpDocNoFix_(dryRun) {
  var actor = getActiveUserEmail();
  if (!isAdminUser(actor)) throw new Error('관리자만 실행할 수 있습니다: ' + actor);

  var log = [];
  var say = function (s) { log.push(s); };
  say((dryRun ? '[DRY RUN] ' : '') + '품의번호 정정 시작 — 실행자 ' + actor);

  // ── 1단계: 시트 (전역 락 안에서, 검증 → 쓰기) ──
  var plans = withLock(function () {
    return _TMP_DOCNO_FIX_20261015.map(function (j) { return _tmpDocNoFixSheets_(j, dryRun, say); });
  });

  // ── 2단계: Drive (락 밖 — PDF 생성은 수 초씩 걸린다) ──
  plans.forEach(function (p) {
    if (!p.ok) return;
    _tmpDocNoFixDrive_(p, dryRun, say);
    // 재실행(PRC PDF 재생성만 일어남)에서는 로그를 다시 쓰지 않는다
    var meaningful = p.changed.length +
      p.driveChanged.filter(function (s) { return s.indexOf('PRC PDF') !== 0; }).length;
    if (!dryRun && meaningful > 0) {
      writeAuditLog({
        eventType: 'ADMIN_DOCNO_CORRECT',
        actor:     actor,
        actorRole: 'admin',
        docNo:     p.newNo,
        docToken:  p.reqToken,
        docType:   'REQ',
        reason:    '품의번호 오기 정정: ' + p.oldNo + ' → ' + p.newNo + ' (결재 완료 후라 데이터 정정, 기안자 김민지)',
        payload:   { oldDocNo: p.oldNo, newDocNo: p.newNo, prcToken: p.prcToken, poNo: p.poNo,
                     sheet: p.changed, drive: p.driveChanged },
      });
    }
  });

  var okCount = plans.filter(function (p) { return p.ok; }).length;
  say('완료: ' + okCount + '/' + plans.length + '건 ' + (dryRun ? '검증 통과 (쓰기 없음)' : '처리'));
  if (!dryRun && okCount > 0) {
    say('▶ 다음: INSP 렌더러로 검수보고서 PDF 재생성 (manifest.json 재작성됨) → 옛 INSP PDF 휴지통');
  }
  var msg = log.join('\n');
  Logger.log(msg);
  console.log(msg);
  return log;
}

/**
 * 1단계 — 한 건의 시트 정정. 어긋나는 값이 하나라도 있으면 그 건은 아무것도 쓰지 않고 SKIP.
 * @returns {Object} 2단계에 넘길 계획 { ok, reqToken, oldNo, newNo, prcToken, poNo, folderId, issueDate, insp[], changed[] }
 */
function _tmpDocNoFixSheets_(j, dryRun, say) {
  var reqToken = j[0], oldNo = j[1], newNo = j[2], prcToken = j[3], poNo = j[4];
  var p = { ok: false, reqToken: reqToken, oldNo: oldNo, newNo: newNo, prcToken: prcToken, poNo: poNo,
            insp: [], changed: [], driveChanged: [] };
  var head = '[' + oldNo + ' → ' + newNo + ' / ' + poNo + '] ';
  var skip = function (why) { say(head + 'SKIP — ' + why); return p; };

  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);

  // ── 검증: REQ 행 ──
  var reqRowNum = findRowNumByToken(sheet, reqToken);
  if (reqRowNum < 0) return skip('REQ 행 없음 (token 확인)');
  var r = readRow(sheet, reqRowNum);
  var curNo = String(r[COL.DOC_NO] || '');
  if (curNo !== oldNo && curNo !== newNo) return skip('품의번호 불일치: 시트=' + curNo);
  if (String(r[COL.DOC_TYPE] || '') !== 'REQ') return skip('REQ 행이 아님: ' + r[COL.DOC_TYPE]);
  if (String(r[COL.STATUS] || '') === '폐기') return skip('폐기된 문서');
  if (isDocNoDuplicated(newNo, reqToken)) return skip('정정 번호 ' + newNo + ' 를 다른 문서가 이미 사용 중');

  // ── 검증: PRC 행 (부모 연결·같은 FINAL 폴더) ──
  var prcRowNum = findRowNumByToken(sheet, prcToken);
  if (prcRowNum < 0) return skip('PRC 행 없음 (token 확인)');
  var pr = readRow(sheet, prcRowNum);
  if (String(pr[COL.DOC_NO] || '') !== poNo) return skip('PO번호 불일치: 시트=' + pr[COL.DOC_NO]);
  if (String(pr[COL.PARENT_DOC_ID] || '') !== reqToken) return skip('PRC의 부모 REQ가 다름');
  var folderId = String(r[COL.DRIVE_ID] || '');
  if (!folderId || String(pr[COL.DRIVE_ID] || '') !== folderId) return skip('REQ/PRC 폴더 ID가 다름 (FINAL 통합 전?)');

  // ── 검증: 검수보고서 (원 REQ token 으로 조인 — 문서번호 = {품의번호}-{회차}) ──
  var inspSheet = ss.getSheetByName(CONFIG.INSP_SHEET_NAME);
  var inspRows  = (inspSheet && inspSheet.getLastRow() >= 2) ? inspSheet.getDataRange().getValues() : [];
  for (var i = 1; i < inspRows.length; i++) {
    var ir = inspRows[i];
    if (String(ir[INSP_COL.REQ_TOKEN] || '') !== reqToken) continue;
    var seq    = parseInt(ir[INSP_COL.SEQ]) || 0;
    var inspNo = String(ir[INSP_COL.DOC_NO] || '');
    var reqNo  = String(ir[INSP_COL.REQ_NO] || '');
    var newInspNo = newNo + '-' + _pad2(seq);
    var oldInspNo = oldNo + '-' + _pad2(seq);
    if ((inspNo !== oldInspNo && inspNo !== newInspNo) || (reqNo !== oldNo && reqNo !== newNo)) {
      return skip('검수보고서 번호 불일치: ' + inspNo + ' / reqNo=' + reqNo);
    }
    p.insp.push({ rowNum: i + 1, token: String(ir[INSP_COL.TOKEN] || ''), oldInspNo: oldInspNo, newInspNo: newInspNo,
                  needWrite: inspNo !== newInspNo || reqNo !== newNo, driveId: String(ir[INSP_COL.DRIVE_ID] || '') });
  }

  // ── 검증: 주문서목록 (PRC token 조인, reqNo 가 옛 번호인 행만 — 빈 칸은 그대로 둔다) ──
  var poSheet = ss.getSheetByName(CONFIG.PO_SHEET_NAME);
  var poRows  = (poSheet && poSheet.getLastRow() >= 2) ? poSheet.getDataRange().getValues() : [];
  var poFix = [];
  for (var k = 1; k < poRows.length; k++) {
    if (String(poRows[k][PO_COL.PRC_TOKEN] || '') !== prcToken) continue;
    if (String(poRows[k][PO_COL.REQ_NO] || '') === oldNo) poFix.push(k + 1);
  }

  // Y열: QUO_{옛 번호}.pdf 항목 이름
  var list = parseAttachments(r[COL.ATTACH_LIST]);
  var quoteRenamed = false;
  list.forEach(function (f) {
    if (f && f.name === quotePdfFileName(oldNo)) { f.name = quotePdfFileName(newNo); quoteRenamed = true; }
  });

  p.ok = true;
  p.folderId  = folderId;
  p.issueDate = r[COL.ISSUE_DATE];

  // ── 쓰기 ──
  if (curNo !== newNo) {
    p.changed.push('품의서목록 ' + reqRowNum + '행 B열 ' + oldNo + ' → ' + newNo);
    if (!dryRun) sheet.getRange(reqRowNum, COL.DOC_NO + 1).setValue(newNo);
  }
  if (quoteRenamed) {
    p.changed.push('품의서목록 ' + reqRowNum + '행 Y열 ' + quotePdfFileName(oldNo) + ' → ' + quotePdfFileName(newNo));
    if (!dryRun) sheet.getRange(reqRowNum, COL.ATTACH_LIST + 1).setValue(JSON.stringify(list));
  }
  p.insp.forEach(function (x) {
    if (!x.needWrite) return;
    p.changed.push('검수보고서목록 ' + x.rowNum + '행 ' + x.oldInspNo + ' → ' + x.newInspNo + ' (reqNo 포함)');
    if (!dryRun) {
      inspSheet.getRange(x.rowNum, INSP_COL.DOC_NO + 1).setValue(x.newInspNo);
      inspSheet.getRange(x.rowNum, INSP_COL.REQ_NO + 1).setValue(newNo);
    }
  });
  poFix.forEach(function (rowNum) {
    p.changed.push('주문서목록 ' + rowNum + '행 reqNo ' + oldNo + ' → ' + newNo);
    if (!dryRun) poSheet.getRange(rowNum, PO_COL.REQ_NO + 1).setValue(newNo);
  });

  say(head + (p.changed.length ? '시트 ' + p.changed.length + '곳' : '시트 이미 정정됨') +
      (p.changed.length ? '\n    · ' + p.changed.join('\n    · ') : ''));
  return p;
}

/** 2단계 — FINAL(PO) 폴더 파일 · STAGING/INSP 폴더 이름 · PDF 재생성 */
function _tmpDocNoFixDrive_(p, dryRun, say) {
  var head = '[' + p.oldNo + ' → ' + p.newNo + '] ';
  var did = function (s) { p.driveChanged.push(s); };

  var folder;
  try { folder = DriveApp.getFolderById(p.folderId); }
  catch (e) { say(head + 'Drive FINAL 폴더 접근 실패: ' + e); return; }

  // ── FINAL 폴더 파일 ──
  var oldReqPdfs = [], hasNewReqPdf = false;
  var oldInspPdfs = [];
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (f.isTrashed()) continue;
    var name = f.getName();

    if (name.indexOf('REQ_' + p.oldNo + '_') === 0 && /\.pdf$/.test(name)) { oldReqPdfs.push(f); continue; }
    if (name.indexOf('REQ_' + p.newNo + '_') === 0 && /\.pdf$/.test(name)) { hasNewReqPdf = true; continue; }

    // 요약 txt ({번호}_품의서요약.txt, {번호}_N차_재상신_*.txt) — 이름 + 본문 '품의번호:' 줄
    if (name.indexOf(p.oldNo + '_') === 0 && /\.txt$/.test(name)) {
      var newName = p.newNo + name.slice(p.oldNo.length);
      did(name + ' → ' + newName);
      if (!dryRun) {
        var body = f.getBlob().getDataAsString('UTF-8');
        f.setContent(body.split('품의번호: ' + p.oldNo).join('품의번호: ' + p.newNo));
        f.setName(newName);
      }
      continue;
    }
    if (name === quotePdfFileName(p.oldNo)) {
      did(name + ' → ' + quotePdfFileName(p.newNo));
      if (!dryRun) f.setName(quotePdfFileName(p.newNo));
      continue;
    }
    if (name === 'po_manifest.json') {
      try {
        var m = JSON.parse(f.getBlob().getDataAsString('UTF-8'));
        if (m && m.reference === p.oldNo) {
          did('po_manifest.json reference ' + p.oldNo + ' → ' + p.newNo);
          if (!dryRun) { m.reference = p.newNo; f.setContent(JSON.stringify(m, null, 2)); }
        }
      } catch (e) { say(head + 'po_manifest.json 읽기 실패 (건너뜀): ' + e); }
      continue;
    }
    p.insp.forEach(function (x) {
      if (name.indexOf('INSP_' + x.oldInspNo + '_') === 0) oldInspPdfs.push(name);
    });
  }

  // ── REQ PDF: 새 번호로 생성 → 성공해야만 옛 번호 PDF 휴지통 ──
  if (!hasNewReqPdf) {
    if (dryRun) did('REQ PDF 재생성 예정 (REQ_' + p.newNo + '_*.pdf)');
    else {
      var rq = generatePdfForDocument(p.reqToken);
      if (rq.ok) did('REQ PDF 생성 ' + rq.fileName);
      else { say(head + 'REQ PDF 생성 실패 — 옛 PDF 유지: ' + rq.message); oldReqPdfs = []; }
    }
  }
  oldReqPdfs.forEach(function (f) {
    did('휴지통: ' + f.getName());
    if (!dryRun) f.setTrashed(true);
  });

  // ── PRC PDF: '참조 원본 품의' 표기 갱신 (같은 PO번호의 이전 PRC PDF는 generatePdfForDocument가 정리) ──
  if (dryRun) did('PRC PDF 재생성 예정 (PRC_' + p.poNo + '_*.pdf, 이전 것은 휴지통)');
  else {
    var pc = generatePdfForDocument(p.prcToken);
    if (pc.ok) did('PRC PDF 생성 ' + pc.fileName + ' (이전 PRC PDF 휴지통)');
    else say(head + 'PRC PDF 생성 실패 — 이전 PDF 유지: ' + pc.message);
  }

  // ── STAGING 빈 폴더 이름 (경로 조회 findFolderByPath 와 일치시키기) ──
  var stg = findFolderByPath(p.oldNo, p.issueDate, 'staging');
  if (stg && stg.getId() !== p.folderId) {
    did('STAGING 폴더 ' + p.oldNo + ' → ' + p.newNo);
    if (!dryRun) stg.setName(p.newNo);
  }

  // ── 검수보고서: INSP_ 폴더 이름 + manifest.json 재작성 (PDF 재렌더 입력) ──
  p.insp.forEach(function (x) {
    if (!x.driveId) return;
    var inspFolder;
    try { inspFolder = DriveApp.getFolderById(x.driveId); } catch (e) { say(head + 'INSP 폴더 접근 실패: ' + e); return; }
    if (inspFolder.getName() === 'INSP_' + x.oldInspNo) {
      did('INSP 폴더 INSP_' + x.oldInspNo + ' → INSP_' + x.newInspNo);
      if (!dryRun) inspFolder.setName('INSP_' + x.newInspNo);
    }
    var mIt = inspFolder.getFilesByName('manifest.json'), cur = null;
    while (mIt.hasNext()) {
      var mf = mIt.next();
      if (mf.isTrashed()) continue;
      try { cur = JSON.parse(mf.getBlob().getDataAsString('UTF-8')); } catch (_) { cur = {}; }
    }
    if (cur && (cur.docNo !== x.newInspNo || cur.reqNo !== p.newNo)) {
      did('INSP manifest.json 재작성 (' + x.newInspNo + ')');
      if (!dryRun) {
        var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
        var info = _readInspRowByToken(ss, x.token);
        if (!info || String(info.row[INSP_COL.DOC_NO] || '') !== x.newInspNo) {
          say(head + 'INSP manifest 재작성 건너뜀 — 시트 docNo가 아직 ' + (info ? info.row[INSP_COL.DOC_NO] : '(행 없음)'));
        } else {
          var manifest = _buildInspManifest(ss, info.row);
          var ex = inspFolder.getFilesByName('manifest.json');
          while (ex.hasNext()) { try { ex.next().setTrashed(true); } catch (_) {} }
          inspFolder.createFile(Utilities.newBlob(JSON.stringify(manifest, null, 2), 'application/json', 'manifest.json'));
        }
      }
    }
  });
  if (oldInspPdfs.length) {
    say(head + '  (렌더러 재생성 후 정리 대상) ' + oldInspPdfs.join(', '));
  }

  say(head + 'Drive ' + p.driveChanged.length + '건' +
      (p.driveChanged.length ? '\n    · ' + p.driveChanged.join('\n    · ') : ''));
}
