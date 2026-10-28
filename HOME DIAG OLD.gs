// [홈 로딩 진단] 이전 코드 사본 — 임시 (2026-10-28). 진단이 끝나면 HOME DIAG.gs와 함께 삭제한다.
//  운영 @110(e6fa2e3)의 홈 관련 함수 6개를 이름만 바꿔(diagOld_…_) 그대로 옮겼다(자동 생성, 손대지 말 것).
//  diagHomeTiming() [6]이 같은 운영 데이터로 '이전 코드 결과 = 현재 코드 결과'인지 대조하는 데만 쓴다.
//  호출하는 다른 함수(toDateStr·getDeptByEmailMap·ensurePoSheet_ 등)는 현재 코드를 쓴다
//  — 날짜 변환은 [4]에서 따로 전수 대조한다.

// ── Code.gs diagOld_getHomeDataForClient_ (e6fa2e3) ──
function diagOld_getHomeDataForClient_() {
  try {
    var actor = getActiveUserEmail();
    var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);

    if (sheet.getLastRow() < 2) {
      return {
        ok: true,
        user: { email: actor, isProcurement: isProcurementUser(actor), isAdmin: isAdminUser(actor), isGlobalViewer: isGlobalViewer(actor) },
        myDrafts: [], myPending: [], myDocs: [],
        approvedInbox: [], myPrc: [],
        completedDocs: [],   // [INSP Step 2]
        myInspPending: [], finalDocs: [],   // [INSP Step 5]
        allPending: [], allInspPending: [],   // [결재 메뉴 세분화] 관리자 전사 조회
        prefs: _readUserPrefs(),   // [설정 유지] 별도 왕복 없이 홈 데이터에 편승
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
    var allPending = [];     // [결재 메뉴 세분화] 관리자 전용: 결재자가 누구든 모든 대기 문서

    var isProc  = isProcurementUser(actor);
    var isAdmin = isAdminUser(actor);
    var isGView = isGlobalViewer(actor);   // 전사조회(읽기 전용) — 조회 게이트에만 반영
    var actorLc = String(actor || '').toLowerCase();

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
        currency:   getRowCurrency(r[COL.ITEMS]),
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
      //    [관리자 조회 권한] 관리자는 부서와 무관하게 전사 기안 문서를 조회한다.
      if (drafterEmail === actor) {
        myDocs.push(Object.assign({}, meta, { mine: true }));
      } else if (isAdmin || isGView || sameDeptAsMe(drafterEmail)) {
        myDocs.push(Object.assign({}, meta, { mine: false }));
      }

      // 2) 결재 대기 문서 — 내 건(myPending) + 관리자용 전사 건(allPending)
      if (curIdx >= 0 && curIdx < apprCount &&
          (status.indexOf('결재') >= 0 || status === '검토중' || status === '재상신')) {
        var curApprName   = String(r[reqApprCol(curIdx, 1)] || '');
        var curApprEmail  = String(r[reqApprCol(curIdx, 2)] || '');
        var curApprStatus = String(r[reqApprCol(curIdx, 3)] || '');
        if (curApprStatus === '대기') {
          // stageIdx 0 = 기안자(자기결재) 블록, 1+ = 결재자 N
          var pendMeta = Object.assign({}, meta, {
            stageIdx:     curIdx,
            stageKind:    (curIdx === 0 ? 'drafter' : 'approver'),
            curApprName:  curApprName,
            curApprEmail: curApprEmail,
            apprCount:    apprCount,
          });
          // ⚠ 이메일 비교는 반드시 소문자끼리 — 결재자목록 시트는 사람이 직접 입력하므로
          //   대소문자가 섞일 수 있다. (아래 amParticipant 루프와 INSP 경로도 소문자 비교)
          if (curApprEmail.toLowerCase() === actorLc) myPending.push(pendMeta);
          if (isAdmin) allPending.push(pendMeta);
        }
      }

      // 3) 구매팀: 구매팀 기안 문서 (1차 최종 승인 + 아직 PRC 미생성)
      if (isProc && docType === 'REQ' && status === '최종승인') {
        approvedInbox.push(meta);
      }

      // [INSP Step 5] 최종 완료(통합) 후보 수집: 최종승인 REQ + 최종승인(PRC)
      //  가시성: 관리자·구매팀 전체 / 일반은 본인 관련(기안자 또는 결재 참여자) + 같은 부서 기안 문서
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
        if (isAdmin || isProc || isGView || amParticipant || sameDeptAsMe(drafterEmail)) {
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

    // 결재 대기 — 오래된 순 (클라이언트가 sortByNewest로 재정렬하지만 페이로드는 결정적으로)
    myPending.sort(function(a, b) {
      return (a.submitAt || '').localeCompare(b.submitAt || '');
    });
    allPending.sort(function(a, b) {
      return (a.submitAt || '').localeCompare(b.submitAt || '');
    });

    // [INSP Step 2] 결재 완료 목록 조립
    // - 가시성: 구매팀은 전체, 일반 사용자는 원본 REQ 기안자 본인 건만 (Q-INSP-05)
    // - 검수보고서 그룹은 INSP.gs의 diagOld__getInspGroupsByPrc_()가 1회 read로 제공
    //   (INSP 파일 미배포 환경에서도 홈이 죽지 않도록 typeof 가드)
    var completedDocs = [];
    if (prcFinals.length > 0) {
      var inspGroups = (typeof diagOld__getInspGroupsByPrc_ === 'function')
        ? diagOld__getInspGroupsByPrc_(ss)
        : {};
      for (var c = 0; c < prcFinals.length; c++) {
        var pf  = prcFinals[c];
        var req = reqMap[pf.parentToken] || null;
        var reqDrafter = req ? req.drafter : '';
        var completedMine = (String(reqDrafter).toLowerCase() === String(actor || '').toLowerCase());
        // 가시성: 관리자·구매팀 전체 / 본인 기안 / 같은 부서 기안 (검수보고서 제출도 같은 부서 허용)
        // ⚠ 관리자는 '조회'만 확장한다 — canSubmit(검수보고서 제출권)은 아래에서 그대로 두어
        //   서버 INSP 게이트(본인/같은 부서)와 기준을 어긋나게 하지 않는다.
        if (!isAdmin && !isProc && !isGView && !completedMine && !sameDeptAsMe(reqDrafter)) continue;

        var insps = inspGroups[pf.token] || [];
        // 종결(최종 승인 완료) / 잠금대기(최종 검수가 반려 외 상태로 진행중) 구분
        var inspClosed = false;        // 최종 검수 '최종승인(INSP)' → 완전 종결
        var inspLockedPending = false; // 최종 검수 제출됐으나 결재 진행중 → 추가 제출만 차단
        for (var ic = 0; ic < insps.length; ic++) {
          // 폐기된 회차는 잠금에서 뺀다 — INSP.gs _isInspLocked와 같은 기준 (폐기 후 다시 제출할 수 있어야 함)
          if (insps[ic].isFinal && insps[ic].status !== '반려' && insps[ic].status !== '폐기') {
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
    var inspMenus = (typeof diagOld__getInspMenusForClient_ === 'function')
      ? diagOld__getInspMenusForClient_(ss, actor, isProc, isAdmin, isGView, deptMap)
      : { myInspPending: [], inspFinals: [], allInspPending: [] };

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
        // 회차 구분 표시용 — 중간 회차도 목록에 포함되므로 화면에서 배지로 가른다.
        isFinal: !!m.isFinal, verdict: m.verdict,
        mine: (String(m.drafter || '').toLowerCase() === String(actor || '').toLowerCase()),
      });
    });
    finalDocs.sort(function(a, b) { return (b.submitAt || '').localeCompare(a.submitAt || ''); });

    return {
      ok: true,
      user: {
        email: actor,
        isProcurement: isProc,
        isAdmin: isAdmin,
        isGlobalViewer: isGView,
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
      // [결재 메뉴 세분화] 관리자 전사 결재 현황.
      // 누산 시점에 이미 isAdmin으로 걸렀지만, 삼항을 한 번 더 두어 서버측 권한 백스톱으로 삼는다.
      allPending:     isAdmin ? allPending : [],
      allInspPending: isAdmin ? (inspMenus.allInspPending || []) : [],
      // [관리자 처리 대기 배너] 견적서 PDF · 주문서 · 검수보고서 PDF — 시트 상태가 근거(메일을 놓쳐도 홈에 남는다)
      adminTasks:     isAdmin ? diagOld__collectAdminTasks_(rows) : null,
      // [설정 유지] 마지막 컬럼 필터 설정. 홈은 서버를 원래 1회만 부르므로 여기에 실어
      //   보내면 왕복이 늘지 않고, 화면이 이 응답을 받은 뒤에야 그려지므로 깜빡임도 없다.
      prefs: _readUserPrefs(),
    };
  } catch(err) {
    return { ok: false, message: err.toString() + ' / ' + (err.stack || '') };
  }
}

// ── Code.gs diagOld__collectAdminTasks_ (e6fa2e3) ──
function diagOld__collectAdminTasks_(rows) {
  var t = { quotePdf: [], po: [], insp: [] };
  try { t.quotePdf = listPendingQuotePdf(rows); }
  catch(e) { Logger.log('[AdminTasks] 견적서 목록 실패: ' + e); }
  try { t.po = (typeof diagOld_listPoPending_ === 'function') ? diagOld_listPoPending_() : []; }
  catch(e) { Logger.log('[AdminTasks] 주문서 목록 실패: ' + e); }
  try {
    t.insp = ((typeof diagOld_listInspAwaitingPdf_ === 'function') ? diagOld_listInspAwaitingPdf_() : [])
      .filter(function(x) { return ADMIN_TASK_EXCLUDE_DOCNOS.indexOf(String(x.docNo || '')) < 0; });
  } catch(e) { Logger.log('[AdminTasks] 검수보고서 목록 실패: ' + e); }
  return t;
}

// ── INSP.gs diagOld__getInspGroupsByPrc_ (e6fa2e3) ──
function diagOld__getInspGroupsByPrc_(ss) {
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
        // 실제 제출자(결재란 0번) — 홈에서 '재상신/폐기' 버튼을 제출자 본인에게만 보이기 위함 (서버가 다시 검증)
        submitter:       String(r[inspApprCol(0, 2)] || '').trim().toLowerCase(),
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
    console.error('[INSP] diagOld__getInspGroupsByPrc_ 실패: ' + e.toString());
  }
  return groups;
}

// ── INSP.gs diagOld__getInspMenusForClient_ (e6fa2e3) ──
function diagOld__getInspMenusForClient_(ss, actor, isProc, isAdmin, isGView, deptMap) {
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

      // 완료된 검수보고서: 최종승인(INSP)이면 회차와 무관하게 전부 수집한다.
      //   ⚠ 과거에는 IS_FINAL='Y'인 최종 회차만 담아, 여러 회차로 나눠 검수한 건의
      //     중간 회차가 '최종 완료' 목록에서 통째로 사라졌다(예: TH-AP-26-044는
      //     -01/-02 두 건을 제출했는데 -02만 보임). 중간 회차도 결재를 마치고 PDF까지
      //     생성된 완결 문서이므로 목록에 나와야 한다.
      //     최종/중간 구분은 meta.isFinal로 내려보내 화면에서 배지로 표시한다.
      // 가시성: 관리자 / 구매팀 / 결재 참여자 / 본인 기안 / 같은 부서(부서 공유)
      if (status === '최종승인(INSP)') {
        if (isAdmin || isProc || isGView || amParticipant ||
            meta.drafter.toLowerCase() === actor ||
            _isSameDeptEmail(ss, meta.drafter, actor, deptMap)) {
          out.inspFinals.push(meta);
        }
      }
    }

    out.myInspPending.sort(function (a, b) { return (a.submitAt || '').localeCompare(b.submitAt || ''); });
    out.allInspPending.sort(function (a, b) { return (a.submitAt || '').localeCompare(b.submitAt || ''); });
    out.inspFinals.sort(function (a, b) { return (b.submitAt || '').localeCompare(a.submitAt || ''); });
  } catch (e) {
    console.error('[INSP] diagOld__getInspMenusForClient_ 실패: ' + e.toString());
  }
  return out;
}

// ── INSP.gs diagOld_listInspAwaitingPdf_ (e6fa2e3) ──
function diagOld_listInspAwaitingPdf_() {
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

// ── PO.gs diagOld_listPoPending_ (e6fa2e3) ──
function diagOld_listPoPending_() {
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
