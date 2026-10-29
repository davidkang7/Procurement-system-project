/**
 * 시스템로그 지연 기록 · 신규 로그 6종 · INSP 폐기 · 락 해제 실패([W]) — 로컬 검증 (배포 없이 실제 .gs 함수 호출)
 * (memory: gas-local-test-harness)
 *
 * 방식: 같은 시나리오를 이전 코드(BASE_REF, 기본 abd290f = 이 기능 직전 커밋)와 작업 트리(수정본)에서 각각 돌려
 *       "반환값 · 시트 상태 · 발송 메일 · 시스템로그 행"이 같은지 비교한다.
 *       달라야 하는 것(신규 기능·보안 수정)은 따로 단언한다.
 *       [W]는 수정본끼리 비교한다 — 락 해제가 실패한 실행이 정상 해제된 실행과 결과가 같아야 한다.
 *
 * 실행: node tests/audit_insp_discard_test.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execSync } = require('child_process');

const SRC = path.join(__dirname, '..');
// WT_DIR: 뮤테이션 검사용 — 일부러 버그를 심은 사본 폴더를 지정하면 그 코드를 '수정본'으로 돌린다
const WT_DIR = process.env.WT_DIR || SRC;
const readWT   = f => fs.readFileSync(path.join(WT_DIR, f), 'utf8');
// 비교 기준(이전 코드) — 이 기능이 들어가기 직전 커밋. 커밋 후에도 '이전과 같은가'를 다시 볼 수 있도록 고정한다.
//  다른 기준으로 보려면 BASE_REF=<커밋> (예: BASE_REF=HEAD → 아직 커밋 안 한 다음 수정본을 직전 커밋과 비교)
const BASE_REF = process.env.BASE_REF || 'abd290f';
const readHEAD = f => execSync('git show ' + BASE_REF + ':"' + f + '"', { cwd: SRC, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });

// ── 고정 시계 (두 환경 공용) ──────────────────────────────────────
const RealDate = Date;
const CLOCK = { now: RealDate.parse('2026-10-14T10:00:00+09:00') };
class FakeDate extends RealDate {
  constructor(...a) { if (a.length === 0) super(CLOCK.now); else super(...a); }
  static now() { return CLOCK.now; }
}
const MIN = 60 * 1000, HOUR = 60 * MIN;

// ── 메모리 시트 ──────────────────────────────────────────────────
class FakeRange {
  constructor(sheet, row, col, nRows, nCols) { Object.assign(this, { sheet, row, col, nRows, nCols }); }
  getValues() {
    const out = [];
    for (let r = 0; r < this.nRows; r++) {
      const row = [];
      for (let c = 0; c < this.nCols; c++) row.push(this.sheet._get(this.row + r, this.col + c));
      out.push(row);
    }
    return out;
  }
  setValues(vals) { this.sheet._guard(); vals.forEach((rv, r) => rv.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v))); return this; }
  setValue(v) { this.sheet._guard(); this.sheet._set(this.row, this.col, v); return this; }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  setFontWeight() { return this; }
  setBackground() { return this; }
  setDataValidation() { return this; }
  createTextFinder(text) {
    const self = this; const opts = { entire: false };
    const finder = {
      matchEntireCell(b) { opts.entire = b; return finder; },
      matchCase() { return finder; },
      findNext() {
        for (let r = 0; r < self.nRows; r++) for (let c = 0; c < self.nCols; c++) {
          const v = String(self.sheet._get(self.row + r, self.col + c) ?? '');
          if (opts.entire ? v === String(text) : v.indexOf(String(text)) >= 0) return new FakeRange(self.sheet, self.row + r, self.col + c, 1, 1);
        }
        return null;
      },
    };
    return finder;
  }
}
class FakeSheet {
  constructor(name) { this.name = name; this.rows = []; this.failWrites = false; this.appendCount = 0; }
  _guard() { if (this.failWrites) throw new Error('시트 쓰기 장애 시뮬레이션 (' + this.name + ')'); }
  _get(r, c) { const row = this.rows[r - 1]; return row ? (row[c - 1] ?? '') : ''; }
  _set(r, c, v) { while (this.rows.length < r) this.rows.push([]); const row = this.rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
  getName() { return this.name; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return this.rows.reduce((m, r) => Math.max(m, r.length), 0); }
  appendRow(vals) { this._guard(); this.appendCount++; this.rows.push(vals.slice()); return this; }
  getRange(row, col, nRows = 1, nCols = 1) { return new FakeRange(this, row, col, nRows, nCols); }
  getDataRange() { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
  setFrozenRows() { return this; }
  setColumnWidth() { return this; }
  createTextFinder(t) { return this.getDataRange().createTextFinder(t); }
}
class FakeSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(n) { return this.sheets[n] || null; }
  insertSheet(n) { this.sheets[n] = new FakeSheet(n); return this.sheets[n]; }
  getSheets() { return Object.values(this.sheets); }
}

// ── 사람 ─────────────────────────────────────────────────────────
const ADMIN = 'davidkang@inlct.com';
const BUYER1 = 'buyer1@inlct.com', BUYER2 = 'buyer2@inlct.com';
const KIM = 'kim@inlct.com', PARK = 'park@inlct.com', LEE = 'lee@inlct.com', CHOI = 'choi@inlct.com';
const NAMES = { [KIM]: '김기안', [PARK]: '박동료', [LEE]: '이결재', [CHOI]: '최결재', [ADMIN]: '강관리', [BUYER1]: '구매1', [BUYER2]: '구매2' };
const DEPT = { [KIM]: '개발팀', [PARK]: '개발팀', [LEE]: '경영팀', [CHOI]: '경영팀', [ADMIN]: '경영총괄', [BUYER1]: '구매팀', [BUYER2]: '구매팀' };

// ── 환경 ─────────────────────────────────────────────────────────
function makeEnv(label, read) {
  const state = {
    label, active: ADMIN, mails: [], adminErrors: [], consoleErrors: [], consoleWarns: [],
    props: new Map(), cache: new Map(), folders: {}, uuid: 0,
    lock: { scriptWait: 0, userTry: 0, userHeld: false, releaseFails: 0 },
    flags: { uploadFail: false, photoFail: false, folderAccessFail: false, propsSetFail: false, releaseFail: false, releaseDelayMs: 0 },
  };
  const ss = new FakeSpreadsheet();

  function folderRec(id, name) {
    if (!state.folders[id]) state.folders[id] = { id, name: name || id, trashed: false, files: [] };
    return state.folders[id];
  }
  function wrapFolder(f) {
    return {
      getId: () => f.id, getName: () => f.name,
      setTrashed: t => { f.trashed = !!t; }, isTrashed: () => f.trashed,
      createFile: (a, b, c) => { const nf = { id: 'file-' + (++state.uuid), name: (a && a.name) || a }; f.files.push(nf); return { getId: () => nf.id, getName: () => nf.name, setTrashed() {} }; },
      getFiles: () => ({ hasNext: () => false, next: () => null }),
      getFilesByName: () => ({ hasNext: () => false, next: () => null }),
      getParents: () => ({ hasNext: () => false, next: () => null }),
    };
  }
  const propsObj = {
    getProperty: k => (state.props.has(k) ? state.props.get(k) : null),
    setProperty: (k, v) => {
      if (state.flags.propsSetFail && String(k).indexOf('AUDIT_BUF_') === 0) throw new Error('Properties 할당량 초과 시뮬레이션');
      state.props.set(k, String(v)); return propsObj;
    },
    deleteProperty: k => { state.props.delete(k); return propsObj; },
    getKeys: () => Array.from(state.props.keys()),
    getProperties: () => Object.fromEntries(state.props),
  };
  const cacheObj = {
    get: k => { const e = state.cache.get(k); return e && e.exp > CLOCK.now ? e.v : null; },
    put: (k, v, ttl) => { state.cache.set(k, { v: String(v), exp: CLOCK.now + (ttl || 600) * 1000 }); },
    remove: k => { state.cache.delete(k); },
  };
  const quiet = { log() {}, info() {}, warn: (...a) => state.consoleWarns.push(a.map(String).join(' ')), error: (...a) => state.consoleErrors.push(a.map(String).join(' ')) };

  const sandbox = {
    console: quiet,
    Date: FakeDate,
    Logger: { log() {} },
    Session: {
      getActiveUser: () => ({ getEmail: () => state.active }),
      getEffectiveUser: () => ({ getEmail: () => ADMIN }),   // 배포 계정
      getScriptTimeZone: () => 'Asia/Seoul',
    },
    SpreadsheetApp: { openById: () => ss, flush() {} },
    DriveApp: {
      getFolderById: id => {
        if (state.flags.folderAccessFail) throw new Error('폴더 접근 실패 시뮬레이션');
        const f = state.folders[id]; if (!f) throw new Error('No item with the given ID could be found: ' + id);
        return wrapFolder(f);
      },
      getFileById: id => ({ setTrashed() {}, getBlob: () => ({ getContentType: () => 'image/jpeg', getBytes: () => [] }) }),
    },
    Utilities: {
      getUuid: () => 'uuid-' + (++state.uuid),
      newBlob: (content, type, name) => ({ content, type, name }),
      formatDate: (d, tz, fmt) => {
        const p = n => String(n).padStart(2, '0'); const t = new RealDate(d.getTime ? d.getTime() : d);
        const ymd = `${t.getFullYear()}-${p(t.getMonth() + 1)}-${p(t.getDate())}`;
        if (fmt === 'yyyyMMdd_HHmmss') return `${t.getFullYear()}${p(t.getMonth() + 1)}${p(t.getDate())}_${p(t.getHours())}${p(t.getMinutes())}${p(t.getSeconds())}`;
        return fmt.indexOf('HH') >= 0 ? `${ymd} ${p(t.getHours())}:${p(t.getMinutes())}` : ymd;
      },
      sleep() {}, base64Decode: () => [], base64Encode: () => '',
    },
    LockService: {
      // releaseFail: 2026-10-15 실제 장애 재현 — 해제가 releaseDelayMs만큼 끌다가 Google 서버 오류를 던진다
      getScriptLock: () => ({ waitLock() { state.lock.scriptWait++; }, tryLock() { state.lock.scriptWait++; return true; }, releaseLock() {
        if (!state.flags.releaseFail) return;
        state.lock.releaseFails++; CLOCK.now += state.flags.releaseDelayMs;
        throw new Error("We're sorry, a server error occurred. Please wait a bit and try again.");
      } }),
      getUserLock: () => ({
        tryLock() { state.lock.userTry++; if (state.lock.userHeld) return false; state.lock.userHeld = true; return true; },
        releaseLock() { state.lock.userHeld = false; },
      }),
    },
    PropertiesService: { getScriptProperties: () => propsObj },
    CacheService: { getScriptCache: () => cacheObj },
    GmailApp: { sendEmail: (to, subj) => { state.mails.push({ kind: 'gmail', to, subj }); } },
    MailApp: { sendEmail() {} },
    HtmlService: { createTemplateFromFile: () => ({ evaluate: () => ({ setTitle() { return this; }, setXFrameOptionsMode() { return this; } }) }), createHtmlOutput: () => ({ setTitle() { return this; }, setXFrameOptionsMode() { return this; } }), XFrameOptionsMode: {} },
    ContentService: { createTextOutput: () => ({ setMimeType() { return this; } }), MimeType: {} },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => ({ timeBased: () => ({ everyMinutes: () => ({ create() {} }) }) }) },
    MimeType: { PDF: 'application/pdf', PLAIN_TEXT: 'text/plain' },
  };
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  ['Code.gs', 'INSP.gs', 'PO.gs'].forEach(f => vm.runInContext(read(f), sandbox, { filename: label + ':' + f }));

  // 저수준 헬퍼 대체 — 두 환경에 똑같이 (메일·Drive 업로드·외부 시트)
  sandbox.__state = state;
  sandbox.__folderRec = folderRec;
  sandbox.__wrapFolder = wrapFolder;
  vm.runInContext(`
    sendApprovalEmailWithRetry   = function (ap, data, token, rowNum, idx) { __state.mails.push({ kind: 'approval', to: ap.email, idx: idx }); return true; };
    sendRejectionNoticeWithRetry = function (to) { __state.mails.push({ kind: 'reject', to: to }); return true; };
    sendEmailWithRetry           = function (to, subj) { __state.mails.push({ kind: 'generic', to: to, subj: subj }); return true; };
    notifyAdminError             = function (m) { __state.adminErrors.push(String(m)); };
    getOrCreateFolder            = function (name) { return __wrapFolder(__folderRec('F_' + name, name)); };
    saveSummaryToDrive           = function () {};
    createFileInFolder           = function () {};
    saveAttachmentsToFolder      = function (folder, atts) { if (__state.flags.uploadFail) throw new Error('업로드 실패 시뮬레이션'); return []; };
    _enqueueQuotePdfIfNeeded_    = function () {};
    _notifyQuoteMissingIfNeeded_ = function () {};
    _prepareInspPdfHandoff       = function () { return { ok: true }; };
    _saveInspPhotos              = function () { if (__state.flags.photoFail) throw new Error('사진 저장 실패 시뮬레이션'); return []; };
    _sendInspApprovalEmail       = function (ap) { __state.mails.push({ kind: 'insp_approval', to: ap.email }); };
    _sendInspRejectEmail         = function (to) { __state.mails.push({ kind: 'insp_reject', to: to }); };
    _sendInspCompletionEmail     = function (to) { __state.mails.push({ kind: 'insp_completion', to: to }); };
    _sendApproverChangedNotice   = function (to) { __state.mails.push({ kind: 'approver_changed', to: to }); return true; };
    getDeptByEmailMap            = function () { return __DEPT; };
    _readUserPrefs               = function () { return {}; };
    _collectAdminTasks_          = function () { return { quotePdf: [], po: [], insp: [] }; };
    _isSameDeptEmail             = function (ss, a, b) { var d = __DEPT; var x = d[String(a||'').toLowerCase()], y = d[String(b||'').toLowerCase()]; return !!x && x === y; };
    getApproverListForClient     = function () { return { ok: true, approvers: __APPROVERS }; };
    getCurrentUserForClient      = function () { var e = Session.getActiveUser().getEmail(); return { ok: true, user: { email: e, name: __NAMES[e] || e, dept: __DEPT[e] || '', rank: '' } }; };
    _processPdfOnlyJob           = function (job) { return __state.jobResult || { ok: true }; };
  `, sandbox);
  sandbox.__DEPT = DEPT;
  sandbox.__NAMES = NAMES;
  sandbox.__APPROVERS = [KIM, PARK, LEE, CHOI].map(e => ({ name: NAMES[e], email: e, dept: DEPT[e], rank: '', active: true }));
  sandbox.CONFIG.PROCUREMENT_TEAM_EMAILS = [ADMIN, BUYER1, BUYER2];
  sandbox.CONFIG.ADMIN_EMAILS = [ADMIN];
  sandbox.CONFIG.GLOBAL_VIEW_EMAILS = [];

  const env = {
    label, sb: sandbox, ss, state,
    hasNew: typeof sandbox.writeAuditLogDeferred_ === 'function',
    as(user) { state.active = user; return env; },
    call(fn, ...args) { return sandbox[fn].apply(sandbox, args); },
    main: () => ss.getSheetByName(sandbox.CONFIG.SHEET_NAME),
    insp: () => ss.getSheetByName(sandbox.CONFIG.INSP_SHEET_NAME),
    auditRows: () => { const s = ss.getSheetByName(sandbox.CONFIG.AUDIT_SHEET_NAME); return s ? s.rows.slice(1) : []; },
    auditSheet: () => ss.getSheetByName(sandbox.CONFIG.AUDIT_SHEET_NAME) || sandbox.ensureAuditLogSheet(),
    bufKeys: () => Array.from(state.props.keys()).filter(k => k.indexOf('AUDIT_BUF_') === 0).sort(),
    heartbeat: ms => { state.props.set('AUDIT_FLUSH_HEARTBEAT', String(ms == null ? CLOCK.now : ms)); },
    flush() { const prev = state.active; state.active = ADMIN; try { return sandbox.flushAuditLogBuffer_(); } finally { state.active = prev; } },   // 트리거 = David 계정
    events: () => env.auditRows().map(r => r[4]),
  };
  return env;
}

// ── 픽스처 ───────────────────────────────────────────────────────
function seed(env) {
  const sb = env.sb, C = sb.COL;
  const main = sb.getOrCreateSheet(env.ss, sb.CONFIG.SHEET_NAME);
  sb.ensureHeaders(main);
  const ap = (label, email, status, at) => ({ label, name: NAMES[email], email, status: status || '대기', at: at || '' });
  function req(o) {
    const row = new Array(C.APPR_START).fill('');
    row[C.SUBMIT_AT] = new FakeDate(o.submitAt || '2026-10-01T09:00:00+09:00');
    row[C.DOC_NO] = o.docNo; row[C.ISSUE_DATE] = '2026-10-01'; row[C.DRAFTER] = NAMES[o.drafter || KIM]; row[C.DEPT] = DEPT[o.drafter || KIM];
    row[C.SUBJECT] = o.subject || '시험 품목'; row[C.VENDOR_NAME] = '라인테크'; row[C.ITEMS] = '1. 볼트 | M8 | 10개 | 1,000 KRW'; row[C.TOTAL_AMT] = 10000;
    row[C.STATUS] = o.status; row[C.TOKEN] = o.token; row[C.DRIVE_ID] = o.driveId || '';
    row[C.APPR_COUNT] = o.approvers.length; row[C.APPR_IDX] = o.apprIdx || 0; row[C.RESUB_COUNT] = 0; row[C.REJECT_LOG] = o.rejectLog || '';
    row[C.ATTACH_LIST] = '[]'; row[C.MOVE_STATUS] = 'STAGING'; row[C.DOC_TYPE] = o.docType || 'REQ'; row[C.PARENT_DOC_ID] = o.parent || '';
    row[C.CLAIMED_BY] = o.claimedBy || ''; row[C.CLAIMED_AT] = o.claimedAt ? new FakeDate(o.claimedAt) : '';
    o.approvers.forEach(a => row.push(a.label, a.name, a.email, a.status, a.at, ''));
    main.appendRow(row);
    if (o.driveId) env.state.folders[o.driveId] = { id: o.driveId, name: o.docNo, trashed: false, files: [] };
  }
  req({ docNo: 'TH-AP-26-001', token: 'tok-req1', status: '검토중', approvers: [ap('기안자', KIM), ap('결재자 1', LEE)], driveId: 'F_TH-AP-26-001', submitAt: '2026-10-01T09:00:00+09:00' });
  req({ docNo: 'TH-AP-26-002', token: 'tok-req2', status: '최종승인', apprIdx: 2, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE, '승인')], submitAt: '2026-10-02T09:00:00+09:00' });
  req({ docNo: 'TH-AP-26-003', token: 'tok-req3', status: '최종승인', apprIdx: 2, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE, '승인')], claimedBy: BUYER2, claimedAt: CLOCK.now - 1 * HOUR, submitAt: '2026-10-03T09:00:00+09:00' });
  req({ docNo: 'TH-AP-26-004', token: 'tok-req4', status: '최종승인', apprIdx: 2, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE, '승인')], claimedBy: BUYER2, claimedAt: CLOCK.now - 30 * HOUR, submitAt: '2026-10-04T09:00:00+09:00' });
  req({ docNo: 'TH-AP-26-005', token: 'tok-req5', status: '반려', apprIdx: 1, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE, '반려')], driveId: 'F_TH-AP-26-005', rejectLog: '[1차 반려] 이결재 / 2026-10-05 / 단가 확인', submitAt: '2026-10-05T09:00:00+09:00' });
  req({ docNo: 'TH-AP-26-006', token: 'tok-req6', status: 'PRC생성됨', apprIdx: 2, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE, '승인')], submitAt: '2026-09-20T09:00:00+09:00' });
  req({ docNo: 'TO-PO-26-901', token: 'tok-prc1', status: '최종승인(PRC)', docType: 'PRC', parent: 'tok-req6', drafter: BUYER1, apprIdx: 2, approvers: [ap('기안자', BUYER1, '승인'), ap('결재자 1', LEE, '승인')], submitAt: '2026-09-21T09:00:00+09:00' });
  req({ docNo: 'TH-AP-26-007', token: 'tok-req7', status: '최종승인', apprIdx: 2, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE, '승인')], claimedBy: BUYER1, claimedAt: CLOCK.now - 1 * HOUR, submitAt: '2026-10-07T09:00:00+09:00' });
  req({ docNo: 'TH-AP-26-008', token: 'tok-req8', status: '결재중 (결재자 1)', apprIdx: 1, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE)], submitAt: '2026-10-08T09:00:00+09:00' });
  req({ docNo: 'TH-AP-26-009', token: 'tok-req9', status: 'PRC생성됨', apprIdx: 2, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE, '승인')], submitAt: '2026-09-22T09:00:00+09:00' });
  req({ docNo: 'TO-PO-26-902', token: 'tok-prc2', status: '최종승인(PRC)', docType: 'PRC', parent: 'tok-req9', drafter: BUYER1, apprIdx: 2, approvers: [ap('기안자', BUYER1, '승인'), ap('결재자 1', LEE, '승인')], submitAt: '2026-09-23T09:00:00+09:00' });

  // 검수보고서 — TO-PO-26-901 (원본 품의 TH-AP-26-006, 기안자 김기안)
  const isheet = sb.ensureInspSheet();
  const I = sb.INSP_COL;
  function insp(o) {
    const row = new Array(sb.INSP_TOTAL_COLS).fill('');
    row[I.SUBMIT_AT] = new FakeDate(o.submitAt); row[I.DOC_NO] = 'TH-AP-26-006-' + String(o.seq).padStart(2, '0'); row[I.SEQ] = o.seq;
    row[I.PRC_TOKEN] = 'tok-prc1'; row[I.REQ_TOKEN] = 'tok-req6'; row[I.PO_NO] = 'TO-PO-26-901'; row[I.REQ_NO] = 'TH-AP-26-006';
    row[I.SUBJECT] = '시험 품목'; row[I.VENDOR_NAME] = '라인테크'; row[I.DRAFTER] = KIM; row[I.DRAFTER_NAME] = NAMES[KIM]; row[I.DEPT] = '개발팀';
    row[I.ISSUE_DATE] = '2026-09-20'; row[I.RECEIVED_DATE] = '2026-10-10'; row[I.RECEIVED_NOTE] = o.seq + '차 입고'; row[I.VERDICT] = '합격';
    row[I.IS_FINAL] = o.isFinal ? 'Y' : ''; row[I.STATUS] = o.status; row[I.TOKEN] = o.token; row[I.DRIVE_ID] = o.driveId || '';
    row[I.APPR_COUNT] = o.approvers.length; row[I.APPR_IDX] = o.apprIdx || 0; row[I.RESUB_COUNT] = 0; row[I.REJECT_LOG] = o.rejectLog || '';
    row[I.PHOTO_LIST] = '[]'; row[I.MOVE_STATUS] = o.moveStatus || 'STAGING';
    o.approvers.forEach((a, k) => { const b = I.APPR_START + k * I.APPR_COLS; row[b] = a.label; row[b + 1] = a.name; row[b + 2] = a.email; row[b + 3] = a.status; row[b + 4] = a.at; row[b + 5] = ''; });
    isheet.appendRow(row);
    if (o.driveId) env.state.folders[o.driveId] = { id: o.driveId, name: o.folderName || ('INSP_' + row[I.DOC_NO]), trashed: false, files: [] };
  }
  insp({ seq: 1, token: 'tok-i1', status: '최종승인(INSP)', isFinal: false, moveStatus: 'FINAL', submitAt: '2026-10-09T09:00:00+09:00', approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE, '승인')], apprIdx: 2 });
  insp({ seq: 2, token: 'tok-i2', status: '반려', isFinal: true, driveId: 'F_INSP_-02', submitAt: '2026-10-10T09:00:00+09:00', approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE, '반려')], apprIdx: 1, rejectLog: '[1차 반려] 이결재 / 2026-10-10 / 사진 누락' });
  insp({ seq: 3, token: 'tok-i3', status: '결재중 (이결재)', isFinal: true, driveId: 'F_INSP_-03', submitAt: '2026-10-11T09:00:00+09:00', approvers: [ap('기안자', PARK, '승인'), ap('결재자 1', LEE)], apprIdx: 1 });
  return env;
}

// ── 단언 ─────────────────────────────────────────────────────────
let pass = 0, fail = 0;
const failures = [];
function check(name, fn) {
  try {
    const r = fn();
    if (r === true) { pass++; console.log('  ✓ ' + name); }
    else { fail++; failures.push(name); console.log('  ✗ ' + name + '  → ' + (typeof r === 'string' ? r : JSON.stringify(r))); }
  } catch (e) { fail++; failures.push(name); console.log('  ✗ ' + name + '  → 예외: ' + (e && e.stack || e)); }
}
const norm = x => JSON.parse(JSON.stringify(x === undefined ? null : x).replace(/uuid-\d+/g, 'UUID'));
const same = (a, b) => JSON.stringify(norm(a)) === JSON.stringify(norm(b)) || ('\n      HEAD=' + JSON.stringify(norm(a)) + '\n      NEW =' + JSON.stringify(norm(b)));
function pair(fn, opts) {
  opts = opts || {};
  const h = seed(makeEnv('HEAD', readHEAD)), n = seed(makeEnv('NEW', readWT));
  if (opts.heartbeat !== false) n.heartbeat();   // 운영 상태: 트리거가 매분 하트비트를 갱신 중
  const rh = fn(h), rn = fn(n);
  return { h, n, rh, rn };
}
const sheetRows = env => ({ main: env.main().rows, insp: env.insp().rows });
const rowOf = (env, sheet, token, col) => sheet.rows.find(r => r[col] === token);
const mainRow = (env, token) => env.main().rows.find(r => r[env.sb.COL.TOKEN] === token);
const mainRowByDoc = (env, docNo) => env.main().rows.find(r => r[env.sb.COL.DOC_NO] === docNo);
const inspRowByDoc = (env, docNo) => env.insp().rows.find(r => r[env.sb.INSP_COL.DOC_NO] === docNo);
const AC = { LOG_ID: 0, TS: 1, ACTOR: 2, ROLE: 3, EV: 4, DOCNO: 5, TOKEN: 6, TYPE: 7, TARGET: 8, REASON: 9, PAYLOAD: 10 };
const auditNoId = rows => rows.map(r => r.slice(1).map(v => (v instanceof RealDate ? v.getTime() : v)));

// =================================================================
console.log('\n[A] 지연 기록 기반 — 쓰기·반영·안전장치');
// =================================================================
{
  const n = seed(makeEnv('NEW', readWT));
  check('A1 하트비트 없음 → 즉시 기록 (버퍼 미사용)', () => {
    n.as(KIM); n.call('writeAuditLogDeferred_', { eventType: 'DOC_SUBMIT', docNo: 'X-1' });
    return (n.auditRows().length === 1 && n.bufKeys().length === 0) || [n.auditRows().length, n.bufKeys().length];
  });
  check('A2 빈 버퍼 반영: 하트비트만 갱신, 락·시트 접근 없음', () => {
    const before = n.auditSheet().appendCount, lt = n.state.lock.userTry, sw = n.state.lock.scriptWait;
    n.state.props.delete('AUDIT_FLUSH_HEARTBEAT');
    const r = n.flush();
    return (r.flushed === 0 && n.state.props.get('AUDIT_FLUSH_HEARTBEAT') === String(CLOCK.now) &&
            n.auditSheet().appendCount === before && n.state.lock.userTry === lt && n.state.lock.scriptWait === sw) || [r, n.state.lock];
  });
  let wroteAt;
  check('A3 하트비트 정상 → 버퍼에만 저장, 시트는 그대로', () => {
    CLOCK.now += 5000; wroteAt = CLOCK.now;
    const before = n.auditRows().length;
    n.as(KIM); const ok = n.call('writeAuditLogDeferred_', { eventType: 'DOC_APPROVE', docNo: 'X-2', actorRole: 'approver' });
    return (ok === true && n.bufKeys().length === 1 && n.auditRows().length === before) || [ok, n.bufKeys(), n.auditRows().length];
  });
  check('A4 [함정①] 처리자·시각은 요청 시점 값 — 트리거(David)가 옮겨도 김기안·쓰던 시각 그대로', () => {
    CLOCK.now += 3 * MIN;
    const r = n.flush();
    const last = n.auditRows()[n.auditRows().length - 1];
    return (r.flushed === 1 && last[AC.ACTOR] === KIM && last[AC.TS] instanceof RealDate && last[AC.TS].getTime() === wroteAt &&
            n.bufKeys().length === 0) || [r, last[AC.ACTOR], last[AC.TS], wroteAt];
  });
  check('A5 [함정①] actor를 안 넘긴 호출도 요청자(이결재)로 남음', () => {
    n.as(LEE); n.call('writeAuditLogDeferred_', { eventType: 'INSP_APPROVE', docNo: 'X-3' });
    n.flush();
    const last = n.auditRows()[n.auditRows().length - 1];
    return (last[AC.ACTOR] === LEE && last[AC.ROLE] === 'unknown') || [last[AC.ACTOR], last[AC.ROLE]];
  });
  check('A6 [함정②] 다른 실행이 반영 중(사용자 락 점유)이면 기다리지 않고 건너뜀', () => {
    n.as(KIM); n.call('writeAuditLogDeferred_', { eventType: 'DOC_SUBMIT', docNo: 'X-4' });
    n.state.lock.userHeld = true;
    const before = n.auditRows().length;
    const r = n.flush();
    n.state.lock.userHeld = false;
    return (r.skipped === 'busy' && n.auditRows().length === before && n.bufKeys().length === 1) || r;
  });
  check('A7 [함정②] 연달아 두 번 반영해도 같은 행이 두 번 들어가지 않음', () => {
    const before = n.auditRows().length;
    n.flush(); n.flush();
    return (n.auditRows().length === before + 1 && n.bufKeys().length === 0) || [before, n.auditRows().length];
  });
  check('A8 [함정②] 반영은 전역(스크립트) 락을 한 번도 잡지 않음', () => {
    n.as(KIM); n.call('writeAuditLogDeferred_', { eventType: 'DOC_SUBMIT', docNo: 'X-5' });
    const sw = n.state.lock.scriptWait; n.flush();
    return n.state.lock.scriptWait === sw || '전역 락 사용됨';
  });
  check('A9 [함정③] 하트비트가 15분 넘게 멈추면 쓰는 쪽이 즉시 기록으로 전환', () => {
    n.heartbeat(CLOCK.now - 16 * MIN);
    const before = n.auditRows().length;
    n.as(KIM); n.call('writeAuditLogDeferred_', { eventType: 'DOC_SUBMIT', docNo: 'X-6' });
    return (n.auditRows().length === before + 1 && n.bufKeys().length === 0) || [n.auditRows().length, n.bufKeys()];
  });
  check('A10 하트비트 14분이면 아직 지연 기록', () => {
    n.heartbeat(CLOCK.now - 14 * MIN);
    n.as(KIM); n.call('writeAuditLogDeferred_', { eventType: 'DOC_SUBMIT', docNo: 'X-7' });
    const ok = n.bufKeys().length === 1; n.flush();
    return ok || n.bufKeys();
  });
  check('A11 [함정③] 버퍼 저장 실패(할당량) → 그 자리에서 즉시 기록, 유실 없음', () => {
    n.heartbeat(); n.state.flags.propsSetFail = true;
    const before = n.auditRows().length;
    n.as(KIM); const ok = n.call('writeAuditLogDeferred_', { eventType: 'DOC_SUBMIT', docNo: 'X-8' });
    n.state.flags.propsSetFail = false;
    const last = n.auditRows()[n.auditRows().length - 1];
    return (ok === true && n.auditRows().length === before + 1 && last[AC.DOCNO] === 'X-8' && last[AC.ACTOR] === KIM) || [ok, n.auditRows().length];
  });
  check('A12 너무 큰 행(한글 2500자 초과) → 즉시 기록', () => {
    const before = n.auditRows().length;
    n.call('writeAuditLogDeferred_', { eventType: 'DOC_SUBMIT', docNo: 'X-9', payload: { big: '가'.repeat(3000) } });
    return (n.auditRows().length === before + 1 && n.bufKeys().length === 0) || [n.auditRows().length, n.bufKeys()];
  });
  check('A13 깨진 버퍼 항목 → AUDIT_BUF_CORRUPT 행으로 원문 보존 + 키 제거', () => {
    n.state.props.set('AUDIT_BUF_0000000000001_bad', '{not json');
    n.flush();
    const last = n.auditRows()[n.auditRows().length - 1];
    return (last[AC.EV] === 'AUDIT_BUF_CORRUPT' && last[AC.PAYLOAD] === '{not json' && n.bufKeys().length === 0) || last;
  });
  check('A14 시트 장애 중 반영 → throw, 버퍼 유지, 하트비트 미갱신 (다음 트리거에서 재시도)', () => {
    n.heartbeat(CLOCK.now);
    n.call('writeAuditLogDeferred_', { eventType: 'DOC_SUBMIT', docNo: 'X-10' });
    const hb = n.state.props.get('AUDIT_FLUSH_HEARTBEAT');
    CLOCK.now += MIN;
    n.auditSheet().failWrites = true;
    let threw = false; try { n.flush(); } catch (e) { threw = true; }
    n.auditSheet().failWrites = false;
    const kept = n.bufKeys().length === 1 && n.state.props.get('AUDIT_FLUSH_HEARTBEAT') === hb;
    const r = n.flush();
    return (threw && kept && r.flushed === 1 && n.bufKeys().length === 0) || [threw, kept, r];
  });
  check('A15 1회 최대 100행, 나머지는 다음 반영 — 시간 순서 유지', () => {
    n.heartbeat();
    const base = CLOCK.now;
    for (let i = 0; i < 150; i++) { CLOCK.now = base + i; n.call('writeAuditLogDeferred_', { eventType: 'DOC_SUBMIT', docNo: 'B-' + i }); }
    const before = n.auditRows().length;
    const r1 = n.flush(); const r2 = n.flush();
    const docNos = n.auditRows().slice(before).map(r => r[AC.DOCNO]);
    const ordered = docNos.every((d, i) => d === 'B-' + i);
    return (r1.flushed === 100 && r1.remaining === 50 && r2.flushed === 50 && ordered) || [r1, r2, docNos.slice(0, 3)];
  });
  check('A16 반영 후 200건 이상 남으면 관리자 알림 1회 (1시간 내 중복 없음)', () => {
    n.heartbeat();
    for (let i = 0; i < 320; i++) n.call('writeAuditLogDeferred_', { eventType: 'DOC_SUBMIT', docNo: 'C-' + i });
    const e0 = n.state.adminErrors.length;
    n.flush(); n.flush();
    const alerts = n.state.adminErrors.slice(e0).filter(m => m.indexOf('밀려 있습니다') >= 0).length;
    n.flush(); n.flush();
    return alerts === 1 || alerts;
  });
  check('A17 하트비트 키는 로그 항목으로 옮겨지지 않음 (접두사 충돌 없음)', () => {
    n.flush();
    return (n.state.props.has('AUDIT_FLUSH_HEARTBEAT') && !n.auditRows().some(r => String(r[AC.REASON]).indexOf('HEARTBEAT') >= 0)) || 'heartbeat 문제';
  });
}

{
  console.log('\n[A-trigger] processQueueTrigger 통합');
  const n = seed(makeEnv('NEW', readWT));
  check('T1 빈 큐 + 빈 버퍼: 전역 락 0회, 하트비트 갱신', () => {
    const sw = n.state.lock.scriptWait;
    n.as(ADMIN); n.call('processQueueTrigger');
    return (n.state.lock.scriptWait === sw && n.state.props.has('AUDIT_FLUSH_HEARTBEAT')) || [n.state.lock.scriptWait - sw];
  });
  check('T2 빈 큐 + 버퍼 2건: 반영되고 전역 락은 여전히 0회', () => {
    n.as(KIM); n.call('writeAuditLogDeferred_', { eventType: 'DOC_SUBMIT', docNo: 'T-1' }); n.call('writeAuditLogDeferred_', { eventType: 'DOC_SUBMIT', docNo: 'T-2' });
    const sw = n.state.lock.scriptWait, before = n.auditRows().length;
    n.as(ADMIN); n.call('processQueueTrigger');
    return (n.auditRows().length === before + 2 && n.state.lock.scriptWait === sw) || [n.auditRows().length - before, n.state.lock.scriptWait - sw];
  });
  check('T3 반영이 실패해도 큐 작업은 그대로 처리 + 관리자 알림', () => {
    n.as(KIM); n.call('writeAuditLogDeferred_', { eventType: 'DOC_SUBMIT', docNo: 'T-3' });
    n.state.props.set('PENDING_JOB_QUEUE', JSON.stringify([{ id: 'job_1', type: 'pdf_only', token: 'tok-req2', docNo: 'TH-AP-26-002', docType: 'REQ', attempts: 0, status: 'pending' }]));
    n.state.jobResult = { ok: true };
    n.auditSheet().failWrites = true;
    const e0 = n.state.adminErrors.length;
    n.as(ADMIN); n.call('processQueueTrigger');
    n.auditSheet().failWrites = false;
    const q = JSON.parse(n.state.props.get('PENDING_JOB_QUEUE'));
    const alerted = n.state.adminErrors.slice(e0).some(m => m.indexOf('지연 기록 반영 실패') >= 0);
    return (q.length === 0 && alerted && n.bufKeys().length === 1) || [q, alerted, n.bufKeys()];
  });
  check('B20 결재 PDF 최종 실패(3회째)만 DRIVE_SAVE_FAIL — 재시도 중 실패는 기록 안 함', () => {
    n.flush();
    n.state.jobResult = { ok: false, message: 'PDF 생성 실패: no servers' };
    n.state.props.set('PENDING_JOB_QUEUE', JSON.stringify([{ id: 'job_2', type: 'pdf_only', token: 'tok-req2', docNo: 'TH-AP-26-002', docType: 'REQ', attempts: 0, status: 'pending' }]));
    n.as(ADMIN); n.call('processQueueTrigger');
    const afterRetry = n.events().filter(e => e === 'DRIVE_SAVE_FAIL').length;
    n.state.props.set('PENDING_JOB_QUEUE', JSON.stringify([{ id: 'job_3', type: 'pdf_and_consolidate', token: 'tok-prc1', docNo: 'TO-PO-26-901', docType: 'PRC', attempts: 2, status: 'pending' }]));
    n.sb._processPdfAndConsolidateJob = function () { return { ok: false, message: 'FINAL 이동 실패' }; };
    n.call('processQueueTrigger');
    const rows = n.auditRows().filter(r => r[AC.EV] === 'DRIVE_SAVE_FAIL');
    const p = rows.length ? JSON.parse(rows[0][AC.PAYLOAD]) : {};
    return (afterRetry === 0 && rows.length === 1 && p.stage === 'queue_pdf_and_consolidate' && p.attempts === 3 && rows[0][AC.ROLE] === 'system') || [afterRetry, rows.length, p];
  });
}

// =================================================================
console.log('\n[R] 즉시 기록(writeAuditLog) 동작 불변 — HEAD와 행 내용 동일');
// =================================================================
{
  const { h, n } = pair(env => { env.as(LEE); env.call('writeAuditLog', { eventType: 'DOC_APPROVE', docNo: 'R-1', docToken: 't', docType: 'REQ', reason: 'r', payload: { a: 1 } }); return null; });
  check('R1 같은 입력 → 같은 12열 행 (logId 제외)', () => same(auditNoId(h.auditRows()), auditNoId(n.auditRows())));
  const p2 = pair(env => { env.as(LEE); return env.call('writeAuditLog', {}); });
  check('R2 eventType 없으면 둘 다 false', () => (p2.rh === false && p2.rn === false) || [p2.rh, p2.rn]);
}

// =================================================================
console.log('\n[B] 기존 기록 지점 전환 + 신규 로그 — HEAD와 같은 결과인지');
// =================================================================
const bufferedEvents = env => env.bufKeys().map(k => JSON.parse(env.state.props.get(k))[4]);
// 시나리오별 의도한 결과 — 성공 경로인지 거부/실패 경로인지
const EXPECT_OK = {
  B1: true, B1b: true, B3: true, B4: true, B5: true, B7: true, B8: true, B9: true, B11: true, B15: true,
  B16: true, B18: true, B18b: true, B19: true, B19b: true, B22b: true, C7: true,
  B2: false, B3b: false, B4b: false, B4c: false, B6: false, B6b: false, B10: false, B17: false, C7b: false, C7c: false,
};
function compareFlow(name, scenario, opts) {
  opts = opts || {};
  const code = name.split(' ')[0];
  if (EXPECT_OK[code] === undefined) throw new Error('EXPECT_OK에 ' + code + ' 없음');
  opts.ok = EXPECT_OK[code];
  const { h, n, rh, rn } = pair(scenario, opts);
  // 반영 전: 지연 기록이어야 할 것은 버퍼에만, 즉시 기록이어야 할 것은 시트에만 있어야 한다
  const buf = bufferedEvents(n), sheetNow = n.events();
  (opts.deferred || []).forEach(ev => check(name + ' — ' + ev + '는 지연 기록 (반영 전 시트에 없음)', () =>
    (buf.indexOf(ev) >= 0 && sheetNow.indexOf(ev) < 0) || { buffer: buf, sheet: sheetNow }));
  (opts.immediate || []).forEach(ev => check(name + ' — ' + ev + '는 즉시 기록 (버퍼 거치지 않음)', () =>
    (sheetNow.indexOf(ev) >= 0 && buf.indexOf(ev) < 0) || { buffer: buf, sheet: sheetNow }));
  n.flush();
  // 둘 다 같은 이유로 실패해도 '동일'이 되므로, 시나리오가 의도한 경로(성공/거부)를 탔는지 먼저 확인한다
  if (opts.ok !== undefined) check(name + ' — 의도한 경로 (ok=' + opts.ok + ')', () => (rn && rn.ok === opts.ok && rh && rh.ok === opts.ok) || [rh && rh.message, rn && rn.message]);
  check(name + ' — 반환값 동일', () => same(rh, rn));
  check(name + ' — 시트 상태 동일', () => same(sheetRows(h), sheetRows(n)));
  check(name + ' — 메일 동일', () => same(h.state.mails, n.state.mails));
  if (!opts.skipAudit) {
    check(name + ' — 시스템로그: HEAD의 행이 지연 기록 반영 후 그대로 있음', () => {
      // fixHeadAudit: 이후 의도한 수정이 바꾼 칸을 HEAD 쪽에 반영해 비교 (나머지 칸은 그대로 같아야 한다)
      const ha = (opts.fixHeadAudit || (x => x))(auditNoId(h.auditRows())), na = auditNoId(n.auditRows());
      const extra = opts.extraEvents || [];
      const naKept = na.filter(r => extra.indexOf(r[AC.EV - 1]) < 0);
      return same(ha, naKept);
    });
  }
  return { h, n, rh, rn };
}

{
  const r = compareFlow('B1 결재 승인(기안자 단계)', env => { env.as(KIM); return env.call('processDecisionFromClient', { token: 'tok-req1', decision: 'approve', comment: '' }); }, { deferred: ['DOC_APPROVE'] });
  check('B1 수정본은 DOC_APPROVE를 버퍼로 썼다가 반영 (처리자=김기안)', () => {
    const row = r.n.auditRows().find(x => x[AC.EV] === 'DOC_APPROVE');
    return (row && row[AC.ACTOR] === KIM && r.n.bufKeys().length === 0) || row;
  });
}
compareFlow('B1b 결재 반려', env => { env.as(LEE); return env.call('processDecisionFromClient', { token: 'tok-req8', decision: 'reject', comment: '단가 재확인' }); }, { deferred: ['DOC_REJECT'] });
{
  const r = compareFlow('B2 비결재자 승인 시도(거부)', env => { env.as(PARK); return env.call('processDecisionFromClient', { token: 'tok-req1', decision: 'approve' }); }, { immediate: ['DECISION_DENIED'] });
  check('B2 거부 기록(DECISION_DENIED)은 버퍼를 거치지 않고 즉시', () => {
    const p = pair(env => { env.as(PARK); return env.call('processDecisionFromClient', { token: 'tok-req1', decision: 'approve' }); });
    return (p.n.events().indexOf('DECISION_DENIED') >= 0 && p.n.bufKeys().length === 0) || [p.n.events(), p.n.bufKeys()];
  });
}
compareFlow('B3 기안자 폐기(반려 후)', env => { env.as(KIM); return env.call('discardRequisitionForClient', { token: 'tok-req5', reason: '기안자 폐기 (반려 후)' }); }, { deferred: ['DOC_DISCARD'] });
compareFlow('B4 재상신(반려 후)', env => { env.as(KIM); const d = submitPayload(KIM); d.subject = '수정 품의'; return env.call('resubmitRequisitionForClient', { token: 'tok-req5', data: d }); }, { deferred: ['DOC_RESUBMIT'] });
compareFlow('B4b 재상신 — 첨부 업로드 실패', env => { env.as(KIM); env.state.flags.uploadFail = true; const x = env.call('resubmitRequisitionForClient', { token: 'tok-req5', data: submitPayload(KIM) }); env.state.flags.uploadFail = false; return x; }, { deferred: ['DOC_RESUBMIT'], immediate: ['DRIVE_SAVE_FAIL'], extraEvents: ['DRIVE_SAVE_FAIL'] });
compareFlow('B4c 남의 문서 재상신 시도(거부)', env => { env.as(PARK); return env.call('resubmitRequisitionForClient', { token: 'tok-req5', data: submitPayload(PARK) }); }, { immediate: ['DECISION_DENIED'] });
compareFlow('B3b 남의 문서 폐기 시도(거부)', env => { env.as(PARK); return env.call('discardRequisitionForClient', { token: 'tok-req5' }); }, { immediate: ['DECISION_DENIED'] });

function submitPayload(who) { return {
  docNo: 'TH-AP-26-100', issueDate: '2026-10-14', drafter: NAMES[who], dept: DEPT[who], subject: '신규 품의', purpose: '시험',
  vendorName: '라인테크', vendorEmail: 'v@x.com', vendorContact: '홍', vendorPhone: '010', deliveryDate: '2026-10-30', deliveryAddr: '대전',
  items: [{ name: '볼트', spec: 'M8', qty: 10, price: 1000, currency: 'KRW' }], remarks: '',
  approvers: [{ label: '기안자', name: NAMES[who], email: who }, { label: '결재자 1', name: NAMES[LEE], email: LEE }], attachments: [],
}; }
{
  const r = compareFlow('B5 품의 제출', env => { env.as(KIM); return env.call('submitRequisitionForClient', submitPayload(KIM)); }, { extraEvents: ['DOC_SUBMIT'], deferred: ['DOC_SUBMIT'] });
  check('B5 DOC_SUBMIT 1건 (처리자=김기안, 토큰=새 행 토큰)', () => {
    const rows = r.n.auditRows().filter(x => x[AC.EV] === 'DOC_SUBMIT');
    const newRow = mainRowByDoc(r.n, 'TH-AP-26-100');
    return (rows.length === 1 && rows[0][AC.ACTOR] === KIM && rows[0][AC.TOKEN] === newRow[r.n.sb.COL.TOKEN] && rows[0][AC.ROLE] === 'requester') || rows;
  });
}
{
  const r = compareFlow('B6 품의 제출 — 첨부 업로드 실패', env => { env.as(KIM); env.state.flags.uploadFail = true; const x = env.call('submitRequisitionForClient', submitPayload(KIM)); env.state.flags.uploadFail = false; return x; }, { extraEvents: ['DOC_SUBMIT', 'DRIVE_SAVE_FAIL'], deferred: ['DOC_SUBMIT'], immediate: ['DRIVE_SAVE_FAIL'] });
  check('B6 DRIVE_SAVE_FAIL 즉시 기록 (stage=attachment_upload) + 상태 업로드오류(HEAD와 동일)', () => {
    const row = r.n.auditRows().find(x => x[AC.EV] === 'DRIVE_SAVE_FAIL');
    const p = row ? JSON.parse(row[AC.PAYLOAD]) : {};
    return (row && p.stage === 'attachment_upload' && mainRowByDoc(r.n, 'TH-AP-26-100')[r.n.sb.COL.STATUS] === '업로드오류') || row;
  });
}
{
  const r = compareFlow('B6b 품의 제출 — 폴더 접근 실패', env => {
    env.as(KIM); const orig = env.sb.getOrCreateFolder;
    // 행 기록(락 안)은 통과시키고, 락 밖 DriveApp.getFolderById만 실패시킨다
    env.state.flags.folderAccessFail = true; const x = env.call('submitRequisitionForClient', submitPayload(KIM)); env.state.flags.folderAccessFail = false; return x;
  }, { extraEvents: ['DOC_SUBMIT', 'DRIVE_SAVE_FAIL'], immediate: ['DRIVE_SAVE_FAIL'] });
  check('B6b DRIVE_SAVE_FAIL (stage=folder_access)', () => {
    const row = r.n.auditRows().find(x => x[AC.EV] === 'DRIVE_SAVE_FAIL');
    return (row && JSON.parse(row[AC.PAYLOAD]).stage === 'folder_access') || row;
  });
}

{
  const r = compareFlow('B7 픽업(신규 점유)', env => { env.as(BUYER1); return env.call('claimRequestForClient', { parentToken: 'tok-req2' }); }, { extraEvents: ['PRC_CLAIM'], deferred: ['PRC_CLAIM'] });
  check('B7 반환값에 _audit 같은 내부 키 없음', () => (!('_audit' in r.rn)) || Object.keys(r.rn));
  check('B7 PRC_CLAIM mode=new, 처리자=구매1', () => {
    const row = r.n.auditRows().find(x => x[AC.EV] === 'PRC_CLAIM');
    return (row && row[AC.ACTOR] === BUYER1 && JSON.parse(row[AC.PAYLOAD]).mode === 'new' && row[AC.DOCNO] === 'TH-AP-26-002') || row;
  });
}
{
  const r = compareFlow('B8 픽업(본인 재점유)', env => { env.as(BUYER1); return env.call('claimRequestForClient', { parentToken: 'tok-req7' }); });
  check('B8 본인 재점유는 기록 안 함', () => r.n.events().indexOf('PRC_CLAIM') < 0 || r.n.events());
}
{
  const r = compareFlow('B9 픽업(만료 점유 인수)', env => { env.as(BUYER1); return env.call('claimRequestForClient', { parentToken: 'tok-req4' }); }, { extraEvents: ['PRC_CLAIM'], deferred: ['PRC_CLAIM'] });
  check('B9 mode=ttl_takeover, 이전 점유자=구매2', () => {
    const row = r.n.auditRows().find(x => x[AC.EV] === 'PRC_CLAIM');
    const p = row ? JSON.parse(row[AC.PAYLOAD]) : {};
    return (p.mode === 'ttl_takeover' && p.previousClaimedBy === BUYER2 && row[AC.TARGET] === BUYER2) || p;
  });
}
{
  const r = compareFlow('B10 픽업(다른 담당자 작업 중 — 거부)', env => { env.as(BUYER1); return env.call('claimRequestForClient', { parentToken: 'tok-req3' }); });
  check('B10 거부는 기록 안 함', () => r.n.events().indexOf('PRC_CLAIM') < 0 || r.n.events());
}
{
  const r = compareFlow('B11 회수(본인 락)', env => { env.as(BUYER1); return env.call('releaseClaimForClient', { parentToken: 'tok-req7' }); }, { extraEvents: ['PRC_RELEASE'], deferred: ['PRC_RELEASE'] });
  check('B11 반환값 {ok,message} 그대로 + PRC_RELEASE 1건', () => {
    const row = r.n.auditRows().find(x => x[AC.EV] === 'PRC_RELEASE');
    return (Object.keys(r.rn).join(',') === 'ok,message' && row && row[AC.ACTOR] === BUYER1 && JSON.parse(row[AC.PAYLOAD]).byAdmin === false) || [r.rn, row];
  });
}
compareFlow('B15 회수(이미 해제됨)', env => { env.as(BUYER1); return env.call('releaseClaimForClient', { parentToken: 'tok-req2' }); });
{
  // 보안 수정 — 여기서는 HEAD와 '달라야' 한다
  const p = pair(env => { env.as(BUYER1); return env.call('releaseClaimForClient', { parentToken: 'tok-req3', force: true }); });
  check('B12 [보안] force=true로 남의 락 해제: HEAD는 풀렸고(취약), 수정본은 거부 + 락 유지', () => {
    const C = p.n.sb.COL;
    return (p.rh.ok === true && mainRow(p.h, 'tok-req3')[C.CLAIMED_BY] === '' &&
            p.rn.ok === false && mainRow(p.n, 'tok-req3')[C.CLAIMED_BY] === BUYER2) || [p.rh, p.rn];
  });
  const q = pair(env => { env.as(''); return env.call('releaseClaimForClient', { parentToken: 'tok-req3' }); });
  check('B14 [보안] 신원 확인 실패: HEAD는 배포 계정(관리자)으로 폴백돼 풀렸고, 수정본은 거부', () => {
    const C = q.n.sb.COL;
    return (q.rh.ok === true && q.rn.ok === false && mainRow(q.n, 'tok-req3')[C.CLAIMED_BY] === BUYER2) || [q.rh, q.rn];
  });
  const a = pair(env => { env.as(ADMIN); return env.call('releaseClaimForClient', { parentToken: 'tok-req3' }); });
  a.n.flush();
  check('B13 관리자가 남의 락 회수 → 둘 다 허용, 수정본은 byAdmin 기록', () => {
    const row = a.n.auditRows().find(x => x[AC.EV] === 'PRC_RELEASE');
    return (a.rh.ok && a.rn.ok && row && row[AC.ROLE] === 'admin' && row[AC.TARGET] === BUYER2) || [a.rh, a.rn, row];
  });
}

const prcPayload = () => ({
  parentToken: 'tok-req7', docNo: 'TO-PO-26-950', issueDate: '2026-10-14', drafter: NAMES[BUYER1], dept: '구매팀', subject: '구매', purpose: '시험',
  vendorName: '라인테크', vendorEmail: 'v@x.com', vendorContact: '홍', vendorPhone: '010', deliveryDate: '2026-10-30', deliveryAddr: '대전',
  items: [{ name: '볼트', spec: 'M8', qty: 10, price: 1000, currency: 'KRW' }], remarks: '', purchaseMethod: '계좌이체', paymentInfo: '익월말',
  approvers: [{ label: '기안자', name: NAMES[BUYER1], email: BUYER1 }, { label: '결재자 1', name: NAMES[LEE], email: LEE }],
  attachments: [], existingAttachments: [], fieldChanges: [{ field: '납기일', from: '2026-10-20', to: '2026-10-30' }],
});
{
  const r = compareFlow('B16 구매 문서 제출', env => { env.as(BUYER1); return env.call('submitPrcForClient', prcPayload()); }, { extraEvents: ['PRC_SUBMIT'], deferred: ['PRC_SUBMIT'] });
  check('B16 반환값에 actor·parentDocNo가 새지 않음', () => (!('actor' in r.rn) && !('parentDocNo' in r.rn)) || Object.keys(r.rn));
  check('B16 PRC_SUBMIT: 원본 품의 TH-AP-26-007, 처리자=구매1', () => {
    const row = r.n.auditRows().find(x => x[AC.EV] === 'PRC_SUBMIT');
    const p = row ? JSON.parse(row[AC.PAYLOAD]) : {};
    return (row && row[AC.ACTOR] === BUYER1 && p.parentDocNo === 'TH-AP-26-007' && p.fieldChanges === 1) || [row, p];
  });
  check('B16 PRC_RELEASE는 따로 남기지 않음', () => r.n.events().indexOf('PRC_RELEASE') < 0 || r.n.events());
}
compareFlow('B17 구매 문서 제출 — 첨부 업로드 실패', env => { env.as(BUYER1); env.state.flags.uploadFail = true; const x = env.call('submitPrcForClient', prcPayload()); env.state.flags.uploadFail = false; return x; }, { extraEvents: ['PRC_SUBMIT', 'DRIVE_SAVE_FAIL'], deferred: ['PRC_SUBMIT'], immediate: ['DRIVE_SAVE_FAIL'] });

const inspPayload = (who, prcToken) => ({
  prcToken: prcToken || 'tok-prc2', resubToken: '', receivedDate: '2026-10-14', receivedNote: '1차 입고', verdict: '합격', isFinal: false, comment: '',
  approvers: [{ label: '기안자', name: NAMES[who], email: who }, { label: '결재자 1', name: NAMES[LEE], email: LEE }], photos: [],
});
{
  const r = compareFlow('B18 검수보고서 제출', env => { env.as(KIM); return env.call('submitInspForClient', inspPayload(KIM)); }, { deferred: ['INSP_SUBMIT'] });
  check('B18 INSP_SUBMIT 처리자=김기안 (지연 기록 경유)', () => {
    const row = r.n.auditRows().find(x => x[AC.EV] === 'INSP_SUBMIT');
    return (row && row[AC.ACTOR] === KIM) || row;
  });
}
{
  const r = compareFlow('B18b 검수보고서 제출 — 사진 저장 실패', env => { env.as(KIM); env.state.flags.photoFail = true; const x = env.call('submitInspForClient', inspPayload(KIM)); env.state.flags.photoFail = false; return x; }, { extraEvents: ['DRIVE_SAVE_FAIL'], deferred: ['INSP_SUBMIT'], immediate: ['DRIVE_SAVE_FAIL'] });
  check('B18b DRIVE_SAVE_FAIL (stage=insp_photo) + 관리자 알림은 HEAD와 동일', () => {
    const row = r.n.auditRows().find(x => x[AC.EV] === 'DRIVE_SAVE_FAIL');
    return (row && JSON.parse(row[AC.PAYLOAD]).stage === 'insp_photo' && same(r.h.state.adminErrors, r.n.state.adminErrors) === true) || row;
  });
}
{
  const r = compareFlow('B19 검수 결재 최종 승인', env => { env.as(LEE); return env.call('processInspDecisionFromClient', { token: 'tok-i3', decision: 'approve', comment: '' }); }, { deferred: ['INSP_FINALIZED'] });
  check('B19 INSP_FINALIZED 처리자=이결재 (actor 미지정 호출인데도)', () => {
    const row = r.n.auditRows().find(x => x[AC.EV] === 'INSP_FINALIZED');
    return (row && row[AC.ACTOR] === LEE) || row;
  });
}
{
  // 2026-10-29 INSP 반려 메일 링크 수정: 반려 결과에 토큰이 실리면서 INSP_REJECT 행의 문서 토큰(그동안 빈칸)도 채워진다
  const tokCol = AC.TOKEN - 1;
  const r = compareFlow('B19b 검수 결재 반려', env => { env.as(LEE); return env.call('processInspDecisionFromClient', { token: 'tok-i3', decision: 'reject', comment: '수량 불일치' }); }, {
    deferred: ['INSP_REJECT'],
    fixHeadAudit: rows => rows.map(x => (x[AC.EV - 1] === 'INSP_REJECT' && x[tokCol] === '') ? Object.assign(x.slice(), { [tokCol]: 'tok-i3' }) : x),
  });
  check('B19b INSP_REJECT 행에 문서 토큰이 남음 (수정 전 빈칸)', () => {
    const nr = r.n.auditRows().find(x => x[AC.EV] === 'INSP_REJECT'), hr = r.h.auditRows().find(x => x[AC.EV] === 'INSP_REJECT');
    return (nr && nr[AC.TOKEN] === 'tok-i3' && hr && hr[AC.TOKEN] === '') || [hr && hr[AC.TOKEN], nr && nr[AC.TOKEN]];
  });
}

{
  console.log('\n[B21] 잘못된 링크 접근');
  const n = seed(makeEnv('NEW', readWT));
  const lia = () => n.auditRows().filter(r => r[AC.EV] === 'LINK_INVALID_ACCESS');
  check('뷰어: 없는 토큰 → 1건, 화면 응답은 HEAD와 같음', () => {
    n.as(KIM); const res = n.call('getRequisitionForViewer', 'tok-gone');
    return (res.ok === false && res.message === '유효하지 않은 토큰입니다.' && lia().length === 1 && lia()[0][AC.ACTOR] === KIM) || [res, lia().length];
  });
  check('같은 사람·같은 토큰 재접속(새로고침) → 추가 기록 없음', () => { n.call('getRequisitionForViewer', 'tok-gone'); return lia().length === 1 || lia().length; });
  check('반려 화면(조회 함수 2개 호출) → 1건만', () => { n.call('getRejectDataForClient', 'tok-gone2'); n.call('getRequisitionForViewer', 'tok-gone2'); return lia().length === 2 || lia().length; });
  check('다른 사람이 같은 토큰 → 별도 1건', () => { n.as(PARK); n.call('getRequisitionForViewer', 'tok-gone'); return lia().length === 3 || lia().length; });
  check('10분 지나면 다시 1건', () => { CLOCK.now += 11 * MIN; n.as(KIM); n.call('getRequisitionForViewer', 'tok-gone'); return lia().length === 4 || lia().length; });
  check('INSP 뷰어·결재 화면·재상신 폼·PRC 작성 폼 각각 기록', () => {
    const b = lia().length; n.as(LEE);
    n.call('getInspViewerDataForClient', 'tok-x1'); n.call('getInspDecisionMetaForClient', 'tok-x2');
    n.as(KIM); n.call('getInspFormDataForClient', 'tok-prc1', 'tok-x3'); n.call('getInspFormDataForClient', 'tok-x4', null);
    n.as(BUYER1); n.call('getPrefillDataForClient', 'tok-x5');
    const w = lia().slice(b).map(r => JSON.parse(r[AC.PAYLOAD]).where);
    return same(['insp_view', 'insp_approve', 'insp_resub', 'insp_form', 'prc_form'], w);
  });
  check('정상 토큰은 기록 안 함', () => { const b = lia().length; n.as(KIM); n.call('getRequisitionForViewer', 'tok-req1'); n.call('getInspViewerDataForClient', 'tok-i1'); return lia().length === b || lia().length; });
  check('시트 장애여도 화면 응답은 그대로 (throw 없음)', () => {
    n.auditSheet().failWrites = true; CLOCK.now += 11 * MIN;
    const res = n.call('getRequisitionForViewer', 'tok-gone9');
    n.auditSheet().failWrites = false;
    return (res.ok === false && res.message === '유효하지 않은 토큰입니다.') || res;
  });
}

{
  console.log('\n[B22] A-3 결재자 변경은 즉시 기록 유지 — 기록 실패 시 원복');
  const p = pair(env => {
    env.as(ADMIN); env.auditSheet().failWrites = true;
    const x = env.call('adminChangeApproverForClient', { docType: 'INSP', token: 'tok-i3', stageIdx: 1, newEmail: CHOI, reason: '휴가' });
    env.auditSheet().failWrites = false; return x;
  });
  check('HEAD·수정본 모두 "변경을 취소했습니다" + 결재자 원복', () => {
    const I = p.n.sb.INSP_COL;
    const rowN = inspRowByDoc(p.n, 'TH-AP-26-006-03');
    return (p.rh.ok === false && p.rn.ok === false && p.rn.message.indexOf('취소') >= 0 && rowN[I.APPR_START + 6 + 2] === LEE && same(p.rh, p.rn) === true) || [p.rh, p.rn];
  });
  const q = compareFlow('B22b 결재자 변경 정상', env => { env.as(ADMIN); return env.call('adminChangeApproverForClient', { docType: 'INSP', token: 'tok-i3', stageIdx: 1, newEmail: CHOI, reason: '휴가' }); }, { immediate: ['ADMIN_CHANGE_APPROVER'] });
  check('B22b 정상 변경 기록은 버퍼 없이 바로 시트에', () => q.n.bufKeys().length === 0 || q.n.bufKeys());
}

// =================================================================
console.log('\n[C] INSP 폐기 — 관리자 강제 폐기');
// =================================================================
{
  const p = pair(env => { env.as(ADMIN); return env.call('adminListDiscardableForClient', {}); });
  const stripTs = d => { const o = Object.assign({}, d); delete o.timestamp; delete o.timestampMs; delete o.rowIndex; return o; };
  check('C1 REQ/PRC 행은 HEAD와 같음 (제출일 표시 버그 수정분 제외)', () => {
    const hDocs = p.rh.docs.map(stripTs).sort((a, b) => a.token < b.token ? -1 : 1);
    const nDocs = p.rn.docs.filter(d => d.docType !== 'INSP').map(stripTs).sort((a, b) => a.token < b.token ? -1 : 1);
    return same(hDocs, nDocs);
  });
  check('C1 [버그 수정] HEAD는 제출일이 전부 비었고, 수정본은 채워지고 최신순 정렬', () => {
    const hEmpty = p.rh.docs.every(d => d.timestamp === '' && d.timestampMs === 0);
    const ms = p.rn.docs.map(d => d.timestampMs);
    const sorted = ms.every((v, i) => i === 0 || ms[i - 1] >= v);
    return (hEmpty && p.rn.docs.every(d => d.timestamp !== '') && sorted) || [hEmpty, ms];
  });
  check('C1 INSP: 반려(-02)·결재중(-03)은 선택 가능, 최종승인(-01)은 보호 대상', () => {
    const i = {}; p.rn.docs.filter(d => d.docType === 'INSP').forEach(d => { i[d.docNo] = d; });
    return (i['TH-AP-26-006-02'] && !i['TH-AP-26-006-02'].isProtected && i['TH-AP-26-006-03'] && !i['TH-AP-26-006-03'].isProtected &&
            i['TH-AP-26-006-01'] && i['TH-AP-26-006-01'].isProtected && i['TH-AP-26-006-03'].drafter === NAMES[PARK] &&
            i['TH-AP-26-006-03'].curApprEmail === LEE && i['TH-AP-26-006-03'].reqDrafterEmail === KIM && i['TH-AP-26-006-03'].amount === null) || i;
  });
  const n = p.n;
  check('C1 유형 필터 INSP → INSP만 / REQ → INSP 없음', () => {
    const a = n.call('adminListDiscardableForClient', { docType: 'INSP' }).docs, b = n.call('adminListDiscardableForClient', { docType: 'REQ' }).docs;
    return (a.length === 3 && a.every(d => d.docType === 'INSP') && b.length > 0 && b.every(d => d.docType === 'REQ')) || [a.length, b.map(d => d.docType)];
  });
  check('C1 [버그 수정] 상태 "결재중" → 진행 중 REQ(결재중 (결재자 1))와 INSP 모두 잡힘 (HEAD는 0건)', () => {
    const hh = p.h.call('adminListDiscardableForClient', { status: '결재중' }).docs;
    const nn = n.call('adminListDiscardableForClient', { status: '결재중' }).docs.map(d => d.docNo).sort();
    return (hh.length === 0 && same(['TH-AP-26-006-03', 'TH-AP-26-008'], nn) === true) || [hh.length, nn];
  });
  check('C1 기간 필터가 이제 동작 (10/05~10/08 → REQ 005·007·008)', () => {
    const d = n.call('adminListDiscardableForClient', { dateFrom: '2026-10-05', dateTo: '2026-10-08', docType: 'REQ' }).docs.map(x => x.docNo).sort();
    return same(['TH-AP-26-005', 'TH-AP-26-007', 'TH-AP-26-008'], d);
  });
}
{
  const n = seed(makeEnv('NEW', readWT)); n.heartbeat();
  const I = n.sb.INSP_COL;
  n.as(ADMIN);
  const res = n.call('adminForceDiscardForClient', { docType: 'INSP', token: 'tok-i3', reason: '중복 제출 정리' });
  const row = inspRowByDoc(n, 'TH-AP-26-006-03');
  check('C3 진행 중 INSP 강제 폐기: 상태 폐기 + 토큰 무효화 + 이력', () => (res.ok && row[I.STATUS] === '폐기' && /^DISCARDED_/.test(row[I.TOKEN]) && /\[관리자 강제폐기\].*중복 제출 정리/.test(row[I.REJECT_LOG])) || [res, row[I.STATUS], row[I.TOKEN]]);
  check('C3 사진 임시폴더(INSP_…) 휴지통', () => n.state.folders['F_INSP_-03'].trashed === true || n.state.folders['F_INSP_-03']);
  check('C3 통지: 제출자(박동료)·품의 기안자(김기안)·현재 결재자(이결재) 3명, 관리자 제외', () => {
    const to = n.state.mails.filter(m => m.kind === 'gmail').map(m => m.to).sort();
    return same([KIM, LEE, PARK].sort(), to);
  });
  check('C3 시스템로그 ADMIN_FORCE_DISCARD 즉시 기록 (docType=INSP)', () => {
    const r = n.auditRows().find(x => x[AC.EV] === 'ADMIN_FORCE_DISCARD');
    const p = r ? JSON.parse(r[AC.PAYLOAD]) : {};
    return (r && r[AC.TYPE] === 'INSP' && r[AC.TOKEN] === 'tok-i3' && p.oldStatus === '결재중 (이결재)' && p.folderTrashed === true && n.bufKeys().length === 0) || [r, p];
  });
  check('C3 REQ/PRC 행은 한 칸도 안 바뀜', () => {
    const fresh = seed(makeEnv('NEW', readWT));
    return same(fresh.main().rows, n.main().rows);
  });
  check('C13 폐기 후 잠금 해제: 같은 품의에 새 검수보고서 제출 가능, 다음 번호 -04', () => {
    n.as(KIM); const f = n.call('getInspFormDataForClient', 'tok-prc1', null);
    return (f.ok && f.form.nextSeq === 4 && f.form.docNo === 'TH-AP-26-006-04') || f;
  });
  check('C13 실제 제출 → -04로 기록 (-03 번호 재사용 없음)', () => {
    const s = n.call('submitInspForClient', inspPayload(KIM, 'tok-prc1'));
    return (s.ok && s.docNo === 'TH-AP-26-006-04' && n.insp().rows.filter(r => r[I.DOC_NO] === 'TH-AP-26-006-03').length === 1) || s;
  });
  check('C14 옛 링크 차단: 뷰어·결재·재상신 모두 "찾을 수 없음"', () => {
    n.as(LEE);
    const v = n.call('getInspViewerDataForClient', 'tok-i3'), m = n.call('getInspDecisionMetaForClient', 'tok-i3');
    const d = n.call('processInspDecisionFromClient', { token: 'tok-i3', decision: 'approve' });
    return (!v.ok && !m.ok && !d.ok && d.message === '유효하지 않은 토큰입니다.') || [v, m, d];
  });
  check('C15 결재 대기·전사 대기·완료 목록에서 빠짐', () => {
    const menus = n.sb._getInspMenusForClient(n.ss, LEE, false, true, false);
    const all = [].concat(menus.myInspPending, menus.allInspPending, menus.inspFinals).map(x => x.docNo);
    return all.indexOf('TH-AP-26-006-03') < 0 || all;
  });
  check('C16 결재자 변경 목록에서 빠짐', () => {
    n.as(ADMIN); const l = n.call('adminListChangeableForClient', { docType: 'INSP' });
    return l.docs.every(d => d.docNo !== 'TH-AP-26-006-03') || l.docs.map(d => d.docNo);
  });
  check('C5 이미 폐기된 건 재시도 → 토큰이 없어 거부, 시트 불변', () => {
    const before = JSON.stringify(n.insp().rows);
    const r = n.call('adminForceDiscardForClient', { docType: 'INSP', token: 'tok-i3', reason: '재시도' });
    return (r.ok === false && JSON.stringify(n.insp().rows) === before) || r;
  });
  check('C4 최종승인(INSP)은 보호 — 거부, 시트·폴더 불변', () => {
    const before = JSON.stringify(n.insp().rows);
    const r = n.call('adminForceDiscardForClient', { docType: 'INSP', token: 'tok-i1', reason: '테스트' });
    return (r.ok === false && r.message.indexOf('보호') >= 0 && JSON.stringify(n.insp().rows) === before) || r;
  });
  check('C4b 사유 없으면 거부', () => { const r = n.call('adminForceDiscardForClient', { docType: 'INSP', token: 'tok-i2', reason: '  ' }); return (r.ok === false && inspRowByDoc(n, 'TH-AP-26-006-02')[I.STATUS] === '반려') || r; });
}
{
  const n = seed(makeEnv('NEW', readWT));
  const I = n.sb.INSP_COL;
  n.insp().rows.find(r => r[I.TOKEN] === 'tok-i2')[I.DRIVE_ID] = 'F_WRONG';
  n.state.folders['F_WRONG'] = { id: 'F_WRONG', name: '02. Purchase', trashed: false, files: [] };
  n.as(ADMIN); n.call('adminForceDiscardForClient', { docType: 'INSP', token: 'tok-i2', reason: '정리' });
  check('C8 폴더명이 INSP_로 시작하지 않으면 휴지통으로 보내지 않음', () => n.state.folders['F_WRONG'].trashed === false || 'trashed!');
  check('C8 반려 건 강제 폐기 → 현재 결재자 없음: 통지는 김기안 1명(제출자=기안자)', () => {
    const to = n.state.mails.filter(m => m.kind === 'gmail').map(m => m.to);
    return same([KIM], to);
  });
}
{
  console.log('\n[C6] 권한 — 관리자가 아니면 막힘');
  const n = seed(makeEnv('NEW', readWT));
  const before = JSON.stringify(n.insp().rows);
  [KIM, BUYER1, LEE].forEach(u => {
    n.as(u);
    check(u + ' 강제 폐기 거부', () => { const r = n.call('adminForceDiscardForClient', { docType: 'INSP', token: 'tok-i3', reason: 'x' }); return r.ok === false || r; });
    check(u + ' INSP 직접 호출(adminForceDiscardInsp) 거부', () => { const r = n.call('adminForceDiscardInsp', { token: 'tok-i3', reason: 'x' }); return r.ok === false || r; });
    check(u + ' 목록 조회 거부', () => { const r = n.call('adminListDiscardableForClient', { docType: 'INSP' }); const r2 = n.call('inspListDiscardableForAdmin', {}); return (r.ok === false && r2.ok === false) || [r, r2]; });
  });
  check('거부 시 시트 불변 + ADMIN_ACCESS_DENIED 기록', () => (JSON.stringify(n.insp().rows) === before && n.events().filter(e => e === 'ADMIN_ACCESS_DENIED').length >= 9) || n.events());
}
{
  console.log('\n[C7] REQ/PRC 강제 폐기는 HEAD와 동일 (docType 없이 호출)');
  compareFlow('C7 REQ 강제 폐기', env => { env.as(ADMIN); return env.call('adminForceDiscardForClient', { token: 'tok-req1', reason: '중복' }); }, { immediate: ['ADMIN_FORCE_DISCARD'] });
  compareFlow('C7b PRC 최종승인 보호', env => { env.as(ADMIN); return env.call('adminForceDiscardForClient', { token: 'tok-prc1', reason: '중복' }); });
  compareFlow('C7c INSP 토큰을 docType 없이 보내면 REQ 경로 → 찾을 수 없음 (HEAD와 동일)', env => { env.as(ADMIN); return env.call('adminForceDiscardForClient', { token: 'tok-i3', reason: '중복' }); });
}

// =================================================================
console.log('\n[C9] INSP 폐기 — 제출자 본인 (반려 후)');
// =================================================================
{
  const n = seed(makeEnv('NEW', readWT)); n.heartbeat();
  const I = n.sb.INSP_COL;
  n.as(PARK);
  check('C10 같은 부서 동료(제출자 아님)는 거부 + DECISION_DENIED 즉시 기록', () => {
    const before = JSON.stringify(n.insp().rows);
    const r = n.call('discardInspForClient', { token: 'tok-i2' });
    const den = n.auditRows().find(x => x[AC.EV] === 'DECISION_DENIED');
    return (r.ok === false && r.message.indexOf('제출자 본인') >= 0 && !('audit' in r) && JSON.stringify(n.insp().rows) === before &&
            den && JSON.parse(den[AC.PAYLOAD]).reasonCode === 'not_submitter') || [r, den];
  });
  check('C11 결재 진행 중(-03)은 제출자 본인(박동료)이어도 거부', () => {
    const r = n.call('discardInspForClient', { token: 'tok-i3' });
    return (r.ok === false && r.message.indexOf('반려된') >= 0 && inspRowByDoc(n, 'TH-AP-26-006-03')[I.STATUS] === '결재중 (이결재)') || r;
  });
  n.as('');
  check('C12 신원 확인 실패 → 거부 (배포 계정으로 폴백하지 않음)', () => { const r = n.call('discardInspForClient', { token: 'tok-i2' }); return (r.ok === false && r.message.indexOf('로그인') >= 0) || r; });
  n.as(KIM);
  const r = n.call('discardInspForClient', { token: 'tok-i2' });
  check('C9 제출자 본인(김기안) 반려 건 폐기 → 성공, 응답에 내부 키 없음', () => (r.ok === true && Object.keys(r).sort().join(',') === 'message,ok') || r);
  check('C9 상태 폐기 + 토큰 무효화 + [폐기] 이력 + 사진 폴더 휴지통', () => {
    const row = inspRowByDoc(n, 'TH-AP-26-006-02');
    return (row[I.STATUS] === '폐기' && /^DISCARDED_/.test(row[I.TOKEN]) && /\[폐기\].*제출자 폐기/.test(row[I.REJECT_LOG]) && n.state.folders['F_INSP_-02'].trashed) || row;
  });
  check('C9 메일 없음 (REQ/PRC 기안자 폐기와 동일)', () => n.state.mails.length === 0 || n.state.mails);
  check('C9 INSP_DISCARD는 지연 기록 → 반영 후 처리자=김기안', () => {
    const buffered = n.bufKeys().length === 1; n.flush();
    const row = n.auditRows().find(x => x[AC.EV] === 'INSP_DISCARD');
    return (buffered && row && row[AC.ACTOR] === KIM && row[AC.ROLE] === 'requester' && JSON.parse(row[AC.PAYLOAD]).folderTrashed === true) || row;
  });
  check('C9 결재 중인 최종(-03)이 그대로라 잠금은 유지', () => { const f = n.call('getInspFormDataForClient', 'tok-prc1', null); return (f.ok === false && f.message.indexOf('결재 진행 중') >= 0) || f; });
}
{
  const n = seed(makeEnv('NEW', readWT));
  n.as(KIM);
  const f = n.call('getInspFormDataForClient', 'tok-prc1', 'tok-i2');
  check('C17 재상신 화면: 제출자 본인에게 canDiscard=true', () => (f.ok && f.resub.canDiscard === true) || f);
  n.as(PARK);
  const g = n.call('getInspFormDataForClient', 'tok-prc1', 'tok-i2');
  check('C17 같은 부서 동료에게는 canDiscard=false (재상신은 그대로 가능)', () => (g.ok && g.resub.canDiscard === false) || g);
  const h = seed(makeEnv('HEAD', readHEAD)); h.as(KIM);
  const fh = h.call('getInspFormDataForClient', 'tok-prc1', 'tok-i2');
  check('C17 나머지 응답은 HEAD와 동일 (canDiscard만 추가)', () => { const x = JSON.parse(JSON.stringify(f)); delete x.resub.canDiscard; return same(fh, x); });
}

// =================================================================
console.log('\n[H] 홈 데이터 — HEAD 대비');
// =================================================================
{
  const p = pair(env => { env.as(KIM); return env.call('getHomeDataForClient'); });
  check('H1 폐기 건이 없으면 홈 응답은 HEAD와 같음 (insps[].submitter만 추가)', () => {
    const x = JSON.parse(JSON.stringify(p.rn));
    (x.completedDocs || []).forEach(c => (c.insps || []).forEach(s => { delete s.submitter; }));
    return same(p.rh, x);
  });
  check('H1 submitter = 결재란 0번 이메일(소문자)', () => {
    const c = p.rn.completedDocs.find(d => d.prcToken === 'tok-prc1');
    return same(['kim@inlct.com', 'kim@inlct.com', 'park@inlct.com'], c.insps.map(s => s.submitter));
  });
  const n = p.n; n.as(ADMIN); n.call('adminForceDiscardForClient', { docType: 'INSP', token: 'tok-i3', reason: '정리' });
  n.as(KIM); const after = n.call('getHomeDataForClient');
  check('H2 최종 검수(-03) 폐기 후: inspLocked=false, 폐기 행은 목록에 남고 상태 "폐기"', () => {
    const c = after.completedDocs.find(d => d.prcToken === 'tok-prc1');
    const s3 = c.insps.find(s => s.seq === 3);
    return (c.inspLocked === false && c.inspClosed === false && s3.status === '폐기') || [c.inspLocked, s3];
  });
  check('H3 관리자 처리 대기(INSP PDF) 목록 영향 없음', () => { n.as(ADMIN); return same([], n.call('listInspAwaitingPdf')); });
}

// =================================================================
console.log('\n[W] 락 해제 실패 (2026-10-15 TO-PO-26-318) — 해제가 실패해도 결과는 정상 해제와 같아야 한다');
// =================================================================
// 수정 직전 코드 — withLock 해제가 실패하면 성공 결과가 버려지던 버전 (재현용으로 고정)
const PREFIX_REF = process.env.PREFIX_REF || 'f322aa5';
const readPreFix = f => execSync('git show ' + PREFIX_REF + ':"' + f + '"', { cwd: SRC, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
const W_CLOCK0 = CLOCK.now;
const SERVER_ERR = "We're sorry, a server error occurred";
{
  const p = seed(makeEnv('PREFIX', readPreFix)); p.heartbeat();
  p.as(KIM); p.state.flags.releaseFail = true;
  let thrown = null;
  try { p.call('processDecisionFromClient', { token: 'tok-req1', decision: 'approve', comment: '' }); } catch (e) { thrown = e; }
  const C = p.sb.COL, row = mainRow(p, 'tok-req1');
  check('W0 [재현] 수정 전: 시트엔 승인이 들어갔는데 서버 오류가 화면으로 올라가고, 메일·DOC_APPROVE는 빠짐', () =>
    (thrown && String(thrown.message).indexOf(SERVER_ERR) === 0 && row[p.sb.reqApprCol(0, 3)] === '승인' && row[C.APPR_IDX] === 1 &&
     p.state.mails.length === 0 && p.bufKeys().length === 0 && p.events().indexOf('DOC_APPROVE') < 0) || [String(thrown), row[C.APPR_IDX], p.state.mails, p.bufKeys()]);
}
function releaseFailFlow(name, expectOk, scenario) {
  CLOCK.now = W_CLOCK0;
  const ok = seed(makeEnv('NEW', readWT)), bad = seed(makeEnv('NEW', readWT));
  ok.heartbeat(); bad.heartbeat();
  const ro = scenario(ok);
  bad.state.flags.releaseFail = true;
  let rb = null, thrown = null;
  try { rb = scenario(bad); } catch (e) { thrown = e; }
  bad.state.flags.releaseFail = false;
  ok.flush(); bad.flush();
  const snap = e => ({ sheets: sheetRows(e), mails: e.state.mails, audit: auditNoId(e.auditRows()), props: Object.fromEntries(e.state.props),
                       folders: e.state.folders, adminErrors: e.state.adminErrors });
  check(name + ' — 해제 실패에도 예외 없이 응답', () => thrown === null || String(thrown));
  check(name + ' — 의도한 경로 (ok=' + expectOk + ')', () => (ro && ro.ok === expectOk && rb && rb.ok === expectOk) || [ro, rb]);
  check(name + ' — 해제 실패가 실제로 일어남', () => bad.state.lock.releaseFails > 0 || 'releaseLock 실패 0회');
  check(name + ' — 반환값 동일', () => same(ro, rb));
  check(name + ' — 시트·메일·시스템로그·큐·Drive 동일', () => same(snap(ok), snap(bad)));
  check(name + ' — 해제 실패 경고 ' + bad.state.lock.releaseFails + '건 남김', () =>
    bad.state.consoleWarns.filter(w => w.indexOf('[Lock] 해제 실패') === 0).length === bad.state.lock.releaseFails || bad.state.consoleWarns);
  return { ok, bad, ro, rb };
}
{
  const r = releaseFailFlow('W1 결재 승인(기안자 단계) — 이번 장애 경로', true, env => { env.as(KIM); return env.call('processDecisionFromClient', { token: 'tok-req1', decision: 'approve', comment: '' }); });
  check('W1 다음 결재자(이결재) 메일 + DOC_APPROVE(처리자=김기안)가 남음', () => {
    const row = r.bad.auditRows().find(x => x[AC.EV] === 'DOC_APPROVE');
    return (r.bad.state.mails.some(m => m.kind === 'approval' && m.to === LEE) && row && row[AC.ACTOR] === KIM) || [r.bad.state.mails, row];
  });
}
{
  const r = releaseFailFlow('W2 결재 최종 승인 — 완료 메일·PDF 큐 등록', true, env => { env.as(LEE); return env.call('processDecisionFromClient', { token: 'tok-req8', decision: 'approve', comment: '' }); });
  check('W2 PDF 큐 등록 + 오탐 "큐 등록 실패" 알림 없음', () => {
    const q = JSON.parse(r.bad.state.props.get('PENDING_JOB_QUEUE') || '[]');
    return (q.some(j => j.token === 'tok-req8') && r.bad.state.adminErrors.every(m => m.indexOf('큐 등록 실패') < 0)) || [q, r.bad.state.adminErrors];
  });
}
releaseFailFlow('W3 결재 반려', true, env => { env.as(LEE); return env.call('processDecisionFromClient', { token: 'tok-req8', decision: 'reject', comment: '단가 재확인' }); });
releaseFailFlow('W4 비결재자 승인 시도(거부 경로)', false, env => { env.as(PARK); return env.call('processDecisionFromClient', { token: 'tok-req1', decision: 'approve' }); });
releaseFailFlow('W5 품의 제출', true, env => { env.as(KIM); return env.call('submitRequisitionForClient', submitPayload(KIM)); });
releaseFailFlow('W6 재상신', true, env => { env.as(KIM); const d = submitPayload(KIM); d.subject = '수정 품의'; return env.call('resubmitRequisitionForClient', { token: 'tok-req5', data: d }); });
releaseFailFlow('W7 기안자 폐기', true, env => { env.as(KIM); return env.call('discardRequisitionForClient', { token: 'tok-req5', reason: '기안자 폐기 (반려 후)' }); });
releaseFailFlow('W8 픽업', true, env => { env.as(BUYER1); return env.call('claimRequestForClient', { parentToken: 'tok-req2' }); });
releaseFailFlow('W9 회수', true, env => { env.as(BUYER1); return env.call('releaseClaimForClient', { parentToken: 'tok-req7' }); });
releaseFailFlow('W10 구매 문서 제출', true, env => { env.as(BUYER1); return env.call('submitPrcForClient', prcPayload()); });
releaseFailFlow('W11 검수보고서 제출', true, env => { env.as(KIM); return env.call('submitInspForClient', inspPayload(KIM)); });
releaseFailFlow('W12 검수 결재 최종 승인', true, env => { env.as(LEE); return env.call('processInspDecisionFromClient', { token: 'tok-i3', decision: 'approve', comment: '' }); });
releaseFailFlow('W13 검수보고서 폐기(제출자)', true, env => { env.as(KIM); return env.call('discardInspForClient', { token: 'tok-i2' }); });
releaseFailFlow('W14 관리자 결재자 변경(INSP)', true, env => { env.as(ADMIN); return env.call('adminChangeApproverForClient', { docType: 'INSP', token: 'tok-i3', stageIdx: 1, newEmail: CHOI, reason: '휴가' }); });
releaseFailFlow('W15 관리자 결재자 변경(REQ)', true, env => { env.as(ADMIN); return env.call('adminChangeApproverForClient', { docType: 'REQ', token: 'tok-req8', stageIdx: 1, newEmail: CHOI, reason: '휴가' }); });
releaseFailFlow('W16 관리자 강제 폐기(REQ)', true, env => { env.as(ADMIN); return env.call('adminForceDiscardForClient', { token: 'tok-req1', reason: '중복' }); });
releaseFailFlow('W17 관리자 강제 폐기(INSP)', true, env => { env.as(ADMIN); return env.call('adminForceDiscardForClient', { docType: 'INSP', token: 'tok-i3', reason: '정리' }); });
releaseFailFlow('W18 관리자 강제 락 해제', true, env => { env.as(ADMIN); return env.call('adminForceReleaseLockForClient', { token: 'tok-req3', reason: '담당자 부재' }); });
releaseFailFlow('W19 큐 트리거 — 최종 승인 PDF 작업 처리', true, env => {
  env.as(LEE); env.call('processDecisionFromClient', { token: 'tok-req8', decision: 'approve', comment: '' });
  const queued = JSON.parse(env.state.props.get('PENDING_JOB_QUEUE') || '[]').filter(j => j.token === 'tok-req8').length;
  env.as(ADMIN); env.call('processQueueTrigger');
  const left = JSON.parse(env.state.props.get('PENDING_JOB_QUEUE') || '[]');   // 끝난 작업은 큐에서 지워진다
  return { ok: queued === 1 && left.length === 0, queued: queued, left: left.map(j => j.status) };
});
{
  // ③ 해제가 실패해도 그 아래 '장시간 점유' 진단 경고는 남아야 한다 (수정 전에는 못 남김)
  const run = read => {
    CLOCK.now = W_CLOCK0;
    const e = seed(makeEnv('X', read)); e.heartbeat(); e.as(KIM);
    e.state.flags.releaseFail = true; e.state.flags.releaseDelayMs = 6000;   // 실제 장애처럼 해제가 약 5초 이상 끌다 실패
    try { e.call('processDecisionFromClient', { token: 'tok-req1', decision: 'approve', comment: '' }); } catch (_) {}
    return e.state.consoleWarns;
  };
  const pre = run(readPreFix), now = run(readWT);
  check('W20 해제가 6초 끌다 실패: 수정본은 "장시간 점유" 진단 경고를 남김 (수정 전엔 없음)', () =>
    (now.some(w => w.indexOf('[Lock] 장시간 점유 6000ms') === 0) && !pre.some(w => w.indexOf('[Lock] 장시간 점유') === 0)) || { pre, now });
}
CLOCK.now = W_CLOCK0;

// ── 화면 테스트용 픽스처 덤프 — 실제 서버 함수가 만든 응답을 HTML 검증(tests/html_insp_discard_test.py)에 그대로 쓴다 ──
if (process.env.DUMP_FIXTURES) {
  const out = {};
  const e1 = seed(makeEnv('NEW', readWT));
  e1.as(KIM);  out.homeKim  = norm(e1.call('getHomeDataForClient'));
  e1.as(PARK); out.homePark = norm(e1.call('getHomeDataForClient'));
  const e2 = seed(makeEnv('NEW', readWT));
  e2.as(ADMIN); e2.call('adminForceDiscardForClient', { docType: 'INSP', token: 'tok-i3', reason: '정리' });
  e2.as(KIM); out.homeKimAfterDiscard = norm(e2.call('getHomeDataForClient'));
  const e3 = seed(makeEnv('NEW', readWT));
  e3.as(ADMIN); out.adminList = norm(e3.call('adminListDiscardableForClient', {}));
  e3.as(KIM);  out.inspFormKimResub  = norm(e3.call('getInspFormDataForClient', 'tok-prc1', 'tok-i2'));
  e3.as(PARK); out.inspFormParkResub = norm(e3.call('getInspFormDataForClient', 'tok-prc1', 'tok-i2'));
  e3.as(KIM);  out.inspFormKimNew    = norm(e3.call('getInspFormDataForClient', 'tok-prc2', null));
  fs.writeFileSync(process.env.DUMP_FIXTURES, JSON.stringify(out, null, 1));
  console.log('픽스처 저장: ' + process.env.DUMP_FIXTURES);
}

// =================================================================
console.log('\n─────────────────────────────');
console.log(`통과 ${pass} / 실패 ${fail}`);
if (fail) console.log('실패 항목:\n  - ' + failures.join('\n  - '));
process.exit(fail ? 1 : 0);
