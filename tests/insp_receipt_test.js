/**
 * 검수보고서 품목별 입고 (INSP v1.1, 2026-11-05) — 로컬 검증 (배포 없이 실제 .gs 함수 호출)
 * (memory: gas-local-test-harness)
 *
 * 무엇을 보나
 *  [R] 품목 장부·제출 검증·시트 기록 — 발주/기입고/잔량, 초과 입고 사유, 최종 검수 미입고 사유, 종결 전용 회차, 잘못된 입력
 *  [S] 반려 회차 누계 제외 + 재상신(본인 행 제외, 53폭 덮어쓰기) — 배포 전 51열 시트에서 열 확장까지
 *  [M] 뷰어·매니페스트 v2 (레거시 회차는 품목 표 없음)
 *  [B] 레거시 회차 백필 임시 래퍼 — dry-run → 실행 → 재실행 멱등, 입고 내역 원문 보존
 *  [N] 비회귀 — 홈 페이로드·레거시 뷰어 응답이 이전 코드(BASE_REF)와 같음
 *
 * 실행: node tests/insp_receipt_test.js
 *       DUMP_FIXTURES=<경로> node tests/insp_receipt_test.js   → tests/html_insp_receipt_test.py 입력
 *       BASE_REF=<커밋> (기본 462cdda = 이 기능 직전 커밋)
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execSync } = require('child_process');

const SRC = path.join(__dirname, '..');
const WT_DIR = process.env.WT_DIR || SRC;
const readWT = f => fs.readFileSync(path.join(WT_DIR, f), 'utf8');
const BASE_REF = process.env.BASE_REF || '462cdda';
const readBASE = f => execSync('git show ' + BASE_REF + ':"' + f + '"', { cwd: SRC, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });

// ── 고정 시계 ────────────────────────────────────────────────────
const RealDate = Date;
const CLOCK = { now: RealDate.parse('2026-11-05T10:00:00+09:00') };
class FakeDate extends RealDate {
  constructor(...a) { if (a.length === 0) super(CLOCK.now); else super(...a); }
  static now() { return CLOCK.now; }
}

// ── 메모리 시트 (audit_insp_discard_test.js와 같은 모형 + 열 수 API) ──
class FakeRange {
  constructor(sheet, row, col, nRows, nCols) { Object.assign(this, { sheet, row, col, nRows, nCols }); }
  getValues() {
    const out = [];
    for (let r = 0; r < this.nRows; r++) { const row = []; for (let c = 0; c < this.nCols; c++) row.push(this.sheet._get(this.row + r, this.col + c)); out.push(row); }
    return out;
  }
  setValues(vals) {
    this.sheet._guard();
    // 실제 Sheets: 시트 최대 열 밖에 setValues → 예외. 배포 전 51열 시트에 53폭 재상신이 바로 쓰면 터지는 경로를 재현한다
    if (this.sheet._strictCols && this.col + this.nCols - 1 > this.sheet.getMaxColumns()) {
      throw new Error('The number of columns in the range must be at most ' + (this.sheet.getMaxColumns() - this.col + 1));
    }
    vals.forEach((rv, r) => rv.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v))); return this;
  }
  setValue(v) { this.sheet._guard(); this.sheet._set(this.row, this.col, v); return this; }
  setFontWeight() { return this; }
  setBackground() { return this; }
  setDataValidation() { return this; }
  createTextFinder(text) {
    const self = this; const opts = { entire: false };
    const finder = {
      matchEntireCell(b) { opts.entire = b; return finder; }, matchCase() { return finder; },
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
  constructor(name) { this.name = name; this.rows = []; this.failWrites = false; this.appendCount = 0; this._maxCols = 0; this._strictCols = false; this.insertCalls = 0; }
  _guard() { if (this.failWrites) throw new Error('시트 쓰기 장애 시뮬레이션 (' + this.name + ')'); }
  _get(r, c) { const row = this.rows[r - 1]; return row ? (row[c - 1] ?? '') : ''; }
  _set(r, c, v) { while (this.rows.length < r) this.rows.push([]); const row = this.rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
  getName() { return this.name; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return this.rows.reduce((m, r) => Math.max(m, r.length), 0); }
  getMaxColumns() { return Math.max(26, this.getLastColumn(), this._maxCols); }
  insertColumnsAfter(after, n) { this._guard(); this.insertCalls++; this._maxCols = this.getMaxColumns() + n; return this; }
  appendRow(vals) { this._guard(); this.appendCount++; this.rows.push(vals.slice()); if (this._strictCols) this._maxCols = Math.max(this._maxCols, vals.length); return this; }
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
const BUYER1 = 'buyer1@inlct.com';
const KIM = 'kim@inlct.com', PARK = 'park@inlct.com', LEE = 'lee@inlct.com';
const NAMES = { [KIM]: '김기안', [PARK]: '박동료', [LEE]: '이결재', [ADMIN]: '강관리', [BUYER1]: '구매1' };
const DEPT = { [KIM]: '개발팀', [PARK]: '개발팀', [LEE]: '경영팀', [ADMIN]: '경영총괄', [BUYER1]: '구매팀' };

// ── 환경 ─────────────────────────────────────────────────────────
function makeEnv(label, read) {
  const state = { label, active: ADMIN, mails: [], adminErrors: [], consoleErrors: [], props: new Map(), cache: new Map(), folders: {}, uuid: 0 };
  const ss = new FakeSpreadsheet();
  function folderRec(id, name) { if (!state.folders[id]) state.folders[id] = { id, name: name || id, trashed: false, files: [] }; return state.folders[id]; }
  function wrapFolder(f) {
    return {
      getId: () => f.id, getName: () => f.name, setTrashed: t => { f.trashed = !!t; }, isTrashed: () => f.trashed,
      createFile: (a) => { const nf = { id: 'file-' + (++state.uuid), name: (a && a.name) || a }; f.files.push(nf); return { getId: () => nf.id, getName: () => nf.name, setTrashed() {} }; },
      getFiles: () => ({ hasNext: () => false, next: () => null }), getFilesByName: () => ({ hasNext: () => false, next: () => null }),
      getParents: () => ({ hasNext: () => false, next: () => null }),
    };
  }
  const propsObj = {
    getProperty: k => (state.props.has(k) ? state.props.get(k) : null),
    setProperty: (k, v) => { state.props.set(k, String(v)); return propsObj; },
    deleteProperty: k => { state.props.delete(k); return propsObj; },
    getKeys: () => Array.from(state.props.keys()), getProperties: () => Object.fromEntries(state.props),
  };
  const cacheObj = { get: k => { const e = state.cache.get(k); return e && e.exp > CLOCK.now ? e.v : null; }, put: (k, v, ttl) => { state.cache.set(k, { v: String(v), exp: CLOCK.now + (ttl || 600) * 1000 }); }, remove: k => { state.cache.delete(k); } };
  const quiet = { log() {}, info() {}, warn() {}, error: (...a) => state.consoleErrors.push(a.map(String).join(' ')) };
  const sandbox = {
    console: quiet, Date: FakeDate, Logger: { log() {} },
    Session: { getActiveUser: () => ({ getEmail: () => state.active }), getEffectiveUser: () => ({ getEmail: () => ADMIN }), getScriptTimeZone: () => 'Asia/Seoul' },
    SpreadsheetApp: { openById: () => ss, flush() {} },
    DriveApp: {
      getFolderById: id => { const f = state.folders[id]; if (!f) throw new Error('No item with the given ID could be found: ' + id); return wrapFolder(f); },
      getFileById: () => ({ setTrashed() {}, getBlob: () => ({ getContentType: () => 'image/jpeg', getBytes: () => [] }) }),
    },
    Utilities: {
      getUuid: () => 'uuid-' + (++state.uuid), newBlob: (content, type, name) => ({ content, type, name }),
      formatDate: (d, tz, fmt) => {
        const p = n => String(n).padStart(2, '0'); const t = new RealDate(d.getTime ? d.getTime() : d);
        const ymd = `${t.getFullYear()}-${p(t.getMonth() + 1)}-${p(t.getDate())}`;
        if (fmt === 'yyyyMMdd_HHmmss') return `${t.getFullYear()}${p(t.getMonth() + 1)}${p(t.getDate())}_${p(t.getHours())}${p(t.getMinutes())}${p(t.getSeconds())}`;
        return fmt.indexOf('HH') >= 0 ? `${ymd} ${p(t.getHours())}:${p(t.getMinutes())}` : ymd;
      },
      sleep() {}, base64Decode: () => [], base64Encode: () => '',
    },
    LockService: {
      getScriptLock: () => ({ waitLock() {}, tryLock() { return true; }, releaseLock() {} }),
      getUserLock: () => ({ tryLock() { return true; }, releaseLock() {} }),
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
  sandbox.__state = state; sandbox.__folderRec = folderRec; sandbox.__wrapFolder = wrapFolder;
  vm.runInContext(`
    sendApprovalEmailWithRetry   = function (ap, data, token, rowNum, idx) { __state.mails.push({ kind: 'approval', to: ap.email, idx: idx }); return true; };
    sendRejectionNoticeWithRetry = function (to) { __state.mails.push({ kind: 'reject', to: to }); return true; };
    sendEmailWithRetry           = function (to, subj) { __state.mails.push({ kind: 'generic', to: to, subj: subj }); return true; };
    notifyAdminError             = function (m) { __state.adminErrors.push(String(m)); };
    getOrCreateFolder            = function (name) { return __wrapFolder(__folderRec('F_' + name, name)); };
    saveSummaryToDrive = function () {}; createFileInFolder = function () {}; saveAttachmentsToFolder = function () { return []; };
    _enqueueQuotePdfIfNeeded_ = function () {}; _notifyQuoteMissingIfNeeded_ = function () {};
    _prepareInspPdfHandoff       = function () { return { ok: true }; };
    _saveInspPhotos              = function () { return []; };
    _sendInspApprovalEmail       = function (ap) { __state.mails.push({ kind: 'insp_approval', to: ap.email }); };
    _sendInspRejectEmail         = function (to) { __state.mails.push({ kind: 'insp_reject', to: to }); };
    _sendInspCompletionEmail     = function (to) { __state.mails.push({ kind: 'insp_completion', to: to }); };
    getDeptByEmailMap            = function () { return __DEPT; };
    _readUserPrefs               = function () { return {}; };
    _collectAdminTasks_          = function () { return { quotePdf: [], po: [], insp: [] }; };
    _isSameDeptEmail             = function (ss, a, b) { var d = __DEPT; var x = d[String(a||'').toLowerCase()], y = d[String(b||'').toLowerCase()]; return !!x && x === y; };
    getApproverListForClient     = function () { return { ok: true, approvers: __APPROVERS }; };
    getCurrentUserForClient      = function () { var e = Session.getActiveUser().getEmail(); return { ok: true, user: { email: e, name: __NAMES[e] || e, dept: __DEPT[e] || '', rank: '' } }; };
  `, sandbox);
  sandbox.__DEPT = DEPT; sandbox.__NAMES = NAMES;
  sandbox.__APPROVERS = [KIM, PARK, LEE].map(e => ({ name: NAMES[e], email: e, dept: DEPT[e], rank: '', active: true }));
  sandbox.CONFIG.PROCUREMENT_TEAM_EMAILS = [ADMIN, BUYER1];
  sandbox.CONFIG.ADMIN_EMAILS = [ADMIN];
  sandbox.CONFIG.GLOBAL_VIEW_EMAILS = [];
  const env = {
    label, sb: sandbox, ss, state,
    as(user) { state.active = user; return env; },
    call(fn, ...args) { return sandbox[fn].apply(sandbox, args); },
    main: () => ss.getSheetByName(sandbox.CONFIG.SHEET_NAME),
    insp: () => ss.getSheetByName(sandbox.CONFIG.INSP_SHEET_NAME),
    inspRow: docNo => env.insp().rows.find(r => r[sandbox.INSP_COL.DOC_NO] === docNo),
    receipt: docNo => { const r = env.inspRow(docNo); return r ? JSON.parse(r[sandbox.INSP_COL.RECEIPT_ITEMS]) : null; },
    auditRows: () => { const s = ss.getSheetByName(sandbox.CONFIG.AUDIT_SHEET_NAME); return s ? s.rows.slice(1) : []; },
  };
  return env;
}

// ── 픽스처 ───────────────────────────────────────────────────────
//  tok-prc1: 품목 1행(볼트 10개). 배포 전 51열 시트에 레거시 회차 3건(최종승인·반려·결재중) — 모두 입고 내역 한 줄만
//  tok-prc2: 품목 2행(Cly 106 / 지그 2). 회차 없음
const ITEMS1 = '1. 볼트 | M8 | 10개 | 1,000 KRW';
const ITEMS2 = '1. Cly(F4PO1-003-A1) | EA | 106개 | 53 USD\n2. 지그 본체 | - | 2개 | 10,000 KRW';
const LEGACY_COLS = 51;   // v1.1 이전 검수보고서목록 폭 (27 + 4×6)
function seed(env, opts) {
  opts = opts || {};
  const sb = env.sb, C = sb.COL;
  const main = sb.getOrCreateSheet(env.ss, sb.CONFIG.SHEET_NAME);
  sb.ensureHeaders(main);
  const ap = (label, email, status, at) => ({ label, name: NAMES[email], email, status: status || '대기', at: at || '' });
  function doc(o) {
    const row = new Array(C.APPR_START).fill('');
    row[C.SUBMIT_AT] = new FakeDate(o.submitAt); row[C.DOC_NO] = o.docNo; row[C.ISSUE_DATE] = '2026-10-01';
    row[C.DRAFTER] = NAMES[o.drafter || KIM]; row[C.DEPT] = DEPT[o.drafter || KIM];
    row[C.SUBJECT] = o.subject || '시험 품목'; row[C.VENDOR_NAME] = '라인테크'; row[C.ITEMS] = o.items || ITEMS1; row[C.TOTAL_AMT] = 10000;
    row[C.STATUS] = o.status; row[C.TOKEN] = o.token; row[C.DRIVE_ID] = '';
    row[C.APPR_COUNT] = o.approvers.length; row[C.APPR_IDX] = o.apprIdx || 0; row[C.RESUB_COUNT] = 0; row[C.REJECT_LOG] = '';
    row[C.ATTACH_LIST] = '[]'; row[C.MOVE_STATUS] = 'STAGING'; row[C.DOC_TYPE] = o.docType || 'REQ'; row[C.PARENT_DOC_ID] = o.parent || '';
    o.approvers.forEach(a => row.push(a.label, a.name, a.email, a.status, a.at, ''));
    main.appendRow(row);
  }
  doc({ docNo: 'TH-AP-26-006', token: 'tok-req6', status: 'PRC생성됨', apprIdx: 2, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE, '승인')], submitAt: '2026-09-20T09:00:00+09:00' });
  doc({ docNo: 'TO-PO-26-901', token: 'tok-prc1', status: '최종승인(PRC)', docType: 'PRC', parent: 'tok-req6', drafter: BUYER1, apprIdx: 2, approvers: [ap('기안자', BUYER1, '승인'), ap('결재자 1', LEE, '승인')], submitAt: '2026-09-21T09:00:00+09:00' });
  doc({ docNo: 'TH-AP-26-009', token: 'tok-req9', status: 'PRC생성됨', apprIdx: 2, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE, '승인')], submitAt: '2026-09-22T09:00:00+09:00', items: ITEMS2 });
  doc({ docNo: 'TO-PO-26-902', token: 'tok-prc2', status: '최종승인(PRC)', docType: 'PRC', parent: 'tok-req9', drafter: BUYER1, apprIdx: 2, approvers: [ap('기안자', BUYER1, '승인'), ap('결재자 1', LEE, '승인')], submitAt: '2026-09-23T09:00:00+09:00', items: ITEMS2 });
  if (opts.extraPrc) opts.extraPrc.forEach(doc);

  // 검수보고서목록 — 배포 전 모습(51열 헤더·행). 수정본 코드가 첫 제출에서 53열로 늘려야 한다
  const I = sb.INSP_COL;
  const isheet = env.ss.insertSheet(sb.CONFIG.INSP_SHEET_NAME);
  const headers = ['submitAt', 'docNo', 'seq', 'prcToken', 'reqToken', 'poNo', 'reqNo', 'subject', 'vendorName', 'drafter', 'drafterName', 'dept', 'issueDate',
    'receivedDate', 'receivedNote', 'verdict', 'isFinal', 'comment', 'status', 'token', 'driveId', 'apprCount', 'apprIdx', 'resubCount', 'rejectLog', 'photoList', 'moveStatus'];
  ['drafter_appr', 'appr1', 'appr2', 'appr3'].forEach(px => headers.push(px + '_label', px + '_name', px + '_email', px + '_status', px + '_processedAt', px + '_comment'));
  isheet.appendRow(headers);
  isheet._strictCols = true; isheet._maxCols = LEGACY_COLS;
  function insp(o) {
    const row = new Array(LEGACY_COLS).fill('');
    row[I.SUBMIT_AT] = new FakeDate(o.submitAt); row[I.DOC_NO] = o.docNo || ('TH-AP-26-006-' + String(o.seq).padStart(2, '0')); row[I.SEQ] = o.seq;
    row[I.PRC_TOKEN] = o.prc || 'tok-prc1'; row[I.REQ_TOKEN] = 'tok-req6'; row[I.PO_NO] = o.poNo || 'TO-PO-26-901'; row[I.REQ_NO] = 'TH-AP-26-006';
    row[I.SUBJECT] = '시험 품목'; row[I.VENDOR_NAME] = '라인테크'; row[I.DRAFTER] = KIM; row[I.DRAFTER_NAME] = NAMES[KIM]; row[I.DEPT] = '개발팀';
    row[I.ISSUE_DATE] = '2026-09-20'; row[I.RECEIVED_DATE] = '2026-10-10'; row[I.RECEIVED_NOTE] = o.note || (o.seq + '차 입고'); row[I.VERDICT] = '합격';
    row[I.IS_FINAL] = o.isFinal ? 'Y' : ''; row[I.STATUS] = o.status; row[I.TOKEN] = o.token; row[I.DRIVE_ID] = o.driveId || '';
    row[I.APPR_COUNT] = o.approvers.length; row[I.APPR_IDX] = o.apprIdx || 0; row[I.RESUB_COUNT] = 0; row[I.REJECT_LOG] = '';
    row[I.PHOTO_LIST] = '[]'; row[I.MOVE_STATUS] = o.moveStatus || 'STAGING';
    o.approvers.forEach((a, k) => { const b = I.APPR_START + k * I.APPR_COLS; row[b] = a.label; row[b + 1] = a.name; row[b + 2] = a.email; row[b + 3] = a.status; row[b + 4] = a.at; row[b + 5] = ''; });
    isheet.appendRow(row);
    if (o.driveId) env.state.folders[o.driveId] = { id: o.driveId, name: 'INSP_' + row[I.DOC_NO], trashed: false, files: [] };
  }
  insp({ seq: 1, token: 'tok-i1', status: '최종승인(INSP)', moveStatus: 'FINAL', submitAt: '2026-10-09T09:00:00+09:00', approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE, '승인')], apprIdx: 2 });
  insp({ seq: 2, token: 'tok-i2', status: '반려', driveId: 'F_INSP_-02', submitAt: '2026-10-10T09:00:00+09:00', approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE, '반려')], apprIdx: 1 });
  insp({ seq: 3, token: 'tok-i3', status: '결재중 (이결재)', driveId: 'F_INSP_-03', submitAt: '2026-10-11T09:00:00+09:00', approvers: [ap('기안자', PARK, '승인'), ap('결재자 1', LEE)], apprIdx: 1 });
  if (opts.extraInsp) opts.extraInsp.forEach(insp);
  return env;
}

// ── 단언 ─────────────────────────────────────────────────────────
let pass = 0, fail = 0; const failures = [];
function check(name, fn) {
  try { const r = fn(); if (r === true) { pass++; console.log('  ✓ ' + name); } else { fail++; failures.push(name); console.log('  ✗ ' + name + '  → ' + (typeof r === 'string' ? r : JSON.stringify(r))); } }
  catch (e) { fail++; failures.push(name); console.log('  ✗ ' + name + '  → 예외: ' + (e && e.stack || e)); }
}
const norm = x => JSON.parse(JSON.stringify(x === undefined ? null : x).replace(/uuid-\d+/g, 'UUID'));
const same = (a, b) => JSON.stringify(norm(a)) === JSON.stringify(norm(b)) || ('\n      BASE=' + JSON.stringify(norm(a)) + '\n      NEW =' + JSON.stringify(norm(b)));
const payload = (who, prcToken, extra) => Object.assign({
  prcToken, resubToken: '', receivedDate: '2026-11-05', verdict: '합격', isFinal: false, comment: '', photos: [],
  approvers: [{ label: '기안자', name: NAMES[who], email: who }, { label: '결재자 1', name: NAMES[LEE], email: LEE }],
}, extra || {});

// =================================================================
console.log('\n[R] 품목 장부 · 제출 검증 · 시트 기록');
// =================================================================
{
  const n = seed(makeEnv('NEW', readWT)); const I = n.sb.INSP_COL;
  n.as(KIM);
  check('R0 구 화면 페이로드(입고 내역 한 줄, receiptItems 없음) → 새로고침 안내, 행 미기록', () => {
    const r = n.call('submitInspForClient', payload(KIM, 'tok-prc2', { receivedNote: '1차 입고' }));
    return (!r.ok && /새로고침/.test(r.message) && n.insp().rows.length === 4 && n.insp().insertCalls === 0) || r;
  });
  let f;
  check('R1 폼: 품목 2행 — 발주 106/2 · 기입고 0 · 잔량 106/2 · 레거시 없음 · 단위 EA', () => {
    f = n.call('getInspFormDataForClient', 'tok-prc2', null);
    const it = f.receipt && f.receipt.items;
    return (f.ok && it && it.length === 2 && it[0].idx === 0 && it[0].name === 'Cly(F4PO1-003-A1)' && it[0].ordered === 106 && it[0].prev === 0 && it[0].remaining === 106
      && it[1].ordered === 2 && it[1].remaining === 2 && f.receipt.legacyRounds.length === 0 && f.receipt.itemsUnavailable === false && f.receipt.unit === 'EA') || f.receipt;
  });
  check('R1b 폼(tok-prc1): 레거시 회차 = 1회차(최종승인)·3회차(결재중) 2건, 반려 2회차 제외, 기입고 0', () => {
    const g = n.call('getInspFormDataForClient', 'tok-prc1', null);
    const lg = g.receipt.legacyRounds;
    return (g.ok && lg.length === 2 && lg.map(x => x.seq).join() === '1,3' && lg[0].receivedNote === '1차 입고' && g.receipt.items[0].prev === 0 && g.receipt.countedRounds === 0) || g.receipt;
  });
  check('R1c 폼: 기존 키(form·resub·currentUser·approverList)는 그대로', () => ['form', 'resub', 'currentUser', 'approverList', 'receipt'].every(k => k in f) || Object.keys(f));
  let s2;
  check('R2 부분 입고(Cly 50, 지그 생략) 중간 회차 → 기록: 53셀 · JSON · 자동 요약 · 사유 없음', () => {
    s2 = n.call('submitInspForClient', payload(KIM, 'tok-prc2', { receiptItems: [{ idx: 0, qty: 50 }] }));
    const row = n.inspRow('TH-AP-26-009-01'); const rc = row && JSON.parse(row[I.RECEIPT_ITEMS]);
    return (s2.ok && row && row.length === 53 && rc.v === 1 && rc.items.length === 2 && rc.items[0].qty === 50 && rc.items[0].prev === 0 && rc.items[0].ordered === 106
      && rc.items[1].qty === 0 && rc.items[1].name === '지그 본체' && row[I.RECEIVED_NOTE] === 'Cly(F4PO1-003-A1) 50EA (50/106)' && row[I.QTY_DIFF_NOTE] === '' && row[I.IS_FINAL] === '') || [s2, row && row.slice(50)];
  });
  check('R2b 배포 전 51열 시트 → 첫 제출이 열 확장 + 헤더(receiptItems·qtyDiffNote) + 완료 표시', () => {
    const sh = n.insp();
    return (sh.insertCalls === 1 && sh.getMaxColumns() === 53 && sh.rows[0][51] === 'receiptItems' && sh.rows[0][52] === 'qtyDiffNote' && n.state.props.get('INSP_EXT_COLS_V1') === '1') || [sh.insertCalls, sh.getMaxColumns(), sh.rows[0].slice(50)];
  });
  check('R2c 메일·감사 로그는 기존대로 (1차 결재자 메일 1통, INSP_SUBMIT)', () => {
    n.call('flushAuditLogBuffer_');
    return (n.state.mails.filter(m => m.kind === 'insp_approval' && m.to === KIM).length === 1 && n.auditRows().some(r => r[4] === 'INSP_SUBMIT' && r[5] === 'TH-AP-26-009-01')) || [n.state.mails, n.auditRows().map(r => r[4])];
  });
  check('R3 다음 회차 폼: 검토중 회차도 누계에 포함 → 기입고 50 · 잔량 56 · countedRounds 1', () => {
    const g = n.call('getInspFormDataForClient', 'tok-prc2', null);
    return (g.ok && g.form.nextSeq === 2 && g.receipt.items[0].prev === 50 && g.receipt.items[0].remaining === 56 && g.receipt.items[1].prev === 0 && g.receipt.countedRounds === 1) || g.receipt;
  });
  check('R4a 초과 입고(50+60=110 > 106) 사유 없음 → 거부, 행 미기록', () => {
    const before = n.insp().rows.length;
    const r = n.call('submitInspForClient', payload(KIM, 'tok-prc2', { receiptItems: [{ idx: 0, qty: 60 }] }));
    return (!r.ok && /초과 입고 사유/.test(r.message) && /110\/106/.test(r.message) && n.insp().rows.length === before) || r;
  });
  check('R4b 초과 입고 + 사유 → 기록, 수량 상이 사유 열에 [초과] 한 줄, 요약은 누계 110/106', () => {
    const r = n.call('submitInspForClient', payload(KIM, 'tok-prc2', { receiptItems: [{ idx: 0, qty: 60, reason: '업체 예비 4EA 추가 납품' }, { idx: 1, qty: 0 }] }));
    const row = n.inspRow('TH-AP-26-009-02');
    return (r.ok && row && row[I.QTY_DIFF_NOTE] === '1. Cly(F4PO1-003-A1) [초과] 발주 106 / 누계 110 — 업체 예비 4EA 추가 납품'
      && row[I.RECEIVED_NOTE] === 'Cly(F4PO1-003-A1) 60EA (110/106)' && JSON.parse(row[I.RECEIPT_ITEMS]).items[0].prev === 50) || [r, row && row.slice(13, 15), row && row[I.QTY_DIFF_NOTE]];
  });
  check('R5a 최종 검수 + 잔량(지그 1/2) 사유 없음 → 거부', () => {
    const r = n.call('submitInspForClient', payload(KIM, 'tok-prc2', { isFinal: true, receiptItems: [{ idx: 1, qty: 1 }] }));
    return (!r.ok && /미입고 사유/.test(r.message) && /잔량 1/.test(r.message)) || r;
  });
  check('R5b 최종 검수 + 잔량 사유 → 기록 isFinal=Y, 사유 열 [미입고], 요약 지그 1EA (1/2) — 이전 회차 초과분(Cly 110/106)은 이번 0이라 사유 불요', () => {
    const r = n.call('submitInspForClient', payload(KIM, 'tok-prc2', { isFinal: true, receiptItems: [{ idx: 1, qty: 1, reason: '생산 취소' }] }));
    const row = n.inspRow('TH-AP-26-009-03');
    return (r.ok && row && row[I.IS_FINAL] === 'Y' && row[I.QTY_DIFF_NOTE] === '2. 지그 본체 [미입고] 발주 2 / 누계 1 — 생산 취소' && row[I.RECEIVED_NOTE] === '지그 본체 1EA (1/2)') || [r, row && row[I.QTY_DIFF_NOTE]];
  });
  check('R5c 최종 회차가 결재 중이면 추가 제출 차단 (기존 잠금 규칙 그대로)', () => {
    const g = n.call('getInspFormDataForClient', 'tok-prc2', null);
    return (!g.ok && /결재 진행 중/.test(g.message)) || g;
  });
}
{
  const n = seed(makeEnv('NEW', readWT)); const I = n.sb.INSP_COL; n.as(KIM);
  check('R6a 입고 수량 전부 0 + 중간 회차 → 거부', () => {
    const r = n.call('submitInspForClient', payload(KIM, 'tok-prc1', { receiptItems: [] }));
    return (!r.ok && /입고 수량이 없습니다/.test(r.message)) || r;
  });
  check('R6b 입고 0 + 최종 + 사유 → 종결 전용 회차 허용, 요약 "입고 없음 — 잔량 종결(사유 기재)"', () => {
    const r = n.call('submitInspForClient', payload(KIM, 'tok-prc1', { isFinal: true, receiptItems: [{ idx: 0, qty: 0, reason: '주문 취소 (업체 통보 11/3)' }] }));
    const row = n.inspRow('TH-AP-26-006-04');
    return (r.ok && row && row[I.RECEIVED_NOTE] === '입고 없음 — 잔량 종결(사유 기재)' && row[I.IS_FINAL] === 'Y' && /\[미입고\] 발주 10 \/ 누계 0 — 주문 취소/.test(row[I.QTY_DIFF_NOTE])) || [r, row && row[I.RECEIVED_NOTE]];
  });
}
{
  const n = seed(makeEnv('NEW', readWT)); n.as(KIM);
  const bad = [
    ['idx 범위 밖', [{ idx: 5, qty: 1 }], /품목 번호/],
    ['소수 수량', [{ idx: 0, qty: 1.5 }], /정수/],
    ['음수 수량', [{ idx: 0, qty: -1 }], /정수/],
    ['문자 수량', [{ idx: 0, qty: 'abc' }], /정수/],
    ['중복 idx', [{ idx: 0, qty: 1 }, { idx: 0, qty: 2 }], /중복/],
    ['receiptItems가 배열 아님', 'x', /새로고침/],
  ];
  bad.forEach(([name, items, re]) => check('R7 잘못된 입력 — ' + name + ' → 거부', () => {
    const r = n.call('submitInspForClient', payload(KIM, 'tok-prc2', { receiptItems: items }));
    return (!r.ok && re.test(r.message) && n.insp().rows.length === 4) || r;
  }));
  check('R7 화면이 보낸 발주 수량·품명은 무시 (서버가 PRC에서 다시 읽음)', () => {
    const r = n.call('submitInspForClient', payload(KIM, 'tok-prc2', { receiptItems: [{ idx: 0, qty: 106, ordered: 1, name: '가짜' }] }));
    const rc = n.receipt('TH-AP-26-009-01');
    return (r.ok && rc.items[0].ordered === 106 && rc.items[0].name === 'Cly(F4PO1-003-A1)') || [r, rc];
  });
  check('R7 사유는 200자로 자른다', () => {
    const r = n.call('submitInspForClient', payload(KIM, 'tok-prc2', { receiptItems: [{ idx: 0, qty: 1, reason: '가'.repeat(300) }] }));
    return (r.ok && n.receipt('TH-AP-26-009-02').items[0].reason.length === 200) || r;
  });
  check('R7 스키마 자가 점검 통과 (확장 열 51·52, 총 53)', () => n.call('_inspSchemaSelfCheck') === true);
  check('R7 runInspReceiptMigration — 이미 적용된 시트에서 변경 없음, 헤더 보고', () => {
    const r = n.call('runInspReceiptMigration');
    return (r.ok && r.changed === false && r.header.join() === 'receiptItems,qtyDiffNote') || r;
  });
}

// =================================================================
console.log('\n[S] 반려 회차 누계 제외 · 재상신 (51열 시트에서 53폭 덮어쓰기)');
// =================================================================
{
  const n = seed(makeEnv('NEW', readWT)); const I = n.sb.INSP_COL;
  n.as(KIM); const s = n.call('submitInspForClient', payload(KIM, 'tok-prc2', { receiptItems: [{ idx: 0, qty: 50 }] }));
  check('S1 1회차 제출 → 기입고 50', () => s.ok && n.call('getInspFormDataForClient', 'tok-prc2', null).receipt.items[0].prev === 50);
  check('S2 기안자 승인 → 결재자 반려 → 반려 회차는 누계에서 빠짐 (기입고 0)', () => {
    n.as(KIM); const a = n.call('processInspDecisionFromClient', { token: s.token, decision: 'approve', comment: '' });
    n.as(LEE); const j = n.call('processInspDecisionFromClient', { token: s.token, decision: 'reject', comment: '수량 확인' });
    n.as(KIM); const g = n.call('getInspFormDataForClient', 'tok-prc2', null);
    return (a.ok && j.ok && n.inspRow('TH-AP-26-009-01')[I.STATUS] === '반려' && g.ok && g.receipt.items[0].prev === 0 && g.form.nextSeq === 2) || [a, j, g.receipt];
  });
  check('S3 재상신 폼: 본인 행 제외한 기입고 0 + 반려 회차의 수량·사유 prefill', () => {
    const g = n.call('getInspFormDataForClient', 'tok-prc2', s.token);
    return (g.ok && g.receipt.items[0].prev === 0 && g.resub.receipt && g.resub.receipt.items[0].qty === 50 && g.resub.receipt.items[0].reason === '') || g;
  });
  check('S4 재상신(수량 30으로 정정) → 같은 행 53폭 덮어쓰기 · RESUB_COUNT 1 · 행 수 불변 · 열 확장은 1회만', () => {
    const r = n.call('submitInspForClient', payload(KIM, 'tok-prc2', { resubToken: s.token, receiptItems: [{ idx: 0, qty: 30 }] }));
    const row = n.inspRow('TH-AP-26-009-01');
    return (r.ok && n.insp().rows.length === 5 && row.length === 53 && row[I.RESUB_COUNT] === 1 && row[I.STATUS] === '검토중' && JSON.parse(row[I.RECEIPT_ITEMS]).items[0].qty === 30
      && row[I.RECEIVED_NOTE] === 'Cly(F4PO1-003-A1) 30EA (30/106)' && n.insp().insertCalls === 1) || [r, row && row.length, n.insp().insertCalls];
  });
  check('S5 재상신 후 폼: 기입고 30', () => n.call('getInspFormDataForClient', 'tok-prc2', null).receipt.items[0].prev === 30);
  check('S6 폐기된 회차도 누계 제외', () => {
    n.as(KIM); n.call('processInspDecisionFromClient', { token: s.token, decision: 'approve', comment: '' });
    n.as(LEE); n.call('processInspDecisionFromClient', { token: s.token, decision: 'reject', comment: '다시' });
    n.as(KIM); const d = n.call('discardInspForClient', { token: s.token, reason: '정리' });
    const g = n.call('getInspFormDataForClient', 'tok-prc2', null);
    return (d.ok && g.receipt.items[0].prev === 0 && g.form.nextSeq === 2) || [d, g.receipt];
  });
}

// =================================================================
console.log('\n[M] 뷰어 · 매니페스트 v2');
// =================================================================
{
  const n = seed(makeEnv('NEW', readWT));
  n.as(KIM); const s = n.call('submitInspForClient', payload(KIM, 'tok-prc2', { isFinal: true, receiptItems: [{ idx: 0, qty: 110, reason: '예비분 포함' }, { idx: 1, qty: 1, reason: '생산 취소' }] }));
  check('M1 뷰어: 품목별 입고 (over/short 상태·누계·잔량·사유) + 수량 상이 사유 텍스트', () => {
    n.as(LEE); const v = n.call('getInspViewerDataForClient', s.token, null);
    const it = v.doc.receipt && v.doc.receipt.items;
    return (v.ok && it && it[0].state === 'over' && it[0].cum === 110 && it[0].remaining === 0 && it[1].state === 'short' && it[1].remaining === 1 && it[1].reason === '생산 취소'
      && /\[초과\]/.test(v.doc.qtyDiffNote) && /\[미입고\]/.test(v.doc.qtyDiffNote) && v.doc.receipt.unit === 'EA') || v.doc;
  });
  check('M2 뷰어(레거시 회차): receipt=null · qtyDiffNote 빈칸 · 입고 내역 한 줄 그대로', () => {
    const v = n.call('getInspViewerDataForClient', 'tok-i1', null);
    return (v.ok && v.doc.receipt === null && v.doc.qtyDiffNote === '' && v.doc.receivedNote === '1차 입고') || v.doc;
  });
  check('M3 매니페스트 v2: receiptItems·qtyUnit·qtyDiffNote + 기존 키 전부', () => {
    const row = n.inspRow('TH-AP-26-009-01');
    const m = n.sb._buildInspManifest(n.ss, row);
    const v1Keys = ['token', 'docNo', 'seq', 'isFinal', 'reqNo', 'poNo', 'subject', 'vendorName', 'drafterName', 'dept', 'issueDate', 'receivedDate', 'receivedNote', 'verdict', 'comment', 'approvers', 'photos', 'finalFolderId', 'generatedAt'];
    return (m.schemaVersion === 'insp-manifest-v2' && v1Keys.every(k => k in m) && m.receiptItems.length === 2 && m.receiptItems[0].state === 'over' && m.qtyUnit === 'EA' && /생산 취소/.test(m.qtyDiffNote)) || m;
  });
  check('M4 매니페스트(레거시 회차): receiptItems=null', () => {
    const m = n.sb._buildInspManifest(n.ss, n.inspRow('TH-AP-26-006-01'));
    return (m.schemaVersion === 'insp-manifest-v2' && m.receiptItems === null && m.receivedNote === '1차 입고') || m;
  });
}

// =================================================================
console.log('\n[B] 레거시 회차 백필 임시 래퍼');
// =================================================================
// 래퍼는 2026-11-05 운영 실행(11행 백필) 후 커밋에서 제거됐다. 다시 넣으면(다른 백필) 이 블록이 그대로 검증한다.
if (typeof makeEnv('PROBE', readWT).sb._tmp_backfillInspReceipt_20261105 !== 'function') {
  console.log('  (생략 — 백필 래퍼 제거됨 2026-11-05)');
} else {
  // 실제 백필 대상 중 TO-PO-26-300 두 회차를 흉내낸다 (품목 4행, 1회차 CLx1 100 / 2회차 CLx2 100)
  const ITEMS300 = '1. Small Mirror (M1) | EA | 400개 | 10.5 USD\n2. Large Mirror (M2) | EA | 200개 | 20.57 USD\n3. Cylindrical Lens X (CLx1) | EA | 100개 | 19.65 USD\n4. Cylindrical Lens (CLx2) | EA | 100개 | 24.2 USD';
  const n = seed(makeEnv('NEW', readWT), {
    extraPrc: [{ docNo: 'TO-PO-26-300', token: 'tok-prc300', status: '최종승인(PRC)', docType: 'PRC', parent: 'tok-req6', drafter: BUYER1, apprIdx: 2,
      approvers: [{ label: '기안자', name: NAMES[BUYER1], email: BUYER1, status: '승인', at: '' }, { label: '결재자 1', name: NAMES[LEE], email: LEE, status: '승인', at: '' }], submitAt: '2026-10-20T09:00:00+09:00', items: ITEMS300 }],
    extraInsp: [
      { docNo: 'TO-AP-26-028-01', seq: 1, prc: 'tok-prc300', poNo: 'TO-PO-26-300', token: 'tok-b1', status: '최종승인(INSP)', moveStatus: 'FINAL', note: 'CLx1 100ea', submitAt: '2026-10-29T09:00:00+09:00', approvers: [{ label: '기안자', name: NAMES[KIM], email: KIM, status: '승인', at: '' }], apprIdx: 1 },
      { docNo: 'TO-AP-26-028-02', seq: 2, prc: 'tok-prc300', poNo: 'TO-PO-26-300', token: 'tok-b2', status: '최종승인(INSP)', moveStatus: 'FINAL', note: 'CLx2 100ea', submitAt: '2026-10-29T10:00:00+09:00', approvers: [{ label: '기안자', name: NAMES[KIM], email: KIM, status: '승인', at: '' }], apprIdx: 1 },
    ],
  });
  const I = n.sb.INSP_COL;
  check('B1 관리자가 아니면 실행 불가', () => { n.as(KIM); try { n.call('_tmp_backfillInspReceipt_20261105_dryRun'); return '예외 없음'; } catch (e) { return /관리자만/.test(String(e)) || String(e); } });
  check('B2 dry-run: 기록 0 · 대상 2행 PLAN · 나머지 9행 SKIP(행 없음) · 시트 불변', () => {
    n.as(ADMIN); const r = n.call('_tmp_backfillInspReceipt_20261105_dryRun');
    const plans = r.log.filter(l => l.indexOf('PLAN ') === 0), skips = r.log.filter(l => l.indexOf('SKIP ') === 0);
    return (r.ok && r.dryRun && r.written === 0 && plans.length === 2 && skips.length === 9 && /PLAN TO-AP-26-028-01 — Cylindrical Lens X \(CLx1\) 100EA \(100\/100\)/.test(plans[0])
      && n.inspRow('TO-AP-26-028-01').length === LEGACY_COLS && n.insp().insertCalls === 0) || r;
  });
  check('B3 실행: 2행 기록 — idx 2/3에 수량 100, prev 0, 입고 내역 원문 보존, 열 확장·헤더, 감사 로그 1건', () => {
    const r = n.call('_tmp_backfillInspReceipt_20261105');
    const a = n.receipt('TO-AP-26-028-01'), b = n.receipt('TO-AP-26-028-02');
    return (r.ok && r.written === 2 && a.items.length === 4 && a.items[2].qty === 100 && a.items[2].prev === 0 && a.items[0].qty === 0 && /backfilled/.test(JSON.stringify(a))
      && b.items[3].qty === 100 && b.items[2].prev === 100 && b.items[2].qty === 0
      && n.inspRow('TO-AP-26-028-01')[I.RECEIVED_NOTE] === 'CLx1 100ea' && n.inspRow('TO-AP-26-028-01')[I.QTY_DIFF_NOTE] === ''
      && n.insp().rows[0][51] === 'receiptItems' && n.auditRows().filter(x => x[4] === 'ADMIN_INSP_RECEIPT_BACKFILL').length === 1) || [r, a, b];
  });
  check('B4 재실행: 이미 기록된 행은 SKIP, 기록 0 (멱등)', () => {
    const r = n.call('_tmp_backfillInspReceipt_20261105');
    return (r.ok && r.written === 0 && r.log.filter(l => /이미 품목별 입고 기록 있음/.test(l)).length === 2) || r;
  });
  check('B5 백필 후 다음 회차 폼: 기입고 CLx1 100·CLx2 100 (잔량 0), M1 400·M2 200 남음, 레거시 배너 없음', () => {
    n.as(KIM); const g = n.call('getInspFormDataForClient', 'tok-prc300', null);
    const it = g.receipt.items;
    return (g.ok && it[2].prev === 100 && it[2].remaining === 0 && it[3].prev === 100 && it[0].remaining === 400 && it[1].remaining === 200 && g.receipt.legacyRounds.length === 0 && g.receipt.countedRounds === 2) || g.receipt;
  });
}

// =================================================================
console.log('\n[N] 비회귀 — 이전 코드(' + BASE_REF + ')와 같은 응답');
// =================================================================
{
  const b = seed(makeEnv('BASE', readBASE)), n = seed(makeEnv('NEW', readWT));
  [KIM, PARK, BUYER1, ADMIN].forEach(u => check('N1 홈 페이로드 동일 — ' + NAMES[u], () => same(b.as(u).call('getHomeDataForClient'), n.as(u).call('getHomeDataForClient'))));
  check('N2 레거시 회차 뷰어 응답 동일 (receipt·qtyDiffNote 키 제외)', () => {
    const vb = b.as(LEE).call('getInspViewerDataForClient', 'tok-i3', null), vn = n.as(LEE).call('getInspViewerDataForClient', 'tok-i3', null);
    const x = JSON.parse(JSON.stringify(vn)); delete x.doc.receipt; delete x.doc.qtyDiffNote;
    return same(vb, x);
  });
  check('N3 레거시 회차 결재(승인·반려)는 시트 행이 51폭 그대로, 결과 동일', () => {
    const rb = b.as(LEE).call('processInspDecisionFromClient', { token: 'tok-i3', decision: 'approve', comment: '' });
    const rn = n.as(LEE).call('processInspDecisionFromClient', { token: 'tok-i3', decision: 'approve', comment: '' });
    return (same(rb, rn) === true && same(b.insp().rows, n.insp().rows) === true && n.inspRow('TH-AP-26-006-03').length === LEGACY_COLS) || [rb, rn];
  });
  check('N4 폼 응답은 receipt 키만 추가 (나머지 동일)', () => {
    const fb = b.as(KIM).call('getInspFormDataForClient', 'tok-prc2', null), fn = n.as(KIM).call('getInspFormDataForClient', 'tok-prc2', null);
    const x = JSON.parse(JSON.stringify(fn)); delete x.receipt;
    return same(fb, x);
  });
  check('N5 레거시 회차 매니페스트: 이전과 같은 값 + v2 키', () => {
    const mb = b.sb._buildInspManifest(b.ss, b.inspRow('TH-AP-26-006-01')), mn = n.sb._buildInspManifest(n.ss, n.inspRow('TH-AP-26-006-01'));
    const x = JSON.parse(JSON.stringify(mn)); delete x.receiptItems; delete x.qtyUnit; delete x.qtyDiffNote; x.schemaVersion = 'insp-manifest-v1';
    return same(mb, x);
  });
}

// ── 화면 테스트용 픽스처 덤프 (tests/html_insp_receipt_test.py) ──
if (process.env.DUMP_FIXTURES) {
  const out = {};
  const e = seed(makeEnv('NEW', readWT));
  e.as(KIM); out.formNew2 = norm(e.call('getInspFormDataForClient', 'tok-prc2', null));            // 2품목, 회차 없음
  out.formLegacy = norm(e.call('getInspFormDataForClient', 'tok-prc1', null));                      // 레거시 회차 2건 배너
  const s = e.call('submitInspForClient', payload(KIM, 'tok-prc2', { receiptItems: [{ idx: 0, qty: 50 }] }));
  out.formPartial = norm(e.call('getInspFormDataForClient', 'tok-prc2', null));                     // 기입고 50 / 잔량 56
  e.as(KIM); e.call('processInspDecisionFromClient', { token: s.token, decision: 'approve', comment: '' });
  e.as(LEE); e.call('processInspDecisionFromClient', { token: s.token, decision: 'reject', comment: '수량 확인' });
  e.as(KIM); out.formResub = norm(e.call('getInspFormDataForClient', 'tok-prc2', s.token));         // 재상신 prefill(qty 50)
  const e2 = seed(makeEnv('NEW', readWT));
  e2.as(KIM); const s2 = e2.call('submitInspForClient', payload(KIM, 'tok-prc2', { isFinal: true, receiptItems: [{ idx: 0, qty: 110, reason: '예비분 포함' }, { idx: 1, qty: 1, reason: '생산 취소' }] }));
  e2.as(LEE); out.viewerStructured = norm(e2.call('getInspViewerDataForClient', s2.token, null));
  out.viewerLegacy = norm(e2.call('getInspViewerDataForClient', 'tok-i1', null));
  fs.writeFileSync(process.env.DUMP_FIXTURES, JSON.stringify(out, null, 1));
  console.log('픽스처 저장: ' + process.env.DUMP_FIXTURES);
}

console.log('\n─────────────────────────────');
console.log(`통과 ${pass} / 실패 ${fail}   (기준 ${BASE_REF})`);
if (fail) console.log('실패 항목:\n  - ' + failures.join('\n  - '));
process.exit(fail ? 1 : 0);
