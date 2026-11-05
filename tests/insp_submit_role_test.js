/**
 * 검수보고서 제출 권한 — SCM 역할 (2026-11-05) — 로컬 검증 (배포 없이 실제 .gs 함수 호출)
 * (memory: gas-local-test-harness)
 *
 * 무엇을 보나
 *  [P] 판정 함수 _inspSubmitBasis_ — 'self'|'dept'|'scm'|'' 진리표 (좁은 근거 우선)
 *  [H] 홈 '결재 완료' — SCM은 전 품의 노출·canSubmit, 관리자·전사조회(비SCM)는 조회만, '제출 가능하면 보인다'
 *  [F] 작성 폼 열기 · [S] 제출 서버 검증 — 진입 경로별로 같은 판정, 감사 로그 payload.submitBasis
 *  [R] 재상신은 SCM 허용 / 폐기는 제출자 본인만 (2026-10-14 결정 유지)
 *  [I] 신원 폴백 차단 — 접속자 확인 실패 시 배포 계정(SCM)으로 관문이 열리지 않음 (+ 뮤테이션으로 관문 효력 확인)
 *  [L] 역할 판정 헬퍼(_isListedEmail_) 리팩터 — isAdmin/isGlobalViewer/isProcurement 결과가 이전 코드와 같음
 *  [N] 비회귀 — SCM이 아닌 사용자의 홈·폼·제출이 이전 코드와 같음 (user.isScm 키만 추가)
 *
 * 실행: node tests/insp_submit_role_test.js
 *       BASE_REF=<커밋> (기본 6b2e9eb = 이 기능 직전 커밋)
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execSync } = require('child_process');

const SRC = path.join(__dirname, '..');
const WT_DIR = process.env.WT_DIR || SRC;
const readWT = f => fs.readFileSync(path.join(WT_DIR, f), 'utf8');
const BASE_REF = process.env.BASE_REF || '6b2e9eb';
const readBASE = f => execSync('git show ' + BASE_REF + ':"' + f + '"', { cwd: SRC, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });

// ── 고정 시계 ────────────────────────────────────────────────────
const RealDate = Date;
const CLOCK = { now: RealDate.parse('2026-11-05T10:00:00+09:00') };
class FakeDate extends RealDate {
  constructor(...a) { if (a.length === 0) super(CLOCK.now); else super(...a); }
  static now() { return CLOCK.now; }
}

// ── 메모리 시트 (insp_receipt_test.js와 같은 모형) ────────────────
class FakeRange {
  constructor(sheet, row, col, nRows, nCols) { Object.assign(this, { sheet, row, col, nRows, nCols }); }
  getValues() {
    const out = [];
    for (let r = 0; r < this.nRows; r++) { const row = []; for (let c = 0; c < this.nCols; c++) row.push(this.sheet._get(this.row + r, this.col + c)); out.push(row); }
    return out;
  }
  setValues(vals) { vals.forEach((rv, r) => rv.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v))); return this; }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
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
  constructor(name) { this.name = name; this.rows = []; this._maxCols = 0; }
  _get(r, c) { const row = this.rows[r - 1]; return row ? (row[c - 1] ?? '') : ''; }
  _set(r, c, v) { while (this.rows.length < r) this.rows.push([]); const row = this.rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
  getName() { return this.name; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return this.rows.reduce((m, r) => Math.max(m, r.length), 0); }
  getMaxColumns() { return Math.max(26, this.getLastColumn(), this._maxCols); }
  insertColumnsAfter(after, n) { this._maxCols = this.getMaxColumns() + n; return this; }
  appendRow(vals) { this.rows.push(vals.slice()); return this; }
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
//  역할은 운영 CONFIG와 같은 모양으로 주입한다 — davidkang = 관리자+구매팀+SCM, logistics = SCM.
//  logistics는 운영에서 전사조회이기도 하지만 여기서는 빼서 'SCM만으로도 결재 완료 목록에 보이는가'를 본다.
const ADMIN  = 'davidkang@inlct.com';   // 관리자 + 구매팀 + SCM (배포 계정 — getEffectiveUser가 돌려주는 값)
const SCM    = 'logistics@inlct.com';   // SCM 전용
const ADM2   = 'admin2@inlct.com';      // 관리자 (SCM 아님)
const GV     = 'bret@inlct.com';        // 전사조회 (SCM 아님)
const BUYER1 = 'buyer1@inlct.com';      // 구매팀
const KIM = 'kim@inlct.com', PARK = 'park@inlct.com', CHOI = 'choi@inlct.com', LEE = 'lee@inlct.com';
const NAMES = { [ADMIN]: '강관리', [SCM]: '물류', [ADM2]: '관리2', [GV]: '강전사', [BUYER1]: '구매1', [KIM]: '김기안', [PARK]: '박동료', [CHOI]: '최영업', [LEE]: '이결재' };
const DEPT  = { [ADMIN]: '경영총괄', [SCM]: '물류팀', [ADM2]: '경영총괄', [GV]: '경영총괄', [BUYER1]: '구매팀', [KIM]: '개발팀', [PARK]: '개발팀', [CHOI]: '영업팀', [LEE]: '경영팀' };

// ── 환경 ─────────────────────────────────────────────────────────
function makeEnv(label, read) {
  //  active: 접속자 이메일. '' = Google이 신원을 주지 않음, THROW = getActiveUser()가 예외
  const state = { label, active: KIM, mails: [], adminErrors: [], props: new Map(), cache: new Map(), folders: {}, uuid: 0, deptMapCalls: 0 };
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
  const quiet = { log() {}, info() {}, warn() {}, error() {} };
  const sandbox = {
    console: quiet, Date: FakeDate, Logger: { log() {} },
    Session: {
      getActiveUser: () => ({ getEmail: () => { if (state.active === 'THROW') throw new Error('신원 조회 실패 시뮬레이션'); return state.active; } }),
      getEffectiveUser: () => ({ getEmail: () => ADMIN }),   // executeAs: USER_DEPLOYING — 항상 배포 계정
      getScriptTimeZone: () => 'Asia/Seoul',
    },
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
  const realScm = sandbox.CONFIG.SCM_EMAILS;   // 운영 설정 확인용 (아래에서 픽스처 값으로 덮기 전)
  sandbox.__state = state; sandbox.__folderRec = folderRec; sandbox.__wrapFolder = wrapFolder;
  sandbox.__DEPT = DEPT; sandbox.__NAMES = NAMES;
  // ⚠ _isSameDeptEmail은 대역으로 바꾸지 않는다 — 실제 함수가 deptMap을 받는지(홈 N+1)를 getDeptByEmailMap 호출 수로 본다
  vm.runInContext(`
    sendApprovalEmailWithRetry   = function (ap) { __state.mails.push({ kind: 'approval', to: ap.email }); return true; };
    sendRejectionNoticeWithRetry = function (to) { __state.mails.push({ kind: 'reject', to: to }); return true; };
    sendEmailWithRetry           = function (to, subj) { __state.mails.push({ kind: 'generic', to: to, subj: subj }); return true; };
    notifyAdminError             = function (m) { __state.adminErrors.push(String(m)); };
    getOrCreateFolder            = function (name) { return __wrapFolder(__folderRec('F_' + name, name)); };
    saveSummaryToDrive = function () {}; createFileInFolder = function () {}; saveAttachmentsToFolder = function () { return []; };
    _prepareInspPdfHandoff       = function () { return { ok: true }; };
    _saveInspPhotos              = function () { return []; };
    _sendInspApprovalEmail       = function (ap) { __state.mails.push({ kind: 'insp_approval', to: ap.email }); };
    _sendInspRejectEmail         = function (to) { __state.mails.push({ kind: 'insp_reject', to: to }); };
    _sendInspCompletionEmail     = function (to) { __state.mails.push({ kind: 'insp_completion', to: to }); };
    getDeptByEmailMap            = function () { __state.deptMapCalls++; return __DEPT; };
    _readUserPrefs               = function () { return {}; };
    _collectAdminTasks_          = function () { return { quotePdf: [], po: [], insp: [] }; };
    getApproverListForClient     = function () { return { ok: true, approvers: [] }; };
    getCurrentUserForClient      = function () { var e = ''; try { e = Session.getActiveUser().getEmail(); } catch (_) {} return { ok: true, user: { email: e, name: __NAMES[e] || e, dept: __DEPT[e] || '', rank: '' } }; };
  `, sandbox);
  sandbox.CONFIG.PROCUREMENT_TEAM_EMAILS = [ADMIN, BUYER1];
  sandbox.CONFIG.ADMIN_EMAILS = [ADMIN, ADM2];
  sandbox.CONFIG.GLOBAL_VIEW_EMAILS = [GV];
  sandbox.CONFIG.SCM_EMAILS = [ADMIN, SCM];   // 이전 코드(BASE)에는 이 설정을 읽는 곳이 없다
  const env = {
    label, sb: sandbox, ss, state, realScm,
    as(user) { state.active = user; return env; },
    call(fn, ...args) { return sandbox[fn].apply(sandbox, args); },
    insp: () => ss.getSheetByName(sandbox.CONFIG.INSP_SHEET_NAME),
    inspRow: docNo => env.insp().rows.find(r => r[sandbox.INSP_COL.DOC_NO] === docNo),
    auditRows: () => { const s = ss.getSheetByName(sandbox.CONFIG.AUDIT_SHEET_NAME); return s ? s.rows.slice(1) : []; },
    flush: () => { if (typeof sandbox.flushAuditLogBuffer_ === 'function') sandbox.flushAuditLogBuffer_(); return env; },
  };
  return env;
}

// ── 픽스처 ───────────────────────────────────────────────────────
//  tok-prcA: 김기안(개발팀) 품의의 구매 — 최종승인(PRC)
//  tok-prcB: 최영업(영업팀) 품의의 구매 — 최종승인(PRC)
//  tok-prcC: 원본 품의가 없는 구매(기안자 이메일 빈칸) — SCM이어도 제출 불가여야 한다
const ITEMS1 = '1. 볼트 | M8 | 10개 | 1,000 KRW';
function seed(env) {
  const sb = env.sb, C = sb.COL;
  const main = sb.getOrCreateSheet(env.ss, sb.CONFIG.SHEET_NAME);
  sb.ensureHeaders(main);
  const ap = (label, email, status) => ({ label, name: NAMES[email], email, status: status || '승인' });
  function doc(o) {
    const row = new Array(C.APPR_START).fill('');
    row[C.SUBMIT_AT] = new FakeDate(o.submitAt); row[C.DOC_NO] = o.docNo; row[C.ISSUE_DATE] = '2026-10-01';
    row[C.DRAFTER] = NAMES[o.drafter]; row[C.DEPT] = DEPT[o.drafter];
    row[C.SUBJECT] = '시험 품목'; row[C.VENDOR_NAME] = '라인테크'; row[C.ITEMS] = ITEMS1; row[C.TOTAL_AMT] = 10000;
    row[C.STATUS] = o.status; row[C.TOKEN] = o.token; row[C.DRIVE_ID] = '';
    row[C.APPR_COUNT] = o.approvers.length; row[C.APPR_IDX] = o.approvers.length; row[C.RESUB_COUNT] = 0; row[C.REJECT_LOG] = '';
    row[C.ATTACH_LIST] = '[]'; row[C.MOVE_STATUS] = 'STAGING'; row[C.DOC_TYPE] = o.docType || 'REQ'; row[C.PARENT_DOC_ID] = o.parent || '';
    o.approvers.forEach(a => row.push(a.label, a.name, a.email, a.status, '', ''));
    main.appendRow(row);
  }
  doc({ docNo: 'TH-AP-26-101', token: 'tok-reqA', status: 'PRC생성됨', drafter: KIM,  approvers: [ap('기안자', KIM),  ap('결재자 1', LEE)], submitAt: '2026-10-01T09:00:00+09:00' });
  doc({ docNo: 'TO-PO-26-901', token: 'tok-prcA', status: '최종승인(PRC)', docType: 'PRC', parent: 'tok-reqA', drafter: BUYER1, approvers: [ap('기안자', BUYER1), ap('결재자 1', LEE)], submitAt: '2026-10-02T09:00:00+09:00' });
  doc({ docNo: 'TS-AP-26-102', token: 'tok-reqB', status: 'PRC생성됨', drafter: CHOI, approvers: [ap('기안자', CHOI), ap('결재자 1', LEE)], submitAt: '2026-10-03T09:00:00+09:00' });
  doc({ docNo: 'TO-PO-26-902', token: 'tok-prcB', status: '최종승인(PRC)', docType: 'PRC', parent: 'tok-reqB', drafter: BUYER1, approvers: [ap('기안자', BUYER1), ap('결재자 1', LEE)], submitAt: '2026-10-04T09:00:00+09:00' });
  doc({ docNo: 'TO-PO-26-903', token: 'tok-prcC', status: '최종승인(PRC)', docType: 'PRC', parent: 'tok-missing', drafter: BUYER1, approvers: [ap('기안자', BUYER1), ap('결재자 1', LEE)], submitAt: '2026-10-05T09:00:00+09:00' });
  sb.ensureInspSheet();
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
  approvers: [{ label: '기안자', name: NAMES[who] || who, email: who }, { label: '결재자 1', name: NAMES[LEE], email: LEE }],
  receiptItems: [{ idx: 0, qty: 5 }],
}, extra || {});
const pair = () => [seed(makeEnv('BASE', readBASE)), seed(makeEnv('NEW', readWT))];
const home = (env, who) => env.as(who).call('getHomeDataForClient');
const noScmFlag = r => { const x = JSON.parse(JSON.stringify(r)); if (x && x.user) delete x.user.isScm; return x; };
const docsOf = r => (r.completedDocs || []).map(d => d.prcToken + ':' + (d.canSubmit ? 'Y' : 'N')).join(',');
const auditOf = (env, ev, docNo) => env.auditRows().find(r => r[4] === ev && r[5] === docNo);
const basisOf = row => { try { return JSON.parse(row[10]).submitBasis; } catch (_) { return undefined; } };
const I_SUBMITTER = sb => sb.inspApprCol(0, 2);

// =================================================================
console.log('\n[C] 운영 설정');
// =================================================================
{
  const n = makeEnv('NEW', readWT);
  check('C1 Code.gs CONFIG.SCM_EMAILS = davidkang · logistics', () => JSON.stringify(n.realScm) === JSON.stringify([ADMIN, SCM]) || n.realScm);
  check('C2 운영 설정의 전사조회·관리자 명단은 그대로 (SCM은 별도 배열)', () => {
    const fresh = new vm.Script(readWT('Code.gs'));
    const box = { console: { log() {} } }; vm.createContext(box); fresh.runInContext(box);
    return (box.CONFIG.GLOBAL_VIEW_EMAILS.indexOf(SCM) >= 0 && JSON.stringify(box.CONFIG.ADMIN_EMAILS) === JSON.stringify([ADMIN])) || [box.CONFIG.GLOBAL_VIEW_EMAILS, box.CONFIG.ADMIN_EMAILS];
  });
}

// =================================================================
console.log('\n[P] 판정 함수 _inspSubmitBasis_ — 좁은 근거부터');
// =================================================================
{
  const n = seed(makeEnv('NEW', readWT));
  const basis = (a, d) => n.call('_inspSubmitBasis_', n.ss, a, d);
  const TABLE = [
    [KIM, KIM, 'self'], [' Kim@INLCT.com ', KIM, 'self'], [KIM, ' KIM@inlct.com', 'self'],
    [PARK, KIM, 'dept'], [CHOI, KIM, ''], [LEE, KIM, ''],
    [SCM, KIM, 'scm'], [SCM, CHOI, 'scm'], [ADMIN, KIM, 'scm'], ['LOGISTICS@inlct.com', KIM, 'scm'],
    [ADMIN, ADMIN, 'self'],          // SCM이어도 본인 건이면 본인 근거
    [ADM2, KIM, ''], [GV, KIM, ''], [BUYER1, KIM, ''],   // 관리자·전사조회·구매팀은 제출 근거가 아니다
    [SCM, '', ''], [SCM, null, ''], ['', KIM, ''], [null, KIM, ''],
  ];
  TABLE.forEach(([a, d, want]) => check(`P ${JSON.stringify(a)} → 기안자 ${JSON.stringify(d)} = ${JSON.stringify(want)}`, () => basis(a, d) === want || basis(a, d)));
  check('P 같은 부서이면서 SCM이면 dept (좁은 근거 우선)', () => {
    n.sb.CONFIG.SCM_EMAILS = [ADMIN, SCM, PARK];
    const r = basis(PARK, KIM); n.sb.CONFIG.SCM_EMAILS = [ADMIN, SCM];
    return r === 'dept' || r;
  });
  check('P 거부 메시지 상수', () => /SCM 담당/.test(n.sb.INSP_SUBMIT_DENIED_MSG));
}

// =================================================================
console.log('\n[H] 홈 결재 완료 — 노출 + 제출 버튼');
// =================================================================
{
  const [b, n] = pair();
  check('H1 SCM 전용(logistics): 이전 0건 → 기안자 있는 전 품의 노출·제출 가능, 원본 없는 PO-903은 안 보임', () => {
    const hb = home(b, SCM), hn = home(n, SCM);
    return (docsOf(hb) === '' && docsOf(hn) === 'tok-prcB:Y,tok-prcA:Y' && hn.user.isScm === true) || [docsOf(hb), docsOf(hn), hn.user];
  });
  check('H1b SCM 전용: 결재 완료 목록 = 제출 가능 건 (보이는데 못 내거나, 낼 수 있는데 안 보이는 건 없음)', () => {
    const hn = home(n, SCM);
    return hn.completedDocs.every(d => d.canSubmit) || docsOf(hn);
  });
  check('H2 관리자+SCM(davidkang): 목록은 이전과 같은 3건, 기안자 있는 건만 canSubmit N→Y', () => {
    const hb = home(b, ADMIN), hn = home(n, ADMIN);
    return (docsOf(hb) === 'tok-prcC:N,tok-prcB:N,tok-prcA:N' && docsOf(hn) === 'tok-prcC:N,tok-prcB:Y,tok-prcA:Y' && hn.user.isScm === true) || [docsOf(hb), docsOf(hn)];
  });
  check('H2b 관리자+SCM: canSubmit·isScm 외에는 이전 응답과 같음', () => {
    const hb = home(b, ADMIN), hn = noScmFlag(home(n, ADMIN));
    hn.completedDocs.forEach(d => { d.canSubmit = false; });
    return same(hb, hn);
  });
  [[ADM2, '관리자(비SCM)'], [GV, '전사조회(비SCM)'], [BUYER1, '구매팀'], [KIM, '기안자 김기안'], [PARK, '같은 부서 박동료'], [CHOI, '타 부서 최영업']].forEach(([who, label]) =>
    check(`H3 ${label}: 홈 응답이 이전과 같음 (user.isScm=false만 추가)`, () => {
      const hn = home(n, who);
      return (hn.user.isScm === false && same(home(b, who), noScmFlag(hn)) === true) || same(home(b, who), noScmFlag(hn));
    }));
  check('H4 관리자·전사조회(비SCM)는 조회만: 타 부서 건 canSubmit=N', () => {
    const a = home(n, ADM2), g = home(n, GV);
    return (docsOf(a) === 'tok-prcC:N,tok-prcB:N,tok-prcA:N' && docsOf(g) === 'tok-prcC:N,tok-prcB:N,tok-prcA:N') || [docsOf(a), docsOf(g)];
  });
  check('H5 결재자목록(부서 맵) 읽기는 홈 1회에 한 번 — SCM·관리자·일반 모두 (N+1 재발 방지)', () => {
    const counts = [SCM, ADMIN, KIM, CHOI].map(who => { n.state.deptMapCalls = 0; home(n, who); return n.state.deptMapCalls; });
    return counts.every(c => c === 1) || counts;
  });
}

// =================================================================
console.log('\n[F] 작성 폼 열기');
// =================================================================
{
  const [b, n] = pair();
  const form = (env, who, tok, resub) => env.as(who).call('getInspFormDataForClient', tok, resub || null);
  check('F1 SCM → 타 부서(개발팀·영업팀) 품의 폼 열림 (이전 코드는 거부)', () => {
    const r = [form(n, SCM, 'tok-prcA'), form(n, SCM, 'tok-prcB'), form(b, SCM, 'tok-prcA')];
    return (r[0].ok && r[1].ok && !r[2].ok && r[0].form.drafterEmail === KIM && r[0].currentUser.email === SCM) || r.map(x => x.message);
  });
  check('F2 SCM이어도 원본 품의 없는 PO(기안자 빈칸)는 거부', () => {
    const r = form(n, SCM, 'tok-prcC');
    return (!r.ok && r.message === n.sb.INSP_SUBMIT_DENIED_MSG) || r.message;
  });
  [[CHOI, 'tok-prcA', '타 부서 최영업'], [ADM2, 'tok-prcA', '관리자(비SCM)'], [GV, 'tok-prcB', '전사조회(비SCM)'], [BUYER1, 'tok-prcA', '구매팀']].forEach(([who, tok, label]) =>
    check(`F3 ${label} → 거부 (이전 코드도 거부), 메시지는 SCM 포함 문구`, () => {
      const rb = form(b, who, tok), rn = form(n, who, tok);
      return (!rb.ok && !rn.ok && rn.message === n.sb.INSP_SUBMIT_DENIED_MSG) || [rb.message, rn.message];
    }));
  [[KIM, 'tok-prcA', '본인'], [PARK, 'tok-prcA', '같은 부서']].forEach(([who, tok, label]) =>
    check(`F4 ${label} → 열림, 응답이 이전과 같음`, () => same(form(b, who, tok), form(n, who, tok))));
}

// =================================================================
console.log('\n[S] 제출 서버 검증 + 감사 로그 근거');
// =================================================================
{
  const [b, n] = pair(); const I = n.sb.INSP_COL;
  check('S1 SCM이 개발팀 품의 검수보고서 제출 → 기록 (이전 코드는 거부)', () => {
    const rb = b.as(SCM).call('submitInspForClient', payload(SCM, 'tok-prcA'));
    const rn = n.as(SCM).call('submitInspForClient', payload(SCM, 'tok-prcA'));
    return (!rb.ok && rn.ok && rn.docNo === 'TH-AP-26-101-01') || [rb.message, rn.message];
  });
  check('S1b 기록: 결재란 0번 = SCM(실제 제출자), 기안자 스냅샷 = 원본 기안자 김기안, 1차 결재자 메일', () => {
    const row = n.inspRow('TH-AP-26-101-01');
    return (row[I_SUBMITTER(n.sb)] === SCM && row[I.DRAFTER] === KIM && row[I.DEPT] === '개발팀'
      && n.state.mails.some(m => m.kind === 'insp_approval' && m.to === SCM)) || [row[I_SUBMITTER(n.sb)], row[I.DRAFTER], n.state.mails];
  });
  check('S1c 감사 로그 INSP_SUBMIT: 처리자 SCM · payload.submitBasis="scm"', () => {
    const row = auditOf(n.flush(), 'INSP_SUBMIT', 'TH-AP-26-101-01');
    return (row && row[2] === SCM && basisOf(row) === 'scm') || row;
  });
  check('S2 같은 부서 박동료 제출 → submitBasis="dept" · 본인 김기안 → "self"', () => {
    // S1의 5개에 이어 1개씩 (누계가 발주 10을 넘으면 초과 사유 검증에 걸린다)
    const p = n.as(PARK).call('submitInspForClient', payload(PARK, 'tok-prcA', { receiptItems: [{ idx: 0, qty: 1 }] }));
    const k = n.as(KIM).call('submitInspForClient', payload(KIM, 'tok-prcA', { receiptItems: [{ idx: 0, qty: 1 }] }));
    n.flush();
    return (p.ok && k.ok && basisOf(auditOf(n, 'INSP_SUBMIT', p.docNo)) === 'dept' && basisOf(auditOf(n, 'INSP_SUBMIT', k.docNo)) === 'self') || [p, k];
  });
  [[CHOI, 'tok-prcA', '타 부서 최영업'], [ADM2, 'tok-prcB', '관리자(비SCM)'], [GV, 'tok-prcA', '전사조회(비SCM)'], [SCM, 'tok-prcC', 'SCM — 원본 없는 PO']].forEach(([who, tok, label]) =>
    check(`S3 ${label} 제출 → 거부, 행·메일 없음`, () => {
      const before = n.insp().rows.length, mails = n.state.mails.length;
      const r = n.as(who).call('submitInspForClient', payload(who, tok));
      return (!r.ok && r.message === n.sb.INSP_SUBMIT_DENIED_MSG && n.insp().rows.length === before && n.state.mails.length === mails) || r.message;
    }));
  check('S4 결재란 0번을 남으로 바꿔 보내면 거부 (본인 검사는 그대로)', () => {
    const r = n.as(SCM).call('submitInspForClient', payload(KIM, 'tok-prcB'));
    return (!r.ok && /기안자는 본인만/.test(r.message)) || r.message;
  });
}
{
  const [b, n] = pair();
  check('S5 본인 제출: 반환값·시트·메일이 이전 코드와 같음 (감사 payload만 추가)', () => {
    const rb = b.as(KIM).call('submitInspForClient', payload(KIM, 'tok-prcA'));
    const rn = n.as(KIM).call('submitInspForClient', payload(KIM, 'tok-prcA'));
    return [same(rb, rn), same(b.insp().rows, n.insp().rows), same(b.state.mails, n.state.mails)].every(x => x === true) || [rb, rn];
  });
}

// =================================================================
console.log('\n[R] 재상신 · 폐기');
// =================================================================
{
  const n = seed(makeEnv('NEW', readWT)); const I = n.sb.INSP_COL;
  const s = n.as(KIM).call('submitInspForClient', payload(KIM, 'tok-prcA'));
  n.as(KIM).call('processInspDecisionFromClient', { token: s.token, decision: 'approve', comment: '' });
  n.as(LEE).call('processInspDecisionFromClient', { token: s.token, decision: 'reject', comment: '수량 확인' });
  check('R0 준비: 김기안 제출 → 이결재 반려', () => n.inspRow(s.docNo)[I.STATUS] === '반려' || n.inspRow(s.docNo)[I.STATUS]);
  check('R1 SCM이 남이 낸 반려 건 재상신 폼을 연다 — 폐기 버튼은 없음(canDiscard=false)', () => {
    const f = n.as(SCM).call('getInspFormDataForClient', 'tok-prcA', s.token);
    return (f.ok && f.resub && f.resub.canDiscard === false) || f;
  });
  check('R2 SCM이 남이 낸 반려 건을 폐기 → 거부 (제출자 본인만)', () => {
    const d = n.as(SCM).call('discardInspForClient', { token: s.token, reason: '정리' });
    return (!d.ok && /제출자 본인만/.test(d.message) && n.inspRow(s.docNo)[I.STATUS] === '반려') || d;
  });
  check('R3 SCM 재상신 → 같은 행 덮어쓰기, 제출자가 SCM으로 바뀜, INSP_RESUBMIT submitBasis="scm"', () => {
    const r = n.as(SCM).call('submitInspForClient', payload(SCM, 'tok-prcA', { resubToken: s.token, receiptItems: [{ idx: 0, qty: 4 }] }));
    const row = n.inspRow(s.docNo); n.flush();
    const a = auditOf(n, 'INSP_RESUBMIT', s.docNo);
    return (r.ok && row[I.STATUS] === '검토중' && row[I.RESUB_COUNT] === 1 && row[I_SUBMITTER(n.sb)] === SCM && a && a[2] === SCM && basisOf(a) === 'scm') || [r, a];
  });
  check('R4 타 부서 비SCM(최영업)은 재상신 폼도 거부', () => {
    const f = n.as(CHOI).call('getInspFormDataForClient', 'tok-prcA', s.token);
    return (!f.ok && f.message === n.sb.INSP_SUBMIT_DENIED_MSG) || f.message;
  });
}

// =================================================================
console.log('\n[I] 신원 폴백 차단 — 접속자 확인 실패 시 배포 계정(davidkang = SCM)으로 열리지 않음');
// =================================================================
{
  const n = seed(makeEnv('NEW', readWT));
  [['', '빈 값'], ['THROW', '예외']].forEach(([act, label]) => {
    check(`I1 접속자 ${label} → 폼 열기 거부 ("로그인 사용자를 확인할 수 없습니다")`, () => {
      const f = n.as(act).call('getInspFormDataForClient', 'tok-prcB', null);
      return (!f.ok && /로그인 사용자를 확인할 수 없습니다/.test(f.message)) || f;
    });
    check(`I2 접속자 ${label} → 제출 거부, 행·메일 없음 (결재란 0번을 배포 계정으로 채워 보내도)`, () => {
      const before = n.insp().rows.length, mails = n.state.mails.length;
      const r = n.as(act).call('submitInspForClient', payload(ADMIN, 'tok-prcB'));
      return (!r.ok && /로그인 사용자를 확인할 수 없습니다/.test(r.message) && n.insp().rows.length === before && n.state.mails.length === mails) || r;
    });
  });
  // 뮤테이션: 두 관문의 신원 함수를 폴백 있는 쪽으로 되돌리면 같은 요청이 '전 품의'로 열린다 — 위 I1·I2가 실제로 그 차이를 잡는지 확인
  check('I3 [뮤테이션] 신원 함수를 폴백 있는 쪽으로 되돌리면 같은 요청으로 영업팀 품의 폼이 열리고 제출까지 된다', () => {
    const m = seed(makeEnv('MUT', readWT));
    m.sb.getRequestUserEmail_ = function () { return String(m.sb.getActiveUserEmail() || '').toLowerCase(); };
    const f = m.as('').call('getInspFormDataForClient', 'tok-prcB', null);
    const r = m.as('').call('submitInspForClient', payload(ADMIN, 'tok-prcB'));
    m.flush(); const a = auditOf(m, 'INSP_SUBMIT', r.docNo);
    return (f.ok && r.ok && a && a[2] === ADMIN && basisOf(a) === 'scm') || [f.message, r.message];
  });
}

// =================================================================
console.log('\n[L] 역할 판정 헬퍼 리팩터 — 이전 코드와 같은 결과');
// =================================================================
{
  const b = makeEnv('BASE', readBASE), n = makeEnv('NEW', readWT);
  const LISTS = [['davidkang@inlct.com', 'Mixed@InLCT.com'], [], ['x@inlct.com']];
  const EMAILS = ['davidkang@inlct.com', 'DavidKang@INLCT.com', 'Mixed@InLCT.com', 'mixed@inlct.com', 'MIXED@INLCT.COM', ' davidkang@inlct.com', 'x@inlct.com', 'nobody@inlct.com', '', null, undefined];
  const FNS = [['isAdminUser', 'ADMIN_EMAILS'], ['isGlobalViewer', 'GLOBAL_VIEW_EMAILS'], ['isProcurementUser', 'PROCUREMENT_TEAM_EMAILS']];
  FNS.forEach(([fn, key]) => {
    const diffs = [];
    let n0 = 0;
    LISTS.forEach(list => EMAILS.forEach(e => {
      b.sb.CONFIG[key] = list.slice(); n.sb.CONFIG[key] = list.slice(); n0++;
      const rb = b.call(fn, e), rn = n.call(fn, e);
      if (rb !== rn) diffs.push(JSON.stringify(list) + ' ' + JSON.stringify(e) + ': ' + rb + '→' + rn);
    }));
    check(`L ${fn}: ${n0}건 모두 이전과 같음`, () => diffs.length === 0 || diffs.slice(0, 4).join(' | '));
  });
  check('L isProcurementUser 빈 목록 = 전원 허용 폴백 유지 / isAdminUser·isGlobalViewer·isScmUser 빈 목록 = 아무도 아님', () => {
    ['ADMIN_EMAILS', 'GLOBAL_VIEW_EMAILS', 'PROCUREMENT_TEAM_EMAILS', 'SCM_EMAILS'].forEach(k => { n.sb.CONFIG[k] = []; });
    const r = [n.call('isProcurementUser', KIM), n.call('isAdminUser', KIM), n.call('isGlobalViewer', KIM), n.call('isScmUser', KIM)];
    return JSON.stringify(r) === '[true,false,false,false]' || r;
  });
}

// ── 화면 테스트용 픽스처 덤프 (tests/html_insp_submit_role_test.py) ──
if (process.env.DUMP_FIXTURES) {
  const out = {};
  // 김기안 제출 → 이결재 반려가 있는 상태 (PO-901에 반려 회차 1건)
  const withRejected = env => {
    const s = env.as(KIM).call('submitInspForClient', payload(KIM, 'tok-prcA'));
    env.as(KIM).call('processInspDecisionFromClient', { token: s.token, decision: 'approve', comment: '' });
    env.as(LEE).call('processInspDecisionFromClient', { token: s.token, decision: 'reject', comment: '수량 확인' });
    return env;
  };
  const [b, n] = pair(); withRejected(b); withRejected(n);
  out.homeScm   = norm(home(n, SCM));     // SCM 전용 — 전 품의 + 남이 낸 반려 회차 '수정 재상신'
  out.homeAdmin = norm(home(n, ADMIN));   // 관리자+SCM
  [[ADM2, 'Adm2'], [GV, 'Gv'], [CHOI, 'Choi'], [KIM, 'Kim']].forEach(([who, k]) => {
    out['home' + k + 'Base'] = norm(home(b, who)); out['home' + k + 'New'] = norm(home(n, who));
  });
  fs.writeFileSync(process.env.DUMP_FIXTURES, JSON.stringify(out, null, 1));
  console.log('픽스처 저장: ' + process.env.DUMP_FIXTURES);
}

console.log('\n─────────────────────────────');
console.log(`통과 ${pass} / 실패 ${fail}   (기준 ${BASE_REF})`);
if (fail) console.log('실패 항목:\n  - ' + failures.join('\n  - '));
process.exit(fail ? 1 : 0);
