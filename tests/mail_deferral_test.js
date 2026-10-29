/**
 * 결재 알림 메일 지연 발송(Code.gs 14-1) — 로컬 검증 (배포 없이 실제 .gs 함수 호출)
 * (memory: gas-local-test-harness, mail-deferral-plan)
 *
 * 방식: 같은 결재를 이전 코드(BASE_REF, 기본 c80574f = 이 기능 직전 커밋)와 작업 트리(수정본)에서 각각 돌린다.
 *  [E] 지연 발송: 결재 응답·시트·큐는 같고, 결재 중에는 메일이 0통, 트리거(flushMailBuffer_)가 보낸 메일은
 *      이전 코드가 즉시 보낸 메일과 수신자·제목·본문(텍스트·HTML)·발신 옵션·순서까지 같아야 한다.
 *  [I] 하트비트 없음 / 오래됨 / 스위치(MAIL_DEFER_OFF) → 이전 코드와 똑같이 즉시 발송.
 *  [D] 중복 금지·유실 알림 장치 — 수정본만.
 *  메일 조립 함수는 실제 코드를 쓰고, 맨 아래 GmailApp.sendEmail만 가로챈다.
 *
 * 실행: node tests/mail_deferral_test.js
 *  BASE_REF=<커밋>  비교 기준
 *  WT_DIR=<폴더>    뮤테이션 검사용 — 일부러 버그를 심은 사본을 '수정본'으로 돌린다
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execSync } = require('child_process');

const SRC = path.join(__dirname, '..');
const WT_DIR = process.env.WT_DIR || SRC;
const BASE_REF = process.env.BASE_REF || 'c80574f';
const readWT   = f => fs.readFileSync(path.join(WT_DIR, f), 'utf8');
const readBASE = f => execSync('git show ' + BASE_REF + ':"' + f + '"', { cwd: SRC, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });

// ── 고정 시계 (두 환경 공용) ──────────────────────────────────────
const RealDate = Date;
const CLOCK = { now: RealDate.parse('2026-10-29T10:00:00+09:00') };
class FakeDate extends RealDate {
  constructor(...a) { if (a.length === 0) super(CLOCK.now); else super(...a); }
  static now() { return CLOCK.now; }
}
const SEC = 1000, MIN = 60 * SEC;

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
  setValues(vals) { vals.forEach((rv, r) => rv.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v))); return this; }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
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
  constructor(name) { this.name = name; this.rows = []; }
  _get(r, c) { const row = this.rows[r - 1]; return row ? (row[c - 1] ?? '') : ''; }
  _set(r, c, v) { while (this.rows.length < r) this.rows.push([]); const row = this.rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
  getName() { return this.name; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return this.rows.reduce((m, r) => Math.max(m, r.length), 0); }
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
const ADMIN = 'davidkang@inlct.com', TEAM = 'admin-inlct@inlct.com';
const BUYER1 = 'buyer1@inlct.com', BUYER2 = 'buyer2@inlct.com';
const KIM = 'kim@inlct.com', LEE = 'lee@inlct.com', CHOI = 'choi@inlct.com';
const NAMES = { [KIM]: '김기안', [LEE]: '이결재', [CHOI]: '최결재', [ADMIN]: '강관리', [BUYER1]: '구매1', [BUYER2]: '구매2' };

// ── 환경 ─────────────────────────────────────────────────────────
function makeEnv(label, read) {
  const state = {
    label, active: ADMIN, mails: [], adminErrors: [], warns: [], uuid: 0,
    props: new Map(), cache: new Map(),
    lock: { scriptWait: 0, userHeld: false },
    // gmailFail: 남은 실패 횟수(별칭 폴백 포함 모든 sendEmail 호출이 소모) / mailSet·mailDel: MAIL_BUF_ 키 쓰기 장애
    flags: { gmailFail: 0, mailSet: null, mailSetDelayMs: 0, mailDel: 0, mailGetFail: false },
    order: [],
  };
  const ss = new FakeSpreadsheet();
  const isMail = k => String(k).indexOf('MAIL_BUF_') === 0;
  const propsObj = {
    getProperty: k => {
      if (state.flags.mailGetFail && isMail(k)) throw new Error('Properties 읽기 장애 시뮬레이션');
      return state.props.has(k) ? state.props.get(k) : null;
    },
    setProperty: (k, v) => {
      if (isMail(k) && state.flags.mailSet) {
        const mode = state.flags.mailSet; state.flags.mailSet = null;
        CLOCK.now += state.flags.mailSetDelayMs;
        if (mode === 'after') state.props.set(k, String(v));   // 저장은 됐는데 오류로 응답
        throw new Error('Properties 쓰기 장애 시뮬레이션 (' + mode + ')');
      }
      state.props.set(k, String(v)); return propsObj;
    },
    deleteProperty: k => {
      if (isMail(k) && state.flags.mailDel > 0) { state.flags.mailDel--; throw new Error('Properties 삭제 장애 시뮬레이션'); }
      state.props.delete(k); return propsObj;
    },
    getKeys: () => Array.from(state.props.keys()),
    getProperties: () => Object.fromEntries(state.props),
  };
  const cacheObj = {
    get: k => { const e = state.cache.get(k); return e && e.exp > CLOCK.now ? e.v : null; },
    put: (k, v, ttl) => { state.cache.set(k, { v: String(v), exp: CLOCK.now + (ttl || 600) * 1000 }); },
    remove: k => { state.cache.delete(k); },
  };
  const sandbox = {
    console: { log() {}, info() {}, warn: (...a) => state.warns.push(a.map(String).join(' ')), error() {} },
    Date: FakeDate,
    Logger: { log() {} },
    Session: {
      getActiveUser: () => ({ getEmail: () => state.active }),
      getEffectiveUser: () => ({ getEmail: () => ADMIN }),
      getScriptTimeZone: () => 'Asia/Seoul',
    },
    SpreadsheetApp: { openById: () => ss, flush() {} },
    DriveApp: { getFolderById: () => { throw new Error('no folder'); }, getFileById: () => ({ setTrashed() {} }) },
    Utilities: {
      getUuid: () => 'uuid-' + String(++state.uuid).padStart(8, '0'),
      newBlob: (content, type, name) => ({ content, type, name }),
      formatDate: (d, tz, fmt) => {
        const p = n => String(n).padStart(2, '0'); const t = new RealDate(d.getTime ? d.getTime() : d);
        const ymd = `${t.getFullYear()}-${p(t.getMonth() + 1)}-${p(t.getDate())}`;
        return fmt.indexOf('HH') >= 0 ? `${ymd} ${p(t.getHours())}:${p(t.getMinutes())}` : ymd;
      },
      sleep() {},
    },
    LockService: {
      getScriptLock: () => ({ waitLock() { state.lock.scriptWait++; }, tryLock() { state.lock.scriptWait++; return true; }, releaseLock() {} }),
      getUserLock: () => ({
        tryLock() { if (state.lock.userHeld) return false; state.lock.userHeld = true; return true; },
        releaseLock() { state.lock.userHeld = false; },
      }),
    },
    PropertiesService: { getScriptProperties: () => propsObj },
    CacheService: { getScriptCache: () => cacheObj },
    GmailApp: {
      sendEmail: (to, subj, plain, opts) => {
        if (state.flags.gmailFail > 0) { state.flags.gmailFail--; throw new Error('Gmail 발송 실패 시뮬레이션'); }
        opts = opts || {};
        state.mails.push({ to, subj, plain, html: opts.htmlBody || '', from: opts.from || '', name: opts.name || '' });
        state.order.push('mail');
      },
    },
    MailApp: { sendEmail() {} },
    HtmlService: { createTemplateFromFile: () => ({ evaluate: () => ({ setTitle() { return this; }, setXFrameOptionsMode() { return this; } }) }), createHtmlOutput: () => ({ setTitle() { return this; }, setXFrameOptionsMode() { return this; } }), XFrameOptionsMode: {} },
    ContentService: { createTextOutput: () => ({ setMimeType() { return this; } }), MimeType: {} },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => ({ timeBased: () => ({ everyMinutes: () => ({ create() {} }) }) }) },
    MimeType: { PDF: 'application/pdf', PLAIN_TEXT: 'text/plain' },
  };
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  ['Code.gs', 'INSP.gs', 'PO.gs'].forEach(f => vm.runInContext(read(f), sandbox, { filename: label + ':' + f }));

  // 메일 조립·발송 함수는 실제 코드 그대로. 범위 밖 부수효과만 두 환경에 똑같이 막는다.
  sandbox.__state = state;
  vm.runInContext(`
    notifyAdminError       = function (m) { __state.adminErrors.push(String(m)); };
    _prepareInspPdfHandoff = function () { return { ok: true }; };
    _processPdfOnlyJob     = function () { __state.order.push('pdf_job'); return { ok: true }; };
  `, sandbox);
  sandbox.CONFIG.ADMIN_EMAILS = [ADMIN];
  sandbox.CONFIG.PROCUREMENT_NOTIFY_EMAILS = [TEAM, BUYER2];   // 수신자별 항목 분리를 보려고 2명

  const env = {
    label, sb: sandbox, ss, state,
    as(user) { state.active = user; return env; },
    call(fn, ...args) { return sandbox[fn].apply(sandbox, args); },
    main: () => ss.getSheetByName(sandbox.CONFIG.SHEET_NAME),
    insp: () => ss.getSheetByName(sandbox.CONFIG.INSP_SHEET_NAME),
    mailKeys: () => Array.from(state.props.keys()).filter(k => k.indexOf('MAIL_BUF_') === 0).sort(),
    mailItem: k => JSON.parse(state.props.get(k)),
    hb: ms => { state.props.set('MAIL_FLUSH_HEARTBEAT', String(ms == null ? CLOCK.now : ms)); },
    // 트리거 = David 계정
    flush() { const prev = state.active; state.active = ADMIN; try { return sandbox.flushMailBuffer_(); } finally { state.active = prev; } },
    queue: () => JSON.parse(state.props.get('PENDING_JOB_QUEUE') || '[]').map(j => [j.type, j.docNo, j.docType, j.status]),
  };
  return env;
}

// ── 픽스처 ───────────────────────────────────────────────────────
function seed(env) {
  const sb = env.sb, C = sb.COL;
  env.state.props.set('AUDIT_FLUSH_HEARTBEAT', String(CLOCK.now));   // 시스템로그는 두 환경 모두 지연 기록 (메일과 무관)
  const main = sb.getOrCreateSheet(env.ss, sb.CONFIG.SHEET_NAME);
  sb.ensureHeaders(main);
  const ap = (label, email, status) => ({ label, name: NAMES[email], email, status: status || '대기' });
  function req(o) {
    const row = new Array(C.APPR_START).fill('');
    row[C.SUBMIT_AT] = new FakeDate('2026-10-20T09:00:00+09:00');
    row[C.DOC_NO] = o.docNo; row[C.ISSUE_DATE] = new FakeDate('2026-10-20T00:00:00+09:00');
    row[C.DRAFTER] = NAMES[o.drafter || KIM]; row[C.DEPT] = '개발팀';
    row[C.SUBJECT] = o.subject || '시험 품목 <특수&문자>'; row[C.PURPOSE] = '라인 증설\n2층 공사'; row[C.VENDOR_NAME] = '라인테크';
    row[C.ITEMS] = '1. 볼트 | M8 | 10개 | 1,000 KRW\n2. 너트 | M8 | 20개 | 500 KRW'; row[C.TOTAL_AMT] = 20000; row[C.REMARKS] = '긴급 \'납기\' 요청';
    row[C.STATUS] = o.status; row[C.TOKEN] = o.token; row[C.APPR_COUNT] = o.approvers.length; row[C.APPR_IDX] = o.apprIdx;
    row[C.RESUB_COUNT] = 0; row[C.REJECT_LOG] = ''; row[C.ATTACH_LIST] = '[]'; row[C.MOVE_STATUS] = 'STAGING';
    row[C.DOC_TYPE] = o.docType || 'REQ'; row[C.PARENT_DOC_ID] = o.parent || '';
    o.approvers.forEach(a => row.push(a.label, a.name, a.email, a.status, '', ''));
    main.appendRow(row);
  }
  req({ docNo: 'TH-AP-26-101', token: 'tok-mid',  status: '결재중 (결재자 1)', apprIdx: 1, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE), ap('결재자 2', CHOI)] });
  req({ docNo: 'TH-AP-26-102', token: 'tok-fin',  status: '결재중 (결재자 1)', apprIdx: 1, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE)] });
  req({ docNo: 'TH-AP-26-103', token: 'tok-rej',  status: '결재중 (결재자 1)', apprIdx: 1, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE)] });
  req({ docNo: 'TH-AP-26-104', token: 'tok-par',  status: 'PRC생성됨', apprIdx: 2, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE, '승인')] });
  req({ docNo: 'TO-PO-26-901', token: 'tok-prc',  status: '결재중 (결재자 1)', apprIdx: 1, docType: 'PRC', parent: 'tok-par', drafter: BUYER1,
        approvers: [ap('기안자', BUYER1, '승인'), ap('결재자 1', LEE)] });

  const isheet = sb.ensureInspSheet();
  const I = sb.INSP_COL;
  function insp(o) {
    const row = new Array(sb.INSP_TOTAL_COLS).fill('');
    row[I.SUBMIT_AT] = new FakeDate('2026-10-25T09:00:00+09:00'); row[I.DOC_NO] = o.docNo; row[I.SEQ] = 1;
    row[I.PRC_TOKEN] = 'tok-prc'; row[I.REQ_TOKEN] = 'tok-par'; row[I.PO_NO] = 'TO-PO-26-901'; row[I.REQ_NO] = 'TH-AP-26-104';
    row[I.SUBJECT] = '시험 품목'; row[I.VENDOR_NAME] = '라인테크'; row[I.DRAFTER] = KIM; row[I.DRAFTER_NAME] = NAMES[KIM];
    row[I.IS_FINAL] = o.isFinal ? 'Y' : ''; row[I.STATUS] = o.status; row[I.TOKEN] = o.token;
    row[I.APPR_COUNT] = o.approvers.length; row[I.APPR_IDX] = o.apprIdx; row[I.RESUB_COUNT] = 0; row[I.REJECT_LOG] = '';
    row[I.PHOTO_LIST] = '[]'; row[I.MOVE_STATUS] = 'STAGING';
    o.approvers.forEach((a, k) => { const b = I.APPR_START + k * I.APPR_COLS; row[b] = a.label; row[b + 1] = a.name; row[b + 2] = a.email; row[b + 3] = a.status; });
    isheet.appendRow(row);
  }
  insp({ docNo: 'TH-AP-26-104-01', token: 'tok-i-mid', status: '결재중 (이결재)', apprIdx: 1, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE), ap('결재자 2', CHOI)] });
  insp({ docNo: 'TH-AP-26-104-02', token: 'tok-i-fin', status: '결재중 (이결재)', apprIdx: 1, isFinal: true, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE)] });
  insp({ docNo: 'TH-AP-26-104-03', token: 'tok-i-rej', status: '결재중 (이결재)', apprIdx: 1, approvers: [ap('기안자', KIM, '승인'), ap('결재자 1', LEE)] });
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
const norm = x => JSON.parse(JSON.stringify(x === undefined ? null : x, (k, v) => (v instanceof RealDate ? 'D' + v.getTime() : v)));
const same = (a, b) => JSON.stringify(norm(a)) === JSON.stringify(norm(b)) || ('\n      BASE=' + JSON.stringify(norm(a)).slice(0, 600) + '\n      NEW =' + JSON.stringify(norm(b)).slice(0, 600));
const sheets = env => ({ main: env.main().rows, insp: env.insp().rows });
const DECIDE = { REQ: 'processDecisionFromClient', INSP: 'processInspDecisionFromClient' };
const SCEN = [
  { name: 'REQ 중간 승인 → 다음 결재자',          fn: DECIDE.REQ,  token: 'tok-mid',   decision: 'approve', mails: 1 },
  { name: 'REQ 최종승인 → 기안자 + 구매팀 2명',    fn: DECIDE.REQ,  token: 'tok-fin',   decision: 'approve', mails: 3 },
  { name: 'REQ 반려 → 기안자',                    fn: DECIDE.REQ,  token: 'tok-rej',   decision: 'reject',  mails: 1 },
  { name: 'PRC 최종승인 → 구매 기안자 + 원 기안자', fn: DECIDE.REQ,  token: 'tok-prc',   decision: 'approve', mails: 2 },
  { name: 'INSP 중간 승인 → 다음 결재자',          fn: DECIDE.INSP, token: 'tok-i-mid', decision: 'approve', mails: 1 },
  { name: 'INSP 최종승인 → 기안자',               fn: DECIDE.INSP, token: 'tok-i-fin', decision: 'approve', mails: 1 },
  { name: 'INSP 반려 → 기안자 (토큰 undefined 링크 보존)', fn: DECIDE.INSP, token: 'tok-i-rej', decision: 'reject', mails: 1 },
];
function decide(env, s) {
  return env.as(LEE).call(s.fn, { token: s.token, decision: s.decision, comment: '확인 <b>완료</b>\n둘째 줄' });
}

// =================================================================
console.log('\n[E] 지연 발송 — 결재 결과는 같고, 메일은 트리거가 이전과 똑같이 보낸다');
// =================================================================
for (const s of SCEN) {
  const b = seed(makeEnv('BASE', readBASE)), n = seed(makeEnv('NEW', readWT));
  n.hb();
  const rb = decide(b, s), rn = decide(n, s);
  check('E ' + s.name + ' — 응답 동일', () => same(rb, rn));
  check('E ' + s.name + ' — 시트·큐 동일', () => same({ s: sheets(b), q: b.queue() }, { s: sheets(n), q: n.queue() }));
  check('E ' + s.name + ' — 결재 중 발송 0통, 보관 ' + s.mails + '건 (이전 코드는 즉시 ' + s.mails + '통)', () =>
    (n.state.mails.length === 0 && n.mailKeys().length === s.mails && b.state.mails.length === s.mails) ||
    [n.state.mails.length, n.mailKeys().length, b.state.mails.length]);
  check('E ' + s.name + ' — 10초 전에는 트리거도 안 보냄', () => { const r = n.flush(); return (r.sent === 0 && n.mailKeys().length === s.mails) || r; });
  CLOCK.now += 15 * SEC;
  const r = n.flush();
  check('E ' + s.name + ' — 트리거 발송 = 이전 코드 즉시 발송 (수신자·제목·본문·HTML·발신·순서)', () =>
    (r.sent === s.mails && n.mailKeys().length === 0 && same(b.state.mails, n.state.mails) === true) || same(b.state.mails, n.state.mails));
  check('E ' + s.name + ' — 관리자 알림 없음', () => (n.state.adminErrors.length === 0 && b.state.adminErrors.length === 0) || [b.state.adminErrors, n.state.adminErrors]);
}

// =================================================================
console.log('\n[I] 즉시 발송으로 돌아가는 조건 — 이전 코드와 똑같이');
// =================================================================
for (const mode of ['하트비트 없음', '하트비트 11분 전', '스위치 MAIL_DEFER_OFF=1']) {
  for (const s of [SCEN[1], SCEN[3], SCEN[6]]) {
    const b = seed(makeEnv('BASE', readBASE)), n = seed(makeEnv('NEW', readWT));
    if (mode === '하트비트 11분 전') n.hb(CLOCK.now - 11 * MIN);
    if (mode.indexOf('스위치') === 0) { n.hb(); n.state.props.set('MAIL_DEFER_OFF', '1'); }
    const rb = decide(b, s), rn = decide(n, s);
    check('I ' + mode + ' / ' + s.name, () =>
      (same(rb, rn) === true && same(sheets(b), sheets(n)) === true && same(b.state.mails, n.state.mails) === true && n.mailKeys().length === 0) ||
      [same(rb, rn), same(b.state.mails, n.state.mails), n.mailKeys()]);
  }
}
{
  const n = seed(makeEnv('NEW', readWT));
  n.hb();
  decide(n, SCEN[0]);                       // 1건 보관된 채로 트리거가 멈춤
  CLOCK.now += 11 * MIN;
  check('I 트리거 멈춤 + 보관함에 남은 메일 → 새 메일은 즉시 발송, 관리자 알림 1회', () => {
    decide(n, SCEN[2]);
    const a = n.state.adminErrors.filter(m => m.indexOf('[메일]') === 0 && m.indexOf('대기 중') >= 0).length;
    return (n.state.mails.length === 1 && n.mailKeys().length === 1 && a === 1) || [n.state.mails.length, n.mailKeys(), n.state.adminErrors];
  });
  check('I 같은 시간 안에 또 결재해도 알림은 1회 (메일 폭풍 방지)', () => {
    decide(n, SCEN[1]);
    const a = n.state.adminErrors.filter(m => m.indexOf('대기 중') >= 0).length;
    return a === 1 || n.state.adminErrors;
  });
  check('I 트리거가 다시 돌면 남은 메일이 나간다 (1통, 중복 없음)', () => {
    const before = n.state.mails.length;
    const r = n.flush();
    return (r.sent === 1 && n.state.mails.length === before + 1 && n.mailKeys().length === 0) || [r, n.state.mails.length - before];
  });
}

// =================================================================
console.log('\n[D] 같은 메일 두 번 금지 · 유실은 알림');
// =================================================================
function fresh() { const n = seed(makeEnv('NEW', readWT)); n.hb(); return n; }
const ARGS = (to, tag) => [to, '김기안', 'TH-AP-26-9' + tag, '제목 ' + tag, '이결재', '사유', 'tok-x' + tag];
function defer(n, tag, to) { return n.call('deferMail_', 'rejection', ARGS(to || KIM, tag), { docNo: 'TH-AP-26-9' + tag, to: to || KIM }); }
const subjects = n => n.state.mails.map(m => m.subj);

{
  const n = fresh();
  check('D1 연달아 두 번 돌려도 한 번만 발송', () => {
    defer(n, '01'); CLOCK.now += 15 * SEC;
    n.flush(); n.flush();
    return (n.state.mails.length === 1 && n.mailKeys().length === 0) || n.state.mails.length;
  });
  check('D2 다른 실행이 발송 중(사용자 락 점유)이면 기다리지 않고 건너뜀', () => {
    defer(n, '02'); CLOCK.now += 15 * SEC;
    n.state.lock.userHeld = true; const r = n.flush(); n.state.lock.userHeld = false;
    return (r.skipped === 'busy' && n.mailKeys().length === 1 && n.state.mails.length === 1) || r;
  });
  check('D3 발송은 전역(스크립트) 락을 잡지 않음', () => {
    const sw = n.state.lock.scriptWait; n.flush();
    return (n.state.lock.scriptWait === sw && n.state.mails.length === 2) || [n.state.lock.scriptWait - sw, n.state.mails.length];
  });
  check('D4 넣은 순서대로 발송 (같은 ms에 넣어도)', () => {
    defer(n, '41'); defer(n, '42'); defer(n, '43'); CLOCK.now += 15 * SEC;
    n.flush();
    return same(subjects(n).slice(-3), ['41', '42', '43'].map(t => '[반려] TH-AP-26-9' + t + ' — 제목 ' + t + ' 가 반려되었습니다.'));
  });
}
{
  const n = fresh();
  defer(n, '05'); CLOCK.now += 15 * SEC;
  const k = n.mailKeys()[0], it = n.mailItem(k);
  it.s = 'sending'; it.t = CLOCK.now - 2 * MIN; n.state.props.set(k, JSON.stringify(it));
  check('D5 [끊긴 실행] 2분 전 \'발송 중\' → 보내지 않고 그대로 둠, 알림 없음', () => {
    const r = n.flush();
    return (r.sent === 0 && n.mailKeys().length === 1 && n.state.adminErrors.length === 0) || [r, n.state.adminErrors];
  });
  check('D6 [끊긴 실행] 7분 넘은 \'발송 중\' → 다시 보내지 않고 \'확인 필요\' 알림 + 정리', () => {
    CLOCK.now += 6 * MIN;
    const r = n.flush();
    const a = n.state.adminErrors.filter(m => m.indexOf('발송 여부 확인 필요') >= 0 && m.indexOf('TH-AP-26-905') >= 0).length;
    return (r.uncertain === 1 && n.state.mails.length === 0 && n.mailKeys().length === 0 && a === 1) || [r, n.state.adminErrors];
  });
}
{
  const n = fresh();
  check('D7 재시도까지 실패 → 관리자 알림, 다시 넣지 않음 (다음 트리거도 재발송 없음)', () => {
    defer(n, '07'); CLOCK.now += 15 * SEC;
    n.state.flags.gmailFail = 99;
    const r = n.flush();
    n.state.flags.gmailFail = 0;
    const r2 = n.flush();
    const a = n.state.adminErrors.filter(m => m.indexOf('지연 발송 실패') >= 0).length;
    return (r.failed === 1 && r2.sent === 0 && n.state.mails.length === 0 && n.mailKeys().length === 0 && a === 1) || [r, r2, n.state.adminErrors];
  });
  check('D8 별칭 1회 실패 → 기존 폴백(별칭 없이)으로 1통만', () => {
    defer(n, '08'); CLOCK.now += 15 * SEC;
    n.state.flags.gmailFail = 1;
    const r = n.flush();
    return (r.sent === 1 && n.state.mails.length === 1 && n.state.mails[0].from === '') || [r, n.state.mails];
  });
}
{
  const n = fresh();
  check('D9 보낸 뒤 지우기 실패 → 발송 1통, 7분 안에는 재발송 없음', () => {
    defer(n, '09'); CLOCK.now += 15 * SEC;
    n.state.flags.mailDel = 1;
    let threw = false; try { n.flush(); } catch (e) { threw = true; }
    const r2 = n.flush();
    return (threw && r2.sent === 0 && n.state.mails.length === 1 && n.mailKeys().length === 1) || [threw, r2, n.state.mails.length];
  });
  check('D10 그 항목은 7분 뒤 \'확인 필요\' 알림으로 정리 — 재발송 없음 (총 1통)', () => {
    CLOCK.now += 8 * MIN;
    const r = n.flush();
    return (r.uncertain === 1 && n.state.mails.length === 1 && n.mailKeys().length === 0) || [r, n.state.mails.length];
  });
  check('D11 발송 표시(쓰기) 실패 → 보내지 않고 throw, 다음 트리거가 1통 발송', () => {
    defer(n, '11'); CLOCK.now += 15 * SEC;
    const before = n.state.mails.length;
    n.state.flags.mailSet = 'before';
    let threw = false; try { n.flush(); } catch (e) { threw = true; }
    const mid = n.state.mails.length;
    n.flush();
    return (threw && mid === before && n.state.mails.length === before + 1 && n.mailKeys().length === 0) || [threw, mid - before, n.state.mails.length - before];
  });
}
// 쓰는 쪽 저장 실패 — 실제로는 저장됐을 수 있다
for (const c of [
  { name: 'D12 저장 실패(안 들어감) → 즉시 1통', set: 'before', del: 0, delay: 0, want: 'sent', total: 1, alert: 0 },
  { name: 'D13 저장 실패(실은 들어감) → 지우고 즉시 1통, 트리거는 안 보냄', set: 'after', del: 0, delay: 0, want: 'sent', total: 1, alert: 0 },
  { name: 'D14 저장 실패(실은 들어감) + 지우기도 실패 → 즉시 발송 안 함, 트리거가 1통', set: 'after', del: 1, delay: 0, want: 'deferred', total: 1, alert: 0 },
  { name: 'D15 저장 실패(실은 들어감)가 10초 넘게 걸림 → 트리거 몫으로 두고 1통', set: 'after', del: 0, delay: 11 * SEC, want: 'deferred', total: 1, alert: 0 },
  { name: 'D16 저장 실패(안 들어감)가 10초 넘게 걸림 → 갔는지 모름: 보내지 않고 알림', set: 'before', del: 0, delay: 11 * SEC, want: 'unknown', total: 0, alert: 1 },
]) {
  check(c.name, () => {
    const n = fresh();
    n.state.flags.mailSet = c.set; n.state.flags.mailDel = c.del; n.state.flags.mailSetDelayMs = c.delay;
    const ret = defer(n, '12');
    n.state.flags.mailDel = 0;
    CLOCK.now += 15 * SEC; n.flush(); CLOCK.now += 15 * SEC; n.flush();
    const a = n.state.adminErrors.filter(m => m.indexOf('보관 상태 불명') >= 0).length;
    return (ret === c.want && n.state.mails.length === c.total && n.mailKeys().length === 0 && a === c.alert) ||
           [ret, n.state.mails.length, n.mailKeys(), n.state.adminErrors];
  });
}
check('D17 [경합] 저장 실패 처리 중 트리거가 낡은 목록으로 그 항목을 되살리지 않음', () => {
  // 트리거가 목록을 읽은 뒤(항목 있음) 쓰는 쪽이 지우고 즉시 발송 — 트리거는 지금 값을 다시 보고 건너뛰어야 한다
  const n = fresh();
  defer(n, '17');
  const k = n.mailKeys()[0];
  CLOCK.now += 15 * SEC;
  const realGetProps = n.state.props;
  const props = n.sb.PropertiesService.getScriptProperties();
  const orig = props.getProperties;
  props.getProperties = () => { const snap = orig(); realGetProps.delete(k); n.state.mails.push({ subj: '쓰는 쪽 즉시 발송' }); return snap; };
  const r = n.flush();
  props.getProperties = orig;
  return (r.sent === 0 && n.state.mails.length === 1 && n.mailKeys().length === 0) || [r, n.state.mails.length, n.mailKeys()];
});
{
  const n = fresh();
  check('D18 너무 큰 메일(한도 초과) → 즉시 발송', () => {
    const ret = n.call('deferMail_', 'rejection', [KIM, '김기안', 'TH-AP-26-918', '제목', '이결재', '가'.repeat(3000), 'tok'], { docNo: 'TH-AP-26-918' });
    return (ret === 'sent' && n.state.mails.length === 1 && n.mailKeys().length === 0) || [ret, n.mailKeys()];
  });
  check('D19 날짜 객체가 든 인자 → 저장하면 바뀌므로 즉시 발송', () => {
    const ret = n.call('deferMail_', 'completion', [KIM, '김기안', 'TH-AP-26-919', '제목', new FakeDate(), 'tok', null], { docNo: 'TH-AP-26-919' });
    return (ret === 'sent' && n.state.mails.length === 2 && n.mailKeys().length === 0) || [ret, n.mailKeys()];
  });
  check('D19b 안쪽 배열에 undefined → 저장하면 null로 바뀌므로 즉시 발송', () => {
    const ret = n.call('deferMail_', 'completion', [KIM, '김기안', 'TH-AP-26-919', '제목', '2026-10-29', 'tok', { list: [1, undefined] }], { docNo: 'TH-AP-26-919' });
    const ok = ret === 'sent' && n.state.mails.length === 3 && n.mailKeys().length === 0;
    n.state.mails.pop();
    return ok || [ret, n.mailKeys()];
  });
  check('D20 모르는 종류 → 아무것도 하지 않고 throw', () => {
    try { n.call('deferMail_', 'nope', [], {}); return '안 던짐'; } catch (e) { return (n.mailKeys().length === 0 && n.state.mails.length === 2) || n.mailKeys(); }
  });
  check('D21 복원 불가 항목 → 원문과 함께 알림 + 정리 (매분 반복 방지)', () => {
    n.state.props.set('MAIL_BUF_0000000000001_001_bad', '{not json');
    n.flush();
    const a = n.state.adminErrors.filter(m => m.indexOf('복원 실패') >= 0 && m.indexOf('{not json') >= 0).length;
    return (a === 1 && n.mailKeys().length === 0) || n.state.adminErrors;
  });
  check('D22 빈 보관함 → 하트비트만 갱신, 락 없음', () => {
    n.state.props.delete('MAIL_FLUSH_HEARTBEAT'); CLOCK.now += SEC;
    const r = n.flush();
    return (r.sent === 0 && n.state.props.get('MAIL_FLUSH_HEARTBEAT') === String(CLOCK.now) && !n.state.lock.userHeld) || r;
  });
  check('D23 1회 30통 한도 → 나머지는 다음 트리거, 15분 넘게 밀리면 관리자 알림 1회', () => {
    for (let i = 0; i < 32; i++) defer(n, String(100 + i));
    CLOCK.now += 16 * MIN; n.hb();
    const r = n.flush();
    const a = n.state.adminErrors.filter(m => m.indexOf('15분 넘게 밀려') >= 0).length;
    const r2 = n.flush();
    return (r.sent === 30 && r.remaining === 2 && a === 1 && r2.sent === 2 && n.mailKeys().length === 0) || [r, r2, a];
  });
}
check('D24 processQueueTrigger — 메일을 PDF job보다 먼저 보내고 하트비트 갱신', () => {
  const n = fresh();
  decide(n, SCEN[1]);   // REQ 최종승인 → 메일 3건 + pdf_only job
  CLOCK.now += 15 * SEC; n.state.props.delete('MAIL_FLUSH_HEARTBEAT');
  n.as(ADMIN).call('processQueueTrigger');
  return (same(n.state.order, ['mail', 'mail', 'mail', 'pdf_job']) === true && n.state.props.get('MAIL_FLUSH_HEARTBEAT') === String(CLOCK.now)) ||
         [n.state.order, n.state.props.get('MAIL_FLUSH_HEARTBEAT')];
});
check('D25 mailBufferStatus — 관리자 점검 함수가 상태를 돌려줌', () => {
  const n = fresh();
  defer(n, '25');
  const st = n.as(ADMIN).call('mailBufferStatus');
  return (st.pending === 1 && st.sending === 0 && st.mode === '지연 발송') || st;
});

// =================================================================
console.log('\n결과: ' + pass + ' 통과 / ' + fail + ' 실패');
if (fail) { console.log('실패 목록:\n  - ' + failures.join('\n  - ')); process.exit(1); }
