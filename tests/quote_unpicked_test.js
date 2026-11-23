/**
 * 견적서 미지정 제출(quoteUnpicked) — 로컬 검증 (배포 없이 실제 .gs 함수 호출)
 * (memory: gas-local-test-harness, quote-pdf-feature)
 *
 * 방식: 같은 시나리오를 이전 코드(BASE_REF, 기본 1d8b398 = 이 기능 직전 커밋)와 작업 트리에서 각각 돌린다.
 *  [S] 제출·재상신 — 견적서 지정(PDF/엑셀)·PRC·첨부 0건은 시트·Drive·메일이 이전과 바이트 단위로 같아야 한다.
 *      첨부 있음·미지정만 달라야 한다: Y열 항목에 quoteUnpicked, 메일 문구. 그 둘을 걷어내면 나머지는 같아야 한다.
 *  [M] 관리자 메일 — 첨부 0건 '[견적서 미첨부 제출]' / 미지정 '[견적서 미지정 제출]' / 지정 시 메일 없음.
 *  [R] 자동 인식(reconcileQuotePdf) — 미지정 건에 QUO_ 업로드·첨부 PDF 이름 변경 인식, 표시 없는 행은 Drive 호출 0.
 *      기존 PENDING 경로(quote-src)는 이전 코드와 같은 결과.
 *  [L] 관리자 홈 목록(listPendingQuotePdf) — UNPICKED 표시·인식 후 사라짐, 기존 PENDING 항목은 이전과 동일.
 *  [D] [견적서 아님](dismissQuoteUnpickedForClient) — 관리자만, 표시만 지움, 감사로그.
 *  [K] markQuoteUnpicked — 기존 제출분 구제(표시만 / PDF 이름 변경 + 등록), 거부 조건.
 *  [Q] 큐 자동 변환(_processQuotePdfJob) — 이전 코드와 같은 결과.
 *
 * 실행: node tests/quote_unpicked_test.js
 *  BASE_REF=<커밋>  비교 기준
 *  WT_DIR=<폴더>    뮤테이션 검사용 — 일부러 버그를 심은 사본을 '수정본'으로 돌린다
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execSync } = require('child_process');

const SRC = path.join(__dirname, '..');
const WT_DIR = process.env.WT_DIR || SRC;
const BASE_REF = process.env.BASE_REF || '1d8b398';
const readWT   = f => fs.readFileSync(path.join(WT_DIR, f), 'utf8');
const readBASE = f => execSync('git show ' + BASE_REF + ':"' + f + '"', { cwd: SRC, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });

// ── 고정 시계 ────────────────────────────────────────────────────
const RealDate = Date;
const CLOCK = { now: RealDate.parse('2026-11-23T10:00:00+09:00') };
class FakeDate extends RealDate {
  constructor(...a) { if (a.length === 0) super(CLOCK.now); else super(...a); }
  static now() { return CLOCK.now; }
}

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
  getValue() { return this.sheet._get(this.row, this.col); }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  setFontWeight() { return this; }
  setBackground() { return this; }
  setDataValidation() { return this; }
  setNumberFormat() { return this; }
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
  getMaxColumns() { return Math.max(26, this.getLastColumn(), this._maxCols || 0); }
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
const ADMIN = 'davidkang@inlct.com', BUYER1 = 'buyer1@inlct.com';
const KIM = 'kim@inlct.com', LEE = 'lee@inlct.com';
const NAMES = { [KIM]: '김기안', [LEE]: '이결재', [ADMIN]: '강관리', [BUYER1]: '구매1' };

// ── 환경 ─────────────────────────────────────────────────────────
function makeEnv(label, read) {
  const state = { label, active: KIM, mails: [], summaries: [], adminErrors: [], uuid: 0, props: new Map(), cache: new Map(),
                  folders: {}, files: {}, driveCalls: 0 };
  const ss = new FakeSpreadsheet();
  const D = () => { state.driveCalls++; };
  function folderRec(id) { if (!state.folders[id]) state.folders[id] = { id }; return state.folders[id]; }
  const iter = arr => { let i = 0; return { hasNext: () => i < arr.length, next: () => arr[i++] }; };
  function wrapFile(f) {
    return {
      getId: () => f.id, getName: () => f.name, getSize: () => f.size,
      setName: n => { D(); f.name = n; return wrapFile(f); },
      isTrashed: () => !!f.trashed, setTrashed: b => { D(); f.trashed = !!b; },
      setSharing: () => { f.shared = true; },
      getBlob: () => ({ getBytes: () => f.bytes, getName: () => f.name }),
    };
  }
  function wrapFolder(fo) {
    return {
      getId: () => fo.id, getName: () => fo.id, isTrashed: () => false, setTrashed() {},
      createFile: blob => {
        D();
        const bytes = Array.isArray(blob.bytes) ? blob.bytes : (blob.getBytes ? blob.getBytes() : []);
        const nf = { id: 'file-' + (++state.uuid), name: blob.name || (blob.getName && blob.getName()), type: blob.type, size: bytes.length, bytes, parent: fo.id, trashed: false };
        state.files[nf.id] = nf;
        return wrapFile(nf);
      },
      getFilesByName: name => { D(); return iter(Object.values(state.files).filter(f => f.parent === fo.id && f.name === name).map(wrapFile)); },
      getFiles: () => { D(); return iter(Object.values(state.files).filter(f => f.parent === fo.id).map(wrapFile)); },
      getFoldersByName: () => iter([]), getParents: () => iter([]),
    };
  }
  const propsObj = {
    getProperty: k => (state.props.has(k) ? state.props.get(k) : null),
    setProperty: (k, v) => { state.props.set(k, String(v)); return propsObj; },
    deleteProperty: k => { state.props.delete(k); return propsObj; },
    getKeys: () => Array.from(state.props.keys()),
    getProperties: () => Object.fromEntries(state.props),
  };
  const cacheObj = {
    get: k => { const e = state.cache.get(k); return e && e.exp > CLOCK.now ? e.v : null; },
    put: (k, v, ttl) => { state.cache.set(k, { v: String(v), exp: CLOCK.now + (ttl || 600) * 1000 }); },
    remove: k => { state.cache.delete(k); },
  };
  const sandbox = {
    console: { log() {}, info() {}, warn() {}, error() {} },
    Date: FakeDate,
    Logger: { log() {} },
    Session: {
      getActiveUser: () => ({ getEmail: () => state.active }),
      getEffectiveUser: () => ({ getEmail: () => ADMIN }),
      getScriptTimeZone: () => 'Asia/Seoul',
    },
    SpreadsheetApp: { openById: () => ss, flush() {} },
    DriveApp: {
      Access: { DOMAIN_WITH_LINK: 'DOMAIN_WITH_LINK' }, Permission: { VIEW: 'VIEW' },
      getFolderById: id => { D(); const f = state.folders[id]; if (!f) throw new Error('No item with the given ID: ' + id); return wrapFolder(f); },
      getFileById: id => { D(); const f = state.files[id]; if (!f) throw new Error('No item with the given ID: ' + id); return wrapFile(f); },
    },
    Drive: { Files: {
      create: () => { throw new Error('변환 스텁 — convertToQuotePdf 를 가로챈다'); },
      remove() {},
      get: id => ({ id, parents: [state.files[id].parent] }),
      update: (meta, id, blob, opt) => { state.files[id].parent = opt.addParents; return {}; },
    } },
    Utilities: {
      getUuid: () => 'uuid-' + String(++state.uuid).padStart(8, '0'),
      base64Decode: s => Array.from(Buffer.from(String(s), 'base64')).map(b => (b > 127 ? b - 256 : b)),
      base64Encode: b => Buffer.from(b).toString('base64'),
      newBlob: (bytes, type, name) => ({ bytes, type, name }),
      formatDate: (d, tz, fmt) => {
        const p = n => String(n).padStart(2, '0'); const t = new RealDate(d.getTime ? d.getTime() : d);
        const ymd = `${t.getFullYear()}-${p(t.getMonth() + 1)}-${p(t.getDate())}`;
        if (fmt === 'yyyyMMdd_HHmmss') return `${t.getFullYear()}${p(t.getMonth() + 1)}${p(t.getDate())}_${p(t.getHours())}${p(t.getMinutes())}${p(t.getSeconds())}`;
        return fmt.indexOf('HH') >= 0 ? `${ymd} ${p(t.getHours())}:${p(t.getMinutes())}` : ymd;
      },
      sleep() {},
    },
    LockService: {
      // 재진입하면 던진다 — withLock 안에서 다시 withLock 을 부르는 실수를 잡는다(quote-pdf-feature 학습)
      getScriptLock: () => ({
        waitLock() { if (state.locked) throw new Error('LOCK_REENTRY'); state.locked = true; },
        tryLock() { if (state.locked) return false; state.locked = true; return true; },
        releaseLock() { state.locked = false; },
      }),
      getUserLock: () => ({ tryLock() { return true; }, releaseLock() {} }),
    },
    PropertiesService: { getScriptProperties: () => propsObj },
    CacheService: { getScriptCache: () => cacheObj },
    GmailApp: { sendEmail: (to, subj, plain, opts) => { state.mails.push({ to, subj, plain, html: (opts || {}).htmlBody || '' }); } },
    MailApp: { sendEmail() {} },
    HtmlService: { createTemplateFromFile: () => ({ evaluate: () => ({ setTitle() { return this; }, setXFrameOptionsMode() { return this; } }) }), createHtmlOutput: () => ({ setTitle() { return this; }, setXFrameOptionsMode() { return this; } }), XFrameOptionsMode: {} },
    ContentService: { createTextOutput: () => ({ setMimeType() { return this; } }), MimeType: {} },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => ({ timeBased: () => ({ everyMinutes: () => ({ create() {} }) }) }) },
    MimeType: { PDF: 'application/pdf', PLAIN_TEXT: 'text/plain', GOOGLE_SHEETS: 'gs', GOOGLE_DOCS: 'gd', GOOGLE_SLIDES: 'gsl', HTML: 'text/html' },
  };
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  ['Code.gs', 'INSP.gs', 'PO.gs'].forEach(f => vm.runInContext(read(f), sandbox, { filename: label + ':' + f }));

  sandbox.__state = state;
  sandbox.__folderRec = folderRec;
  sandbox.__wrapFolder = wrapFolder;
  vm.runInContext(`
    notifyAdminError   = function (m) { __state.adminErrors.push(String(m)); };
    getOrCreateFolder  = function (name) { return __wrapFolder(__folderRec('F_' + name)); };
    createFileInFolder = function (folder, name, content) { __state.summaries.push({ name: name }); };
    findFolderByPath   = function () { return null; };
    convertToQuotePdf  = function (blob, ext, fileName) {
      return { bytes: [1, 2, 3, 4], type: 'application/pdf', name: fileName, getBytes: function () { return [1, 2, 3, 4]; }, getName: function () { return fileName; } };
    };
  `, sandbox);
  sandbox.CONFIG.PROCUREMENT_TEAM_EMAILS = [ADMIN, BUYER1];
  sandbox.CONFIG.ADMIN_EMAILS = [ADMIN];
  sandbox.CONFIG.QUOTE_PDF_NOTIFY_ON_SUCCESS = false;
  sandbox.getOrCreateSheet(ss, sandbox.CONFIG.SHEET_NAME);
  sandbox.ensureHeaders(ss.getSheetByName(sandbox.CONFIG.SHEET_NAME));

  return {
    label, sb: sandbox, ss, state,
    as(user) { state.active = user; return this; },
    call(fn, ...args) { return sandbox[fn].apply(sandbox, args); },
    run(code) { return vm.runInContext(code, sandbox); },
    main: () => ss.getSheetByName(sandbox.CONFIG.SHEET_NAME),
    C: sandbox.COL,
    row(i = 1) { return this.main().rows[i]; },
    y(i = 1) { return JSON.parse(this.row(i)[sandbox.COL.ATTACH_LIST] || '[]'); },
    audit() { const s = ss.getSheetByName(sandbox.CONFIG.AUDIT_SHEET_NAME); return s ? s.rows.slice(1) : []; },
    folderOf(i = 1) { return String(this.row(i)[sandbox.COL.DRIVE_ID]); },
    liveFiles(folderId) { return Object.values(state.files).filter(f => f.parent === folderId && !f.trashed).map(f => f.name).sort(); },
    upload(folderId, name) {   // 관리자가 Drive 에 직접 올린 파일
      const f = { id: 'up-' + (++state.uuid), name, type: 'application/pdf', size: 9, bytes: [9], parent: folderId, trashed: false };
      state.files[f.id] = f; return f;
    },
  };
}

// ── 결과 정규화 ──────────────────────────────────────────────────
const norm = v => JSON.parse(JSON.stringify(v, (k, x) => (x instanceof RealDate ? 'D:' + x.toISOString() : x)));
function snap(env) {
  return norm({
    rows: env.main().rows,
    files: Object.values(env.state.files).map(f => ({ id: f.id, name: f.name, parent: f.parent, trashed: !!f.trashed, size: f.size })),
    summaries: env.state.summaries,
    mails: env.state.mails,
    adminErrors: env.state.adminErrors,
    queue: env.state.props.get('PENDING_JOB_QUEUE') || '',
  });
}

// ── 단언 ─────────────────────────────────────────────────────────
let pass = 0, fail = 0;
const failures = [];
function check(name, cond, detail) {
  if (cond === true) { pass++; console.log('  ✓ ' + name); }
  else { fail++; failures.push(name); console.log('  ✗ ' + name + (detail !== undefined ? '  → ' + JSON.stringify(detail).slice(0, 900) : '')); }
}
function firstDiff(a, b, p = '') {
  if (JSON.stringify(a) === JSON.stringify(b)) return null;
  if (typeof a !== 'object' || typeof b !== 'object' || !a || !b) return { at: p, base: a, wt: b };
  for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) { const d = firstDiff(a[k], b[k], p + '.' + k); if (d) return d; }
  return null;
}

// ── 픽스처 ───────────────────────────────────────────────────────
const ITEMS = [{ name: '정전기 측정기 검교정', spec: 'EA', qty: 2, price: 30500, currency: 'KRW' }];
const b64 = s => Buffer.from(s).toString('base64');
const att = (name, isQuote) => {
  const a = { name, type: name.endsWith('.pdf') ? 'application/pdf' : 'application/octet-stream', size: 7, data: b64('content') };
  if (isQuote) a.isQuote = true;
  return a;
};
const QPDF = '20261123-2 인엘씨테크놀러지 견적서_6485 외.pdf';
const XLSX = '견적서_원본.xlsx';
const ETC  = '사양서.docx';

function reqPayload(docNo, attachments, drafter = KIM) {
  return {
    docNo, issueDate: '2026-11-23', drafter: NAMES[drafter], dept: '품질팀', subject: 'LC공정 계측기 검교정', purpose: '정기 검교정',
    vendorName: '(주)검교정', vendorEmail: 'sales@example.com', vendorContact: '박담당', vendorPhone: '042-000-0000',
    deliveryDate: '2026-12-01', deliveryAddr: '대전광역시 유성구', items: JSON.parse(JSON.stringify(ITEMS)), remarks: '',
    approvers: [{ label: '기안자', name: NAMES[drafter], email: drafter }, { label: '결재자 1', name: NAMES[LEE], email: LEE }],
    attachments: attachments.map(a => Object.assign({}, a)),
  };
}
const submit = (env, docNo, atts) => env.as(KIM).call('submitRequisitionForClient', reqPayload(docNo, atts));
const quoteMails = env => env.state.mails.filter(m => /^\[견적서 미(지정|첨부)/.test(m.subj));

function runBoth(fn) {
  const b = makeEnv('BASE', readBASE), w = makeEnv('WT', readWT);
  const rb = norm(fn(b)), rw = norm(fn(w));
  return { b, w, rb, rw, sb: snap(b), sw: snap(w) };
}
// 미지정 건 비교용: quoteUnpicked 와 견적서 안내 메일을 걷어낸 스냅샷
function stripUnpicked(s, C) {
  const c = JSON.parse(JSON.stringify(s));
  c.rows = c.rows.map((r, i) => {
    if (i === 0 || !r[C.ATTACH_LIST]) return r;
    try { const l = JSON.parse(r[C.ATTACH_LIST]); l.forEach(f => { delete f.quoteUnpicked; }); r[C.ATTACH_LIST] = JSON.stringify(l); } catch (_) {}
    return r;
  });
  c.mails = c.mails.filter(m => !/^\[견적서 미(지정|첨부)/.test(m.subj));
  return c;
}

// ═════════════════════════════════════════════════════════════════
console.log('\n[S] 제출·재상신 — 지정·PRC·첨부 0건은 이전과 동일, 미지정만 표시 추가');
[
  ['S1 PDF 견적서 지정', [att(QPDF, true), att(ETC)]],
  ['S2 엑셀 견적서 지정(큐 변환)', [att(XLSX, true), att(ETC)]],
  ['S3 첨부 0건', []],
].forEach(([name, atts]) => {
  const r = runBoth(env => submit(env, 'TQ-AP-26-004', atts));
  check(name + ' — 제출 성공', r.rb.ok === true && r.rw.ok === true, [r.rb, r.rw]);
  check(name + ' — 응답·시트·Drive·큐 동일', JSON.stringify(r.rb) === JSON.stringify(r.rw) && JSON.stringify(stripUnpicked(r.sb, r.b.C)) === JSON.stringify(stripUnpicked(r.sw, r.w.C)),
    firstDiff(stripUnpicked(r.sb, r.b.C), stripUnpicked(r.sw, r.w.C)));
  check(name + ' — 시트 행 바이트 동일(표시 없음)', JSON.stringify(r.sb.rows) === JSON.stringify(r.sw.rows), firstDiff(r.sb.rows, r.sw.rows));
  check(name + ' — quoteUnpicked 없음', r.w.y().every(f => !f.quoteUnpicked), r.w.y());
});

(function () {
  const r = runBoth(env => submit(env, 'TQ-AP-26-004', [att(QPDF), att(ETC)]));
  const y = r.w.y();
  check('S4 첨부 2건·미지정 — 제출 성공', r.rb.ok === true && r.rw.ok === true, [r.rb, r.rw]);
  check('S4 — 모든 항목에 quoteUnpicked(제출 시각)', y.length === 2 && y.every(f => f.quoteUnpicked === '2026-11-23T01:00:00.000Z'), y);
  check('S4 — 이전 코드엔 표시 없음', r.b.y().every(f => !f.quoteUnpicked), r.b.y());
  check('S4 — 표시·안내 메일을 걷어내면 나머지(행·Drive·결재 메일·큐) 동일',
    JSON.stringify(stripUnpicked(r.sb, r.b.C)) === JSON.stringify(stripUnpicked(r.sw, r.w.C)), firstDiff(stripUnpicked(r.sb, r.b.C), stripUnpicked(r.sw, r.w.C)));
  check('S4 — 큐 등록 없음(변환할 원본 없음)', !r.w.state.props.get('PENDING_JOB_QUEUE') || JSON.parse(r.w.state.props.get('PENDING_JOB_QUEUE')).length === 0);
})();

(function () {
  // 재상신: 반려(Y 비움) → 미지정 재상신 → 표시, 그다음 지정 재상신 → 표시 사라짐
  const w = makeEnv('WT', readWT);
  submit(w, 'TQ-AP-26-004', [att(QPDF), att(ETC)]);
  const C = w.C; const row = w.row();
  row[C.STATUS] = '반려'; row[C.ATTACH_LIST] = '[]';   // 반려 처리와 같은 상태(Y열 비움)
  w.state.mails.length = 0;
  const r1 = w.as(KIM).call('resubmitRequisitionForClient', { token: row[C.TOKEN], data: reqPayload('TQ-AP-26-004', [att(QPDF)]) });
  check('S5 미지정 재상신 — 성공·표시', r1.ok === true && w.y().length === 1 && !!w.y()[0].quoteUnpicked, [r1, w.y()]);
  check('S5 — 관리자 메일 [견적서 미지정 재상신]', quoteMails(w).length === 1 && quoteMails(w)[0].subj.indexOf('[견적서 미지정 재상신] TQ-AP-26-004') === 0, quoteMails(w).map(m => m.subj));
  const row2 = w.row(); row2[C.STATUS] = '반려'; row2[C.ATTACH_LIST] = '[]';
  const r2 = w.as(KIM).call('resubmitRequisitionForClient', { token: row2[C.TOKEN], data: reqPayload('TQ-AP-26-004', [att(QPDF, true)]) });
  check('S6 지정 재상신 — 견적서 항목·표시 없음', r2.ok === true && w.y().length === 1 && w.y()[0].role === 'quote' && !w.y()[0].quoteUnpicked, [r2, w.y()]);
})();

(function () {
  // PRC 제출(quoteDocNo 없음) — 첨부 있어도 표시 없음, 이전과 동일
  const r = runBoth(env => {
    submit(env, 'TQ-AP-26-004', [att(QPDF, true)]);
    const C = env.C; const parent = env.row();
    parent[C.STATUS] = '최종승인'; parent[C.CLAIMED_BY] = BUYER1;
    env.state.mails.length = 0;
    const p = reqPayload('TO-PO-26-901', [att('발주서.pdf'), att(ETC)], BUYER1);
    Object.assign(p, { parentToken: parent[C.TOKEN], dept: '구매팀', purchaseMethod: '계좌이체', paymentInfo: '익월말', existingAttachments: [], fieldChanges: [] });
    return env.as(BUYER1).call('submitPrcForClient', p);
  });
  check('S7 PRC 제출 — 성공', r.rb.ok === true && r.rw.ok === true, [r.rb, r.rw]);
  check('S7 PRC — 시트·Drive·메일 바이트 동일', JSON.stringify(r.sb) === JSON.stringify(r.sw), firstDiff(r.sb, r.sw));
})();

// ═════════════════════════════════════════════════════════════════
console.log('\n[M] 관리자 메일 구분');
(function () {
  const w0 = makeEnv('WT', readWT); submit(w0, 'TQ-AP-26-010', []);
  const m0 = quoteMails(w0);
  check('M1 첨부 0건 → [견적서 미첨부 제출] 1통', m0.length === 1 && m0[0].subj === '[견적서 미첨부 제출] TQ-AP-26-010 - LC공정 계측기 검교정', m0.map(m => m.subj));
  check('M1 — 본문: 첨부 없음·기안자 확인 안내, 배너·[견적서 아님] 언급 없음',
    m0[0].html.indexOf('첨부 없음') >= 0 && m0[0].html.indexOf('견적서가 필요한 품의라면 기안자에게 확인') >= 0 && m0[0].html.indexOf('견적서 아님') < 0 && m0[0].html.indexOf('견적서 없이 제출') >= 0);
  const w1 = makeEnv('WT', readWT); submit(w1, 'TQ-AP-26-011', [att(QPDF), att(ETC)]);
  const m1 = quoteMails(w1);
  check('M2 미지정 → [견적서 미지정 제출] 1통', m1.length === 1 && m1[0].subj === '[견적서 미지정 제출] TQ-AP-26-011 - LC공정 계측기 검교정', m1.map(m => m.subj));
  check('M2 — 본문: 파일명·QUO_ 이름·이름 변경 안내·[견적서 아님]·관리자 처리 대기',
    m1[0].html.indexOf(QPDF) >= 0 && m1[0].html.indexOf('QUO_TQ-AP-26-011.pdf') >= 0 && m1[0].html.indexOf('파일 이름만') >= 0
      && m1[0].html.indexOf('[견적서 아님]') >= 0 && m1[0].html.indexOf('관리자 처리 대기') >= 0, m1[0].html);
  check('M2 — 수신 = 관리자 알림 목록', m1[0].to === w1.sb.CONFIG.ADMIN_NOTIFY_EMAILS.join(','));
  const w2 = makeEnv('WT', readWT); submit(w2, 'TQ-AP-26-012', [att(QPDF, true)]);
  check('M3 지정 → 안내 메일 없음', quoteMails(w2).length === 0, w2.state.mails.map(m => m.subj));
})();

// ═════════════════════════════════════════════════════════════════
console.log('\n[R] 자동 인식');
function unpickedEnv(docNo = 'TQ-AP-26-004', atts = [att(QPDF), att(ETC)]) {
  const w = makeEnv('WT', readWT); submit(w, docNo, atts); w.state.mails.length = 0; return w;
}
const recon = (env, opts) => env.call('reconcileQuotePdf', env.main(), 2, env.row(), opts);
(function () {
  const w = unpickedEnv();
  const before = JSON.stringify(w.row());
  check('R1 미지정·QUO_ 없음 → false, 행 그대로', recon(w) === false && JSON.stringify(w.row()) === before);
  check('R1 quick 도 false', recon(w, { quick: true }) === false && JSON.stringify(w.row()) === before);

  const up = w.upload(w.folderOf(), 'QUO_TQ-AP-26-004.pdf');
  check('R2 QUO_ 업로드 후 quick 인식 → true', recon(w, { quick: true }) === true);
  const y = w.y();
  check('R2 — 원래 첨부 2건 유지 + 견적서 항목(admin, origName 빈값)', y.length === 3 && y[0].name === QPDF && y[1].name === ETC
    && y[2].role === 'quote' && y[2].id === up.id && y[2].source === 'admin' && y[2].origName === '' && y[2].quotePdf === 'DONE', y);
  check('R2 — 표시 모두 지워짐', y.every(f => !f.quoteUnpicked), y);
  check('R2 — 감사로그 QUOTE_PDF_GENERATED', w.audit().some(r => r.indexOf('QUOTE_PDF_GENERATED') >= 0));
  check('R2 — 다시 호출하면 false(재등록 없음)', recon(w, { quick: true }) === false);
})();
(function () {
  const w = unpickedEnv();
  const pdfEntry = w.y()[0];
  w.state.files[pdfEntry.id].name = 'QUO_TQ-AP-26-004.pdf';   // 관리자가 Drive 에서 첨부 PDF 이름을 바꿈
  check('R3 첨부 PDF 이름 변경 → 인식', recon(w, { quick: true }) === true);
  const y = w.y();
  check('R3 — 그 항목이 견적서 항목이 됨(항목 2건, 같은 id, origName = 원래 이름)', y.length === 2 && y[0].name === ETC
    && y[1].id === pdfEntry.id && y[1].name === 'QUO_TQ-AP-26-004.pdf' && y[1].role === 'quote' && y[1].origName === QPDF, y);
  check('R3 — 표시 지워짐', y.every(f => !f.quoteUnpicked), y);
})();
(function () {
  // 표시도 견적서도 없는 행(이 기능 이전 제출분) — Drive 호출 0, false
  const w = makeEnv('WT', readWT); submit(w, 'TQ-AP-26-003', [att(QPDF)]);
  const C = w.C; const l = w.y(); l.forEach(f => delete f.quoteUnpicked); w.row()[C.ATTACH_LIST] = JSON.stringify(l);
  w.upload(w.folderOf(), 'QUO_TQ-AP-26-003.pdf');
  const calls = w.state.driveCalls;
  check('R4 표시 없는 옛 행 → false, Drive 호출 0', recon(w) === false && recon(w, { quick: true }) === false && w.state.driveCalls === calls);
})();
(function () {
  // 기존 PENDING 경로 — 이전 코드와 같은 결과
  const r = runBoth(env => {
    submit(env, 'TQ-AP-26-020', [att(XLSX, true)]);
    const C = env.C; const l = JSON.parse(env.row()[C.ATTACH_LIST]); l[0].quotePdf = 'PENDING'; env.row()[C.ATTACH_LIST] = JSON.stringify(l);
    env.upload(env.folderOf(), 'QUO_TQ-AP-26-020.pdf');
    return [recon(env, { quick: true }), JSON.parse(env.row()[C.ATTACH_LIST])];
  });
  check('R5 PENDING + 업로드 — 인식·Y열 이전과 동일', r.rb[0] === true && JSON.stringify(r.rb) === JSON.stringify(r.rw), firstDiff(r.rb, r.rw));
  check('R5 — 시트 전체 이전과 동일', JSON.stringify(r.sb.rows) === JSON.stringify(r.sw.rows), firstDiff(r.sb.rows, r.sw.rows));
  const q = runBoth(env => {
    submit(env, 'TQ-AP-26-021', [att(XLSX, true)]);
    return [recon(env, { quick: true }), recon(env)];   // QUEUED: quick 은 안 보고, full 은 폴더 확인(없음)
  });
  check('R6 QUEUED — quick/full 결과 이전과 동일', JSON.stringify(q.rb) === JSON.stringify(q.rw) && q.rw[0] === false && q.rw[1] === false, [q.rb, q.rw]);
})();
(function () {
  // FINAL 이동 경로(checkReplaced·lockHeld) — 미지정 건도 잡고, 락 재진입 없음
  const w = unpickedEnv();
  w.upload(w.folderOf(), 'QUO_TQ-AP-26-004.pdf');
  w.state.locked = true;   // 호출자가 락을 쥔 상태
  let ok, err;
  try { ok = recon(w, { checkReplaced: true, lockHeld: true }); } catch (e) { err = e.message; }
  w.state.locked = false;
  check('R7 FINAL 이동 경로(lockHeld) — 인식, 락 재진입 없음', ok === true && !err, [ok, err]);
})();

// ═════════════════════════════════════════════════════════════════
console.log('\n[L] 관리자 홈 목록');
(function () {
  const w = unpickedEnv();
  let list = w.call('listPendingQuotePdf');
  check('L1 미지정 건 표시(state UNPICKED·since·파일·QUO_ 이름)', list.length === 1 && list[0].state === 'UNPICKED' && list[0].docNo === 'TQ-AP-26-004'
    && list[0].since === '2026-11-23T01:00:00.000Z' && list[0].files.length === 2 && list[0].files[0].name === QPDF && list[0].fileCount === 2
    && list[0].fileName === 'QUO_TQ-AP-26-004.pdf' && list[0].token === w.row()[w.C.TOKEN], list);
  w.upload(w.folderOf(), 'QUO_TQ-AP-26-004.pdf');
  list = w.call('listPendingQuotePdf');
  check('L2 QUO_ 업로드 후 목록에서 빠지고 등록됨', list.length === 0 && w.y().some(f => f.role === 'quote'), [list, w.y()]);
})();
(function () {
  // 기존 PENDING/QUEUED 항목은 이전 코드와 같은 모양
  const r = runBoth(env => {
    submit(env, 'TQ-AP-26-030', [att(XLSX, true)]);
    const C = env.C; const l = JSON.parse(env.row()[C.ATTACH_LIST]); l[0].quotePdf = 'PENDING'; l[0].quoteReason = '용량 초과'; l[0].quotePendingAt = '2026-11-22T00:00:00.000Z';
    env.row()[C.ATTACH_LIST] = JSON.stringify(l);
    submit(env, 'TQ-AP-26-031', [att(QPDF, true)]);   // DONE — 목록 제외
    submit(env, 'TQ-AP-26-032', []);                  // 첨부 0건 — 목록 제외
    return env.call('listPendingQuotePdf');
  });
  check('L3 PENDING 1건만 — 이전 코드와 동일', r.rb.length === 1 && JSON.stringify(r.rb) === JSON.stringify(r.rw), firstDiff(r.rb, r.rw));
})();
(function () {
  const w = unpickedEnv();
  w.row()[w.C.STATUS] = '폐기';
  check('L4 폐기 행은 미지정이어도 제외', w.call('listPendingQuotePdf').length === 0);
  const w2 = unpickedEnv();
  w2.sb.ADMIN_TASK_EXCLUDE_DOCNOS.push('TQ-AP-26-004');
  check('L5 제외 목록(ADMIN_TASK_EXCLUDE_DOCNOS) 적용', w2.call('listPendingQuotePdf').length === 0);
})();

// ═════════════════════════════════════════════════════════════════
console.log('\n[D] [견적서 아님]');
(function () {
  const w = unpickedEnv();
  const token = w.row()[w.C.TOKEN];
  const before = JSON.stringify(w.row());
  const r0 = w.as(KIM).call('dismissQuoteUnpickedForClient', { token });
  check('D1 비관리자 → 거부, 행 그대로', r0.ok === false && JSON.stringify(w.row()) === before, r0);
  const r00 = w.as(ADMIN).call('dismissQuoteUnpickedForClient', {});
  check('D2 token 없음 → 거부', r00.ok === false, r00);
  const r1 = w.as(ADMIN).call('dismissQuoteUnpickedForClient', { token });
  const y = w.y();
  check('D3 관리자 → ok, 표시만 지움(파일·이름·id 그대로)', r1.ok === true && !r1.already && y.length === 2 && y.every(f => !f.quoteUnpicked) && y[0].name === QPDF, [r1, y]);
  check('D3 — 감사로그 QUOTE_UNPICKED_DISMISSED (actor=관리자)', w.audit().some(r => r.indexOf('QUOTE_UNPICKED_DISMISSED') >= 0 && r.indexOf(ADMIN) >= 0));
  check('D3 — 목록에서 빠짐', w.call('listPendingQuotePdf').length === 0);
  const r2 = w.as(ADMIN).call('dismissQuoteUnpickedForClient', { token });
  check('D4 다시 → ok + already, 로그 추가 없음', r2.ok === true && r2.already === true && w.audit().filter(r => r.indexOf('QUOTE_UNPICKED_DISMISSED') >= 0).length === 1, r2);
  const r3 = w.as(ADMIN).call('dismissQuoteUnpickedForClient', { token: 'no-such-token' });
  check('D5 없는 토큰 → 거부 메시지', r3.ok === false && /찾을 수 없습니다/.test(r3.message), r3);
  check('D6 락 해제됨', w.state.locked === false);
})();

// ═════════════════════════════════════════════════════════════════
console.log('\n[K] markQuoteUnpicked (기존 제출분 구제)');
function oldRowEnv(docNo, atts) {
  // 이전 코드로 제출된 것과 같은 행(표시 없음)
  const w = makeEnv('WT', readWT); submit(w, docNo, atts);
  const l = w.y(); l.forEach(f => delete f.quoteUnpicked); w.row()[w.C.ATTACH_LIST] = JSON.stringify(l);
  w.state.mails.length = 0; return w;
}
const throws = fn => { try { fn(); return ''; } catch (e) { return e.message; } };
(function () {
  const w = oldRowEnv('TQ-AP-26-004', [att(QPDF), att(ETC)]);
  check('K1 비관리자 → 예외', /관리자만/.test(throws(() => w.as(KIM).call('markQuoteUnpicked', 'TQ-AP-26-004', QPDF))));
  const pdfId = w.y()[0].id;
  const msg = w.as(ADMIN).call('markQuoteUnpicked', 'TQ-AP-26-004', QPDF);
  const y = w.y();
  check('K2 PDF 지정 → 이름 QUO_ 로 변경·등록', /OK/.test(msg) && w.state.files[pdfId].name === 'QUO_TQ-AP-26-004.pdf'
    && y.length === 2 && y[1].id === pdfId && y[1].role === 'quote' && y[1].origName === QPDF && y.every(f => !f.quoteUnpicked), [msg, y]);
  check('K2 — 감사로그 MARKED + GENERATED', w.audit().some(r => r.indexOf('QUOTE_UNPICKED_MARKED') >= 0) && w.audit().some(r => r.indexOf('QUOTE_PDF_GENERATED') >= 0));
  check('K2 — 목록에 남지 않음', w.call('listPendingQuotePdf').length === 0);
  check('K3 이미 등록된 행 → 예외', /이미 견적서/.test(throws(() => w.as(ADMIN).call('markQuoteUnpicked', 'TQ-AP-26-004'))));
  check('K4 락 해제됨', w.state.locked === false);
})();
(function () {
  const w = oldRowEnv('TQ-AP-26-003', [att('거래명세서.pdf')]);
  const msg = w.as(ADMIN).call('markQuoteUnpicked', 'TQ-AP-26-003');
  check('K5 표시만 → 목록 UNPICKED, 파일 그대로', /견적서 미지정 표시/.test(msg) && w.call('listPendingQuotePdf').length === 1 && w.liveFiles(w.folderOf()).indexOf('거래명세서.pdf') >= 0, msg);
  check('K6 두 번 → 이미 있음', /이미 있음/.test(w.as(ADMIN).call('markQuoteUnpicked', 'TQ-AP-26-003')));
})();
(function () {
  const w = oldRowEnv('TQ-AP-26-005', [att(ETC), att(QPDF)]);
  check('K7 없는 파일명 → 예외', /그 파일이 없습니다/.test(throws(() => w.as(ADMIN).call('markQuoteUnpicked', 'TQ-AP-26-005', '없는.pdf'))));
  check('K8 PDF 아닌 첨부 → 예외', /PDF 첨부만/.test(throws(() => w.as(ADMIN).call('markQuoteUnpicked', 'TQ-AP-26-005', ETC))));
  check('K9 없는 품의번호 → 예외', /특정할 수 없습니다/.test(throws(() => w.as(ADMIN).call('markQuoteUnpicked', 'TQ-AP-26-999'))));
  w.upload(w.folderOf(), 'QUO_TQ-AP-26-005.pdf');
  const e = throws(() => w.as(ADMIN).call('markQuoteUnpicked', 'TQ-AP-26-005', QPDF));
  check('K10 폴더에 다른 QUO_ 가 있으면 이름 변경 중단', /이미 다른/.test(e) && w.liveFiles(w.folderOf()).indexOf(QPDF) >= 0, e);
  check('K11 락 해제됨(예외 경로)', w.state.locked === false);
})();
(function () {
  const w = makeEnv('WT', readWT); submit(w, 'TQ-AP-26-006', []);
  check('K12 첨부 0건 행 → 예외', /첨부가 없는/.test(throws(() => w.as(ADMIN).call('markQuoteUnpicked', 'TQ-AP-26-006'))));
  const w2 = oldRowEnv('TQ-AP-26-007', [att(QPDF)]); w2.row()[w2.C.STATUS] = '폐기';
  check('K13 폐기 행 → 예외', /특정할 수 없습니다/.test(throws(() => w2.as(ADMIN).call('markQuoteUnpicked', 'TQ-AP-26-007'))));
})();

// ═════════════════════════════════════════════════════════════════
console.log('\n[Q] 큐 자동 변환 — 이전 코드와 동일');
(function () {
  const r = runBoth(env => {
    submit(env, 'TQ-AP-26-040', [att(XLSX, true), att(ETC)]);
    return env.call('_processQuotePdfJob', { type: 'quote_pdf', token: env.row()[env.C.TOKEN], docNo: 'TQ-AP-26-040' });
  });
  check('Q1 변환 job — 성공', r.rb.ok === true && r.rw.ok === true, [r.rb, r.rw]);
  check('Q1 — 응답·시트·Drive·로그 메일 동일', JSON.stringify(r.rb) === JSON.stringify(r.rw) && JSON.stringify(r.sb) === JSON.stringify(r.sw), firstDiff(r.sb, r.sw));
})();

// ═════════════════════════════════════════════════════════════════
console.log('\n[H] 관리자 홈 응답(getHomeDataForClient)');
function homeEnv() {
  const w = makeEnv('WT', readWT);
  submit(w, 'TQ-AP-26-004', [att(QPDF), att(ETC)]);                  // 미지정
  submit(w, 'TQ-AP-26-030', [att(XLSX, true)]);                      // PENDING 으로 만든다
  const C = w.C; const r = w.main().rows[2]; const l = JSON.parse(r[C.ATTACH_LIST]);
  l[0].quotePdf = 'PENDING'; l[0].quoteReason = '용량 초과 12MB'; l[0].quotePendingAt = '2026-11-22T00:00:00.000Z'; r[C.ATTACH_LIST] = JSON.stringify(l);
  submit(w, 'TQ-AP-26-031', [att(QPDF, true)]);                      // 지정 완료 — 배너 없음
  return w;
}
(function () {
  const w = homeEnv();
  const h = w.as(ADMIN).call('getHomeDataForClient');
  const q = (h.adminTasks || {}).quotePdf || [];
  check('H1 관리자 홈 — 견적서 대기 2건(UNPICKED·PENDING)', h.ok === true && q.length === 2 && q.map(x => x.state).sort().join() === 'PENDING,UNPICKED', q);
  const k = w.as(KIM).call('getHomeDataForClient');
  check('H2 일반 사용자 홈 — adminTasks 없음', k.ok === true && !k.adminTasks, k.adminTasks);
  if (process.env.DUMP_FIXTURES) {
    const out = { homeAdmin: norm(h), dismissToken: q.find(x => x.state === 'UNPICKED').token };
    fs.writeFileSync(process.env.DUMP_FIXTURES, JSON.stringify(out, null, 1));
    console.log('  픽스처 저장: ' + process.env.DUMP_FIXTURES);
  }
})();

// ═════════════════════════════════════════════════════════════════
console.log('\n' + (fail ? '✗ 실패 ' + fail + '건 / ' : '✓ ') + '통과 ' + pass + '건');
if (fail) { console.log('실패 목록:\n - ' + failures.join('\n - ')); process.exit(1); }
