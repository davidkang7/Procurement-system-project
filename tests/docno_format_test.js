/**
 * 품의번호 형식 점검 — 로컬 검증 (배포 없이 실제 .gs 함수 호출)
 * (memory: gas-local-test-harness)
 *
 * 규칙: 영문 대문자 2자리-영문 대문자 2자리-발행 일자 연도 끝 2자리-일련번호 3자리 (예: TO-AP-26-001)
 *  [U] reqDocNoFormatError_ 판정표 — 운영 DB에서 실제로 나온 틀린 번호들 + 연말·연초 경계
 *  [B] 품의 최초 제출: 틀린 번호는 화면 경로·doPost 경로 모두 시트·폴더·메일 0건으로 막힌다
 *  [S] 올바른 번호의 제출은 이전 코드(BASE_REF)와 시트 행·요약·메일이 바이트 단위로 같다
 *  [R] 옛 형식 번호의 반려 문서도 재상신은 그대로 된다 (재상신은 시트 번호를 유지)
 *
 * 실행: node tests/docno_format_test.js
 *  BASE_REF=<커밋>  비교 기준 (기본 362803c = 이 기능 직전 커밋)
 *  WT_DIR=<폴더>    뮤테이션 검사용 — 일부러 버그를 심은 사본을 '수정본'으로 돌린다
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execSync } = require('child_process');

const SRC = path.join(__dirname, '..');
const WT_DIR = process.env.WT_DIR || SRC;
const BASE_REF = process.env.BASE_REF || '362803c';
const readWT   = f => fs.readFileSync(path.join(WT_DIR, f), 'utf8');
const readBASE = f => execSync('git show ' + BASE_REF + ':"' + f + '"', { cwd: SRC, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });

// ── 고정 시계 (두 환경 공용) ──────────────────────────────────────
const RealDate = Date;
const CLOCK = { now: RealDate.parse('2026-10-30T10:00:00+09:00') };
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
  const state = { label, active: KIM, mails: [], files: [], adminErrors: [], uuid: 0, props: new Map(), cache: new Map(), folders: {} };
  const ss = new FakeSpreadsheet();
  function folderRec(id) { if (!state.folders[id]) state.folders[id] = { id, files: [] }; return state.folders[id]; }
  function wrapFolder(f) {
    const empty = () => ({ hasNext: () => false, next: () => null });
    return {
      getId: () => f.id, getName: () => f.id, isTrashed: () => false, setTrashed() {},
      createFile: blob => {
        const bytes = Buffer.from((blob.bytes || []).map(b => (b < 0 ? b + 256 : b)));
        const nf = { id: 'file-' + (++state.uuid), name: blob.name, type: blob.type, bytes };
        f.files.push(nf);
        return { getId: () => nf.id, getName: () => nf.name, setTrashed() {}, setSharing() {} };
      },
      getFiles: empty, getFilesByName: empty, getFoldersByName: empty, getParents: empty,
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
      getFolderById: id => { const f = state.folders[id]; if (!f) throw new Error('No item with the given ID: ' + id); return wrapFolder(f); },
      getFileById: () => ({ setTrashed() {} }),
    },
    Utilities: {
      getUuid: () => 'uuid-' + String(++state.uuid).padStart(8, '0'),
      // GAS와 같이 부호 있는 바이트 배열(Byte[])
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
      getScriptLock: () => ({ waitLock() {}, tryLock() { return true; }, releaseLock() {} }),
      getUserLock: () => ({ tryLock() { return true; }, releaseLock() {} }),
    },
    PropertiesService: { getScriptProperties: () => propsObj },
    CacheService: { getScriptCache: () => cacheObj },
    GmailApp: {
      sendEmail: (to, subj, plain, opts) => { state.mails.push({ to, subj, plain, html: (opts || {}).htmlBody || '' }); },
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

  // 범위 밖 부수효과만 두 환경에 똑같이 막는다 — 계산·요약·메일 조립은 실제 코드
  sandbox.__state = state;
  sandbox.__folderRec = folderRec;
  sandbox.__wrapFolder = wrapFolder;
  vm.runInContext(`
    notifyAdminError             = function (m) { __state.adminErrors.push(String(m)); };
    getOrCreateFolder            = function (name) { return __wrapFolder(__folderRec('F_' + name)); };
    createFileInFolder           = function (folder, name, content) { __state.files.push({ name: name, content: String(content) }); };
    saveAttachmentsToFolder      = function () { return []; };
    _enqueueQuotePdfIfNeeded_    = function () {};
    _notifyQuoteMissingIfNeeded_ = function () {};
  `, sandbox);
  sandbox.CONFIG.PROCUREMENT_TEAM_EMAILS = [ADMIN, BUYER1];
  sandbox.CONFIG.ADMIN_EMAILS = [ADMIN];
  sandbox.getOrCreateSheet(ss, sandbox.CONFIG.SHEET_NAME);
  sandbox.ensureHeaders(ss.getSheetByName(sandbox.CONFIG.SHEET_NAME));

  return {
    label, sb: sandbox, ss, state,
    as(user) { state.active = user; return this; },
    call(fn, ...args) { return sandbox[fn].apply(sandbox, args); },
    main: () => ss.getSheetByName(sandbox.CONFIG.SHEET_NAME),
    rowByToken(token) { const C = sandbox.COL; return this.main().rows.find(r => r[C.TOKEN] === token); },
  };
}

// ── 결과 정규화 (Date → 문자열) ─────────────────────────────────
const norm = v => JSON.parse(JSON.stringify(v, (k, x) => (x instanceof RealDate ? 'D:' + x.toISOString() : x)));
function snap(env) {
  return norm({
    rows: env.main().rows,
    files: env.state.files,
    mails: env.state.mails,
    adminErrors: env.state.adminErrors,
  });
}

// ── 단언 ─────────────────────────────────────────────────────────
let pass = 0, fail = 0;
const failures = [];
function check(name, cond, detail) {
  if (cond === true) { pass++; console.log('  ✓ ' + name); }
  else { fail++; failures.push(name); console.log('  ✗ ' + name + (detail !== undefined ? '  → ' + JSON.stringify(detail).slice(0, 700) : '')); }
}
function firstDiff(a, b, p = '') {
  if (JSON.stringify(a) === JSON.stringify(b)) return null;
  if (typeof a !== 'object' || typeof b !== 'object' || !a || !b) return { at: p, base: a, wt: b };
  for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) { const d = firstDiff(a[k], b[k], p + '.' + k); if (d) return d; }
  return null;
}


// ── 픽스처 ───────────────────────────────────────────────────────
const KRW_ITEMS = [
  { name: '지그 본체', spec: 'SUS304 / EA', qty: 3, price: 15000, currency: 'KRW' },
  { name: '볼트 세트', spec: 'M8x20 / SET', qty: 10, price: 1100, currency: 'KRW' },
];

function reqPayload(docNo, issueDate = '2026-10-30', drafter = KIM) {
  return {
    docNo, issueDate, drafter: NAMES[drafter], dept: '개발팀', subject: '검사 지그 제작', purpose: '라인 검사',
    vendorName: '(주)태성테크', vendorEmail: 'sales@example.com', vendorContact: '박담당', vendorPhone: '042-000-0000',
    deliveryDate: '2026-11-20', deliveryAddr: '대전광역시 유성구', items: JSON.parse(JSON.stringify(KRW_ITEMS)), remarks: '비고',
    approvers: [{ label: '기안자', name: NAMES[drafter], email: drafter }, { label: '결재자 1', name: NAMES[LEE], email: LEE }],
    attachments: [], vatMode: 'excl',
  };
}

function runBoth(fn) {
  const b = makeEnv('BASE', readBASE), w = makeEnv('WT', readWT);
  const rb = norm(fn(b)), rw = norm(fn(w));
  return { b, w, rb, rw, sb: snap(b), sw: snap(w) };
}
const C_ = env => env.sb.COL;

const MSG_SHAPE = yy => '품의번호 형식이 올바르지 않습니다. 영문 대문자 2자리-영문 대문자 2자리-발행연도 2자리-일련번호 3자리로 입력해 주세요. (예: TO-AP-' + yy + '-001)';
const MSG_YEAR = (got, issue, yy) => '품의번호의 연도(' + got + ')가 발행 일자(' + issue + ')의 연도(' + yy + ')와 다릅니다.';
const MSG_DATE = '발행 일자를 입력해 주세요.';

// [docNo, issueDate, 기대 문구('' = 통과)]
const TABLE = [
  ['TO-AP-26-045', '2026-10-30', ''],
  ['TH-AP-26-060', '2026-10-29', ''],
  ['LC-PR-26-021', '2026-10-30', ''],
  [' TO-AP-26-045 ', '2026-10-30', ''],                          // 앞뒤 공백은 화면이 지우지만 서버도 무시
  ['TR-AP-026-56', '2026-10-15', MSG_SHAPE('26')],               // 운영 DB 실제 사례 — 연도 3자리·번호 2자리
  ['TO-AP-026-045', '2026-10-30', MSG_SHAPE('26')],              // 연도 3자리
  ['LC-2026-020', '2026-10-28', MSG_SHAPE('26')],                // 운영 DB 실제 사례
  ['TP-AP-26-24', '2026-10-29', MSG_SHAPE('26')],                // 운영 DB 실제 사례 — 번호 2자리
  ['PRQ-2026-001', '2026-10-30', MSG_SHAPE('26')],               // 이전 화면 기본값
  ['to-ap-26-045', '2026-10-30', MSG_SHAPE('26')],               // 소문자
  ['TO-AP-26-0451', '2026-10-30', MSG_SHAPE('26')],
  ['TOO-AP-26-045', '2026-10-30', MSG_SHAPE('26')],
  ['TO-A1-26-045', '2026-10-30', MSG_SHAPE('26')],
  ['TO_AP_26_045', '2026-10-30', MSG_SHAPE('26')],
  ['TO-AP-26-045-01', '2026-10-30', MSG_SHAPE('26')],
  ['', '2026-10-30', MSG_SHAPE('26')],
  ['TR-AP-20-056', '2026-10-15', MSG_YEAR('20', '2026-10-15', '26')],   // 운영 DB 실제 사례 — 연도 오타
  ['TO-AP-27-001', '2027-01-02', ''],                                    // 연초
  ['TO-AP-27-001', '2026-12-31', MSG_YEAR('27', '2026-12-31', '26')],   // 연말에 다음 해 번호
  ['TO-AP-26-001', '2027-01-02', MSG_YEAR('26', '2027-01-02', '27')],   // 연초에 지난해 번호
  ['TO-AP-26-045', '', MSG_DATE],
  ['TO-AP-26-045', undefined, MSG_DATE],
  ['TO-AP-26-045', '2026/10/30', ''],                                    // 서버 toDateStr가 날짜로 읽음
];

// ═════════════════════════════════════════════════════════════════
console.log('\n[U] 판정표 — reqDocNoFormatError_ (서버) · docNoFormatError (화면)');
(function () {
  const w = makeEnv('WT', readWT);
  TABLE.forEach(([docNo, issue, want]) => {
    const got = w.call('reqDocNoFormatError_', docNo, issue);
    check(`U 서버 ${JSON.stringify(docNo)} @ ${JSON.stringify(issue)} → ${want ? '막힘' : '통과'}`, got === want, { got, want });
  });
  const d = w.call('reqDocNoFormatError_', 'TO-AP-26-045', new FakeDate('2026-10-30T00:00:00+09:00'));
  check('U 서버 — 발행 일자가 Date 객체여도 같은 판정', d === '', d);

  // 화면 함수 — 같은 규칙·문구 (발행 일자는 <input type=date> 값이라 YYYY-MM-DD 또는 빈 값만 들어온다)
  const html = readWT('Procurement_Form.html');
  const src = html.slice(html.indexOf('var DOC_NO_RE'), html.indexOf('// ─── 품의번호 중복 검사'));
  const ctx = {}; vm.createContext(ctx); vm.runInContext(src, ctx);
  TABLE.filter(([, issue]) => issue === undefined || issue === '' || /^\d{4}-\d{2}-\d{2}$/.test(issue)).forEach(([docNo, issue, want]) => {
    const dn = docNo.trim();   // 화면은 trim 후 검사
    const got = ctx.docNoFormatError(dn, issue);
    check(`U 화면 ${JSON.stringify(dn)} @ ${JSON.stringify(issue)} = 서버 문구`, got === want, { got, want });
  });
  check('U 화면 짧은 안내 — 형식 / 연도 / 발행 일자 없음',
    ctx.docNoFormatError('LC-2026-020', '2026-10-28', true) === '형식 오류 — 예: TO-AP-26-001'
    && ctx.docNoFormatError('TR-AP-20-056', '2026-10-15', true) === '연도는 26 (발행 일자 기준)'
    && ctx.docNoFormatError('TO-AP-26-045', '', true) === '발행 일자를 먼저 입력해 주세요'
    && ctx.docNoFormatError('TO-AP-26-045', '2026-10-30', true) === '');
})();

// ═════════════════════════════════════════════════════════════════
console.log('\n[B] 틀린 번호의 품의 제출 — 화면 경로·doPost 경로 모두 아무것도 남기지 않고 막힌다');
(function () {
  const bad = [['TO-AP-026-045', '2026-10-30'], ['LC-2026-021', '2026-10-30'], ['TR-AP-20-056', '2026-10-15'], ['PRQ-2026-001', '2026-10-30']];
  bad.forEach(([docNo, issue]) => {
    ['client', 'doPost'].forEach(path => {
      const w = makeEnv('WT', readWT).as(KIM);
      w.sb.ContentService.createTextOutput = s => ({ body: s, setMimeType() { return this; } });
      const props0 = JSON.stringify([...w.state.props]);
      const p = reqPayload(docNo, issue);
      const out = path === 'client'
        ? w.call('submitRequisitionForClient', p)
        : JSON.parse(w.call('doPost', { postData: { contents: JSON.stringify(Object.assign({ action: 'submit' }, p)) } }).body);
      const want = w.call('reqDocNoFormatError_', docNo, issue);
      check(`B ${path} ${docNo} → ok:false · 형식 문구`, out && out.ok === false && out.message === want && want !== '', out);
      check(`B ${path} ${docNo} → 시트 행·폴더·파일·메일·속성 변화 0건`,
        w.main().rows.length === 1 && Object.keys(w.state.folders).length === 0 && w.state.files.length === 0
        && w.state.mails.length === 0 && JSON.stringify([...w.state.props]) === props0,
        { rows: w.main().rows.length, folders: Object.keys(w.state.folders), mails: w.state.mails.length });
    });
    // 이전 코드는 받아 줬다 (이번 변경으로 막히는 것)
    const b = makeEnv('BASE', readBASE).as(KIM);
    const rb = b.call('submitRequisitionForClient', reqPayload(docNo, issue));
    check(`B ${docNo} — 이전 코드는 제출됨`, rb.ok === true, rb);
  });

  // 기안자 본인 확인이 먼저 — 남의 이름으로 낸 제출은 형식과 상관없이 이전과 같은 문구
  const w = makeEnv('WT', readWT).as(LEE);
  const r = w.call('submitRequisitionForClient', reqPayload('LC-2026-021'));
  check('B 기안자 본인 확인이 먼저 — 이전과 같은 문구', r.ok === false && /기안자는 본인만/.test(r.message), r);
})();

// ═════════════════════════════════════════════════════════════════
console.log('\n[S] 올바른 번호의 품의 제출 — 이전 코드와 바이트 단위로 같아야 한다');
[
  ['S1 TO-AP-26-045', 'TO-AP-26-045', '2026-10-30'],
  ['S2 연초 TO-AP-27-001 @ 2027-01-02', 'TO-AP-27-001', '2027-01-02'],
].forEach(([name, docNo, issue]) => {
  const r = runBoth(env => env.as(KIM).call('submitRequisitionForClient', reqPayload(docNo, issue)));
  check(name + ' — 제출 성공', r.rb.ok === true && r.rw.ok === true, [r.rb, r.rw]);
  check(name + ' — 응답 동일', JSON.stringify(r.rb) === JSON.stringify(r.rw), firstDiff(r.rb, r.rw));
  check(name + ' — 시트 행·요약 파일·결재 메일 동일', JSON.stringify(r.sb) === JSON.stringify(r.sw), firstDiff(r.sb, r.sw));
});
(function () {
  // 중복 번호는 이전처럼 중복 문구 (형식은 통과)
  const r = runBoth(env => { env.as(KIM).call('submitRequisitionForClient', reqPayload('TO-AP-26-046')); return env.call('submitRequisitionForClient', reqPayload('TO-AP-26-046')); });
  check('S3 중복 번호 — 이전과 같은 중복 문구', r.rw.ok === false && JSON.stringify(r.rb) === JSON.stringify(r.rw) && /이미 사용 중/.test(r.rw.message), [r.rb, r.rw]);
})();

// ═════════════════════════════════════════════════════════════════
console.log('\n[R] 옛 형식 번호 — 재상신·구매(PRC)는 막지 않는다');
(function () {
  const r = runBoth(env => {
    env.as(KIM).call('submitRequisitionForClient', reqPayload('TO-AP-26-047'));
    const C = C_(env), row = env.main().rows[1];
    row[C.DOC_NO] = 'LC-2026-020'; row[C.STATUS] = '반려';       // 배포 전에 옛 형식으로 제출돼 반려된 문서
    env.state.mails.length = 0; env.state.files.length = 0;
    const d = reqPayload('LC-2026-020'); d.subject = '수정 품의';
    return env.call('resubmitRequisitionForClient', { token: row[C.TOKEN], data: d });
  });
  check('R1 옛 형식 반려 문서 재상신 — 성공', r.rb.ok === true && r.rw.ok === true, [r.rb, r.rw]);
  check('R1 품의번호 LC-2026-020 그대로 유지', r.w.main().rows[1][C_(r.w).DOC_NO] === 'LC-2026-020');
  check('R1 시트·요약·메일 이전과 동일', JSON.stringify(r.sb) === JSON.stringify(r.sw), firstDiff(r.sb, r.sw));

  const p = runBoth(env => {
    env.as(KIM).call('submitRequisitionForClient', reqPayload('TO-AP-26-048'));
    const C = C_(env), parent = env.main().rows[1];
    parent[C.DOC_NO] = 'LC-2026-019'; parent[C.STATUS] = '최종승인'; parent[C.CLAIMED_BY] = BUYER1;
    env.state.mails.length = 0; env.state.files.length = 0;
    const q = reqPayload('LC-PO-2026-01', '2026-10-30', BUYER1);   // PO No.는 이번 규칙 대상 아님
    Object.assign(q, { parentToken: parent[C.TOKEN], dept: '구매팀', purchaseMethod: '계좌이체', paymentInfo: '익월말', existingAttachments: [], fieldChanges: [] });
    return env.as(BUYER1).call('submitPrcForClient', q);
  });
  check('R2 옛 형식 품의의 구매(PRC) 제출 · 형식 다른 PO No. — 성공', p.rb.ok === true && p.rw.ok === true, [p.rb, p.rw]);
  check('R2 시트·요약·메일 이전과 동일', JSON.stringify(p.sb) === JSON.stringify(p.sw), firstDiff(p.sb, p.sw));
})();

console.log('\n─────────────────────────────');
console.log('통과 ' + pass + ' / 실패 ' + fail + '   (기준 ' + BASE_REF + (WT_DIR !== SRC ? ', 수정본 ' + WT_DIR : '') + ')');
if (fail) { console.log('실패 항목:\n  - ' + failures.join('\n  - ')); process.exitCode = 1; }
