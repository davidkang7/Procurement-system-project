/**
 * 부가세 포함/별도(단가 기준) · 검수 사진 합계 한도 — 로컬 검증 (배포 없이 실제 .gs 함수 호출)
 * (memory: gas-local-test-harness, vat-toggle-photo-limit-plan)
 *
 * 방식: 같은 제출을 이전 코드(BASE_REF, 기본 ef65788 = 이 기능 직전 커밋)와 작업 트리에서 각각 돌린다.
 *  [S] '부가세 별도' 제출(REQ·재상신·PRC) — 시트 행·요약 파일·결재 메일이 이전 코드와 바이트 단위로 같아야 한다.
 *      화면이 vatMode를 안 보내는 옛 화면, 'excl'을 보내는 새 화면, 외화인데 'incl'을 보낸 조작 요청 모두.
 *  [V] '부가세 포함' 제출 — O열 = round(합계 ÷ 1.1), P열 = 품목 합, 요약·메일 금액. 이전 코드와는 달라야 한다.
 *  [D] 판별(detectVatMode) — 조회·구매 작성·재상신·PDF·주문서 매니페스트가 같은 답을 내는지.
 *  [P] 검수 사진 — 장당 5MB 검사 삭제, 실제 바이트 합계 15MB, 6장 한도 유지.
 *  메일 조립 함수는 실제 코드를 쓰고, 맨 아래 GmailApp.sendEmail만 가로챈다.
 *
 * 실행: node tests/vat_photo_test.js
 *  BASE_REF=<커밋>  비교 기준
 *  WT_DIR=<폴더>    뮤테이션 검사용 — 일부러 버그를 심은 사본을 '수정본'으로 돌린다
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execSync } = require('child_process');

const SRC = path.join(__dirname, '..');
const WT_DIR = process.env.WT_DIR || SRC;
const BASE_REF = process.env.BASE_REF || 'ef65788';
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
];                                                                        // 합 56,000
const ODD_ITEMS = [{ name: '센서', spec: 'S-1 / EA', qty: 7, price: 1234, currency: 'KRW' }];   // 합 8,638 (반올림 경계 확인)
const USD_ITEMS = [{ name: 'Laser module', spec: 'LM-532 / EA', qty: 2, price: 1250.5, currency: 'USD' }];

function reqPayload(docNo, items, vatMode, drafter = KIM) {
  const p = {
    docNo, issueDate: '2026-10-30', drafter: NAMES[drafter], dept: '개발팀', subject: '검사 지그 제작', purpose: '라인 검사',
    vendorName: '(주)태성테크', vendorEmail: 'sales@example.com', vendorContact: '박담당', vendorPhone: '042-000-0000',
    deliveryDate: '2026-11-20', deliveryAddr: '대전광역시 유성구', items: JSON.parse(JSON.stringify(items)), remarks: '비고',
    approvers: [{ label: '기안자', name: NAMES[drafter], email: drafter }, { label: '결재자 1', name: NAMES[LEE], email: LEE }],
    attachments: [],
  };
  if (vatMode !== undefined) p.vatMode = vatMode;
  return p;
}

function runBoth(fn) {
  const b = makeEnv('BASE', readBASE), w = makeEnv('WT', readWT);
  const rb = norm(fn(b)), rw = norm(fn(w));
  return { b, w, rb, rw, sb: snap(b), sw: snap(w) };
}
const C_ = env => env.sb.COL;

// ═════════════════════════════════════════════════════════════════
console.log('\n[S] 부가세 별도 — 이전 코드와 바이트 단위로 같아야 한다');
[
  ['S1 REQ 원화 · 옛 화면(vatMode 없음)', KRW_ITEMS, undefined],
  ['S2 REQ 원화 · 새 화면 vatMode=excl', KRW_ITEMS, 'excl'],
  ['S3 REQ 원화 · 반올림 경계 합 8,638', ODD_ITEMS, 'excl'],
  ['S4 REQ 외화(USD) · vatMode=incl 조작 요청', USD_ITEMS, 'incl'],
  ['S5 REQ 원화 · 알 수 없는 값 vatMode="yes"', KRW_ITEMS, 'yes'],
].forEach(([name, items, mode]) => {
  const r = runBoth(env => env.as(KIM).call('submitRequisitionForClient', reqPayload('TH-AP-26-501', items, mode)));
  check(name + ' — 제출 성공', r.rb.ok === true && r.rw.ok === true, [r.rb, r.rw]);
  check(name + ' — 응답 동일', JSON.stringify(r.rb) === JSON.stringify(r.rw), firstDiff(r.rb, r.rw));
  check(name + ' — 시트 행·요약 파일·결재 메일 동일', JSON.stringify(r.sb) === JSON.stringify(r.sw), firstDiff(r.sb, r.sw));
});

// 재상신 (반려된 REQ) — 별도
(function () {
  const r = runBoth(env => {
    env.as(KIM).call('submitRequisitionForClient', reqPayload('TH-AP-26-502', KRW_ITEMS));
    const row = env.main().rows[1]; row[C_(env).STATUS] = '반려';
    const tok = row[C_(env).TOKEN];
    env.state.mails.length = 0; env.state.files.length = 0;
    return env.call('resubmitRequisitionForClient', { token: tok, data: reqPayload('TH-AP-26-502', ODD_ITEMS) });
  });
  check('S6 재상신 원화 별도 — 성공', r.rb.ok === true && r.rw.ok === true, [r.rb, r.rw]);
  check('S6 재상신 — 시트·요약·메일 동일', JSON.stringify(r.sb) === JSON.stringify(r.sw), firstDiff(r.sb, r.sw));
})();

// 재상신 화면이 통화를 보내게 된 것 — 원화 품목에 currency:'KRW'가 붙어도 결과는 이전(통화 없음)과 같아야 한다
(function () {
  const noCur = KRW_ITEMS.map(({ currency, ...x }) => x);
  const b = makeEnv('BASE', readBASE), w = makeEnv('WT', readWT);
  [b, w].forEach(env => {
    env.as(KIM).call('submitRequisitionForClient', reqPayload('TH-AP-26-503', KRW_ITEMS));
    const row = env.main().rows[1]; row[C_(env).STATUS] = '반려';
    env.state.mails.length = 0; env.state.files.length = 0;
    const p = reqPayload('TH-AP-26-503', env === b ? noCur : KRW_ITEMS, env === b ? undefined : 'excl');
    env.res = env.call('resubmitRequisitionForClient', { token: row[C_(env).TOKEN], data: p });
  });
  check('S7 재상신 원화 — 옛 화면(통화 없음) vs 새 화면(통화 KRW·excl) 시트·메일 동일',
    b.res.ok && w.res.ok && JSON.stringify(snap(b)) === JSON.stringify(snap(w)), firstDiff(snap(b), snap(w)));
})();

// PRC 제출 — 별도
function prcFlow(env, items, mode, parentItems, parentMode) {
  env.as(KIM).call('submitRequisitionForClient', reqPayload('TH-AP-26-510', parentItems || KRW_ITEMS, parentMode));
  const C = C_(env); const parent = env.main().rows[1];
  parent[C.STATUS] = '최종승인'; parent[C.CLAIMED_BY] = BUYER1;
  env.state.mails.length = 0; env.state.files.length = 0;
  const p = reqPayload('TO-PO-26-910', items, mode, BUYER1);
  Object.assign(p, { parentToken: parent[C.TOKEN], dept: '구매팀', purchaseMethod: '계좌이체', paymentInfo: '익월말', existingAttachments: [], fieldChanges: [] });
  return env.as(BUYER1).call('submitPrcForClient', p);
}
(function () {
  const r = runBoth(env => prcFlow(env, KRW_ITEMS, 'excl'));
  check('S8 PRC 원화 별도 — 성공', r.rb.ok === true && r.rw.ok === true, [r.rb, r.rw]);
  check('S8 PRC — 시트·요약·메일 동일', JSON.stringify(r.sb) === JSON.stringify(r.sw), firstDiff(r.sb, r.sw));
})();
// 외화 + 'incl' 조작 요청 — 재상신·PRC 경로에서도 서버가 무시해야 한다 (요약 파일까지 이전과 같음)
(function () {
  const r = runBoth(env => {
    env.as(KIM).call('submitRequisitionForClient', reqPayload('TH-AP-26-504', USD_ITEMS));
    const row = env.main().rows[1]; row[C_(env).STATUS] = '반려';
    env.state.mails.length = 0; env.state.files.length = 0;
    return env.call('resubmitRequisitionForClient', { token: row[C_(env).TOKEN], data: reqPayload('TH-AP-26-504', USD_ITEMS, 'incl') });
  });
  check('S9 재상신 외화 · vatMode=incl 조작 — 시트·요약·메일 동일', r.rb.ok && r.rw.ok && JSON.stringify(r.sb) === JSON.stringify(r.sw), firstDiff(r.sb, r.sw));
  const p = runBoth(env => prcFlow(env, USD_ITEMS, 'incl', USD_ITEMS));
  check('S10 PRC 외화 · vatMode=incl 조작 — 시트·요약·메일 동일', p.rb.ok && p.rw.ok && JSON.stringify(p.sb) === JSON.stringify(p.sw), firstDiff(p.sb, p.sw));
})();

// ═════════════════════════════════════════════════════════════════
console.log('\n[V] 부가세 포함 — 새 계산');
(function () {
  const w = makeEnv('WT', readWT).as(KIM);
  const res = w.call('submitRequisitionForClient', reqPayload('TH-AP-26-520', KRW_ITEMS, 'incl'));
  const C = C_(w), row = w.main().rows[1];
  check('V1 제출 성공', res.ok === true, res);
  check('V1 O열 공급가액 = round(56,000 ÷ 1.1) = 50,909', row[C.TOTAL_AMT] === 50909, row[C.TOTAL_AMT]);
  check('V1 P열 합계 = 품목 합 56,000', row[C.TOTAL_VAT] === 56000, row[C.TOTAL_VAT]);
  check('V1 품목내역 단가는 입력 그대로 (15,000 · 1,100)', /3개 \| 15,000 KRW/.test(row[C.ITEMS]) && /10개 \| 1,100 KRW/.test(row[C.ITEMS]), row[C.ITEMS]);
  const sum = w.state.files.find(f => /품의서요약/.test(f.name)) || {};
  check('V1 요약 파일 — 공급가액 50,909 · 부가세 5,091 · 합계 56,000 · "단가 기준: 부가세 포함"',
    /공급가액:\s+50,909 KRW/.test(sum.content) && /부가세:\s+5,091 KRW/.test(sum.content) && /합계:\s+56,000 KRW/.test(sum.content)
    && /단가 기준:\s+부가세 포함/.test(sum.content), sum.content);
  const mail = w.state.mails[0] || {};
  check('V1 결재 메일 "부가세 제외 금액" = 공급가액 50,909 원 (품목 합 56,000 아님)',
    /부가세 제외 금액: 50,909 원/.test(mail.plain) && !/56,000 원/.test(mail.plain), mail.plain);

  const b = makeEnv('BASE', readBASE).as(KIM);
  b.call('submitRequisitionForClient', reqPayload('TH-AP-26-520', KRW_ITEMS, 'incl'));
  check('V1 이전 코드는 같은 요청을 별도로 계산했다 (O 56,000 / P 61,600) — 의도한 차이',
    b.main().rows[1][C.TOTAL_AMT] === 56000 && b.main().rows[1][C.TOTAL_VAT] === 61600);

  const w2 = makeEnv('WT', readWT).as(KIM);
  w2.call('submitRequisitionForClient', reqPayload('TH-AP-26-521', ODD_ITEMS, 'incl'));
  const r2 = w2.main().rows[1];
  check('V2 반올림 경계: 합 8,638 → 공급가액 7,853 · P 8,638 (부가세 785)', r2[C.TOTAL_AMT] === 7853 && r2[C.TOTAL_VAT] === 8638, [r2[C.TOTAL_AMT], r2[C.TOTAL_VAT]]);

  // 재상신 — 포함으로 바꿔 다시 올림
  const w3 = makeEnv('WT', readWT).as(KIM);
  w3.call('submitRequisitionForClient', reqPayload('TH-AP-26-522', KRW_ITEMS));
  const r3 = w3.main().rows[1]; r3[C.STATUS] = '반려';
  const rr = w3.call('resubmitRequisitionForClient', { token: r3[C.TOKEN], data: reqPayload('TH-AP-26-522', KRW_ITEMS, 'incl') });
  check('V3 재상신 별도 → 포함: O 50,909 · P 56,000', rr.ok === true && r3[C.TOTAL_AMT] === 50909 && r3[C.TOTAL_VAT] === 56000, [rr, r3[C.TOTAL_AMT], r3[C.TOTAL_VAT]]);

  // PRC — 포함
  const w4 = makeEnv('WT', readWT);
  const rp = prcFlow(w4, KRW_ITEMS, 'incl', KRW_ITEMS, 'incl');
  const prc = w4.main().rows[2];
  check('V4 PRC 포함: O 50,909 · P 56,000', rp.ok === true && prc[C.TOTAL_AMT] === 50909 && prc[C.TOTAL_VAT] === 56000, [rp, prc && prc[C.TOTAL_AMT]]);
  check('V4 PRC 결재 메일 금액 = 공급가액 50,909 원', /부가세 제외 금액: 50,909 원/.test((w4.state.mails[0] || {}).plain), (w4.state.mails[0] || {}).plain);

  // 외화에 incl — 무시
  const w5 = makeEnv('WT', readWT).as(KIM);
  w5.call('submitRequisitionForClient', reqPayload('TH-AP-26-523', USD_ITEMS, 'incl'));
  const r5 = w5.main().rows[1];
  check('V5 외화는 incl을 보내도 부가세 0 (O = P = 2,501)', r5[C.TOTAL_AMT] === 2501 && r5[C.TOTAL_VAT] === 2501, [r5[C.TOTAL_AMT], r5[C.TOTAL_VAT]]);
})();

// ═════════════════════════════════════════════════════════════════
console.log('\n[D] 판별 — 화면·PDF·주문서가 같은 답');
(function () {
  const w = makeEnv('WT', readWT).as(KIM);
  w.call('submitRequisitionForClient', reqPayload('TH-AP-26-530', KRW_ITEMS, 'incl'));
  w.call('submitRequisitionForClient', reqPayload('TH-AP-26-531', KRW_ITEMS, 'excl'));
  w.call('submitRequisitionForClient', reqPayload('TH-AP-26-532', USD_ITEMS, 'excl'));
  w.call('submitRequisitionForClient', reqPayload('TH-AP-26-533', ODD_ITEMS, 'incl'));
  const C = C_(w), rows = w.main().rows.slice(1);
  const modes = rows.map(r => w.sb.detectRowVatMode_(r));
  check('D1 판별: 포함 → incl, 별도 → excl, 외화 → excl, 반올림 경계 포함 → incl', JSON.stringify(modes) === '["incl","excl","excl","incl"]', modes);

  const tokIncl = rows[0][C.TOKEN], tokExcl = rows[1][C.TOKEN];
  const v1 = w.as(KIM).call('getRequisitionForViewer', tokIncl), v2 = w.call('getRequisitionForViewer', tokExcl);
  check('D2 결재자 상세 데이터 vatMode (포함 incl / 별도 excl)', v1.ok && v1.data.vatMode === 'incl' && v2.ok && v2.data.vatMode === 'excl', [v1.message, v1.data && v1.data.vatMode, v2.data && v2.data.vatMode]);

  rows[0][C.STATUS] = '반려';
  const rj = w.call('getRejectDataForClient', tokIncl);
  check('D3 재상신 화면 데이터 vatMode incl + 품목 통화 KRW', rj.ok && rj.data.vatMode === 'incl' && rj.data.items[0].currency === 'KRW', rj);

  rows[0][C.STATUS] = '최종승인';
  const pf = w.as(KIM).call('getPrefillDataForClient', tokIncl);
  check('D4 구매 작성 화면 데이터 vatMode incl (품의서 이어받기)', pf.ok && pf.data.vatMode === 'incl', pf);

  const pdf = w.sb.buildPdfPayload(rows[0]);
  check('D5 PDF 데이터 vatMode incl · totalAmt 50,909 · totalVat 56,000', pdf.vatMode === 'incl' && pdf.totalAmt === 50909 && pdf.totalVat === 56000, [pdf.vatMode, pdf.totalAmt, pdf.totalVat]);
  check('D5 PDF 데이터 별도 건 vatMode excl', w.sb.buildPdfPayload(rows[1]).vatMode === 'excl');

  const m = w.sb._buildPoManifest_(rows[0], null);
  check('D6 주문서 매니페스트 vatMode incl · totalVat 56,000 · totalAmt 50,909', m.vatMode === 'incl' && m.totalVat === 56000 && m.totalAmt === 50909, [m.vatMode, m.totalVat, m.totalAmt]);

  // 이 기능 전에 쓰인 행 — 이전 코드로 제출한 행을 새 코드가 판별
  const b = makeEnv('BASE', readBASE).as(KIM);
  [['TH-AP-26-540', KRW_ITEMS], ['TH-AP-26-541', ODD_ITEMS], ['TH-AP-26-542', USD_ITEMS],
   ['TH-AP-26-543', [{ name: '소액', spec: 'x', qty: 1, price: 4, currency: 'KRW' }]],
   ['TH-AP-26-544', [{ name: '영', spec: 'x', qty: 0, price: 0, currency: 'KRW' }]]].forEach(([no, it]) => b.call('submitRequisitionForClient', reqPayload(no, it)));
  const legacy = b.main().rows.slice(1).map(r => w.sb.detectRowVatMode_(r));
  check('D7 이전 코드가 쓴 행은 모두 excl (원화·경계·외화·소액·0원)', legacy.every(x => x === 'excl'), legacy);
  // 옛 하위호환 포맷 '1,000원' / '/' 구분자
  const old = new Array(C.APPR_START).fill('');
  old[C.ITEMS] = '1. 볼트 M8 / 10개 / 1,000원\n2. 너트 / M8 / 20개 / 500 KRW'; old[C.TOTAL_AMT] = 20000; old[C.TOTAL_VAT] = 22000;
  check('D7 옛 품목 포맷 행도 excl', w.sb.detectRowVatMode_(old) === 'excl');
})();

// ═════════════════════════════════════════════════════════════════
console.log('\n[P] 검수 사진 — 장당 5MB 삭제 · 합계 15MB · 6장');
(function () {
  const MB = 1024 * 1024;
  const photo = (bytes, declared) => ({ name: 'p.jpg', type: 'image/jpeg', size: declared == null ? bytes : declared, data: Buffer.alloc(bytes, 7).toString('base64') });
  const inspPayload = photos => ({
    prcToken: 'tok-none', resubToken: '', receivedDate: '2026-10-30', receivedNote: '2EA', verdict: '합격', isFinal: false, comment: '',
    approvers: [{ label: '기안자', name: NAMES[KIM], email: KIM }, { label: '결재자 1', name: NAMES[LEE], email: LEE }], photos,
  });
  // 사진 검사를 통과하면 락 안의 '대상 구매 문서 없음'에서 멈춘다 — 그 메시지가 곧 '사진 검사 통과'
  const PASSED = '대상 구매 문서를 찾을 수 없습니다.';
  const run = (env, photos) => env.as(KIM).call('submitInspForClient', inspPayload(photos)).message;
  const cases = [
    ['P1 7장 → 장수 초과 (이전과 같음)', Array.from({ length: 7 }, () => photo(1000)), m => /최대 6장/.test(m), m => /최대 6장/.test(m)],
    ['P2 6장 × 1.55MB(9.3MB, 압축본 이론상 최악) → 통과 (이전과 같음)', Array.from({ length: 6 }, () => photo(Math.round(1.55 * MB))), m => m === PASSED, m => m === PASSED],
    ['P3 장당 신고값 20MB · 실제 300KB → 통과 (이전은 장당 5MB로 거부)', [photo(300 * 1024, 20 * MB)], m => m === PASSED, m => /장당 5MB/.test(m)],
    ['P4 6장 × 2.6MB(15.6MB) → 합계 초과 (이전은 장당 5MB 미만이라 통과)', Array.from({ length: 6 }, () => photo(Math.round(2.6 * MB))), m => /합계\(15\.6MB\)가 한도 15MB/.test(m), m => m === PASSED],
    ['P5 정확히 15MB → 통과', [photo(15 * MB)], m => m === PASSED, m => /장당 5MB/.test(m)],
    ['P6 15MB + 1바이트 → 합계 초과', [photo(15 * MB + 1)], m => /한도 15MB/.test(m), m => /장당 5MB/.test(m)],
    ['P7 사진 없음 → 통과 (이전과 같음)', [], m => m === PASSED, m => m === PASSED],
  ];
  cases.forEach(([name, photos, okW, okB]) => {
    const mw = run(makeEnv('WT', readWT), photos), mb = run(makeEnv('BASE', readBASE), photos);
    check(name, okW(mw) && okB(mb), { wt: mw, base: mb });
  });

  const w = makeEnv('WT', readWT);
  const bad = [];
  for (let n = 0; n <= 64; n++) { const s = Buffer.alloc(n, n).toString('base64'); if (w.sb._base64Bytes_(s) !== n) bad.push(n); }
  check('P8 _base64Bytes_ = 실제 디코딩 길이 (0~64바이트 전부, 패딩 0/1/2)', bad.length === 0, bad);

  // 저장: 받은 바이트가 그대로 Drive에 써지는지 (사진 저장 함수는 바뀌지 않았지만 흐름 끝까지 확인)
  const folder = w.sb.__wrapFolder(w.sb.__folderRec('F_X'));
  const src = [photo(123457), photo(2000)];
  w.sb._saveInspPhotos(folder, src);
  const written = w.state.folders.F_X.files.map(f => f.bytes.toString('base64'));
  check('P9 저장 바이트 = 받은 데이터 (2/2)', written.length === 2 && written.every((d, i) => d === src[i].data));
})();

console.log('\n─────────────────────────────');
console.log('통과 ' + pass + ' / 실패 ' + fail + '   (기준 ' + BASE_REF + (WT_DIR !== SRC ? ', 수정본 ' + WT_DIR : '') + ')');
if (fail) { console.log('실패 항목:\n  - ' + failures.join('\n  - ')); process.exitCode = 1; }
