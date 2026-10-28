/**
 * 홈 로딩 성능 개선 — 로컬 검증 (배포 없이 실제 .gs 함수 호출)
 * (memory: gas-local-test-harness)
 *
 * 원칙: 성능 수정은 결과를 바꾸면 안 된다.
 *  같은 시트 픽스처로 이전 코드(BASE_REF)와 작업 트리(수정본)의 getHomeDataForClient()를
 *  역할별(관리자/구매팀/전사조회/일반/부서 미등록 등)로 호출해 응답 페이로드가 완전히 같은지 비교한다.
 *  시트 읽기 횟수도 세어 '같은 결과를 더 적은 읽기로' 냈는지 확인한다.
 *
 * 실행: node tests/home_perf_test.js
 *  BASE_REF=<커밋>  비교 기준(기본 eb04a44 = 성능 개선 직전)
 *  WT_DIR=<폴더>    뮤테이션 검사용 — 일부러 버그를 심은 사본을 '수정본'으로 돌린다
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execSync } = require('child_process');

const SRC = path.join(__dirname, '..');
const WT_DIR = process.env.WT_DIR || SRC;
const BASE_REF = process.env.BASE_REF || 'eb04a44';
const readWT   = f => fs.readFileSync(path.join(WT_DIR, f), 'utf8');
const readBASE = f => execSync('git show ' + BASE_REF + ':"' + f + '"', { cwd: SRC, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });

// ── 고정 시계 ────────────────────────────────────────────────────
const RealDate = Date;
const NOW = RealDate.parse('2026-10-28T10:00:00+09:00');
class FakeDate extends RealDate {
  constructor(...a) { if (a.length === 0) super(NOW); else super(...a); }
  static now() { return NOW; }
}
const MIN = 60 * 1000;

// Utilities.formatDate 대역 — 실제 시간대 DB(Node ICU)로 tz를 적용한다. 고정 +9 스텁이면 기준·수정본이 같은 가정을 공유해
//  날짜 변환 수정을 검증하지 못한다. 토큰은 이 시스템이 쓰는 것만(yyyy MM dd HH mm ss).
const icuFmts = {};
function icuFormat(d, tz, fmt) {
  const f = icuFmts[tz] || (icuFmts[tz] = new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit',
    day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' }));
  const p = Object.fromEntries(f.formatToParts(d).map(x => [x.type, x.value]));
  return fmt.replace(/yyyy|MM|dd|HH|mm|ss/g, k => ({ yyyy: p.year, MM: p.month, dd: p.day, HH: p.hour, mm: p.minute, ss: p.second })[k]);
}

// ── 메모리 시트 (읽기 횟수 계측) ───────────────────────────────────
class FakeRange {
  constructor(sheet, row, col, nRows, nCols) { Object.assign(this, { sheet, row, col, nRows, nCols }); }
  getValues() {
    this.sheet._reads.n++;
    const out = [];
    for (let r = 0; r < this.nRows; r++) {
      const row = [];
      for (let c = 0; c < this.nCols; c++) row.push(this.sheet._get(this.row + r, this.col + c));
      out.push(row);
    }
    return out;
  }
  setValues() { this.sheet._writes.n++; return this; }
  setValue() { this.sheet._writes.n++; return this; }
}
// _meta: 크기 확인·범위 계산 호출(getLastRow / getLastColumn / getDataRange) — 운영에서 회당 0.1~0.2초인 서버 호출
class FakeSheet {
  constructor(name, reads, writes, meta) { this.name = name; this.rows = []; this._reads = reads; this._writes = writes; this._meta = meta; }
  _get(r, c) { const row = this.rows[r - 1]; return row ? (row[c - 1] ?? '') : ''; }
  _lastCol() { return this.rows.reduce((m, r) => Math.max(m, r.length), 0); }
  getName() { return this.name; }
  getLastRow() { this._meta.n++; return this.rows.length; }
  getLastColumn() { this._meta.n++; return this._lastCol(); }
  getMaxRows() { return Math.max(this.rows.length, 1000); }
  getMaxColumns() { return Math.max(this._lastCol(), 26); }
  getRange(row, col, nRows = 1, nCols = 1) { return new FakeRange(this, row, col, nRows, nCols); }
  getDataRange() { this._meta.n++; return this.getRange(1, 1, Math.max(this.rows.length, 1), Math.max(this._lastCol(), 1)); }
  appendRow() { this._writes.n++; return this; }
}
class FakeSpreadsheet {
  constructor(counters) { this.sheets = {}; this.counters = counters; this.opens = 0; }
  getSheetByName(n) { return this.sheets[n] || null; }
  getSheets() { return Object.values(this.sheets); }
  insertSheet(n) { throw new Error('홈 경로가 시트를 만들면 안 된다: ' + n); }
  _add(n, rows) {
    const c = this.counters[n] = { reads: { n: 0 }, writes: { n: 0 }, meta: { n: 0 } };
    const s = new FakeSheet(n, c.reads, c.writes, c.meta);
    s.rows = rows.map(r => r.slice());
    this.sheets[n] = s;
  }
}

// ── 사람 ─────────────────────────────────────────────────────────
const ADMIN  = 'davidkang@inlct.com';     // 관리자 + 구매팀
const BUYER  = 'admin-inlct@inlct.com';   // 구매팀 전용
const GVIEW  = 'bret@inlct.com';          // 전사조회
const KIM    = 'kim@inlct.com';           // 개발팀
const PARK   = 'park@inlct.com';          // 개발팀
const LEE    = 'lee@inlct.com';           // 경영팀
const CHOI   = 'choi@inlct.com';          // 경영팀
const NODEPT = 'nodept@inlct.com';        // 결재자목록 미등록
const BLANK  = 'blank@inlct.com';         // 등록됐으나 부서 빈칸
const NAMES = { [ADMIN]: '강관리', [BUYER]: '구매담당', [GVIEW]: '강전사', [KIM]: '김기안', [PARK]: '박동료',
  [LEE]: '이경영', [CHOI]: '최경영', [NODEPT]: '노부서', [BLANK]: '빈부서' };

// ── 픽스처 ───────────────────────────────────────────────────────
function buildFixtures(ctx) {
  const { COL, INSP_COL, PO_TOTAL_COLS, PO_COL } = ctx;

  // 결재자목록: [이름, 부서, 직급, 이메일] — 사람이 직접 입력하므로 대소문자·공백이 섞인다
  const approvers = [
    ['이름', '부서', '직급', '이메일'],
    ['김기안', '개발팀', '대리', ' KIM@inlct.com '],
    ['박동료', '개발팀', '사원', PARK],
    ['이경영', '경영팀', '과장', LEE],
    ['최경영', '경영팀', '차장', 'Choi@INLCT.com'],
    ['빈부서', '', '사원', BLANK],
    ['강관리', '경영총괄', '매니저', ADMIN],
    ['강전사', '경영총괄팀', '부장', GVIEW],
  ];

  // 품의서목록 — 운영 시트처럼 날짜 칸은 Date(시트 시간대 Asia/Seoul)로 둔다. raw:true 행만 문자열 그대로(문자열 경로 검사).
  //  sandbox의 Date(FakeDate)로 만들어야 코드의 `instanceof Date`가 참이 된다.
  const kst = s => (typeof s === 'string' ? new FakeDate(s.slice(0, 10) + 'T' + (s.slice(11, 16) || '00:00') + ':00+09:00') : s);
  const REQ_LEN = COL.APPR_START + 4 * COL.APPR_COLS;
  function req(o) {
    const r = new Array(REQ_LEN).fill('');
    const asCell = v => (o.raw ? v : kst(v));
    r[COL.SUBMIT_AT] = asCell(o.submitAt); r[COL.DOC_NO] = o.docNo;
    r[COL.ISSUE_DATE] = o.issueDate !== undefined ? o.issueDate : asCell(typeof o.submitAt === 'string' ? o.submitAt.slice(0, 10) : o.submitAt);
    r[COL.DRAFTER] = NAMES[o.drafter] || o.drafter; r[COL.DEPT] = o.dept || ''; r[COL.SUBJECT] = o.subject || o.docNo + ' 품명';
    r[COL.VENDOR_NAME] = o.vendor || '업체' + o.docNo.slice(-1);
    r[COL.ITEMS] = o.items || '1. 볼트 | M6 | 10 EA | 1,000 KRW';
    r[COL.TOTAL_AMT] = o.amt || 10000;
    r[COL.STATUS] = o.status; r[COL.TOKEN] = o.token;
    r[COL.APPR_COUNT] = o.appr.length; r[COL.APPR_IDX] = o.curIdx ?? o.appr.length;
    r[COL.ATTACH_LIST] = o.attach || '';
    r[COL.DOC_TYPE] = o.type || 'REQ'; r[COL.PARENT_DOC_ID] = o.parent || '';
    r[COL.CLAIMED_BY] = o.claimedBy || ''; r[COL.CLAIMED_AT] = o.claimedAt ? asCell(o.claimedAt) : '';
    o.appr.forEach((a, i) => {
      const base = COL.APPR_START + i * COL.APPR_COLS;
      r[base] = i === 0 ? '기안' : '결재' + i; r[base + 1] = NAMES[a.email] || a.email; r[base + 2] = a.email; r[base + 3] = a.status;
    });
    return r;
  }
  const quoteAttach = (state, queuedAgoMin) => JSON.stringify([{ role: 'quote-src', name: '견적.xlsx', quotePdf: state,
    quoteQueuedAt: queuedAgoMin == null ? undefined : new RealDate(NOW - queuedAgoMin * MIN).toISOString() }]);
  const ok = e => ({ email: e, status: '승인' }), wait = e => ({ email: e, status: '대기' }), none = e => ({ email: e, status: '' });
  const reqRows = [
    new Array(REQ_LEN).fill('헤더'),
    req({ docNo: 'TH-AP-26-001', token: 'tR1', drafter: KIM, dept: '개발팀', status: '최종승인', submitAt: '2026-10-01 09:00', appr: [ok(KIM), ok(LEE)] }),
    req({ docNo: 'TH-AP-26-002', token: 'tR2', drafter: PARK, dept: '개발팀', status: 'PRC생성됨', submitAt: '2026-10-02 09:00', appr: [ok(PARK), ok(CHOI)], raw: true }),
    req({ docNo: 'TH-PO-26-002', token: 'tP2', type: 'PRC', parent: 'tR2', drafter: BUYER, status: '최종승인(PRC)', submitAt: '2026-10-03 09:00', appr: [ok(BUYER), ok(ADMIN)] }),
    req({ docNo: 'TH-PO-26-001', token: 'tP1', type: 'PRC', parent: 'tR1', drafter: ADMIN, status: '최종승인(PRC)', submitAt: '2026-10-04 09:00', appr: [ok(ADMIN), ok(CHOI)] }),
    // 날짜 경계: KST 23:59:59.999 / 다음날 00:00 / 1989년 이전(시간만 든 칸은 1899-12-30) / 잘못된 Date
    req({ docNo: 'TH-AP-26-003', token: 'tR3', drafter: LEE, dept: '경영팀', status: '결재중 (김기안)', submitAt: new FakeDate(RealDate.UTC(2026, 9, 5, 14, 59, 59, 999)), appr: [ok(LEE), wait(KIM)], curIdx: 1 }),
    req({ docNo: 'TH-AP-26-004', token: 'tR4', drafter: CHOI, dept: '경영팀', status: '검토중', submitAt: new FakeDate(RealDate.UTC(2026, 9, 5, 15, 0, 0)), appr: [wait(CHOI), none(LEE)], curIdx: 0 }),
    req({ docNo: 'TH-AP-26-005', token: 'tR5', drafter: NODEPT, dept: '', status: '최종승인', submitAt: '2026-10-07 09:00', issueDate: new FakeDate(RealDate.UTC(1899, 11, 29, 15, 30)), appr: [ok(NODEPT), ok(LEE)], items: '1. 센서 | X | 1 EA | 120 USD', amt: 120, claimedAt: new FakeDate(NaN) }),
    req({ docNo: 'TH-AP-26-006', token: 'tR6', drafter: KIM, dept: '개발팀', status: '폐기', submitAt: '2026-10-08 09:00', appr: [ok(KIM), none(LEE)] }),
    req({ docNo: 'TH-AP-26-007', token: 'tR7', drafter: KIM, dept: '개발팀', status: '반려', submitAt: '2026-10-09 09:00', appr: [ok(KIM), { email: LEE, status: '반려' }], curIdx: 1 }),
    req({ docNo: 'TH-AP-26-008', token: 'tR8', drafter: PARK, dept: '개발팀', status: '최종승인', submitAt: '2026-10-10 09:00', appr: [ok(PARK), ok(LEE)], claimedBy: BUYER, claimedAt: '2026-10-11 09:00' }),
    req({ docNo: 'TH-PO-26-008', token: 'tP8', type: 'PRC', parent: 'tR8', drafter: BUYER, status: '결재중 (강관리)', submitAt: '2026-10-11 10:00', appr: [ok(BUYER), wait(ADMIN)], curIdx: 1 }),
    req({ docNo: 'TH-AP-26-009', token: 'tR9', drafter: BLANK, dept: '', status: '재상신', submitAt: '2026-10-12 09:00', appr: [ok(BLANK), wait('LEE@inlct.com')], curIdx: 1 }),
    // 관리자 배너(견적서) — 완료 / 갓 QUEUED / 10분 넘은 QUEUED(정체) / 테스트 행
    req({ docNo: 'TH-AP-26-010', token: 'tR10', drafter: KIM, dept: '개발팀', status: '최종승인', submitAt: '2026-10-13 09:00', appr: [ok(KIM), ok(LEE)], attach: quoteAttach('DONE') }),
    req({ docNo: 'TH-AP-26-011', token: 'tR11', drafter: LEE, dept: '경영팀', status: '검토중', submitAt: '2026-10-14 09:00', appr: [wait(LEE), none(CHOI)], curIdx: 0, attach: quoteAttach('QUEUED', 5) }),
    req({ docNo: 'TH-AP-26-012', token: 'tR12', drafter: CHOI, dept: '경영팀', status: '최종승인', submitAt: '2026-10-15 09:00', appr: [ok(CHOI), ok(LEE)], attach: quoteAttach('QUEUED', 12) }),
    req({ docNo: 'PRQ-2026-001-01', token: 'tTest', drafter: ADMIN, dept: '경영총괄', status: '최종승인', submitAt: '2026-10-16 09:00', appr: [ok(ADMIN)], attach: quoteAttach('QUEUED', 12) }),
  ];

  // 검수보고서목록 — 일반 사용자 N+1이 드러나도록 최종승인(INSP) 행을 넉넉히 둔다
  const INSP_LEN = INSP_COL.APPR_START + INSP_COL.MAX_APPROVERS * INSP_COL.APPR_COLS;
  let seqNo = 0;
  function insp(o) {
    const r = new Array(INSP_LEN).fill('');
    seqNo++;
    const genAt = '2026-10-' + String(10 + (seqNo % 18)).padStart(2, '0') + ' ' + String(seqNo % 24).padStart(2, '0') + ':00';
    r[INSP_COL.SUBMIT_AT] = o.submitAt || (seqNo % 5 === 0 ? genAt : kst(genAt));   // 5건 중 1건은 문자열
    r[INSP_COL.DOC_NO] = o.docNo || 'TH-AP-26-1' + String(seqNo).padStart(2, '0') + '-01';
    r[INSP_COL.SEQ] = o.seq || 1; r[INSP_COL.PRC_TOKEN] = o.prc || ''; r[INSP_COL.REQ_TOKEN] = o.reqTok || '';
    r[INSP_COL.PO_NO] = 'PO-' + seqNo; r[INSP_COL.REQ_NO] = 'REQ-' + seqNo; r[INSP_COL.SUBJECT] = '검수품 ' + seqNo; r[INSP_COL.VENDOR_NAME] = '업체';
    r[INSP_COL.DRAFTER] = o.drafter; r[INSP_COL.DRAFTER_NAME] = NAMES[String(o.drafter).trim().toLowerCase()] || '';
    r[INSP_COL.RECEIVED_DATE] = seqNo % 4 === 0 ? '2026-10-01' : kst('2026-10-0' + (1 + seqNo % 9)); r[INSP_COL.VERDICT] = '합격'; r[INSP_COL.IS_FINAL] = o.isFinal === false ? '' : 'Y';
    r[INSP_COL.STATUS] = o.status; r[INSP_COL.TOKEN] = o.token === undefined ? 'tI' + seqNo : o.token;
    r[INSP_COL.APPR_COUNT] = o.apprCount ?? o.appr.length; r[INSP_COL.APPR_IDX] = o.curIdx ?? o.appr.length;
    r[INSP_COL.PHOTO_LIST] = o.photos || '[]'; r[INSP_COL.MOVE_STATUS] = o.move || 'FINAL';
    o.appr.forEach((a, i) => {
      const base = INSP_COL.APPR_START + i * INSP_COL.APPR_COLS;
      r[base + 1] = NAMES[a.email] || a.email; r[base + 2] = a.email; r[base + 3] = a.status;
    });
    return r;
  }
  const inspRows = [new Array(INSP_LEN).fill('헤더')];
  // 다양한 기안자의 최종승인(INSP) 40건 — 대소문자·공백 섞인 이메일, 미등록, 빈칸 포함
  const drafters = [KIM, PARK, LEE, CHOI, NODEPT, ' Park@INLCT.com ', BLANK, 'Choi@inlct.com', 'ghost@inlct.com', ''];
  for (let i = 0; i < 40; i++) {
    const d = drafters[i % drafters.length];
    const approver = (i % 7 === 0) ? LEE : CHOI;   // 일부는 LEE가 결재 참여자
    inspRows.push(insp({ drafter: d, status: '최종승인(INSP)', isFinal: i % 3 !== 0, prc: i % 2 ? 'tP1' : 'tP2',
      appr: [ok(String(d).trim() || KIM), ok(approver)], move: i % 9 === 0 ? 'PENDING_PDF' : 'FINAL',
      photos: i % 9 === 0 ? '[{"id":"a"},{"id":"b"}]' : '[]' }));
  }
  inspRows.push(
    insp({ drafter: PARK, status: '결재중 (김기안)', prc: 'tP1', appr: [ok(PARK), wait(KIM)], curIdx: 1, isFinal: true }),
    insp({ drafter: KIM, status: '검토중', prc: 'tP2', appr: [wait(KIM), none(LEE)], curIdx: 0, isFinal: false }),
    insp({ drafter: LEE, status: '결재중 (최경영)', prc: 'tP2', appr: [ok(LEE), wait('CHOI@inlct.com')], curIdx: 1 }),
    insp({ drafter: KIM, status: '반려', prc: 'tP1', appr: [ok(KIM), { email: LEE, status: '반려' }], curIdx: 1 }),
    insp({ drafter: PARK, status: '폐기', prc: 'tP1', appr: [ok(PARK), none(LEE)] }),
    insp({ drafter: LEE, status: '최종승인(INSP)', prc: 'tP2', appr: [ok(LEE), ok(CHOI)], apprCount: 99 }),   // 깨진 apprCount — 클램프 경로
    insp({ drafter: KIM, status: '최종승인(INSP)', prc: 'tP1', appr: [ok(KIM), ok(LEE)], token: '' }),      // 토큰 없음 — 메뉴에서 제외
    insp({ drafter: NODEPT, status: '최종승인(INSP)', prc: 'tP2', appr: [ok(KIM), ok(CHOI)] }),             // 결재란 0번 ≠ 기안자 — '본인 기안' 조건만으로 보이는 행
    insp({ docNo: 'PRQ-2026-001-01', drafter: ADMIN, status: '최종승인(INSP)', prc: 'tP1', appr: [ok(ADMIN)], move: 'PENDING_PDF' }),
  );

  // 주문서목록
  const poRow = (poNo, status) => { const r = new Array(PO_TOTAL_COLS).fill(''); r[PO_COL.CREATED_AT] = '2026-10-20 09:00'; r[PO_COL.PO_NO] = poNo;
    r[PO_COL.PRC_TOKEN] = 'tok-' + poNo; r[PO_COL.CURRENCY] = 'KRW'; r[PO_COL.TOTAL_AMT] = 5000; r[PO_COL.STATUS] = status; return r; };
  const poHeader = new Array(PO_TOTAL_COLS).fill(''); poHeader[0] = 'createdAt'; poHeader[1] = 'poNo'; poHeader[2] = 'prcToken';
  const poRows = [poHeader, poRow('TH-PO-26-001', '생성대기'), poRow('TH-PO-26-002', '생성완료'), poRow('TH-PO-26-003', '생성불요'), poRow('TH-PO-26-004', '')];

  return { approvers, reqRows, inspRows, poRows };
}

// ── 환경 ─────────────────────────────────────────────────────────
// extraFiles: Code/INSP/PO 외에 함께 올릴 .gs (예: 진단 함수 파일) — 같은 read로 읽는다
function makeEnv(label, read, active, prefsByEmail, extraFiles = []) {
  const counters = {};
  const svc = { fmt: 0, tz: 0 };   // 서버 호출로 비싼 날짜 함수 호출 수 (Utilities.formatDate / Session.getScriptTimeZone)
  const errors = [];
  const ss = new FakeSpreadsheet(counters);
  let opens = 0;
  const sandbox = {
    console: { log() {}, info() {}, warn() {}, error: (...a) => errors.push(a.map(String).join(' ')) },
    Date: FakeDate,
    Logger: { log: (...a) => { const s = a.map(String).join(' '); if (/실패/.test(s)) errors.push(s); } },
    Session: {
      getActiveUser: () => ({ getEmail: () => active }),
      getEffectiveUser: () => ({ getEmail: () => ADMIN }),
      getScriptTimeZone: () => { svc.tz++; return 'Asia/Seoul'; },
    },
    SpreadsheetApp: { openById: () => { opens++; return ss; }, flush() {} },
    DriveApp: { getFolderById: () => { throw new Error('홈 경로에서 Drive 호출 금지'); }, getFileById: () => { throw new Error('홈 경로에서 Drive 호출 금지'); } },
    Utilities: {
      getUuid: () => 'uuid',
      formatDate: (d, tz, fmt) => { svc.fmt++; return icuFormat(d, tz, fmt); },
    },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (prefsByEmail[k] || null) }) },
    CacheService: { getScriptCache: () => ({ get: () => null, put() {}, remove() {} }) },
    LockService: { getScriptLock: () => { throw new Error('홈 경로에 락 금지'); }, getUserLock: () => { throw new Error('홈 경로에 락 금지'); } },
    MimeType: { PDF: 'application/pdf' },
  };
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  ['Code.gs', 'INSP.gs', 'PO.gs', ...extraFiles].forEach(f => vm.runInContext(read(f), sandbox, { filename: label + ':' + f }));
  // 홈 경로에서 쓰기·메일이 일어나면 안 된다 — 큐·감사로그·메일은 기록만 하고 실패로 센다
  sandbox.__errors = errors;
  vm.runInContext(`
    writeAuditLog = function (o) { __errors.push('감사로그 기록 발생: ' + JSON.stringify(o)); };
    enqueueJob    = function (o) { __errors.push('큐 등록 발생: ' + JSON.stringify(o)); };
  `, sandbox);

  const ctx = vm.runInContext('({ COL: COL, INSP_COL: INSP_COL, PO_COL: PO_COL, PO_TOTAL_COLS: PO_TOTAL_COLS, CONFIG: CONFIG })', sandbox);
  const fx = buildFixtures(ctx);
  ss._add(ctx.CONFIG.SHEET_NAME, fx.reqRows);
  ss._add('결재자목록', fx.approvers);
  ss._add(ctx.CONFIG.INSP_SHEET_NAME, fx.inspRows);
  ss._add(ctx.CONFIG.PO_SHEET_NAME, fx.poRows);

  return {
    sandbox, counters, errors, svc, ss,
    // 검수보고서 작성·제출 관문(INSP.gs)은 deptMap 없이 _isSameDeptEmail을 부른다 — 그 경로도 그대로여야 한다
    sameDept(a, b, map) {
      sandbox.__args = [a, b, map];
      return vm.runInContext('_isSameDeptEmail(SpreadsheetApp.openById(CONFIG.SHEET_ID), __args[0], __args[1], __args[2] === "BUILD" ? getDeptByEmailMap(SpreadsheetApp.openById(CONFIG.SHEET_ID)) : undefined)', sandbox);
    },
    run() {
      const res = vm.runInContext('getHomeDataForClient()', sandbox);
      const reads = {};
      Object.keys(counters).forEach(k => { reads[k] = counters[k].reads.n; });
      const writes = Object.keys(counters).reduce((s, k) => s + counters[k].writes.n, 0);
      const svcCalls = { fmt: svc.fmt, tz: svc.tz };
      svc.fmt = 0; svc.tz = 0;
      // 시트 호출 = 값 읽기(getValues) + 크기 확인·범위 계산(getLastRow/getLastColumn/getDataRange)
      const sheetCalls = Object.keys(counters).reduce((s, k) => s + counters[k].reads.n + counters[k].meta.n, 0);
      return { res: JSON.parse(JSON.stringify(res)), reads, writes, opens, errors, svc: svcCalls, sheetCalls };
    },
  };
}

// ── 비교 도구 ─────────────────────────────────────────────────────
function firstDiff(a, b, p = '') {
  if (a === b) return null;
  if (typeof a !== typeof b || a === null || b === null || typeof a !== 'object') return `${p || '(root)'}: ${JSON.stringify(a)} ≠ ${JSON.stringify(b)}`;
  if (Array.isArray(a) !== Array.isArray(b)) return `${p}: 배열/객체 불일치`;
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  for (const k of keys) { const d = firstDiff(a[k], b[k], p + (Array.isArray(a) ? `[${k}]` : `.${k}`)); if (d) return d; }
  return null;
}

let pass = 0, fail = 0;
function check(name, cond, detail) {
  if (cond) { pass++; console.log('  ✓ ' + name); }
  else { fail++; console.log('  ✗ ' + name + (detail ? '\n      ' + detail : '')); }
}

// 다른 테스트(진단 함수·후속 수정 검증)가 같은 환경·픽스처를 쓰도록 공개한다. require로 불리면 시나리오는 돌리지 않는다.
module.exports = { makeEnv, firstDiff, readWT, readBASE, icuFormat, FakeDate, RealDate, ADMIN, BUYER, GVIEW, KIM, PARK, LEE, NODEPT };
if (require.main !== module) return;

// ── 시나리오 ─────────────────────────────────────────────────────
const PREFS = { ['UPREF_' + KIM]: JSON.stringify({ v: 1, cols: { a: 1 } }) };
const ROLES = [
  { name: '관리자(관리자+구매팀)', email: ADMIN, general: false },
  { name: '구매팀',               email: BUYER, general: false },
  { name: '전사조회',             email: GVIEW, general: false },
  { name: '일반 · 개발팀 김기안',   email: KIM,   general: true },
  { name: '일반 · 개발팀 박동료',   email: PARK,  general: true },
  { name: '일반 · 경영팀 이경영',   email: LEE,   general: true },
  { name: '일반 · 부서 미등록',     email: NODEPT, general: true },
  { name: '일반 · 부서 빈칸',       email: BLANK, general: true },
];

console.log(`홈 데이터 페이로드 동일성 — 기준 ${BASE_REF} vs 작업 트리${WT_DIR !== SRC ? ' (' + WT_DIR + ')' : ''}\n`);
const table = [];
const results = {};
for (const role of ROLES) {
  console.log(`[${role.name}] ${role.email}`);
  const base = makeEnv('BASE', readBASE, role.email, PREFS).run();
  const wt   = makeEnv('WT',   readWT,   role.email, PREFS).run();
  results[role.email] = wt.res;

  check('기준·수정본 모두 ok=true', base.res.ok === true && wt.res.ok === true, base.res.message || wt.res.message);
  const d = firstDiff(base.res, wt.res);
  check('응답 페이로드 완전 동일', d === null, d);
  check('오류 로그·홈 경로 쓰기 없음 (수정본)', wt.errors.length === 0 && wt.writes === 0, wt.errors.join(' / ') + (wt.writes ? ` 쓰기 ${wt.writes}회` : ''));

  const apprReadsBase = base.reads['결재자목록'], apprReadsWt = wt.reads['결재자목록'];
  if (role.general) check('결재자목록 읽기 1회 (수정본)', apprReadsWt === 1, `수정본 ${apprReadsWt}회`);
  else check('결재자목록 읽기 늘지 않음', apprReadsWt <= apprReadsBase, `기준 ${apprReadsBase} → 수정본 ${apprReadsWt}`);

  const sum = r => Object.values(r).reduce((s, n) => s + n, 0);
  table.push({ role: role.name, base: base.reads, wt: wt.reads, baseSum: sum(base.reads), wtSum: sum(wt.reads), opensBase: base.opens, opensWt: wt.opens,
    svcBase: base.svc, svcWt: wt.svc, callsBase: base.sheetCalls, callsWt: wt.sheetCalls });
  console.log('');
}

// 시나리오가 '빈 결과끼리 같다'로 통과하지 않도록 가시성 자체를 확인한다
console.log('[가시성 확인 — 시나리오가 실제로 부서 공유 경로를 지나는가]');
const inspFinalsOf = email => results[email].finalDocs.filter(x => x.docType === 'INSP');
const drafterLc = x => String(x.drafterName);
const kim = inspFinalsOf(KIM), park = inspFinalsOf(PARK), lee = inspFinalsOf(LEE), nodept = inspFinalsOf(NODEPT), adminList = inspFinalsOf(ADMIN);
check('김기안: 같은 부서 박동료의 검수보고서가 보인다', kim.some(x => drafterLc(x) === '박동료'));
check('김기안: 다른 부서(경영팀) 검수보고서는 결재 참여가 아니면 안 보인다', !kim.some(x => drafterLc(x) === '최경영'));
check('박동료: 같은 부서 김기안의 검수보고서가 보인다(시트 이메일 대소문자·공백)', park.some(x => drafterLc(x) === '김기안'));
check('이경영: 같은 부서 최경영(대소문자 섞인 이메일) 건이 보인다', lee.some(x => drafterLc(x) === '최경영'));
check('부서 미등록: 본인 건만 보인다(결재란 0번이 본인이 아닌 기안 포함)', nodept.length === 5 && nodept.every(x => drafterLc(x) === '노부서'), `${nodept.length}건`);
check('관리자: 토큰 있는 최종승인(INSP) 전부가 보인다', adminList.length === 43, `${adminList.length}건`);
check('관리자 배너: 주문서 2건·검수 PDF 대기·견적서 정체 1건', (() => {
  const t = results[ADMIN].adminTasks;
  return t && t.po.length === 2 && t.insp.length > 0 && t.insp.every(x => x.docNo !== 'PRQ-2026-001-01')
    && t.quotePdf.length === 1 && t.quotePdf[0].docNo === 'TH-AP-26-012';
})(), JSON.stringify(results[ADMIN].adminTasks && { po: results[ADMIN].adminTasks.po.length, insp: results[ADMIN].adminTasks.insp.length, quote: results[ADMIN].adminTasks.quotePdf.map(x => x.docNo) }));
check('일반 사용자에게 adminTasks 없음', results[KIM].adminTasks === null);
check('설정(prefs) 편승 유지', results[KIM].prefs && results[KIM].prefs.v === 1);

console.log('\n[검수보고서 작성·제출 관문 — _isSameDeptEmail 단독 호출(부서 맵 없이)]');
{
  const PAIRS = [[KIM, PARK], [' Park@INLCT.com ', KIM], [KIM, ' Park@INLCT.com '], [KIM, LEE], [LEE, 'Choi@inlct.com'], [NODEPT, KIM], [BLANK, BLANK],
    ['ghost@inlct.com', KIM], ['', KIM], [KIM, '']];
  const base = makeEnv('BASE', readBASE, KIM, {}), wt = makeEnv('WT', readWT, KIM, {});
  const b = PAIRS.map(p => base.sameDept(p[0], p[1])), w = PAIRS.map(p => wt.sameDept(p[0], p[1]));
  const wm = PAIRS.map(p => wt.sameDept(p[0], p[1], 'BUILD'));
  check('맵 없이 호출: 기준과 같은 판정 ' + JSON.stringify(w), JSON.stringify(b) === JSON.stringify(w), '기준 ' + JSON.stringify(b));
  check('맵을 넘겨 호출: 맵 없이 호출한 것과 같은 판정', JSON.stringify(w) === JSON.stringify(wm), '맵 전달 ' + JSON.stringify(wm));
  check('판정이 참·거짓을 모두 포함(무의미 통과 방지)', w.includes(true) && w.includes(false));
}

console.log('\n[날짜 변환 — toDateStr 기준 vs 수정본 (Utilities.formatDate 대역은 시간대 DB 기준)]');
{
  const base = makeEnv('BASE', readBASE, KIM, {}), wt = makeEnv('WT', readWT, KIM, {});
  const call = (env, v, f) => { env.sandbox.__a = [v, f]; return vm.runInContext('toDateStr(__a[0], __a[1])', env.sandbox); };
  const D = ms => new FakeDate(ms);
  const inputs = [];
  // 1900~2100년 매년: 월초·2월 말·연말의 KST 자정 직전/직후, 한낮
  for (let y = 1900; y <= 2100; y++) {
    for (const [m, d, h, mi] of [[0, 1, 0, 0], [1, 28, 14, 59], [1, 28, 15, 0], [11, 31, 14, 59], [11, 31, 15, 0], [5, 15, 3, 30]]) inputs.push(D(RealDate.UTC(y, m, d, h, mi, 59, 999)));
  }
  let seed = 11; const rnd = () => (seed = (seed * 48271) % 2147483647) / 2147483647;
  for (let i = 0; i < 20000; i++) inputs.push(D(RealDate.UTC(1950, 0, 1) + rnd() * (RealDate.UTC(2100, 0, 1) - RealDate.UTC(1950, 0, 1))));
  inputs.push(D(NaN), D(RealDate.UTC(1899, 11, 29, 15, 30)), D(RealDate.UTC(1988, 6, 1)), D(RealDate.UTC(9998, 11, 31, 20, 0)), D(RealDate.UTC(9999, 5, 1)));
  const nDates = inputs.length;
  inputs.push('2026-10-28', '2026-10-28 10:00', ' 2026-10-28 ', 'Oct 28, 2026 10:00:00', 'October 1, 2026', '2026/10/28', 'garbage', ' ', '', null, undefined, 0, 45000, true);
  const fmts = [undefined, 'yyyy-MM-dd', 'yyyy-MM-dd HH:mm', 'yyyy.MM.dd', 'yyyyMMdd_HHmmss'];
  let n = 0; const diffs = [];
  for (const v of inputs) for (const f of fmts) {
    n++;
    const b = call(base, v, f), w = call(wt, v, f);
    if (b !== w) diffs.push(`${v instanceof RealDate ? v.toISOString() : JSON.stringify(v)} [${f}] 기준 ${b} / 수정본 ${w}`);
  }
  check(`기준과 결과 동일 — ${n}건 (Date ${nDates}개 × 형식 5종 + 문자열·빈값·숫자)`, diffs.length === 0, diffs.slice(0, 3).join(' | '));
  // 수정본이 Utilities.formatDate로 넘겨야 하는 건수를 사양대로 따로 센다:
  //  Date로 바뀌는 값 중 (형식이 홈의 두 형식이 아니거나) (UTC 연도가 1989~9998 밖)인 것만
  const FAST = [undefined, 'yyyy-MM-dd', 'yyyy-MM-dd HH:mm'];
  const asDate = v => {
    if (v instanceof RealDate) return isNaN(v.getTime()) ? null : v;
    if (v === null || v === undefined || v === '') return null;
    const s = String(v).trim();
    if (!s || /^\d{4}-\d{2}-\d{2}/.test(s)) return null;
    const d = new RealDate(s);
    return isNaN(d.getTime()) ? null : d;
  };
  let expectFallback = 0;
  for (const v of inputs) for (const f of fmts) {
    const d = asDate(v);
    if (d && !(FAST.includes(f) && d.getUTCFullYear() >= 1989 && d.getUTCFullYear() <= 9998)) expectFallback++;
  }
  check(`서버 호출: formatDate ${base.svc.fmt}→${wt.svc.fmt}회(사양상 기존 방식 ${expectFallback}회), 시간대 조회 ${base.svc.tz}→${wt.svc.tz}회`,
    wt.svc.tz === 1 && wt.svc.fmt === expectFallback);

  // 스크립트 시간대가 바뀌면(appsscript.json timeZone) 순수 계산을 쓰지 않고 기존 방식과 같아야 한다
  const base2 = makeEnv('BASE', readBASE, KIM, {}), wt2 = makeEnv('WT', readWT, KIM, {});
  [base2, wt2].forEach(e => { e.sandbox.Session.getScriptTimeZone = () => 'America/New_York'; });
  const sample = inputs.filter((v, i) => v instanceof RealDate && i % 50 === 0);
  const tzDiffs = [];
  for (const v of sample) for (const f of fmts) { const b = call(base2, v, f), w = call(wt2, v, f); if (b !== w) tzDiffs.push(`${v.toISOString()} [${f}] ${b} / ${w}`); }
  check(`스크립트 시간대가 서울이 아니면 기존 방식과 동일 (${sample.length * fmts.length}건)`, tzDiffs.length === 0, tzDiffs.slice(0, 3).join(' | '));
}
check('홈 1회 날짜 서버 호출: 모든 역할에서 시간대 조회 1회 이하·formatDate는 범위 밖 날짜만',
  table.every(t => t.svcWt.tz <= 1 && t.svcWt.fmt <= 2) && table.every(t => t.svcBase.fmt > 50),
  table.map(t => `${t.role} fmt ${t.svcBase.fmt}→${t.svcWt.fmt} tz ${t.svcBase.tz}→${t.svcWt.tz}`).join(' / '));

console.log('\n[시트가 없거나 비었거나 읽기가 실패해도 기준과 같은 응답]');
{
  const INSP = '검수보고서목록', MAIN = '품의서목록';
  const scenarios = [
    ['검수보고서목록 읽기 실패', s => { s.sheets[INSP].getDataRange = () => { throw new Error('읽기 장애 시뮬레이션'); }; }],
    ['검수보고서목록 시트 없음', s => { delete s.sheets[INSP]; }],
    ['검수보고서목록 머리글만', s => { s.sheets[INSP].rows.length = 1; }],
    ['검수보고서목록 완전히 빔', s => { s.sheets[INSP].rows = []; }],
    ['품의서목록 머리글만', s => { s.sheets[MAIN].rows.length = 1; }],
    ['품의서목록 완전히 빔', s => { s.sheets[MAIN].rows = []; }],
  ];
  for (const [name, mutate] of scenarios) for (const who of [ADMIN, KIM]) {
    const b = makeEnv('BASE', readBASE, who, PREFS), w = makeEnv('WT', readWT, who, PREFS);
    mutate(b.ss); mutate(w.ss);
    const rb = b.run().res, rw = w.run().res;
    const d = firstDiff(rb, rw);
    check(`${name} (${who === ADMIN ? '관리자' : '일반'}): 기준과 같은 응답`, rb.ok === true && rw.ok === true && d === null, d || (rb.message + ' / ' + rw.message));
  }
}

check('검수보고서목록은 모든 역할에서 홈 1회에 한 번만 읽는다', table.every(t => t.wt['검수보고서목록'] === 1),
  table.map(t => `${t.role} ${t.base['검수보고서목록']}→${t.wt['검수보고서목록']}`).join(' / '));
check('시트 호출 수가 모든 역할에서 줄었다', table.every(t => t.callsWt < t.callsBase),
  table.map(t => `${t.role} ${t.callsBase}→${t.callsWt}`).join(' / '));

console.log('\n[시트 읽기 횟수 — getValues 호출 수 (기준 → 수정본)]');
const sheetNames = ['품의서목록', '결재자목록', '검수보고서목록', '주문서목록'];
console.log('  (시트 호출 = 값 읽기 + 크기 확인·범위 계산. 운영에서 회당 0.1~0.3초)');
console.log('  ' + ['역할'.padEnd(18), ...sheetNames.map(s => s.padStart(10)), '합계'.padStart(9), '시트 호출'.padStart(8), 'openById'.padStart(9), 'formatDate'.padStart(11), '시간대'.padStart(7)].join(' '));
table.forEach(t => {
  const cell = s => `${t.base[s] || 0}→${t.wt[s] || 0}`.padStart(10);
  console.log('  ' + [t.role.padEnd(18), ...sheetNames.map(cell), `${t.baseSum}→${t.wtSum}`.padStart(9), `${t.callsBase}→${t.callsWt}`.padStart(10), `${t.opensBase}→${t.opensWt}`.padStart(9),
    `${t.svcBase.fmt}→${t.svcWt.fmt}`.padStart(11), `${t.svcBase.tz}→${t.svcWt.tz}`.padStart(7)].join(' '));
});

console.log(`\n결과: ${pass} 통과 / ${fail} 실패`);
process.exit(fail ? 1 : 0);
