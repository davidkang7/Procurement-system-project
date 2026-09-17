/**
 * 주문서목록 시트 로컬 검증 — 배포 없이 Code.gs + PO.gs 를 node vm 에 로드해 실제 함수를 호출한다.
 * (memory: gas-local-test-harness)
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC = path.join(__dirname, '..');

// ── 메모리 시트 ────────────────────────────────────────────────
class FakeRange {
  constructor(sheet, row, col, nRows, nCols) {
    Object.assign(this, { sheet, row, col, nRows, nCols });
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.nRows; r++) {
      const row = [];
      for (let c = 0; c < this.nCols; c++) row.push(this.sheet._get(this.row + r, this.col + c));
      out.push(row);
    }
    return out;
  }
  setValues(vals) {
    vals.forEach((rowVals, r) =>
      rowVals.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v)));
    return this;
  }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  setFontWeight() { return this; }
  setBackground() { return this; }
  setDataValidation(rule) { this.sheet.validation = { range: [this.row, this.col, this.nRows], rule }; return this; }
  createTextFinder(text) {
    const self = this;
    const opts = { entire: false, matchCase: false };
    const finder = {
      matchEntireCell(b) { opts.entire = b; return finder; },
      matchCase(b) { opts.matchCase = b; return finder; },
      findNext() {
        for (let r = 0; r < self.nRows; r++) {
          for (let c = 0; c < self.nCols; c++) {
            const v = String(self.sheet._get(self.row + r, self.col + c) ?? '');
            const hit = opts.entire ? v === String(text) : v.indexOf(String(text)) >= 0;
            if (hit) return new FakeRange(self.sheet, self.row + r, self.col + c, 1, 1);
          }
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
  _set(r, c, v) {
    while (this.rows.length < r) this.rows.push([]);
    const row = this.rows[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v;
  }
  getName() { return this.name; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return this.rows.reduce((m, r) => Math.max(m, r.length), 0); }
  getMaxRows() { return Math.max(this.rows.length, 1000); }
  appendRow(vals) { this.rows.push(vals.slice()); return this; }
  getRange(row, col, nRows = 1, nCols = 1) { return new FakeRange(this, row, col, nRows, nCols); }
  setFrozenRows() { return this; }
  setColumnWidth() { return this; }
  createTextFinder(text) { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)).createTextFinder(text); }
}

class FakeSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(n) { return this.sheets[n] || null; }
  insertSheet(n) { this.sheets[n] = new FakeSheet(n); return this.sheets[n]; }
  getSheets() { return Object.values(this.sheets); }
}

// ── 드라이브 스텁 ──────────────────────────────────────────────
const driveFiles = {};       // folderId → [{id,name,trashed}]
let fileSeq = 0;
function makeFolder(id) {
  driveFiles[id] = driveFiles[id] || [];
  const list = () => driveFiles[id].filter(f => !f.trashed);
  const iter = arr => { let i = 0; return { hasNext: () => i < arr.length, next: () => arr[i++] }; };
  const wrap = f => ({
    getId: () => f.id, getName: () => f.name,
    setTrashed: t => { f.trashed = t; },
    getParents: () => iter([{ getName: () => '02. Purchase' }]),
  });
  return {
    getId: () => id,
    getName: () => 'folder:' + id,
    getParents: () => iter([{ getName: () => '02. Purchase' }]),
    getFiles: () => iter(list().map(wrap)),
    getFilesByName: n => iter(list().filter(f => f.name === n).map(wrap)),
    createFile: blob => {
      const f = { id: 'file' + (++fileSeq), name: blob.name, content: blob.content, trashed: false };
      driveFiles[id].push(f);
      return wrap(f);
    },
  };
}

// ── vm 컨텍스트 ───────────────────────────────────────────────
const ss = new FakeSpreadsheet();
const mails = [];
const adminErrors = [];
let activeUser = 'davidkang@inlct.com';

const sandbox = {
  console,
  Logger: { log: () => {} },
  Session: {
    getActiveUser: () => ({ getEmail: () => activeUser }),
    getEffectiveUser: () => ({ getEmail: () => 'davidkang@inlct.com' }),   // 배포 계정
    getScriptTimeZone: () => 'Asia/Seoul',
  },
  SpreadsheetApp: {
    openById: () => ss,
    flush: () => {},
    newDataValidation: () => {
      const r = { values: null, allowInvalid: true };
      const b = {
        requireValueInList(list) { r.values = list; return b; },
        setAllowInvalid(v) { r.allowInvalid = v; return b; },
        build: () => r,
      };
      return b;
    },
  },
  DriveApp: { getFolderById: id => makeFolder(id) },
  Utilities: {
    getUuid: () => 'uuid-' + Math.random().toString(16).slice(2),
    newBlob: (content, type, name) => ({ content, type, name }),
    formatDate: (d, tz, fmt) => {
      const p = n => String(n).padStart(2, '0');
      const t = new Date(d);
      return fmt.indexOf('HH') >= 0
        ? `${t.getFullYear()}-${p(t.getMonth() + 1)}-${p(t.getDate())} ${p(t.getHours())}:${p(t.getMinutes())}:${p(t.getSeconds())}`
        : `${t.getFullYear()}-${p(t.getMonth() + 1)}-${p(t.getDate())}`;
    },
    sleep: () => {},
  },
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  PropertiesService: {
    getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {}, deleteProperty: () => {} }),
  },
  ScriptApp: { getProjectTriggers: () => [], newTrigger: () => ({ timeBased: () => ({ everyMinutes: () => ({ create: () => {} }) }) }) },
  MailApp: { sendEmail: () => {} },
  GmailApp: { sendEmail: () => {} },
  HtmlService: { createTemplateFromFile: () => ({ evaluate: () => ({}) }) },
  MimeType: { PDF: 'application/pdf' },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) },
};
sandbox.globalThis = sandbox;
vm.createContext(sandbox);

vm.runInContext(fs.readFileSync(path.join(SRC, 'Code.gs'), 'utf8'), sandbox, { filename: 'Code.gs' });
vm.runInContext(fs.readFileSync(path.join(SRC, 'PO.gs'), 'utf8'), sandbox, { filename: 'PO.gs' });

// 저수준 헬퍼 대체 (메일/폴더/감사로그)
vm.runInContext(`
  sendEmailWithRetry = function (to, subj, plain, html) { __mails.push({ to: to, subj: subj }); return true; };
  notifyAdminError   = function (msg) { __adminErrors.push(msg); };
  getOrCreateFolder  = function (poNo, issueDate, kind) { return DriveApp.getFolderById('FINAL_' + poNo); };
`, sandbox);
sandbox.__mails = mails;
sandbox.__adminErrors = adminErrors;

// ── 테스트 픽스처 ─────────────────────────────────────────────
const COL = sandbox.COL;
const PO_COL = sandbox.PO_COL;
const CONFIG = sandbox.CONFIG;

function makePrcRow({ docNo, token, status = '최종승인(PRC)', vendor = '라인테크', items, total, moveStatus = 'FINAL' }) {
  const row = new Array(32).fill('');
  row[COL.SUBMIT_AT] = new Date('2026-09-01T10:00:00');
  row[COL.DOC_NO] = docNo;
  row[COL.ISSUE_DATE] = new Date('2026-09-01T00:00:00');
  row[COL.DRAFTER] = '강기종';
  row[COL.DEPT] = '구매팀';
  row[COL.SUBJECT] = 'LCOS Artwork';
  row[COL.VENDOR_NAME] = vendor;
  row[COL.VENDOR_EMAIL] = 'v@example.com';
  row[COL.DELIVERY_DATE] = new Date('2026-09-04T00:00:00');
  row[COL.DELIVERY_ADDR] = '대전 유성구';
  row[COL.ITEMS] = items || '1. LCOS_IN_CONNECTION | ea | 1개 | 100,000 KRW';
  row[COL.TOTAL_AMT] = total ?? 100000;
  row[COL.STATUS] = status;
  row[COL.TOKEN] = token;
  row[COL.DRIVE_ID] = 'FINAL_' + docNo;
  row[COL.MOVE_STATUS] = moveStatus;
  row[COL.DOC_TYPE] = 'PRC';
  row[COL.PURCHASE_METHOD] = '현금';
  row[COL.PAYMENT_INFO] = '익월말';
  return row;
}

const mainSheet = ss.insertSheet(CONFIG.SHEET_NAME);
mainSheet.appendRow(new Array(32).fill('header'));
mainSheet.appendRow(makePrcRow({ docNo: 'TO-PO-26-901', token: 'tok-901' }));
mainSheet.appendRow(makePrcRow({ docNo: 'TO-PO-26-902', token: 'tok-902', vendor: 'KMCO',
  items: '1. Fiber array | 1 ea | 30개 | 143,700 JPY', total: 4311000 }));
mainSheet.appendRow(makePrcRow({ docNo: 'TO-PO-26-903', token: 'tok-903', status: '결재진행중' }));

// ── 단언 헬퍼 ─────────────────────────────────────────────────
let pass = 0, fail = 0;
function check(name, fn) {
  try {
    const r = fn();
    if (r === true) { pass++; console.log('  ✓ ' + name); }
    else { fail++; console.log('  ✗ ' + name + '  → ' + r); }
  } catch (e) {
    fail++; console.log('  ✗ ' + name + '  → 예외: ' + e.message);
  }
}
const poSheet = () => ss.getSheetByName(CONFIG.PO_SHEET_NAME);
const poRows = () => { const s = poSheet(); return s ? s.rows.slice(1) : []; };
const call = (fnName, ...args) => sandbox[fnName].apply(sandbox, args);

console.log('\n[1] 스키마·시트 생성');
check('PO_COL 인덱스 연속', () => call('_poSchemaSelfCheck_') === true || '실패');
check('주문서목록 시트 생성 + 헤더 20열', () => {
  const s = call('ensurePoSheet_');
  return (s.getName() === '주문서목록' && s.rows[0].length === 20) || JSON.stringify(s.rows[0]);
});

console.log('\n[2] 결재완료 → 핸드오프 등록');
const r1 = call('_preparePoHandoff_', 'tok-901');
check('핸드오프 ok', () => r1.ok === true || JSON.stringify(r1));
check('매니페스트 파일 생성', () => (driveFiles['FINAL_TO-PO-26-901'] || []).some(f => f.name === 'po_manifest.json') || '없음');
check('주문서목록 1행 등록', () => poRows().length === 1 || poRows().length);
check('상태 = 생성대기', () => poRows()[0][PO_COL.STATUS] === '생성대기' || poRows()[0][PO_COL.STATUS]);
check('PO번호·업체·통화 스냅샷', () => {
  const r = poRows()[0];
  return (r[PO_COL.PO_NO] === 'TO-PO-26-901' && r[PO_COL.VENDOR_NAME] === '라인테크'
    && r[PO_COL.CURRENCY] === 'KRW' && r[PO_COL.TOTAL_AMT] === 100000) || JSON.stringify(r.slice(0, 9));
});
check('manifestId·folderId 기록', () => (!!poRows()[0][PO_COL.MANIFEST_ID] && !!poRows()[0][PO_COL.FOLDER_ID]) || '누락');
check('작업요청 메일 1건', () => mails.length === 1 || mails.length);

console.log('\n[3] 주문서 생성 완료 마감');
const r2 = call('markPoDone', 'tok-901', 'XLSXID1');
check('markPoDone ok', () => r2.ok === true || JSON.stringify(r2));
check('상태 = 생성완료', () => poRows()[0][PO_COL.STATUS] === '생성완료' || poRows()[0][PO_COL.STATUS]);
check('파일ID·링크·시각·실행자 기록', () => {
  const r = poRows()[0];
  return (r[PO_COL.XLSX_FILE_ID] === 'XLSXID1' && String(r[PO_COL.XLSX_URL]).indexOf('XLSXID1') > 0
    && !!r[PO_COL.GENERATED_AT] && r[PO_COL.GENERATED_BY] === 'davidkang@inlct.com') || JSON.stringify(r.slice(15, 19));
});
check('행 수 그대로(중복 생성 없음)', () => poRows().length === 1 || poRows().length);

console.log('\n[4] 재핸드오프(재발행) 처리');
const r3 = call('_preparePoHandoff_', 'tok-901');
check('행 추가 없이 갱신', () => poRows().length === 1 || poRows().length);
check('상태가 생성대기로 복귀', () => poRows()[0][PO_COL.STATUS] === '생성대기' || poRows()[0][PO_COL.STATUS]);
check('비고에 재핸드오프 이력', () => String(poRows()[0][PO_COL.NOTE]).indexOf('재핸드오프') === 0 || poRows()[0][PO_COL.NOTE]);
check('이전 산출물 정보 비움', () => (poRows()[0][PO_COL.XLSX_FILE_ID] === '' && poRows()[0][PO_COL.XLSX_URL] === '') || '남아있음');

console.log('\n[5] 미결재 건은 등록하지 않음');
const r4 = call('_preparePoHandoff_', 'tok-903');
check('최종승인 아니면 거부', () => r4.ok === false || JSON.stringify(r4));
check('주문서목록 변화 없음', () => poRows().length === 1 || poRows().length);

console.log('\n[6] 미생성 목록');
call('_preparePoHandoff_', 'tok-902');
const pending = call('listPoPending');
check('미생성 2건(901 재발행 + 902)', () => pending.length === 2 || pending.map(p => p.poNo).join(','));

console.log('\n[6-1] 생성불요 처리');
check('상태 드롭다운 3종 설정', () => {
  const v = poSheet().validation;
  return (v && v.rule.values.join(',') === '생성대기,생성완료,생성불요' && v.rule.allowInvalid === false) || JSON.stringify(v && v.rule);
});
check('사유 없으면 거부', () => call('markPoNotRequired', 'tok-902', '  ').ok === false || '통과해 버림');
const r6 = call('markPoNotRequired', 'tok-902', '샘플 입고분 — 발주서 발행 불요');
check('markPoNotRequired ok', () => r6.ok === true || JSON.stringify(r6));
check('상태 = 생성불요', () => {
  const r = poRows().find(x => x[PO_COL.PO_NO] === 'TO-PO-26-902');
  return r[PO_COL.STATUS] === '생성불요' || r[PO_COL.STATUS];
});
check('비고에 사유·처리자 기록', () => {
  const r = poRows().find(x => x[PO_COL.PO_NO] === 'TO-PO-26-902');
  return (String(r[PO_COL.NOTE]).indexOf('발행 불요') > 0 && r[PO_COL.GENERATED_BY] === 'davidkang@inlct.com')
    || JSON.stringify([r[PO_COL.NOTE], r[PO_COL.GENERATED_BY]]);
});
check('미생성 목록에서 빠짐', () => {
  const p = call('listPoPending');
  return (p.length === 1 && p[0].poNo === 'TO-PO-26-901') || p.map(x => x.poNo).join(',');
});
check('오타 상태는 미생성으로 남음(조용히 사라지지 않음)', () => {
  const s = poSheet();
  const rowNum = call('_findPoRowNum_', s, 'tok-902');
  s.getRange(rowNum, PO_COL.STATUS + 1).setValue('생성 완료');   // 공백 오타
  const p = call('listPoPending');
  const back = p.some(x => x.poNo === 'TO-PO-26-902');
  s.getRange(rowNum, PO_COL.STATUS + 1).setValue('생성불요');
  return back || '오타인데도 완료로 처리됨';
});
check('생성불요 건에 재핸드오프 시 다시 대기 + 이력', () => {
  call('_preparePoHandoff_', 'tok-902');
  const r = poRows().find(x => x[PO_COL.PO_NO] === 'TO-PO-26-902');
  return (r[PO_COL.STATUS] === '생성대기' && String(r[PO_COL.NOTE]).indexOf('이전 상태 생성불요') > 0)
    || JSON.stringify([r[PO_COL.STATUS], r[PO_COL.NOTE]]);
});

console.log('\n[7] 권한 관문 (google.script.run 노출 대비)');
activeUser = 'someone@inlct.com';
const before = JSON.stringify(poSheet().rows);
['markPoDone', 'listPoPending', 'backfillPoSheet', 'rerunPoHandoff', 'markPoNotRequired'].forEach(fn => {
  check(fn + ' 비관리자 차단', () => {
    try { call(fn, 'tok-901', 'X'); return '차단 실패(호출됨)'; }
    catch (e) { return e.message.indexOf('관리자 권한') >= 0 || e.message; }
  });
});
check('거부 시 시트 미변경', () => JSON.stringify(poSheet().rows) === before || '시트가 변경됨');
activeUser = 'davidkang@inlct.com';

console.log('\n[8] 백필(과거 건)');
driveFiles[sandbox.PO_CONFIG.ORDER_FOLDER_ID] = [
  { id: 'OLD902', name: 'TO-PO-26-902_KMCO.xlsx', trashed: false },
  { id: 'OLD904', name: 'TO-PO-26-904_빅터스.xlsx', trashed: false },
];
mainSheet.appendRow(makePrcRow({ docNo: 'TO-PO-26-904', token: 'tok-904', vendor: '빅터스' }));
const bf = call('backfillPoSheet');
check('백필 ok', () => bf.ok === true || JSON.stringify(bf));
check('신규 1건만 추가(기존 행 보존)', () => bf.added === 1 || JSON.stringify(bf));
check('폴더에 파일 있으면 생성완료', () => {
  const r = poRows().find(x => x[PO_COL.PO_NO] === 'TO-PO-26-904');
  return (r && r[PO_COL.STATUS] === '생성완료' && r[PO_COL.XLSX_FILE_ID] === 'OLD904') || JSON.stringify(r);
});
check('이미 등록된 902는 덮어쓰지 않음', () => {
  const r = poRows().filter(x => x[PO_COL.PO_NO] === 'TO-PO-26-902');
  return (r.length === 1 && r[0][PO_COL.STATUS] === '생성대기') || JSON.stringify(r);
});

console.log('\n[9] 동명이인 시트 방어 (수기 시트가 이미 있는 경우)');
{
  const saved = ss.sheets[CONFIG.PO_SHEET_NAME];
  const bogus = new FakeSheet(CONFIG.PO_SHEET_NAME);
  bogus.appendRow(['날짜', '메모', '기타']);          // 사람이 만든 다른 시트
  ss.sheets[CONFIG.PO_SHEET_NAME] = bogus;
  check('머리글이 다르면 중단', () => {
    try { call('ensurePoSheet_'); return '통과해 버림(데이터 오염 위험)'; }
    catch (e) { return e.message.indexOf('머리글') >= 0 || e.message; }
  });
  check('남의 시트를 건드리지 않음', () => bogus.rows.length === 1 || JSON.stringify(bogus.rows));
  ss.sheets[CONFIG.PO_SHEET_NAME] = saved;
}

console.log('\n[10] 비치명성 — 대장 쓰기가 실패해도 핸드오프는 계속');
vm.runInContext('ensurePoSheet_ = function () { throw new Error("시트 장애 시뮬레이션"); };', sandbox);
mainSheet.appendRow(makePrcRow({ docNo: 'TO-PO-26-905', token: 'tok-905' }));
const mailsBefore = mails.length;
const r5 = call('_preparePoHandoff_', 'tok-905');
check('핸드오프 자체는 성공', () => r5.ok === true || JSON.stringify(r5));
check('매니페스트 기록됨', () => (driveFiles['FINAL_TO-PO-26-905'] || []).some(f => f.name === 'po_manifest.json') || '없음');
check('작업요청 메일 계속 발송', () => mails.length === mailsBefore + 1 || mails.length);
check('관리자 오류 알림 남김', () => adminErrors.some(m => m.indexOf('주문서목록 등록 실패') >= 0) || adminErrors.join('|'));

console.log('\n─────────────────────────────');
console.log(`통과 ${pass} / 실패 ${fail}`);
process.exit(fail ? 1 : 0);
