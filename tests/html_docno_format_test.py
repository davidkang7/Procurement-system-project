"""
화면(HTML) 검증 — 품의 작성 화면의 품의번호 형식 점검
(memory: html-ui-playwright-harness)
규칙: 영문 대문자 2자리-영문 대문자 2자리-발행 일자 연도 끝 2자리-일련번호 3자리 (예: TO-AP-26-001)
방식: 이전 코드(BASE_REF, 기본 362803c)와 작업 트리의 같은 화면을 같은 스텁 응답으로 띄워
      ① 틀린 번호는 다른 점검보다 먼저 배너로 막히고 서버 호출(제출·중복 조회)이 0건인지
      ② 올바른 번호의 제출 값은 이전 코드와 같은지
      ③ 입력 중 안내·발행 일자 변경·초기화·기본값 제거를 단언한다. 브라우저 기본 대화상자·스크립트 오류는 0건이어야 한다.
실행:
  insp-renderer/.venv/Scripts/python.exe tests/html_docno_format_test.py
"""
import json
import os
import pathlib
import re
import subprocess
import sys
from playwright.sync_api import sync_playwright

SRC = pathlib.Path(__file__).resolve().parent.parent
BASE_REF = os.environ.get('BASE_REF', '362803c')
BASE = 'http://t.local'
WEBAPP = BASE + '/exec'

passed, failed, failures = 0, 0, []


def check(name, cond, detail=''):
    global passed, failed
    if cond is True:
        passed += 1
        print('  ✓ ' + name)
    else:
        failed += 1
        failures.append(name)
        print('  ✗ ' + name + ('  → ' + str(detail)[:700] if detail != '' else ''))


def source(name, rev):
    if rev == 'NEW':
        return (SRC / name).read_text(encoding='utf-8')
    return subprocess.run(['git', 'show', BASE_REF + ':' + name], cwd=SRC, capture_output=True,
                          encoding='utf-8', check=True).stdout


STUB = """<script>
window.__calls = []; window.__resp = %s;
(function () {
  function runner(succ, fail) {
    return new Proxy({}, { get: function (_, prop) {
      if (prop === 'withSuccessHandler') return function (f) { return runner(f, fail); };
      if (prop === 'withFailureHandler') return function (f) { return runner(succ, f); };
      if (prop === 'withUserObject') return function () { return runner(succ, fail); };
      return function () {
        var args = Array.prototype.slice.call(arguments);
        window.__calls.push({ fn: prop, args: JSON.parse(JSON.stringify(args)) });
        var r = window.__resp[prop];
        setTimeout(function () {
          if (!r) return;
          var v = JSON.parse(JSON.stringify(r.value));
          if (r.mode === 'fail') { if (fail) fail(v); } else { if (succ) succ(v); }
        }, 0);
      };
    } });
  }
  window.google = { script: { run: runner(null, null) } };
})();
</script>"""

RESP = {
    'getCurrentUserForClient': {'mode': 'ok', 'value': {'ok': True, 'user': {'email': 'kim@inlct.com', 'name': '김기안', 'dept': '개발팀'}}},
    'getApproverListForClient': {'mode': 'ok', 'value': {'ok': True, 'approvers': []}},
    'checkDocNoForClient': {'mode': 'ok', 'value': {'ok': True, 'duplicated': False}},
}


class Page:
    def __init__(self, browser, name, rev, tvars, resp):
        self.ctx = browser.new_context(viewport={'width': 1400, 'height': 1000})
        self.page = self.ctx.new_page()
        self.page.set_default_timeout(15000)
        self.dialogs, self.errors = [], []
        self.page.on('dialog', lambda d: (self.dialogs.append(d.type + ':' + d.message), d.dismiss()))
        self.page.on('pageerror', lambda e: self.errors.append(str(e)))
        html = source(name, rev)
        html = re.sub(r'<\?!?=\s*(\w+)\s*\?>', lambda m: str(tvars.get(m.group(1), '')), html)
        html = html.replace('<head>', '<head>' + (STUB % json.dumps(resp, ensure_ascii=False)), 1)
        self.page.route(BASE + '/**', lambda r: r.fulfill(
            status=200, content_type='text/html; charset=utf-8',
            body=html if r.request.url.startswith(BASE + '/page') else '<html><body>NAV</body></html>'))
        self.page.goto(BASE + '/page')

    def ev(self, js, arg=None):
        return self.page.evaluate(js, arg) if arg is not None else self.page.evaluate(js)

    def calls(self, fn):
        return [x for x in self.ev('window.__calls') if x['fn'] == fn]

    def wait_calls(self, fn, n=1):
        self.page.wait_for_function('([f, n]) => window.__calls.filter(c => c.fn === f).length >= n', arg=[fn, n])

    def text(self, sel):
        return self.page.inner_text(sel).strip()

    def close(self):
        self.ctx.close()


CLOSE_PROCESSING = "var b = document.getElementById('pm-backdrop'); if (b) b.classList.remove('show');"
SHAPE = '품의번호 형식이 올바르지 않습니다. 영문 대문자 2자리-영문 대문자 2자리-발행연도 2자리-일련번호 3자리로 입력해 주세요. (예: TO-AP-26-001)'

FILL_VALID = """() => {
  document.getElementById('issue-date').value = '2026-10-30';
  document.getElementById('subject').value = '검사 지그 제작';
  document.getElementById('purpose').value = '라인 검사';
  document.getElementById('vendorName').value = '(주)태성테크';
  document.getElementById('vendorContact').value = '박담당';
  document.getElementById('vendorEmail').value = 'sales@example.com';
  var people = [['이결재', 'lee@inlct.com'], ['최결재', 'choi@inlct.com'], ['박결재', 'park@inlct.com']];
  for (var i = 1; i < apprIds.length; i++) {
    document.getElementById('trig-' + apprIds[i]).textContent = people[i - 1][0];
    document.getElementById('email-' + apprIds[i]).value = people[i - 1][1];
  }
  var tr = document.querySelector('#items-body tr');
  tr.querySelector('.item-name').value = '지그 본체'; tr.querySelector('.item-spec').value = 'SUS304 / EA';
  tr.querySelector('.qty').value = 3; tr.querySelector('.price').value = '15,000';
  calcRow(tr.querySelector('.price'));
}"""


def open_form(browser, rev):
    p = Page(browser, 'Procurement_Form.html', rev, {'webappUrl': WEBAPP}, RESP)
    p.page.wait_for_selector('#items-body tr')
    p.page.wait_for_function("document.getElementById('drafter').value === '김기안'")
    return p


def type_docno(p, v):
    """입력란에 실제로 타이핑 → 600ms 뒤 형식 안내 또는 중복 조회"""
    p.page.fill('#doc-no', v)
    p.page.wait_for_timeout(800)


def dup_msg(p):
    return [p.text('#dup-msg'), p.ev("document.getElementById('dup-msg').className")]


def submit_until_server(p):
    """submitForm → (견적서 미지정 확인 모달 '계속 제출') → 서버 제출 호출 값"""
    n = len(p.calls('submitRequisitionForClient'))
    p.ev('submitForm()')
    if p.ev("document.getElementById('pm-backdrop').classList.contains('show')"):
        p.page.click('#pm-ok-btn')
    p.wait_calls('submitRequisitionForClient', n + 1)
    p.ev(CLOSE_PROCESSING)
    return p.calls('submitRequisitionForClient')[-1]['args'][0]


# ─────────────────────────────────────────────────────────────
def first_view(browser):
    print('\n[H] 처음 화면 · 입력 중 안내')
    n = open_form(browser, 'NEW')
    yy = n.ev("document.getElementById('issue-date').value").split('-')[0][2:]
    check('H1 품의번호 빈칸 · 안내 문구 placeholder · 상단 번호 "-"',
          n.ev("document.getElementById('doc-no').value") == ''
          and n.ev("document.getElementById('doc-no').placeholder") == '예: TO-AP-' + yy + '-001'
          and n.text('#dno-disp') == '-',
          [n.ev("document.getElementById('doc-no').placeholder"), n.text('#dno-disp')])

    n.ev("document.getElementById('issue-date').value = '2026-10-30'")
    type_docno(n, 'lc-2026-020')
    check('H2 소문자로 쳐도 대문자로 · 상단 번호 따라감', n.ev("document.getElementById('doc-no').value") == 'LC-2026-020'
          and n.text('#dno-disp') == 'LC-2026-020')
    check('H2 형식 오류 안내 (빨간색) · 중복 조회 0건',
          dup_msg(n) == ['형식 오류 — 예: TO-AP-26-001', 'dup-msg err'] and not n.calls('checkDocNoForClient'), [dup_msg(n), n.calls('checkDocNoForClient')])

    type_docno(n, 'TR-AP-20-056')
    check('H3 연도 오타 → "연도는 26 (발행 일자 기준)" · 중복 조회 0건',
          dup_msg(n) == ['연도는 26 (발행 일자 기준)', 'dup-msg err'] and not n.calls('checkDocNoForClient'), dup_msg(n))

    type_docno(n, 'TO-AP-26-045')
    c = n.calls('checkDocNoForClient')
    check('H4 올바른 번호 → 중복 조회 1건 · "사용 가능한 품의번호입니다."',
          len(c) == 1 and c[0]['args'] == ['TO-AP-26-045'] and dup_msg(n) == ['사용 가능한 품의번호입니다.', 'dup-msg ok'], [c, dup_msg(n)])

    n.page.fill('#issue-date', '2027-01-02')
    n.page.dispatch_event('#issue-date', 'change')
    check('H5 발행 일자를 2027년으로 바꾸면 즉시 "연도는 27"', dup_msg(n) == ['연도는 27 (발행 일자 기준)', 'dup-msg err'], dup_msg(n))
    n.page.fill('#issue-date', '2026-10-30')
    n.page.dispatch_event('#issue-date', 'change')
    n.page.wait_for_timeout(200)
    check('H5 2026년으로 되돌리면 다시 "사용 가능"', dup_msg(n) == ['사용 가능한 품의번호입니다.', 'dup-msg ok'], dup_msg(n))

    n.ev('doResetAll()')
    check('H6 초기화 → 품의번호 빈칸 · 상단 "-" · 안내 지움',
          n.ev("document.getElementById('doc-no').value") == '' and n.text('#dno-disp') == '-' and dup_msg(n) == ['', 'dup-msg'],
          [n.text('#dno-disp'), dup_msg(n)])
    check('H 브라우저 대화상자 0건 · 스크립트 오류 0건', not n.dialogs and not n.errors, [n.dialogs, n.errors])
    n.close()


# ─────────────────────────────────────────────────────────────
def submit_block(browser):
    print('\n[B] 제출 — 틀린 번호는 다른 점검보다 먼저 막힌다')
    for docno, want in (('LC-2026-021', SHAPE), ('TO-AP-026-045', SHAPE), ('PRQ-2026-001', SHAPE),
                        ('TR-AP-20-056', '품의번호의 연도(20)가 발행 일자(2026-10-30)의 연도(26)와 다릅니다.')):
        n = open_form(browser, 'NEW')
        n.ev(FILL_VALID)
        n.ev("([v]) => { var e = document.getElementById('doc-no'); e.value = v; }", [docno])
        n.ev('submitForm()')
        check(f'B {docno} → 배너 문구', n.text('#status-banner') == want and n.page.is_visible('#status-banner'), n.text('#status-banner'))
        check(f'B {docno} → 제출 호출 0건 · 확인 모달 안 뜸 · 입력란 빨간 안내',
              not n.calls('submitRequisitionForClient') and not n.ev("document.getElementById('pm-backdrop').classList.contains('show')")
              and 'dup-err' in n.ev("document.getElementById('doc-no').className"),
              [len(n.calls('submitRequisitionForClient')), n.ev("document.getElementById('doc-no').className")])
        n.close()

    # 점검 순서 — 공급업체를 비워 둬도 품의번호 형식이 먼저
    n = open_form(browser, 'NEW')
    n.ev(FILL_VALID)
    n.ev("document.getElementById('vendorName').value = ''; document.getElementById('doc-no').value = 'LC-2026-021';")
    n.ev('submitForm()')
    check('B 공급업체 미입력이어도 품의번호 형식 문구가 먼저', n.text('#status-banner') == SHAPE, n.text('#status-banner'))
    # 필수 항목(번호 빈칸)은 이전처럼 필수 문구
    n.ev("document.getElementById('doc-no').value = ''")
    n.ev('submitForm()')
    check('B 번호 빈칸 → 이전과 같은 필수 항목 문구', n.text('#status-banner') == '필수 항목(품의번호, 기안자, 품의제목)을 입력해 주세요.',
          n.text('#status-banner'))
    # 발행 일자 빈칸
    n.ev("document.getElementById('doc-no').value = 'TO-AP-26-045'; document.getElementById('issue-date').value = ''")
    n.ev('submitForm()')
    check('B 발행 일자 빈칸 → "발행 일자를 입력해 주세요."', n.text('#status-banner') == '발행 일자를 입력해 주세요.', n.text('#status-banner'))
    check('B 대화상자 0건 · 스크립트 오류 0건 · 제출 호출 0건', not n.dialogs and not n.errors and not n.calls('submitRequisitionForClient'),
          [n.dialogs, n.errors])
    n.close()


# ─────────────────────────────────────────────────────────────
def submit_valid(browser):
    print('\n[S] 올바른 번호의 제출 — 서버로 가는 값이 이전 코드와 같다')
    pg = {rev: open_form(browser, rev) for rev in ('HEAD', 'NEW')}
    for p in pg.values():
        p.ev(FILL_VALID)
        type_docno(p, 'TO-AP-26-045')
    ph, pn = submit_until_server(pg['HEAD']), submit_until_server(pg['NEW'])
    check('S1 제출 값 이전과 같음', ph == pn, [ph, pn])
    check('S1 docNo · issueDate', pn['docNo'] == 'TO-AP-26-045' and pn['issueDate'] == '2026-10-30', pn.get('docNo'))
    for rev, p in pg.items():
        check(f'S {rev} 대화상자 0건 · 스크립트 오류 0건', not p.dialogs and not p.errors, [p.dialogs, p.errors])
        p.close()


with sync_playwright() as pw:
    browser = pw.chromium.launch()
    try:
        first_view(browser)
        submit_block(browser)
        submit_valid(browser)
    finally:
        browser.close()

print('\n─────────────────────────────')
print(f'통과 {passed} / 실패 {failed}   (기준 {BASE_REF})')
if failed:
    print('실패 항목:\n  - ' + '\n  - '.join(failures))
    sys.exit(1)
