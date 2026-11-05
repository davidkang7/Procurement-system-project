"""
화면(HTML) 검증 — 검수보고서 제출 권한 SCM 역할 (2026-11-05) / 홈 결재 완료
(memory: html-ui-playwright-harness)

방식: 이전 코드(BASE_REF, 기본 6b2e9eb)와 작업 트리의 홈 화면을 서버 하네스가 만든 응답(픽스처)으로 띄워
      ① SCM 사용자 — 배지·안내 문구·타 부서 품의의 제출/재상신 버튼과 이동 URL을 단언하고
      ② SCM이 아닌 사용자 — 바뀐 안내 문구 한 줄 외에는 화면이 이전 코드와 같은지 비교한다.
      브라우저 기본 대화상자(alert/confirm)·스크립트 오류는 0건이어야 한다.

픽스처: DUMP_FIXTURES=<경로> node tests/insp_submit_role_test.js
실행:   insp-renderer/.venv/Scripts/python.exe tests/html_insp_submit_role_test.py <fixtures.json>
"""
import json
import os
import pathlib
import re
import subprocess
import sys

from playwright.sync_api import sync_playwright

SRC = pathlib.Path(__file__).resolve().parent.parent
FIX = json.loads(pathlib.Path(sys.argv[1]).read_text(encoding='utf-8'))
BASE_REF = os.environ.get('BASE_REF', '6b2e9eb')
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
        print('  ✗ ' + name + ('  → ' + str(detail)[:600] if detail != '' else ''))


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


def build(rev, data):
    html = source('Procurement_Home.html', rev)
    tvars = {'webappUrl': WEBAPP, 'initialMenu': 'completed'}
    html = re.sub(r'<\?!?=\s*(\w+)\s*\?>', lambda m: str(tvars.get(m.group(1), '')), html)
    resp = {'getHomeDataForClient': {'mode': 'ok', 'value': data}}
    return html.replace('<head>', '<head>' + (STUB % json.dumps(resp, ensure_ascii=False)), 1)


class Home:
    """홈 화면 + 서버 스텁. 이동(window.top.location)은 더미 페이지로 받는다."""

    def __init__(self, browser, rev, data, open_rows=()):
        self.ctx = browser.new_context(viewport={'width': 1400, 'height': 1000})
        self.page = self.ctx.new_page()
        self.page.set_default_timeout(8000)
        self.dialogs, self.errors = [], []
        self.page.on('dialog', lambda d: (self.dialogs.append(d.type + ':' + d.message), d.dismiss()))
        self.page.on('pageerror', lambda e: self.errors.append(str(e)))
        html = build(rev, data)

        def handler(route):
            if route.request.url.startswith(BASE + '/page'):
                route.fulfill(status=200, content_type='text/html; charset=utf-8', body=html)
            else:
                route.fulfill(status=200, content_type='text/html; charset=utf-8', body='<html><body>NAV</body></html>')
        self.page.route(BASE + '/**', handler)
        self.page.goto(BASE + '/page')
        self.page.wait_for_selector('.page-sub')
        for tok in open_rows:
            self.page.click('tr[data-toggle-insp="%s"]' % tok)
            self.page.wait_for_selector('tr[data-toggle-insp="%s"].open' % tok)

    def pills(self):
        return self.page.locator('#user-roles .role-pill').all_inner_texts()

    def sub(self):
        return self.page.locator('.page-sub').first.inner_text()

    def close(self):
        self.ctx.close()


SCM_NOTE = 'SCM 담당 권한으로 모든 품의의 검수보고서를 제출할 수 있습니다.'
OLD_FOOT = re.compile(r'\) 또는 같은 부서 담당자만 가능합니다\.')
NEW_FOOT = '), 같은 부서 담당자 또는 SCM 담당만 가능합니다.'


def scm_user(browser):
    print('\n[S] SCM 전용 사용자(logistics) — 타 부서 품의')
    p = Home(browser, 'NEW', FIX['homeScm'], open_rows=('tok-prcA', 'tok-prcB'))
    check('S1 역할 배지: 부서 + "SCM"', p.pills() == ['물류팀', 'SCM'], p.pills())
    check('S2 안내 문구: 전사 표시 + SCM 제출 안내 (같은 문장 반복 없음)',
          p.sub().count('SCM') == 1 and '전사 문서가 표시됩니다.' in p.sub() and SCM_NOTE in p.sub(), p.sub())
    rows = p.page.locator('tr[data-toggle-insp]')
    check('S3 결재 완료 목록 = 개발팀·영업팀 품의 2건 (원본 없는 PO-903 없음)',
          sorted(rows.evaluate_all('els => els.map(e => e.getAttribute("data-toggle-insp"))')) == ['tok-prcA', 'tok-prcB'])
    btn = p.page.locator('[data-insp-form="tok-prcB"]:not([data-insp-resub])')
    check('S4 영업팀 품의(검수 없음): "+ 검수보고서 제출 (01회차)" 버튼', btn.count() == 1 and '01회차' in btn.inner_text(),
          btn.count())
    resub = p.page.locator('[data-insp-resub]')
    check('S5 남(김기안)이 낸 반려 회차: "수정 재상신" 버튼 — 제출자가 아니므로 "재상신/폐기" 아님',
          resub.count() == 1 and resub.inner_text() == '수정 재상신', resub.all_inner_texts())
    check('S6 미권한 안내("…만 가능합니다") 없음', '만 가능합니다' not in p.page.inner_text('#content'))
    btn.click()
    p.page.wait_for_url('**/exec?**')
    check('S7 제출 버튼 → insp_form 화면(token=PRC)', 'action=insp_form' in p.page.url and 'token=tok-prcB' in p.page.url, p.page.url)
    check('S 브라우저 대화상자 0건 · 스크립트 오류 0건', not p.dialogs and not p.errors, [p.dialogs, p.errors])
    p.close()

    print('\n[A] 관리자 + SCM(davidkang)')
    p = Home(browser, 'NEW', FIX['homeAdmin'], open_rows=('tok-prcB',))
    check('A1 역할 배지에 SCM·관리자 함께 표시', 'SCM' in p.pills() and '관리자' in p.pills() and '구매팀' in p.pills(), p.pills())
    check('A2 안내 문구: 관리자 "(제출은 본인·같은 부서만)" 단서 빠지고 SCM 제출 안내',
          '본인·같은 부서 건만 가능' not in p.sub() and SCM_NOTE in p.sub(), p.sub())
    check('A3 타 부서 품의에 제출 버튼', p.page.locator('[data-insp-form="tok-prcB"]').count() == 1)
    check('A 브라우저 대화상자 0건 · 스크립트 오류 0건', not p.dialogs and not p.errors, [p.dialogs, p.errors])
    p.close()


def non_scm(browser):
    print('\n[N] SCM이 아닌 사용자 — 이전 코드와 같은 화면 (미권한 안내 한 줄만 문구 변경)')
    for key, label, rows in (('Adm2', '관리자(비SCM)', ('tok-prcA', 'tok-prcB')), ('Gv', '전사조회(비SCM)', ('tok-prcA', 'tok-prcB')),
                             ('Choi', '타 부서 최영업', ('tok-prcB',)), ('Kim', '기안자 김기안', ('tok-prcA',))):
        b = Home(browser, 'BASE', FIX['home' + key + 'Base'], open_rows=rows)
        n = Home(browser, 'NEW', FIX['home' + key + 'New'], open_rows=rows)
        hb = OLD_FOOT.sub(NEW_FOOT, b.page.inner_html('#content'))
        hn = n.page.inner_html('#content')
        check('N %s: 결재 완료 화면 동일' % label, hb == hn, '')
        check('N %s: 역할 배지 동일 (SCM 없음)' % label, b.pills() == n.pills() and 'SCM' not in n.pills(), [b.pills(), n.pills()])
        check('N %s: 대화상자·스크립트 오류 0건' % label, not n.dialogs and not n.errors, [n.dialogs, n.errors])
        if key == 'Adm2':
            check('N 관리자(비SCM): 미권한 안내에 "SCM 담당" 포함', NEW_FOOT in hn and '본인·같은 부서 건만 가능' in n.sub(), n.sub())
        b.close(); n.close()


with sync_playwright() as pw:
    browser = pw.chromium.launch()
    for fn in (scm_user, non_scm):
        try:
            fn(browser)
        except Exception as e:  # 한 화면이 죽어도 나머지는 돌린다 — 실패로 집계
            failed += 1
            failures.append(fn.__name__ + ' 예외')
            print('  ✗ ' + fn.__name__ + ' 예외: ' + repr(e)[:800])
    browser.close()

print('\n─────────────────────────────')
print('통과 %d / 실패 %d   (기준 %s)' % (passed, failed, BASE_REF))
if failed:
    print('실패 항목:\n  - ' + '\n  - '.join(failures))
sys.exit(1 if failed else 0)
