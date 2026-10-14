"""
화면(HTML) 검증 — 관리자 강제 폐기 / 검수보고서 작성(재상신) / 홈 결재 완료 목록
(memory: html-ui-playwright-harness)

방식: HEAD와 작업 트리의 같은 화면을 같은 서버 응답(픽스처)으로 띄워
      ① 서버로 가는 페이로드가 같은지 ② 화면이 같은지를 비교하고,
      새 기능(INSP 폐기)은 따로 단언한다. 브라우저 기본 대화상자(alert/confirm)는 0건이어야 한다.

픽스처: 서버 하네스가 실제 .gs 함수로 만든 응답
  DUMP_FIXTURES=<경로> node tests/audit_insp_discard_test.js
실행:
  insp-renderer/.venv/Scripts/python.exe tests/html_insp_discard_test.py <fixtures.json>
"""
import json
import pathlib
import re
import subprocess
import sys

from playwright.sync_api import sync_playwright

SRC = pathlib.Path(__file__).resolve().parent.parent
FIX = json.loads(pathlib.Path(sys.argv[1]).read_text(encoding='utf-8'))
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
    return subprocess.run(['git', 'show', 'HEAD:' + name], cwd=SRC, capture_output=True,
                          encoding='utf-8', check=True).stdout


STUB = """<script>
window.__calls = []; window.__resp = %s; window.__dialogs = 0;
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


def build(name, rev, tvars, resp):
    html = source(name, rev)
    html = re.sub(r'<\?!?=\s*(\w+)\s*\?>', lambda m: str(tvars.get(m.group(1), '')), html)
    return html.replace('<head>', '<head>' + (STUB % json.dumps(resp, ensure_ascii=False)), 1)


class Page:
    """한 화면 + 서버 스텁. goto 이동(window.top.location)은 더미 페이지로 받는다."""

    def __init__(self, browser, name, rev, tvars, resp):
        self.ctx = browser.new_context(viewport={'width': 1400, 'height': 1000})
        self.page = self.ctx.new_page()
        self.page.set_default_timeout(8000)
        self.dialogs = []
        self.page.on('dialog', lambda d: (self.dialogs.append(d.type + ':' + d.message), d.dismiss()))
        self.errors = []
        self.page.on('pageerror', lambda e: self.errors.append(str(e)))
        html = build(name, rev, tvars, resp)

        def handler(route):
            if route.request.url.startswith(BASE + '/page'):
                route.fulfill(status=200, content_type='text/html; charset=utf-8', body=html)
            else:
                route.fulfill(status=200, content_type='text/html; charset=utf-8', body='<html><body>NAV</body></html>')
        self.page.route(BASE + '/**', handler)
        self.page.goto(BASE + '/page')

    def calls(self, fn=None):
        c = self.page.evaluate('window.__calls')
        return [x for x in c if fn is None or x['fn'] == fn]

    def resp(self, fn, value, mode='ok'):
        self.page.evaluate('([f, v, m]) => { window.__resp[f] = { mode: m, value: v }; }', [fn, value, mode])

    def wait_calls(self, fn, n=1):
        self.page.wait_for_function('([f, n]) => window.__calls.filter(c => c.fn === f).length >= n', arg=[fn, n])

    def close(self):
        self.ctx.close()


def admin_discard(browser):
    print('\n[D] 관리자 — 강제 폐기 화면')
    resp = {'adminListDiscardableForClient': {'mode': 'ok', 'value': FIX['adminList']},
            'adminForceDiscardForClient': {'mode': 'ok', 'value': {'ok': True, 'message': '문서가 강제 폐기되었습니다.'}}}
    pg = {rev: Page(browser, 'Procurement_Admin_Discard.html', rev, {'webappUrl': WEBAPP}, resp) for rev in ('HEAD', 'NEW')}
    for rev, p in pg.items():
        p.page.wait_for_selector('#table-container table')
    h, n = pg['HEAD'], pg['NEW']

    check('D1 첫 로딩 요청 페이로드 HEAD와 동일 ({})', h.calls('adminListDiscardableForClient') == n.calls('adminListDiscardableForClient'),
          [h.calls(), n.calls()])

    # 유형 '전체'로 검색 → 서버로 가는 값이 이전과 같아야 한다
    for p in (h, n):
        p.page.fill('#f-docno', 'TH-AP')
        p.page.click('button:has-text("검색")')
        p.wait_calls('adminListDiscardableForClient', 2)
    check('D3 검색(유형=전체) 페이로드 HEAD와 동일 — docType 키 없음',
          h.calls('adminListDiscardableForClient')[-1] == n.calls('adminListDiscardableForClient')[-1]
          and 'docType' not in n.calls('adminListDiscardableForClient')[-1]['args'][0],
          [h.calls()[-1], n.calls()[-1]])

    n.page.select_option('#f-doctype', 'INSP')
    n.page.click('button:has-text("검색")')
    n.wait_calls('adminListDiscardableForClient', 3)
    check('D4 유형=INSP 검색 → docType:"INSP" 전달', n.calls('adminListDiscardableForClient')[-1]['args'][0].get('docType') == 'INSP',
          n.calls()[-1])
    n.page.click('button:has-text("초기화")')
    n.wait_calls('adminListDiscardableForClient', 4)
    check('D8 초기화 → 유형도 전체로, 페이로드 {}',
          n.page.eval_on_selector('#f-doctype', 'e => e.value') == '' and n.calls('adminListDiscardableForClient')[-1]['args'] == [{}])
    n.page.wait_for_selector('#table-container table')

    row3 = n.page.locator('tr', has_text='TH-AP-26-006-03')
    check('D2 INSP 행: 문서번호 청록(docno insp)·금액 "-"·★최종·결재중 색',
          row3.locator('span.docno.insp').count() == 1
          and row3.locator('td').nth(4).inner_text().strip() == '-'
          and '★최종' in row3.inner_text()
          and row3.locator('span.status.in_appr').count() == 1, row3.inner_html())
    row1 = n.page.locator('tr', has_text='TH-AP-26-006-01')
    check('D2 최종승인(INSP) 행은 보호 — 회색·클릭 불가',
          'row-protected' in (row1.get_attribute('class') or '') and row1.get_attribute('onclick') is None
          and '보호된 문서' in row1.inner_text())
    req8 = n.page.locator('tr', has_text='TH-AP-26-008')
    check('D2 [버그 수정] REQ "결재중 (결재자 1)"도 결재중 색으로 (HEAD는 회색)',
          req8.locator('span.status.in_appr').count() == 1
          and h.page.locator('tr', has_text='TH-AP-26-008').locator('span.status.unknown').count() == 1)

    # REQ 강제 폐기 — 페이로드가 이전과 같아야 한다
    for p in (h, n):
        p.page.locator('tr', has_text='TH-AP-26-001').click()
        p.page.fill('#reason-input', '중복 제출')
        p.page.click('#execute-btn')
        p.page.wait_for_selector('#confirm-modal.show')
    check('D5 REQ 확인 모달 문구 HEAD와 동일',
          h.page.inner_text('#confirm-modal-body') == n.page.inner_text('#confirm-modal-body'),
          [h.page.inner_text('#confirm-modal-body'), n.page.inner_text('#confirm-modal-body')])
    for p in (h, n):
        p.page.click('#confirm-exec-btn')
        p.wait_calls('adminForceDiscardForClient', 1)
        p.page.wait_for_selector('#result-modal.show')
    check('D5 REQ 강제 폐기 페이로드 HEAD와 동일 {token, reason}',
          h.calls('adminForceDiscardForClient') == n.calls('adminForceDiscardForClient')
          and n.calls('adminForceDiscardForClient')[0]['args'] == [{'token': 'tok-req1', 'reason': '중복 제출'}],
          [h.calls('adminForceDiscardForClient'), n.calls('adminForceDiscardForClient')])
    n.page.click('#result-ok')

    # INSP 강제 폐기
    n.page.locator('tr', has_text='TH-AP-26-006-03').click()
    n.page.fill('#reason-input', '중복 제출 정리')
    n.page.click('#execute-btn')
    n.page.wait_for_selector('#confirm-modal.show')
    body = n.page.inner_text('#confirm-modal-body')
    check('D6 INSP 확인 모달: 제출자 표기·3명 통지(제출자·품의 기안자·현재 결재 대기자)·회차/PO 안내',
          '제출자' in body and '3명' in body and all(e in body for e in ('park@inlct.com', 'kim@inlct.com', 'lee@inlct.com'))
          and '현재 결재 대기자' in body and '3회차' in body and 'TO-PO-26-901' in body and '다시 제출' in body
          and '휴지통' in body, body)
    n.page.click('#confirm-exec-btn')
    n.wait_calls('adminForceDiscardForClient', 2)
    n.page.wait_for_selector('#result-modal.show')
    check('D6 INSP 강제 폐기 페이로드 {token, reason, docType:"INSP"}',
          n.calls('adminForceDiscardForClient')[-1]['args'] == [{'token': 'tok-i3', 'reason': '중복 제출 정리', 'docType': 'INSP'}],
          n.calls('adminForceDiscardForClient')[-1])
    check('D6 결과 모달 + 목록에서 사라짐',
          '처리 완료' in n.page.inner_text('#result-title') and n.page.locator('#table-container tr', has_text='TH-AP-26-006-03').count() == 0)
    check('D7 브라우저 기본 대화상자 0건 · 스크립트 오류 0건', not n.dialogs and not h.dialogs and not n.errors, [n.dialogs, n.errors])
    for p in pg.values():
        p.close()


def insp_form(browser):
    print('\n[E] 검수보고서 작성 / 재상신 화면')
    ok_submit = {'mode': 'ok', 'value': {'ok': True, 'message': '제출 완료', 'docNo': 'TH-AP-26-009-01', 'token': 't', 'photos': 0}}

    # E1 신규 제출 — 폐기 버튼 없음, 제출 페이로드 동일
    resp = {'getInspFormDataForClient': {'mode': 'ok', 'value': FIX['inspFormKimNew']}, 'submitInspForClient': ok_submit}
    pg = {rev: Page(browser, 'Procurement_INSP_Form.html', rev, {'prcToken': 'tok-prc2', 'resubToken': '', 'webappUrl': WEBAPP}, resp)
          for rev in ('HEAD', 'NEW')}
    for p in pg.values():
        p.page.wait_for_selector('#submit-btn')
    h, n = pg['HEAD'], pg['NEW']
    check('E1 신규 제출 화면에는 폐기 버튼 없음 (HEAD와 같은 화면)',
          n.page.locator('#discard-btn').count() == 0 and h.page.inner_html('#wrap') == n.page.inner_html('#wrap'))
    for p in (h, n):
        p.page.click('.verdict.pass')
        p.page.fill('#f-note', '지그 본체 2EA')
        p.page.evaluate("STATE.appr[0] = {name:'이결재', dept:'경영팀', rank:'', email:'lee@inlct.com'}; renderApprTable();")
        p.page.click('#submit-btn')
        p.page.click('#fin-no')
        p.wait_calls('submitInspForClient', 1)
    check('E1 제출 페이로드 HEAD와 동일', h.calls('submitInspForClient') == n.calls('submitInspForClient'),
          [h.calls('submitInspForClient'), n.calls('submitInspForClient')])
    for p in pg.values():
        p.close()

    # E3 재상신 — 같은 부서 동료(canDiscard=false): 폐기 버튼 없음, 화면·재상신 페이로드 HEAD와 동일
    resp = {'getInspFormDataForClient': {'mode': 'ok', 'value': FIX['inspFormParkResub']}, 'submitInspForClient': ok_submit}
    pg = {rev: Page(browser, 'Procurement_INSP_Form.html', rev, {'prcToken': 'tok-prc1', 'resubToken': 'tok-i2', 'webappUrl': WEBAPP}, resp)
          for rev in ('HEAD', 'NEW')}
    for p in pg.values():
        p.page.wait_for_selector('#submit-btn')
    h, n = pg['HEAD'], pg['NEW']
    check('E3 동료의 재상신 화면: 폐기 버튼 없음 + HEAD와 같은 화면',
          n.page.locator('#discard-btn').count() == 0 and h.page.inner_html('#wrap') == n.page.inner_html('#wrap'))
    for p in (h, n):
        p.page.evaluate("STATE.appr[0] = {name:'이결재', dept:'경영팀', rank:'', email:'lee@inlct.com'}; renderApprTable();")
        p.page.click('#submit-btn')
        p.page.click('#fin-yes')
        p.wait_calls('submitInspForClient', 1)
    check('E6 재상신 페이로드 HEAD와 동일', h.calls('submitInspForClient') == n.calls('submitInspForClient'),
          [h.calls('submitInspForClient'), n.calls('submitInspForClient')])
    for p in pg.values():
        p.close()

    # E2 재상신 — 제출자 본인(canDiscard=true): 폐기 흐름
    resp = {'getInspFormDataForClient': {'mode': 'ok', 'value': FIX['inspFormKimResub']},
            'discardInspForClient': {'mode': 'ok', 'value': {'ok': True, 'message': '검수보고서가 폐기 처리되었습니다.'}}}
    n = Page(browser, 'Procurement_INSP_Form.html', 'NEW', {'prcToken': 'tok-prc1', 'resubToken': 'tok-i2', 'webappUrl': WEBAPP}, resp)
    n.page.wait_for_selector('#submit-btn')
    btns = n.page.locator('.actions button').all_inner_texts()
    check('E2 제출자 본인에게 [폐기] 버튼 — 취소·인쇄·폐기·재상신 순', btns == ['취소', '인쇄', '폐기', '재상신'], btns)
    n.page.click('#discard-btn')
    n.page.wait_for_selector('#overlay.on')
    card = n.page.inner_text('#ov-card')
    check('E2 확인 모달(브라우저 confirm 아님): 문서번호·회차·복구 불가·사진 휴지통·번호 재사용 안 함',
          '폐기하시겠습니까' in card and 'TH-AP-26-006-02' in card and '2회차' in card and '복구할 수 없습니다' in card
          and '휴지통' in card and '다시 쓰지 않습니다' in card and not n.dialogs, card)
    focused = n.page.evaluate('document.activeElement && document.activeElement.id')
    n.page.keyboard.press('Enter')
    n.page.wait_for_timeout(200)
    check('E2 기본 포커스 없음 — Enter 한 번으로 폐기되지 않음', focused != 'dc-ok' and not n.calls('discardInspForClient'), [focused, n.calls()])
    n.page.click('#dc-cancel')
    check('E2 취소 → 모달 닫힘, 서버 호출 없음', not n.page.locator('#overlay.on').count() and not n.calls('discardInspForClient'))
    n.page.click('#discard-btn')
    n.page.click('#dc-ok')
    n.wait_calls('discardInspForClient', 1)
    n.page.wait_for_selector('.ok-icon')
    check('E2 폐기 확인 → discardInspForClient({token, reason}) 1회',
          n.calls('discardInspForClient') == [{'fn': 'discardInspForClient', 'args': [{'token': 'tok-i2', 'reason': '제출자 폐기 (반려 후)'}]}],
          n.calls('discardInspForClient'))
    check('E2 완료 카드 "폐기 완료" → 홈으로 이동', '폐기 완료' in n.page.inner_text('#ov-card'))
    n.page.click('#ov-card button:has-text("홈으로")')
    n.page.wait_for_url(WEBAPP)
    check('E2 홈 이동 URL = WEBAPP_URL', n.page.url == WEBAPP, n.page.url)
    check('E2 브라우저 대화상자 0건 · 스크립트 오류 0건', not n.dialogs and not n.errors, [n.dialogs, n.errors])
    n.close()

    # E4 서버 거부 / E5 네트워크 오류 → 오류 모달, 버튼 다시 활성
    for code, mode, value, expect in (
            ('E4', 'ok', {'ok': False, 'message': '반려된 검수보고서만 폐기할 수 있습니다. (상태: 결재중)'}, '반려된 검수보고서만'),
            ('E5', 'fail', {'message': 'timeout'}, '네트워크 오류')):
        resp = {'getInspFormDataForClient': {'mode': 'ok', 'value': FIX['inspFormKimResub']}, 'discardInspForClient': {'mode': mode, 'value': value}}
        n = Page(browser, 'Procurement_INSP_Form.html', 'NEW', {'prcToken': 'tok-prc1', 'resubToken': 'tok-i2', 'webappUrl': WEBAPP}, resp)
        n.page.wait_for_selector('#discard-btn')
        n.page.click('#discard-btn')
        n.page.click('#dc-ok')
        n.page.wait_for_selector('.x-icon >> nth=0')
        n.page.wait_for_function("document.getElementById('ov-card').innerText.indexOf('폐기에 실패했습니다') >= 0")
        card = n.page.inner_text('#ov-card')
        enabled = n.page.evaluate("!document.getElementById('discard-btn').disabled && !document.getElementById('submit-btn').disabled")
        check(code + ' 실패 시 오류 모달("' + expect + '") + 버튼 다시 활성', expect in card and enabled and not n.dialogs, [card, enabled])
        n.close()


def home(browser):
    print('\n[F] 홈 — 결재 완료 · 회차 목록')

    def open_home(rev, data):
        p = Page(browser, 'Procurement_Home.html', rev, {'webappUrl': WEBAPP, 'initialMenu': 'completed'},
                 {'getHomeDataForClient': {'mode': 'ok', 'value': data}})
        p.page.wait_for_selector('tr[data-toggle-insp="tok-prc1"]')
        p.page.click('tr[data-toggle-insp="tok-prc1"]')
        p.page.wait_for_selector('.insp-sub-table')
        return p

    # F1 동료(박동료) 화면 — 제출자가 아니므로 HEAD와 완전히 같아야 한다
    h, n = open_home('HEAD', FIX['homePark']), open_home('NEW', FIX['homePark'])
    check('F1 폐기 건 없음 + 제출자 아님 → 결재 완료 화면 HEAD와 동일', h.page.inner_html('#content') == n.page.inner_html('#content'))
    h.close(); n.close()

    # F2 제출자 본인(김기안) — 버튼 이름만 '재상신/폐기', 가는 곳은 같다
    h, n = open_home('HEAD', FIX['homeKim']), open_home('NEW', FIX['homeKim'])
    bh = h.page.locator('[data-insp-resub="tok-i2"]').inner_text()
    bn = n.page.locator('[data-insp-resub="tok-i2"]').inner_text()
    check('F2 반려 회차 버튼: HEAD "수정 재상신" → 제출자 본인에게 "재상신/폐기"', bh == '수정 재상신' and bn == '재상신/폐기', [bh, bn])
    for p in (h, n):
        p.page.click('[data-insp-resub="tok-i2"]')
        p.page.wait_for_url('**/exec?**')
    check('F2 버튼 이동 URL HEAD와 동일 (같은 재상신 화면)', h.page.url == n.page.url and 'resub=tok-i2' in n.page.url, [h.page.url, n.page.url])
    h.close(); n.close()

    # F3 최종 검수(-03) 폐기 후
    n = open_home('NEW', FIX['homeKimAfterDiscard'])
    r3 = n.page.locator('.insp-sub-table tr', has_text='TH-AP-26-006-03')
    check('F3 폐기 회차: 행은 남고(번호 소모 표시) 흐리게·클릭 불가·"폐기" 배지',
          r3.count() == 1 and r3.get_attribute('data-insp-view') is None and 'opacity' in (r3.get_attribute('style') or '')
          and r3.locator('span.status').inner_text() == '폐기', r3.inner_html() if r3.count() else 'none')
    r3.click()
    n.page.wait_for_timeout(200)
    check('F3 폐기 행 클릭해도 이동 없음', n.page.url == BASE + '/page', n.page.url)
    po = n.page.locator('tr[data-toggle-insp="tok-prc1"]')
    cells = po.locator('td').all_inner_texts()
    check('F3 검수 건수에서 폐기 제외(3→2건) · 집계 "검수 반려 있음"', cells[-1] == '2건' and cells[-2] == '검수 반려 있음', cells)
    foot = n.page.locator('.insp-sub-box').inner_text()
    check('F3 잠금 해제 → "+ 검수보고서 제출 (04회차)" 버튼 (-03 번호 재사용 없음)', '검수보고서 제출 (04회차)' in foot, foot)
    check('F 브라우저 대화상자 0건 · 스크립트 오류 0건', not n.dialogs and not n.errors, [n.dialogs, n.errors])
    n.close()


with sync_playwright() as pw:
    browser = pw.chromium.launch()
    for fn in (admin_discard, insp_form, home):
        try:
            fn(browser)
        except Exception as e:  # 한 화면이 죽어도 나머지는 돌린다 — 실패로 집계
            failed += 1
            failures.append(fn.__name__ + ' 예외')
            print('  ✗ ' + fn.__name__ + ' 예외: ' + repr(e)[:800])
    browser.close()

print('\n─────────────────────────────')
print('통과 %d / 실패 %d' % (passed, failed))
if failed:
    print('실패 항목:\n  - ' + '\n  - '.join(failures))
sys.exit(1 if failed else 0)
