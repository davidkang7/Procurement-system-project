"""
화면(HTML) 검증 — 검수보고서 품목별 입고 (INSP v1.1, 2026-11-05)
(memory: html-ui-playwright-harness)

방식: 서버 하네스(tests/insp_receipt_test.js)가 실제 .gs 함수로 만든 응답(픽스처)으로 작성·뷰어 화면을 띄워
      표 렌더 → 체크/수량/사유 입력 → 최종 확인 모달 → 서버로 가는 페이로드를 확인한다.
      브라우저 기본 대화상자(alert/confirm)는 0건, 스크립트 오류 0건이어야 한다.

픽스처:  DUMP_FIXTURES=<경로> node tests/insp_receipt_test.js
실행:    insp-renderer/.venv/Scripts/python.exe tests/html_insp_receipt_test.py <fixtures.json>
"""
import json
import pathlib
import re
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
        print('  ✗ ' + name + ('  → ' + str(detail)[:700] if detail != '' else ''))


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


def build(name, tvars, resp):
    html = (SRC / name).read_text(encoding='utf-8')
    html = re.sub(r'<\?!?=\s*(\w+)\s*\?>', lambda m: str(tvars.get(m.group(1), '')), html)
    return html.replace('<head>', '<head>' + (STUB % json.dumps(resp, ensure_ascii=False)), 1)


class Page:
    def __init__(self, browser, name, tvars, resp):
        self.ctx = browser.new_context(viewport={'width': 1400, 'height': 1000})
        self.page = self.ctx.new_page()
        self.page.set_default_timeout(8000)
        self.dialogs, self.errors = [], []
        self.page.on('dialog', lambda d: (self.dialogs.append(d.type + ':' + d.message), d.dismiss()))
        self.page.on('pageerror', lambda e: self.errors.append(str(e)))
        html = build(name, tvars, resp)

        def handler(route):
            if route.request.url.startswith(BASE + '/page'):
                route.fulfill(status=200, content_type='text/html; charset=utf-8', body=html)
            else:
                route.fulfill(status=200, content_type='text/html; charset=utf-8', body='<html><body>NAV</body></html>')
        self.page.route(BASE + '/**', handler)
        self.page.goto(BASE + '/page')

    def calls(self, fn=None):
        return [x for x in self.page.evaluate('window.__calls') if fn is None or x['fn'] == fn]

    def wait_calls(self, fn, n=1):
        self.page.wait_for_function('([f, n]) => window.__calls.filter(c => c.fn === f).length >= n', arg=[fn, n])

    def card(self):
        return self.page.inner_text('#ov-card')

    def clean(self):
        return not self.dialogs and not self.errors

    def close(self):
        self.ctx.close()


OK_SUBMIT = {'mode': 'ok', 'value': {'ok': True, 'message': '제출 완료', 'docNo': 'TH-AP-26-009-01', 'token': 't', 'photos': 0}}


def form(browser, fixture, resub=''):
    p = Page(browser, 'Procurement_INSP_Form.html', {'prcToken': 'tok-prc2', 'resubToken': resub, 'webappUrl': WEBAPP},
             {'getInspFormDataForClient': {'mode': 'ok', 'value': FIX[fixture]}, 'submitInspForClient': OK_SUBMIT})
    p.page.wait_for_selector('#submit-btn')
    return p


def pick_approver(p):
    p.page.evaluate("STATE.appr[0] = {name:'이결재', dept:'경영팀', rank:'', email:'lee@inlct.com'}; renderApprTable();")


def payload(p):
    return p.calls('submitInspForClient')[-1]['args'][0]


def table_and_prefill(browser):
    print('\n[T] 입고 품목 표 — 렌더 · 전체 선택 · 잔량 prefill')
    p = form(browser, 'formNew2')
    check('T1 입고 내역 한 줄 입력(#f-note) 없음 · 입고 품목 표 있음', p.page.locator('#f-note').count() == 0 and p.page.locator('#rc-table').count() == 1)
    rows = p.page.locator('#rc-body tr')
    cells0 = rows.nth(0).locator('td').all_inner_texts()
    check('T2 행 2개 · 1행: No 1 / Cly / 발주 106 / 기입고 0 / 잔량 106 / 상태 대기',
          rows.count() == 2 and cells0[1] == '1' and cells0[2] == 'Cly(F4PO1-003-A1)' and cells0[4] == '106' and cells0[5] == '0' and cells0[6] == '106' and cells0[9] == '대기', cells0)
    check('T3 체크 전: 수량 입력 비활성 · 사유 칸 비활성',
          p.page.is_disabled('#rc-qty-0') and p.page.is_disabled('#rc-rs-0'))
    p.page.check('#rc-chk-0')
    check('T4 체크 → 이번 입고 = 잔량 106 자동 입력 · 상태 완료 · 합계 "입고 1개 품목 · 완료 1/2"',
          p.page.input_value('#rc-qty-0') == '106' and p.page.inner_text('#rc-st-0') == '완료' and '1개 품목' in p.page.inner_text('#rc-sum') and '1</b>/2' in p.page.inner_html('#rc-sum'),
          [p.page.input_value('#rc-qty-0'), p.page.inner_text('#rc-sum')])
    p.page.fill('#rc-qty-0', '40')
    check('T5 수량 40으로 수정 → 상태 부분 · 사유 칸은 여전히 비활성', p.page.inner_text('#rc-st-0') == '부분' and p.page.is_disabled('#rc-rs-0'))
    p.page.check('#rc-all')
    check('T6 전체 선택 → 두 행 모두 잔량으로 (106, 2)', p.page.input_value('#rc-qty-0') == '106' and p.page.input_value('#rc-qty-1') == '2')
    p.page.uncheck('#rc-all')
    check('T7 전체 해제 → 0 · 비활성', p.page.input_value('#rc-qty-0') == '0' and p.page.is_disabled('#rc-qty-1'))
    check('T 대화상자·오류 0건', p.clean(), [p.dialogs, p.errors])
    p.close()


def over_receipt(browser):
    print('\n[O] 초과 입고 — 사유 칸 즉시 노출 · 제출 게이트')
    p = form(browser, 'formPartial')     # Cly 기입고 50 / 잔량 56
    cells0 = p.page.locator('#rc-body tr').nth(0).locator('td').all_inner_texts()
    check('O1 기입고 50 · 잔량 56 표시', cells0[5] == '50' and cells0[6] == '56', cells0)
    p.page.check('#rc-chk-0')
    p.page.fill('#rc-qty-0', '60')
    check('O2 50+60 > 106 → 상태 초과 · 사유 칸 활성 + 빨간 표시(need)',
          p.page.inner_text('#rc-st-0') == '초과' and not p.page.is_disabled('#rc-rs-0') and 'need' in (p.page.get_attribute('#rc-rs-0', 'class') or ''))
    p.page.click('.verdict.pass'); pick_approver(p)
    p.page.click('#submit-btn')
    p.page.wait_for_selector('#overlay.on')
    check('O3 사유 없이 제출 → 오류 모달(초과 입고 사유) · 서버 호출 0',
          '초과 입고 품목의 수량 상이 사유' in p.card() and 'Cly' in p.card() and not p.calls('submitInspForClient'), p.card())
    p.page.click('#ov-ok')
    p.page.fill('#rc-rs-0', '업체 예비 4EA 추가 납품')
    check('O4 사유 입력 → need 표시 해제', 'need' not in (p.page.get_attribute('#rc-rs-0', 'class') or ''))
    p.page.click('#submit-btn')
    p.page.wait_for_selector('#fin-yes')
    card = p.card()
    check('O5 최종 확인 모달에 입고 현황: 초과 — Cly 누계 110/106 (사유) · 미입고 — 지그 잔량 2',
          '입고 현황' in card and '초과 — 1. Cly' in card and '110/106' in card and '업체 예비 4EA' in card and '미입고 — 2. 지그' in card and '잔량 2' in card, card)
    p.page.click('#fin-no')
    p.wait_calls('submitInspForClient', 1)
    a = payload(p)
    check('O6 중간 회차 제출 페이로드: receivedNote 없음 · receiptItems [{0,60,사유},{1,0,""}] · isFinal false',
          'receivedNote' not in a and a['receiptItems'] == [{'idx': 0, 'qty': 60, 'reason': '업체 예비 4EA 추가 납품'}, {'idx': 1, 'qty': 0, 'reason': ''}] and a['isFinal'] is False, a.get('receiptItems'))
    check('O 대화상자·오류 0건', p.clean(), [p.dialogs, p.errors])
    p.close()


def final_shortfall(browser):
    print('\n[F] 최종 검수 — 미입고 사유 2단계 · 종결 전용 회차 · 전 품목 완료 되묻기')
    p = form(browser, 'formNew2')
    p.page.click('.verdict.pass'); pick_approver(p)
    p.page.check('#rc-chk-0')                       # Cly 106 전량, 지그 0
    p.page.check('#f-final')
    check('F1 최종 체크 → 잔량 남은 지그 행에 미입고 사유 칸 활성(need) · Cly는 완료',
          not p.page.is_disabled('#rc-rs-1') and 'need' in (p.page.get_attribute('#rc-rs-1', 'class') or '') and p.page.inner_text('#rc-st-1') == '미입고' and p.page.is_disabled('#rc-rs-0'))
    p.page.click('#submit-btn')
    p.page.wait_for_selector('#fin-yes')
    p.page.click('#fin-yes')
    p.page.wait_for_selector('#rs-ok')
    check('F2 사유 없이 "최종" → 같은 창 2단계: 지그 잔량 2 사유 입력, 제출 버튼 비활성 · 서버 호출 0',
          '종결 사유' in p.card() and '잔량 2' in p.card() and p.page.is_disabled('#rs-ok') and not p.calls('submitInspForClient'), p.card())
    p.page.fill('#rs-in-1', '생산 취소')
    check('F3 사유 입력 → 제출 버튼 활성', not p.page.is_disabled('#rs-ok'))
    p.page.click('#rs-ok')
    p.wait_calls('submitInspForClient', 1)
    a = payload(p)
    check('F4 페이로드: isFinal true · 지그 사유 "생산 취소" 반영 · 표의 사유 칸에도 동기화',
          a['isFinal'] is True and a['receiptItems'] == [{'idx': 0, 'qty': 106, 'reason': ''}, {'idx': 1, 'qty': 0, 'reason': '생산 취소'}] and p.page.input_value('#rc-rs-1') == '생산 취소', a.get('receiptItems'))
    check('F 대화상자·오류 0건', p.clean(), [p.dialogs, p.errors])
    p.close()

    # 종결 전용 회차 (입고 0) — 중간은 막고, 최종+사유는 통과
    p = form(browser, 'formNew2')
    p.page.click('.verdict.pass'); pick_approver(p)
    p.page.click('#submit-btn')
    p.page.wait_for_selector('#fin-no')
    check('F5 입고 0 → 현황에 "종결 전용 회차" 안내', '종결 전용 회차' in p.card(), p.card())
    p.page.click('#fin-no')
    p.page.wait_for_selector('#ov-ok')
    check('F6 입고 0 + 중간 회차 → 오류 모달(중간 회차로 제출할 수 없습니다) · 서버 호출 0',
          '중간 회차로 제출할 수 없습니다' in p.card() and not p.calls('submitInspForClient'), p.card())
    p.page.click('#ov-ok')
    p.page.click('#submit-btn'); p.page.wait_for_selector('#fin-yes'); p.page.click('#fin-yes')
    p.page.wait_for_selector('#rs-ok')
    p.page.fill('#rs-in-0', '주문 취소'); p.page.fill('#rs-in-1', '주문 취소')
    p.page.click('#rs-ok')
    p.wait_calls('submitInspForClient', 1)
    a = payload(p)
    check('F7 입고 0 + 최종 + 두 품목 사유 → 제출 (qty 0, 사유 2건)', a['isFinal'] is True and [x['qty'] for x in a['receiptItems']] == [0, 0] and all(x['reason'] == '주문 취소' for x in a['receiptItems']), a.get('receiptItems'))
    p.close()

    # 전 품목 완료인데 '중간' → 한 번 되묻고 진행
    p = form(browser, 'formNew2')
    p.page.click('.verdict.pass'); pick_approver(p)
    p.page.check('#rc-all')
    p.page.click('#submit-btn'); p.page.wait_for_selector('#fin-no')
    check('F8 전 품목 완료 → 현황이 녹색 안내 "최종 검수를 권장"', '최종 검수를 권장' in p.card() and p.page.locator('#fin-status.ov-ok').count() == 1, p.card())
    p.page.click('#fin-no')
    p.page.wait_for_selector('#ad-ok')
    check('F9 중간 선택 → 되묻는 카드, 아직 서버 호출 0', '중간 회차로 제출합니다' in p.card() and not p.calls('submitInspForClient'))
    p.page.click('#ad-ok')
    p.wait_calls('submitInspForClient', 1)
    check('F10 계속 → 중간 회차로 제출 (isFinal false)', payload(p)['isFinal'] is False)
    check('F 대화상자·오류 0건 (2)', p.clean(), [p.dialogs, p.errors])
    p.close()


def legacy_and_resub(browser):
    print('\n[L] 레거시 배너 · 재상신 prefill')
    p = form(browser, 'formLegacy')
    banner = p.page.locator('.rc-legacy')
    check('L1 레거시 회차 2건 배너: 회차 번호·입고 내역 원문·"기입고에 반영되지 않았습니다"',
          banner.count() == 1 and '이전 2회차' in banner.inner_text() and 'TH-AP-26-006-01' in banner.inner_text() and '1차 입고' in banner.inner_text() and '반영되지 않았습니다' in banner.inner_text(), banner.inner_text() if banner.count() else 'none')
    p.close()
    p = form(browser, 'formNew2')
    check('L2 레거시 회차 없으면 배너 없음', p.page.locator('.rc-legacy').count() == 0)
    p.close()

    p = form(browser, 'formResub', resub=FIX['formResub']['resub']['token'])
    check('L3 재상신: 반려 회차의 수량 50 prefill · 체크 상태 · 기입고 0(본인 행 제외)',
          p.page.is_checked('#rc-chk-0') and p.page.input_value('#rc-qty-0') == '50' and p.page.locator('#rc-body tr').nth(0).locator('td').all_inner_texts()[5] == '0')
    p.page.click('.verdict.pass'); pick_approver(p)
    p.page.fill('#rc-qty-0', '30')
    p.page.click('#submit-btn'); p.page.wait_for_selector('#fin-no'); p.page.click('#fin-no')
    p.wait_calls('submitInspForClient', 1)
    a = payload(p)
    check('L4 재상신 페이로드: resubToken · receiptItems qty 30', a['resubToken'] == FIX['formResub']['resub']['token'] and a['receiptItems'][0] == {'idx': 0, 'qty': 30, 'reason': ''}, a)
    check('L 대화상자·오류 0건', p.clean(), [p.dialogs, p.errors])
    p.close()


def server_reject(browser):
    print('\n[X] 서버 거절 메시지 — 품목별 한 줄씩')
    p = Page(browser, 'Procurement_INSP_Form.html', {'prcToken': 'tok-prc2', 'resubToken': '', 'webappUrl': WEBAPP},
             {'getInspFormDataForClient': {'mode': 'ok', 'value': FIX['formNew2']},
              'submitInspForClient': {'mode': 'ok', 'value': {'ok': False, 'message': '1. Cly: 누계 110/106 — 초과 입고 사유를 입력해 주세요.\n2. 지그 본체: 잔량 1 — 최종 검수로 종결하려면 미입고 사유를 입력해 주세요.'}}})
    p.page.wait_for_selector('#submit-btn')
    p.page.click('.verdict.pass'); pick_approver(p); p.page.check('#rc-chk-0'); p.page.fill('#rc-qty-0', '5')
    p.page.click('#submit-btn'); p.page.wait_for_selector('#fin-no'); p.page.click('#fin-no')
    p.wait_calls('submitInspForClient', 1)
    p.page.wait_for_function("document.getElementById('ov-card').innerText.indexOf('제출에 실패했습니다') >= 0")
    html = p.page.inner_html('#ov-card')
    check('X1 두 줄 메시지가 <br>로 나뉘어 표시 · 제출 버튼 다시 활성',
          '초과 입고 사유' in html and '미입고 사유' in html and '<br>' in html and not p.page.is_disabled('#submit-btn'), html)
    check('X 대화상자·오류 0건', p.clean(), [p.dialogs, p.errors])
    p.close()


def viewer(browser):
    print('\n[V] 뷰어 — 품목 표 / 레거시 한 줄')
    p = Page(browser, 'Procurement_INSP_Viewer.html', {'token': 't', 'hintIdx': '', 'webappUrl': WEBAPP},
             {'getInspViewerDataForClient': {'mode': 'ok', 'value': FIX['viewerStructured']}})
    p.page.wait_for_selector('table.rct')
    rows = p.page.locator('table.rct tbody tr')
    r0 = rows.nth(0).locator('td').all_inner_texts(); r1 = rows.nth(1).locator('td').all_inner_texts()
    check('V1 품목 표 2행: Cly 발주106/기입고0/이번110/누계110/"+4 초과"/사유 · 지그 1/2 잔량 1 "생산 취소"',
          rows.count() == 2 and r0[3] == '106' and r0[5] == '110' and r0[6] == '110' and r0[7] == '+4 초과' and '예비분' in r0[8]
          and r1[5] == '1' and r1[7] == '1' and r1[8] == '생산 취소', [r0, r1])
    check('V2 수량 상이 사유 블록 표시', '수량 상이 사유' in p.page.inner_text('.rct-note') and '[초과]' in p.page.inner_text('.rct-note'))
    check('V3 입고 내역 요약 행은 별도로 보이지 않음(품목 표가 대신) · 오류 0', p.page.locator('th:has-text("입 고 내 역")').count() == 0 and p.clean(), [p.dialogs, p.errors])
    p.close()
    p = Page(browser, 'Procurement_INSP_Viewer.html', {'token': 'tok-i1', 'hintIdx': '', 'webappUrl': WEBAPP},
             {'getInspViewerDataForClient': {'mode': 'ok', 'value': FIX['viewerLegacy']}})
    p.page.wait_for_selector('table.meta')
    check('V4 레거시 회차: 입고 내역 한 줄 그대로, 품목 표 없음', p.page.locator('th:has-text("입 고 내 역")').count() == 1 and p.page.locator('table.rct').count() == 0 and '1차 입고' in p.page.inner_text('table.meta'))
    check('V 대화상자·오류 0건', p.clean(), [p.dialogs, p.errors])
    p.close()


with sync_playwright() as pw:
    browser = pw.chromium.launch()
    for fn in (table_and_prefill, over_receipt, final_shortfall, legacy_and_resub, server_reject, viewer):
        try:
            fn(browser)
        except Exception as e:
            failed += 1
            failures.append(fn.__name__ + ' 예외')
            print('  ✗ ' + fn.__name__ + ' 예외: ' + repr(e))
    browser.close()

print('\n─────────────────────────────')
print(f'통과 {passed} / 실패 {failed}')
if failed:
    print('실패 항목:\n  - ' + '\n  - '.join(failures))
sys.exit(1 if failed else 0)
