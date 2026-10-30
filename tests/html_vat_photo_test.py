"""
화면(HTML) 검증 — 단가 기준(부가세 별도/포함) 버튼 · 검수 사진 합계 한도
(memory: html-ui-playwright-harness, vat-toggle-photo-limit-plan)
방식: 이전 코드(BASE_REF, 기본 ef65788)와 작업 트리의 같은 화면을 같은 서버 응답으로 띄워
      ① '부가세 별도' 흐름은 화면 금액과 서버로 가는 값이 이전과 같은지(새로 붙는 vatMode·currency 키만 예외)
      ② '부가세 포함'·외화·사진 기능은 새 동작을 따로 단언한다. 브라우저 기본 대화상자·스크립트 오류는 0건이어야 한다.
픽스처: 서버 하네스가 실제 .gs 함수로 만든 검수보고서 작성 화면 응답 (inspFormKimNew)
  DUMP_FIXTURES=<경로> node tests/audit_insp_discard_test.js
실행:
  insp-renderer/.venv/Scripts/python.exe tests/html_vat_photo_test.py <fixtures.json>
"""
import base64
import json
import os
import pathlib
import re
import subprocess
import sys
from playwright.sync_api import sync_playwright

SRC = pathlib.Path(__file__).resolve().parent.parent
FIX = json.loads(pathlib.Path(sys.argv[1]).read_text(encoding='utf-8'))
BASE_REF = os.environ.get('BASE_REF', 'ef65788')
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


def without(d, *keys):
    return {k: v for k, v in d.items() if k not in keys}


KRW_ITEMS = [['지그 본체', 'SUS304 / EA', 3, '15,000'], ['볼트 세트', 'M8x20 / SET', 10, '1,100']]   # 합 56,000
CLOSE_PROCESSING = "var b = document.getElementById('pm-backdrop'); if (b) b.classList.remove('show');"
APPR = [{'label': '기안자', 'name': '김기안', 'email': 'kim@inlct.com'}, {'label': '결재자 1', 'name': '이결재', 'email': 'lee@inlct.com'}]


# ─────────────────────────────────────────────────────────────
def req_form(browser):
    print('\n[F] 품의 작성 화면 (Procurement_Form.html)')

    def fill(p, items, cur='KRW'):
        p.ev("""([items, cur]) => {
          document.getElementById('items-body').innerHTML = ''; rowCount = 0;
          items.forEach(function (it) {
            addRow(cur);
            var tr = document.querySelector('#items-body tr:last-child');
            tr.querySelector('.item-name').value = it[0]; tr.querySelector('.item-spec').value = it[1];
            tr.querySelector('.qty').value = it[2]; tr.querySelector('.price').value = it[3];
            calcRow(tr.querySelector('.price'));
          });
          if (typeof updateHeaderCurrency === 'function') updateHeaderCurrency(cur);
          reNum(); calcTotal();
        }""", [items, cur])

    def totals(p):
        return [p.text('#supply-total'), p.text('#vat-amt'), p.text('#grand-total'), p.text('#vat-label'), p.text('#total-vat')]

    def submit(p):
        n = len(p.calls('submitRequisitionForClient'))
        p.ev("([a]) => submitFormConfirmed('PRQ-2026-501', '김기안', '검사 지그 제작', a)", [APPR])
        p.wait_calls('submitRequisitionForClient', n + 1)
        p.ev(CLOSE_PROCESSING)   # 스텁은 응답하지 않으므로 처리 중 모달을 닫아 다음 조작을 이어 간다
        return p.calls('submitRequisitionForClient')[-1]['args'][0]

    pg = {rev: Page(browser, 'Procurement_Form.html', rev, {'webappUrl': WEBAPP}, {}) for rev in ('HEAD', 'NEW')}
    for p in pg.values():
        p.page.wait_for_selector('#items-body tr')
        p.ev("document.getElementById('issue-date').value = '2026-10-30'")
        fill(p, KRW_ITEMS)
    h, n = pg['HEAD'], pg['NEW']
    check('F1 처음 열면 "부가세 별도" 선택 · 버튼 보임', n.ev("VAT_MODE") == 'excl'
          and n.ev("document.querySelector('#vat-mode button.on').dataset.mode") == 'excl'
          and n.page.is_visible('#vat-mode') and not n.page.is_visible('#price-note'))
    check('F1 별도 금액 표시 이전과 같음 (56,000 / 5,600 / 61,600)', totals(h) == totals(n) and totals(n)[:3] == ['56,000', '5,600', '61,600'],
          [totals(h), totals(n)])
    ph, pn = submit(h), submit(n)
    check('F2 제출 값 이전과 같음 + vatMode:"excl"만 추가', without(pn, 'vatMode') == ph and pn.get('vatMode') == 'excl',
          [ph, pn])

    n.page.click('#vat-mode button[data-mode=incl]')
    check('F3 "부가세 포함" → 공급가액(합계 ÷ 1.1) 50,909 · 부가세 5,091 · 합계 56,000 · 단가 머리글 "부가세 포함"',
          totals(n)[:3] == ['50,909', '5,091', '56,000'] and n.text('#supply-row-label') == '공급가액 (합계 ÷ 1.1)'
          and n.text('#vat-row-label') == '부가세 (10% · 단가에 포함)' and n.page.is_visible('#price-note')
          and n.text('#total-vat') == '56,000 원', [totals(n), n.text('#supply-row-label')])
    pi = submit(n)
    check('F3 제출 vatMode:"incl" · 품목 단가는 입력 그대로', pi['vatMode'] == 'incl' and pi['items'][0]['price'] == 15000, pi)

    fill(n, [['Laser module', 'LM-532 / EA', 2, '1,250.00']], 'USD')
    check('F4 외화로 바꾸면 버튼 숨김 · "외화 · 부가세 없음" · 부가세 해당없음',
          not n.page.is_visible('#vat-mode') and n.page.is_visible('#vat-mode-na') and n.text('#vat-amt') == '-'
          and not n.page.is_visible('#price-note'), totals(n))
    pu = submit(n)
    check('F4 외화 제출은 버튼이 "포함"이어도 vatMode:"excl"', pu['vatMode'] == 'excl' and n.ev('VAT_MODE') == 'incl', pu.get('vatMode'))

    n.ev("doResetAll()")
    check('F5 초기화하면 "부가세 별도"로 돌아감', n.ev('VAT_MODE') == 'excl' and n.page.is_visible('#vat-mode'))
    check('F 브라우저 대화상자 0건 · 스크립트 오류 0건', not n.dialogs and not n.errors and not h.errors, [n.dialogs, n.errors, h.errors])
    for p in pg.values():
        p.close()


# ─────────────────────────────────────────────────────────────
def prc_form(browser):
    print('\n[P] 구매 작성 화면 (Procurement_PRC_Form.html)')

    def prefill(vat_mode=None, cur='KRW'):
        d = {'parentDocNo': 'TH-AP-26-061', 'parentToken': 'tok-req', 'subject': '검사 지그 제작', 'drafter': '김기안',
             'purpose': '라인 검사용', 'vendorName': '(주)태성테크', 'vendorEmail': 'sales@example.com', 'vendorContact': '박담당',
             'vendorPhone': '042-000-0000', 'deliveryDate': '2026-11-20', 'deliveryAddr': '대전광역시 유성구', 'remarks': '',
             'attachments': [], 'parentApprovers': [], 'attachmentsSize': 0, 'maxTotalSize': 31457280,
             'items': [{'name': a, 'spec': b, 'qty': str(c), 'price': d.replace(',', ''), 'currency': cur} for a, b, c, d in KRW_ITEMS]}
        if vat_mode is not None:
            d['vatMode'] = vat_mode
        return {'getPrefillDataForClient': {'mode': 'ok', 'value': {'ok': True, 'data': d}}}

    def open_(rev, resp):
        p = Page(browser, 'Procurement_PRC_Form.html', rev, {'webappUrl': WEBAPP, 'parentToken': 'tok-req'}, resp)
        p.page.wait_for_selector('#items-body tr')
        p.page.wait_for_timeout(200)
        return p

    def totals(p):
        return [p.text('#supply-total'), p.text('#vat-amt'), p.text('#total-vat'), p.text('#total-amt')]

    def submit(p):
        n = len(p.calls('submitPrcForClient'))
        p.ev("([a]) => submitFormConfirmed('TO-PO-26-910', '검사 지그 제작', a)", [APPR])
        p.wait_calls('submitPrcForClient', n + 1)
        p.ev(CLOSE_PROCESSING)
        return p.calls('submitPrcForClient')[-1]['args'][0]

    # 서버가 vatMode를 안 주거나(옛 서버) 'excl' → 이전과 같은 화면·값
    h = open_('HEAD', prefill())
    for label, resp in (('옛 서버(vatMode 없음)', prefill()), ('excl', prefill('excl'))):
        n = open_('NEW', resp)
        check(f'P1 {label}: "부가세 별도"로 열림 · 금액 이전과 같음 (61,600)', n.ev('VAT_MODE') == 'excl' and totals(n) == totals(h)
              and totals(n)[2] == '61,600 원' and not n.page.is_visible('#vat-changed'), [totals(h), totals(n)])
        ph, pn = submit(h), submit(n)
        check(f'P1 {label}: 제출 값 이전과 같음 + vatMode:"excl"만 추가 · 변경 이력에 단가 기준 없음',
              without(pn, 'vatMode') == ph and pn['vatMode'] == 'excl'
              and not any(c['field'] == 'vatMode' for c in pn['fieldChanges']), [ph.get('fieldChanges'), pn.get('fieldChanges')])
        n.close()
    h.close()

    n = open_('NEW', prefill('incl'))
    check('P2 품의서가 "부가세 포함" → 포함으로 열림 · 합계 56,000 · 공급가액 50,909',
          n.ev('VAT_MODE') == 'incl' and totals(n)[:3] == ['50,909 원', '5,091 원', '56,000 원'] and not n.page.is_visible('#vat-changed'), totals(n))
    p2 = submit(n)
    check('P2 제출 vatMode:"incl" · 변경 이력 없음', p2['vatMode'] == 'incl' and not any(c['field'] == 'vatMode' for c in p2['fieldChanges']), p2['fieldChanges'])
    n.page.click('#vat-mode button[data-mode=excl]')
    check('P3 구매팀이 "별도"로 바꾸면 주황 표시 + "품의서: 부가세 포함" · 합계 61,600',
          n.page.is_visible('#vat-changed') and n.text('#vat-changed') == '품의서: 부가세 포함'
          and 'changed' in n.ev("document.getElementById('vat-mode').className") and totals(n)[2] == '61,600 원')
    p3 = submit(n)
    vc = [c for c in p3['fieldChanges'] if c['field'] == 'vatMode']
    check('P3 제출 vatMode:"excl" + 변경 이력 "단가 기준: 부가세 포함 → 부가세 별도"',
          p3['vatMode'] == 'excl' and vc == [{'field': 'vatMode', 'label': '단가 기준', 'from': '부가세 포함', 'to': '부가세 별도'}], p3['fieldChanges'])
    n.page.click('#vat-mode button[data-mode=incl]')
    p4 = submit(n)
    check('P3 다시 "포함"으로 되돌리면 표시·변경 이력 사라짐', not n.page.is_visible('#vat-changed')
          and not any(c['field'] == 'vatMode' for c in p4['fieldChanges']), p4['fieldChanges'])
    check('P 대화상자 0건 · 스크립트 오류 0건', not n.dialogs and not n.errors, [n.dialogs, n.errors])
    n.close()

    u = open_('NEW', prefill('excl', 'USD'))
    check('P4 외화 품의: 버튼 숨김 · "외화 · 부가세 없음"', not u.page.is_visible('#vat-mode') and u.page.is_visible('#vat-mode-na'))
    check('P4 외화 제출 vatMode:"excl"', submit(u)['vatMode'] == 'excl')
    u.close()


# ─────────────────────────────────────────────────────────────
def reject_form(browser):
    print('\n[R] 반려 후 재상신 화면 (Procurement_Reject.html)')

    def resp(cur='KRW', vat_mode=None):
        items = [{'name': a, 'spec': b, 'qty': str(c), 'price': d.replace(',', ''), 'currency': cur} for a, b, c, d in KRW_ITEMS]
        data = {'docNo': 'TH-AP-26-061', 'issueDate': '2026-10-30', 'drafter': '김기안', 'dept': '개발팀', 'subject': '검사 지그 제작',
                'purpose': '라인', 'vendorName': '(주)태성테크', 'vendorEmail': 's@example.com', 'vendorContact': '박', 'vendorPhone': '042',
                'deliveryDate': '2026-11-20', 'deliveryAddr': '대전', 'remarks': '', 'rejectHistory': '', 'rejectApprover': '이결재',
                'rejectTime': '2026-10-30 09:00', 'rejectReason': '단가 확인', 'items': items}
        if vat_mode is not None:
            data['vatMode'] = vat_mode
        return {'getRejectDataForClient': {'mode': 'ok', 'value': {'ok': True, 'data': data}},
                'getRequisitionForViewer': {'mode': 'ok', 'value': {'ok': True, 'data': {'approvers': APPR}}}}

    def open_(rev, r):
        p = Page(browser, 'Procurement_Reject.html', rev, {'webappUrl': WEBAPP, 'token': 'tok-req', 'docNo': 'TH-AP-26-061'}, r)
        p.page.wait_for_selector('#items-body tr')
        p.wait_calls('getRequisitionForViewer')
        p.page.wait_for_timeout(150)
        return p

    def submit(p):
        n = len(p.calls('resubmitRequisitionForClient'))
        p.ev('proceedResubmit()')
        p.wait_calls('resubmitRequisitionForClient', n + 1)
        return p.calls('resubmitRequisitionForClient')[-1]['args'][0]

    h, n = open_('HEAD', resp()), open_('NEW', resp())
    check('R1 원화 별도: 총 금액 이전과 같음 (61,600 원)', h.text('#total-vat') == n.text('#total-vat') == '61,600 원', [h.text('#total-vat'), n.text('#total-vat')])
    check('R1 합계 3줄 (공급가액 56,000 · 부가세 5,600 · 합계 61,600)',
          [n.text('#supply-total'), n.text('#vat-amt'), n.text('#grand-total')] == ['56,000', '5,600', '61,600'])
    ph, pn = submit(h), submit(n)
    items_nocur = [without(i, 'currency') for i in pn['data']['items']]
    check('R2 재상신 값 이전과 같음 + vatMode:"excl" · 품목마다 currency:"KRW"만 추가',
          without(pn['data'], 'vatMode', 'items') == without(ph['data'], 'items') and items_nocur == ph['data']['items']
          and pn['data']['vatMode'] == 'excl' and all(i['currency'] == 'KRW' for i in pn['data']['items']), [ph, pn])
    h.close(); n.close()

    u = open_('NEW', resp('USD'))
    pu = submit(u)
    check('R3 [결함 수정] 외화 품의 재상신 — 통화 USD 유지 (이전엔 통화가 빠져 원화로 저장됨)',
          all(i['currency'] == 'USD' for i in pu['data']['items']) and pu['data']['vatMode'] == 'excl', pu['data']['items'])
    check('R3 외화: 버튼 숨김 · "총 금액 (부가세 없음)" 61,…→ 56,000 USD · 머리글 (USD)',
          not u.page.is_visible('#vat-mode') and u.text('#total-label') == '총 금액 (부가세 없음)'
          and u.text('#total-vat') == '56,000 USD' and '(USD)' in u.text('table.items thead'), [u.text('#total-label'), u.text('#total-vat')])
    hu = open_('HEAD', resp('USD'))
    phu = submit(hu)
    check('R3 (비교) 이전 화면은 통화를 빼고 보냈다', all('currency' not in i for i in phu['data']['items']))
    hu.close(); u.close()

    i = open_('NEW', resp('KRW', 'incl'))
    check('R4 품의서가 "포함" → 포함으로 열림 · 56,000 원 · 공급가액 50,909',
          i.ev('VAT_MODE') == 'incl' and i.text('#total-vat') == '56,000 원' and i.text('#supply-total') == '50,909')
    check('R4 재상신 vatMode:"incl"', submit(i)['data']['vatMode'] == 'incl')
    check('R 대화상자 0건 · 스크립트 오류 0건', not i.dialogs and not i.errors, [i.dialogs, i.errors])
    i.close()


# ─────────────────────────────────────────────────────────────
def viewer(browser):
    print('\n[V] 결재자 상세 화면 (Procurement_Viewer.html)')

    def data(vat_mode):
        d = {'docNo': 'TH-AP-26-061', 'docType': 'REQ', 'status': '결재중 (결재자 1)', 'drafter': '김기안', 'dept': '개발팀',
             'issueDate': '2026-10-30', 'subject': '검사 지그 제작', 'purpose': '라인', 'vendorName': '(주)태성테크',
             'vendorEmail': 's@example.com', 'vendorContact': '박', 'vendorPhone': '042', 'deliveryDate': '2026-11-20',
             'deliveryAddr': '대전', 'items': [{'name': a, 'spec': b, 'qty': str(c), 'price': d.replace(',', ''), 'currency': 'KRW'} for a, b, c, d in KRW_ITEMS],
             'approvers': [{'label': '기안자', 'name': '김기안', 'email': 'kim@inlct.com', 'status': '승인', 'processedAt': '2026-10-30 09:12'},
                           {'label': '결재자 1', 'name': '이결재', 'email': 'lee@inlct.com', 'status': '대기', 'processedAt': ''}],
             'attachments': [], 'parentAttachments': [], 'parentApprovers': [], 'myApproverIdx': 1}
        if vat_mode is not None:
            d['vatMode'] = vat_mode
        return {'getRequisitionForViewer': {'mode': 'ok', 'value': {'ok': True, 'data': d}}}

    def open_(rev, r):
        p = Page(browser, 'Procurement_Viewer.html', rev, {'webappUrl': WEBAPP, 'token': 'tok-req', 'decision': '', 'approverIndex': '1'}, r)
        p.page.wait_for_selector('table.items')
        return p

    sec = "document.querySelector('table.items').closest('.section').outerHTML"
    h = open_('HEAD', data(None))
    for label, mode in (('옛 서버(vatMode 없음)', None), ('excl', 'excl')):
        n = open_('NEW', data(mode))
        check(f'V1 {label}: 품목 내역 영역 HTML이 이전과 똑같음', h.ev(sec) == n.ev(sec), [h.ev(sec)[-300:], n.ev(sec)[-300:]])
        n.close()
    h.close()
    n = open_('NEW', data('incl'))
    body = n.ev(sec)
    check('V2 포함 건: "단가 = 부가세 포함" 표시 · 공급가액(합계 ÷ 1.1) 50,909 · 부가세 5,091 · 합계 56,000',
          '단가 = 부가세 포함' in body and '공급가액 (합계 ÷ 1.1)' in body and '>50,909<' in body and '>5,091<' in body
          and '>56,000<' in body and '부가세 (10% · 단가에 포함)' in body, body[-600:])
    check('V 스크립트 오류 0건', not n.errors, n.errors)
    n.close()


# ─────────────────────────────────────────────────────────────
NOISE_JPEG = """([w, h, q, seed]) => {
  const c = document.createElement('canvas'); c.width = w; c.height = h;
  const ctx = c.getContext('2d'); const img = ctx.createImageData(w, h);
  let s = seed >>> 0;
  for (let i = 0; i < img.data.length; i += 4) {
    s = (s * 1103515245 + 12345) >>> 0;
    const v = (s >>> 16) & 255, g = (i / 4 / w) | 0;
    img.data[i] = (v + g) & 255; img.data[i + 1] = v; img.data[i + 2] = (v + 2 * g) & 255; img.data[i + 3] = 255;
  }
  ctx.putImageData(img, 0, 0);
  ctx.fillStyle = 'red'; ctx.fillRect(0, 0, w / 6, h / 6);   // 순서 확인용 표식
  const u = c.toDataURL('image/jpeg', q); return u.slice(u.indexOf(',') + 1);
}"""


def insp_form(browser):
    print('\n[I] 검수보고서 작성 화면 (Procurement_INSP_Form.html) — 사진')
    resp = {'getInspFormDataForClient': {'mode': 'ok', 'value': FIX['inspFormKimNew']},
            'submitInspForClient': {'mode': 'ok', 'value': {'ok': True, 'message': '제출 완료', 'docNo': 'X', 'token': 't', 'photos': 0}}}

    def open_(rev):
        p = Page(browser, 'Procurement_INSP_Form.html', rev, {'prcToken': 'tok-prc2', 'resubToken': '', 'webappUrl': WEBAPP}, resp)
        p.page.wait_for_selector('#submit-btn')
        return p

    def settle(p):
        p.page.wait_for_function('STATE.photos.every(function (x) { return !x.pending; })', timeout=60000)
        p.page.wait_for_timeout(300)

    def submit(p):
        p.page.click('.verdict.pass')
        p.page.fill('#f-note', '지그 본체 2EA')
        p.ev("STATE.appr[0] = {name:'이결재', dept:'경영팀', rank:'', email:'lee@inlct.com'}; renderApprTable();")
        n = len(p.calls('submitInspForClient'))
        p.page.click('#submit-btn')
        p.page.click('#fin-no')
        p.wait_calls('submitInspForClient', n + 1)
        return p.calls('submitInspForClient')[-1]['args'][0]

    # 실험용 사진 — 순서를 알 수 있게 크기를 다르게(큰 것이 먼저 → 압축이 늦게 끝남)
    gen = open_('NEW')
    big = [base64.b64decode(gen.ev(NOISE_JPEG, [4032, 3024, 1.0, 11])),   # 휴대폰 12MP 최고화질 — 5MB 넘음
           base64.b64decode(gen.ev(NOISE_JPEG, [3000, 4000, 0.97, 22])),
           base64.b64decode(gen.ev(NOISE_JPEG, [800, 600, 0.8, 33]))]       # 작은 사진
    gen.close()
    sizes = [len(b) / 1048576 for b in big]
    files = [{'name': f'p{i + 1}.jpg', 'mimeType': 'image/jpeg', 'buffer': b} for i, b in enumerate(big)]
    print('    (실험 사진 원본 크기: ' + ', '.join(f'{s:.1f}MB' for s in sizes) + ')')

    n = open_('NEW')
    n.page.set_input_files('#f-photos', files)
    pend = n.ev('STATE.photos.map(function (x) { return x.pending; })')
    check('I1 고른 즉시 3자리가 고른 순서대로 잡힘(처리 중 표시)', len(pend) == 3 and '사진 줄이는 중' in n.text('#pg'), pend)
    settle(n)
    names = n.ev('STATE.photos.map(function (x) { return x.name; })')
    check(f'I1 원본 {sizes[0]:.1f}MB·{sizes[1]:.1f}MB 사진도 첨부됨 (장당 5MB 검사 없음) · 경고 없음',
          sizes[0] > 5 and len(names) == 3 and not n.ev("document.getElementById('overlay').classList.contains('on')"), names)
    check('I2 사진 순서 = 고른 순서 (p1, p2, p3)', names == ['p1.jpg', 'p2.jpg', 'p3.jpg'], names)
    payload = submit(n)
    ph = payload['photos']
    check('I3 제출 사진 3장 · 고른 순서 · 처리 중 표시 값(pending) 안 섞임 · JPEG',
          [x['name'] for x in ph] == ['p1.jpg', 'p2.jpg', 'p3.jpg'] and all(set(x) == {'name', 'size', 'type', 'data'} for x in ph)
          and all(base64.b64decode(x['data'])[:2] == b'\xff\xd8' for x in ph), [(x['name'], x['size']) for x in ph])
    check('I3 줄인 사본 합계가 한도 안 (장당 2MB 미만)', all(x['size'] < 2 * 1048576 for x in ph), [x['size'] for x in ph])
    check('I 대화상자 0건 · 스크립트 오류 0건', not n.dialogs and not n.errors, [n.dialogs, n.errors])
    n.close()

    h = open_('HEAD')
    h.page.set_input_files('#f-photos', files)
    h.page.wait_for_timeout(2500)
    hw = h.ev("document.getElementById('overlay').classList.contains('on') ? document.getElementById('ov-card').innerText : ''")
    check('I1 (비교) 이전 화면은 5MB 넘는 원본을 거부했다', '장당 5MB 초과' in hw and h.ev('STATE.photos.length') == 1, hw)
    h.close()

    # HEIC — 예전엔 조용히 사라짐
    n = open_('NEW')
    n.page.set_input_files('#f-photos', files=[{'name': 'IMG_0001.HEIC', 'mimeType': 'image/heic', 'buffer': b'\x00\x00\x00\x18ftypheic' + b'\x00' * 4096}])
    n.page.wait_for_function("document.getElementById('overlay').classList.contains('on')", timeout=10000)
    msg = n.text('#ov-card')
    check('I4 [결함 수정] HEIC → 경고 모달 "사진을 열 수 없습니다 … JPG로" · 첨부 안 됨',
          'IMG_0001.HEIC' in msg and '사진을 열 수 없습니다' in msg and 'JPG' in msg and n.ev('STATE.photos.length') == 0, msg)
    n.close()
    h = open_('HEAD')
    h.page.set_input_files('#f-photos', files=[{'name': 'IMG_0001.HEIC', 'mimeType': 'image/heic', 'buffer': b'\x00\x00\x00\x18ftypheic' + b'\x00' * 4096}])
    h.page.wait_for_timeout(2500)
    check('I4 (비교) 이전 화면은 HEIC를 첨부도 경고도 없이 버렸다',
          h.ev('STATE.photos.length') == 0 and not h.ev("document.getElementById('overlay').classList.contains('on')"))
    h.close()

    # 장수 — 7장 고르면 6장만, 경고 1장
    n = open_('NEW')
    small = files[2]['buffer']
    n.page.set_input_files('#f-photos', [{'name': f's{i}.jpg', 'mimeType': 'image/jpeg', 'buffer': small} for i in range(1, 8)])
    settle(n)
    msg = n.text('#ov-card') if n.ev("document.getElementById('overlay').classList.contains('on')") else ''
    check('I5 7장 → 6장만 첨부 + "최대 6장" 경고', n.ev('STATE.photos.length') == 6 and '최대 6장' in msg, msg)
    n.close()

    # 합계 — 한도를 낮춰 경로 확인 (실제 사진 6장으로는 15MB에 닿지 않음)
    n = open_('NEW')
    n.ev('PHOTO_MAX_TOTAL = 400 * 1024')
    n.page.set_input_files('#f-photos', files)
    settle(n)
    msg = n.text('#ov-card') if n.ev("document.getElementById('overlay').classList.contains('on')") else ''
    total = n.ev('photoBytes()')
    check('I6 합계 한도 넘는 사진은 빼고 경고 · 남은 합계는 한도 안', '합계가' in msg and total <= 400 * 1024 and n.ev('STATE.photos.length') < 3,
          [msg, total])
    n.close()

    # 처리 중에 제출 누르면 막음
    n = open_('NEW')
    n.page.click('.verdict.pass'); n.page.fill('#f-note', 'x')
    n.ev("STATE.appr[0] = {name:'이결재', dept:'경영팀', rank:'', email:'lee@inlct.com'}; renderApprTable();")
    n.ev("STATE.photos.push({name:'w.jpg', size:0, dataUrl:'', pending:true}); drawPhotos();")
    n.page.click('#submit-btn')
    msg = n.text('#ov-card')
    check('I7 사진 처리 중 제출 → "사진을 줄이는 중" 안내, 서버 호출 없음', '사진을 줄이는 중' in msg and not n.calls('submitInspForClient'), msg)
    n.close()

    # 작은 사진만 — 이전과 같은 결과 (안내 문구 외 동작 변화 없음)
    pay = {}
    for rev in ('HEAD', 'NEW'):
        p = open_(rev)
        p.page.set_input_files('#f-photos', [files[2]])
        p.page.wait_for_function('STATE.photos.length === 1 && !STATE.photos[0].pending && STATE.photos[0].dataUrl', timeout=20000)
        pay[rev] = submit(p)
        p.close()
    check('I8 5MB 미만 사진 1장 제출 값 이전과 같음 (바이트까지)', pay['HEAD'] == pay['NEW'],
          [without(pay['HEAD'], 'photos'), without(pay['NEW'], 'photos')])


with sync_playwright() as pw:
    browser = pw.chromium.launch()
    for fn in (req_form, prc_form, reject_form, viewer, insp_form):
        try:
            fn(browser)
        except Exception as e:
            failed += 1
            failures.append(fn.__name__ + ' 예외')
            print('  ✗ ' + fn.__name__ + ' 예외: ' + repr(e)[:900])
    browser.close()

print('\n─────────────────────────────')
print('통과 %d / 실패 %d   (기준 %s)' % (passed, failed, BASE_REF))
if failed:
    print('실패 항목:\n  - ' + '\n  - '.join(failures))
    sys.exit(1)
