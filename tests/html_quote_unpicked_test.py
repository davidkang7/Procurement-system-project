"""
화면(HTML) 검증 — 제출·재상신 전 견적서 안내 모달 분리 · 관리자 홈 [견적서 아님]
(memory: html-ui-playwright-harness, quote-pdf-feature, ui-modal-not-browser-dialog)
방식: 이전 코드(BASE_REF, 기본 1d8b398)와 작업 트리의 같은 화면을 같은 서버 응답으로 띄워
      ① 모달 문구·버튼은 새 동작을 단언하고
      ② [계속] 으로 넘어갔을 때 서버로 가는 값은 이전과 바이트 단위로 같은지 비교한다(서버 페이로드 불변).
      브라우저 기본 대화상자·스크립트 오류는 0건이어야 한다.
픽스처(관리자 홈): DUMP_FIXTURES=<경로> node tests/quote_unpicked_test.js
실행:
  insp-renderer/.venv/Scripts/python.exe tests/html_quote_unpicked_test.py <fixtures.json>
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
BASE_REF = os.environ.get('BASE_REF', '1d8b398')
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
        print('  ✗ ' + name + ('  → ' + str(detail)[:900] if detail != '' else ''))


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
    def __init__(self, browser, name, rev, tvars, resp, width=1400):
        self.ctx = browser.new_context(viewport={'width': width, 'height': 1000})
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

    def modal(self):
        return self.ev("""() => ({
          show: document.getElementById('pm-backdrop').classList.contains('show'),
          wide: document.querySelector('#pm-backdrop .pm-card').classList.contains('wide'),
          msg: document.getElementById('pm-msg').textContent,
          sub: document.getElementById('pm-sub').innerHTML,
          btns: Array.from(document.querySelectorAll('#pm-actions button')).map(b => b.textContent),
          tops: Array.from(document.querySelectorAll('#pm-actions button')).map(b => Math.round(b.getBoundingClientRect().top)),
          fit: Array.from(document.querySelectorAll('#pm-actions button')).every(b => {
            var c = document.querySelector('#pm-backdrop .pm-card').getBoundingClientRect(), r = b.getBoundingClientRect();
            return r.left >= c.left - 1 && r.right <= c.right + 1; })
        })""")

    def close(self):
        self.ctx.close()


CLOSE_PROCESSING = "var b = document.getElementById('pm-backdrop'); if (b) b.classList.remove('show');"
FILES = {
    'quote': {'name': '견적서_6485.pdf', 'mimeType': 'application/pdf', 'buffer': b'%PDF-1.4 quote'},
    'spec': {'name': '사양서.docx', 'mimeType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'buffer': b'PK spec'},
    'hwp': {'name': '견적.hwp', 'mimeType': 'application/x-hwp', 'buffer': b'HWP doc'},
}

# ═════════════════════════════════════════════════════════════
FORM_RESP = {
    'getCurrentUserForClient': {'mode': 'ok', 'value': {'ok': True, 'user': {'email': 'kim@inlct.com', 'name': '김기안', 'dept': '품질팀'}}},
    'getApproverListForClient': {'mode': 'ok', 'value': {'ok': True, 'approvers': []}},
    'checkDocNoForClient': {'mode': 'ok', 'value': {'ok': True, 'duplicated': False}},
}
FILL_VALID = """([dup]) => {
  document.getElementById('issue-date').value = '2026-11-23';
  document.getElementById('subject').value = 'LC공정 계측기 검교정';
  document.getElementById('purpose').value = '정기 검교정';
  document.getElementById('vendorName').value = '(주)검교정';
  document.getElementById('vendorContact').value = '박담당';
  document.getElementById('vendorEmail').value = 'sales@example.com';
  var people = [['이결재', 'lee@inlct.com'], ['최결재', 'choi@inlct.com'], ['박결재', 'park@inlct.com']];
  if (dup) people[1] = ['이결재', 'lee@inlct.com'];
  for (var i = 1; i < apprIds.length; i++) {
    document.getElementById('trig-' + apprIds[i]).textContent = people[i - 1][0];
    document.getElementById('email-' + apprIds[i]).value = people[i - 1][1];
  }
  var tr = document.querySelector('#items-body tr');
  tr.querySelector('.item-name').value = '정전기 측정기'; tr.querySelector('.item-spec').value = 'EA';
  tr.querySelector('.qty').value = 2; tr.querySelector('.price').value = '30,500';
  calcRow(tr.querySelector('.price'));
  document.getElementById('doc-no').value = 'TQ-AP-26-004';
}"""


def open_form(browser, rev, files=(), pick=None, dup=False):
    p = Page(browser, 'Procurement_Form.html', rev, {'webappUrl': WEBAPP}, FORM_RESP)
    p.page.wait_for_selector('#items-body tr')
    p.page.wait_for_function("document.getElementById('drafter').value === '김기안'")
    p.ev(FILL_VALID, [dup])
    if files:
        p.page.set_input_files('#file-in', [FILES[k] for k in files])
        p.page.wait_for_function('(n) => document.querySelectorAll("#attach-list .attach-item").length >= n', arg=len(files))
    if pick is not None:
        p.page.check('input[name="quote-pick"][value="%d"]' % pick)
    return p


def go_on(p, fn):
    """모달의 [계속] → 서버 호출 값"""
    n = len(p.calls(fn))
    p.page.click('#pm-ok-btn')
    p.wait_calls(fn, n + 1)
    p.ev(CLOSE_PROCESSING)
    return p.calls(fn)[-1]['args'][0]


def form(browser):
    print('\n[F] 품의 작성 화면 (Procurement_Form.html)')

    # F1 첨부 0건
    n = open_form(browser, 'NEW')
    n.ev('submitForm()')
    m = n.modal()
    check('F1 첨부 0건 → 제목 "첨부된 파일이 없습니다." · 본문 질문', m['show'] and m['msg'] == '첨부된 파일이 없습니다.'
          and m['sub'] == '견적서 없이 기안을 올리시겠습니까?', m)
    check('F1 버튼 [첨부하러 가기][견적서 없이 제출] · 넓은 카드 · 한 줄 · 카드 안', m['btns'] == ['첨부하러 가기', '견적서 없이 제출']
          and m['wide'] and (max(m['tops']) - min(m['tops'])) <= 2 and m['fit'], m)
    n.page.click('#pm-cancel-btn')
    check('F2 [첨부하러 가기] → 모달 닫힘 · 제출 0건 · 첨부 영역 강조 · 넓은 카드 해제',
          not n.modal()['show'] and not n.calls('submitRequisitionForClient')
          and 'hint' in n.ev("document.getElementById('attach-zone').className") and not n.modal()['wide'],
          [n.modal(), n.ev("document.getElementById('attach-zone').className")])
    n.page.wait_for_timeout(3200)
    check('F2 강조는 3초 뒤 사라짐', 'hint' not in n.ev("document.getElementById('attach-zone').className"))
    h = open_form(browser, 'HEAD')
    h.ev('submitForm()')
    n.ev('submitForm()')
    ph, pn = go_on(h, 'submitRequisitionForClient'), go_on(n, 'submitRequisitionForClient')
    check('F3 [견적서 없이 제출] → 서버 값 이전과 동일', ph == pn and pn['attachments'] == [], [ph, pn])
    for p in (h, n):
        check('F3 대화상자 0건 · 스크립트 오류 0건', not p.dialogs and not p.errors, [p.dialogs, p.errors])
        p.close()

    # F4 첨부 있음·미지정
    n = open_form(browser, 'NEW', files=['quote', 'spec'])
    n.ev('submitForm()')
    m = n.modal()
    check('F4 미지정 → 제목 "견적서가 지정되지 않았습니다." · 본문 안내', m['show'] and m['msg'] == '견적서가 지정되지 않았습니다.'
          and m['sub'] == "첨부 중 견적서가 있다면 파일 옆 '견적서'를 선택해 주세요.", m)
    check('F4 버튼 [지정하러 가기][견적서 아님 — 계속 제출] · 한 줄 · 카드 안', m['btns'] == ['지정하러 가기', '견적서 아님 — 계속 제출']
          and m['wide'] and (max(m['tops']) - min(m['tops'])) <= 2 and m['fit'], m)
    n.page.click('#pm-cancel-btn')
    hint = n.ev("Array.from(document.querySelectorAll('#attach-list .quote-pick')).map(e => e.classList.contains('hint'))")
    check('F5 [지정하러 가기] → 제출 0건 · 견적서 라디오 2개 모두 강조', not n.modal()['show'] and not n.calls('submitRequisitionForClient')
          and hint == [True, True], hint)
    n.page.check('input[name="quote-pick"][value="0"]')
    n.ev('submitForm()')
    n.wait_calls('submitRequisitionForClient', 1)
    n.ev(CLOSE_PROCESSING)
    sent = n.calls('submitRequisitionForClient')[-1]['args'][0]
    check('F6 지정 후 다시 제출 → 모달 없이 바로 제출 · isQuote 1건', [a.get('isQuote') for a in sent['attachments']] == [True, None], sent['attachments'])
    n.close()

    h = open_form(browser, 'HEAD', files=['quote', 'spec'])
    n = open_form(browser, 'NEW', files=['quote', 'spec'])
    h.ev('submitForm()')
    n.ev('submitForm()')
    ph, pn = go_on(h, 'submitRequisitionForClient'), go_on(n, 'submitRequisitionForClient')
    check('F7 [견적서 아님 — 계속 제출] → 서버 값 이전과 동일(isQuote 없음)', ph == pn and all('isQuote' not in a for a in pn['attachments']), [ph, pn])
    for p in (h, n):
        p.close()

    # F8 지정 — 이전과 같이 모달 없이 제출, 값 동일
    h = open_form(browser, 'HEAD', files=['quote', 'spec'], pick=0)
    n = open_form(browser, 'NEW', files=['quote', 'spec'], pick=0)
    for p in (h, n):
        p.ev('submitForm()')
        p.wait_calls('submitRequisitionForClient', 1)
        p.ev(CLOSE_PROCESSING)
    ph, pn = h.calls('submitRequisitionForClient')[0]['args'][0], n.calls('submitRequisitionForClient')[0]['args'][0]
    check('F8 견적서 지정 → 모달 없이 제출 · 서버 값 이전과 동일', ph == pn and pn['attachments'][0].get('isQuote') is True, [ph, pn])
    for p in (h, n):
        p.close()

    # F9 결재자 중복 + 미지정 — 한 모달, 일반 버튼
    n = open_form(browser, 'NEW', files=['quote'], dup=True)
    if n.ev('apprIds.length') >= 3:
        n.ev('submitForm()')
        m = n.modal()
        check('F9 중복+미지정 → "제출 전 확인해 주세요." · 두 안내 모두 · [취소][계속 제출] · 기본 카드',
              m['msg'] == '제출 전 확인해 주세요.' and '동일한 결재자가 여러 단계에 배정되어 있습니다.' in m['sub']
              and '견적서가 지정되지 않았습니다.' in m['sub'] and '이대로 제출하시겠습니까?' in m['sub']
              and m['btns'] == ['취소', '계속 제출'] and not m['wide'], m)
        h = open_form(browser, 'HEAD', files=['quote'], dup=True)
        h.ev('submitForm()')
        ph, pn = go_on(h, 'submitRequisitionForClient'), go_on(n, 'submitRequisitionForClient')
        check('F9 [계속 제출] → 서버 값 이전과 동일', ph == pn, [ph, pn])
        h.close()
        # F10 중복만 — 문구·버튼이 이전과 글자까지 같다
        h = open_form(browser, 'HEAD', files=['quote'], pick=0, dup=True)
        n2 = open_form(browser, 'NEW', files=['quote'], pick=0, dup=True)
        h.ev('submitForm()')
        n2.ev('submitForm()')
        mh, mn = h.modal(), n2.modal()
        check('F10 결재자 중복만 → 제목·본문·버튼 이전과 동일', [mh['msg'], mh['sub'], mh['btns']] == [mn['msg'], mn['sub'], mn['btns']], [mh, mn])
        h.close()
        n2.close()
    else:
        check('F9 결재자 칸 3개 이상(픽스처 전제)', False, n.ev('apprIds.length'))
    n.close()

    # F11 첨부가 전부 hwp(지정 불가) — 미지정 안내, [지정하러 가기]는 목록으로만 이동(오류 없음)
    n = open_form(browser, 'NEW', files=['hwp'])
    n.ev('submitForm()')
    check('F11 hwp 만 첨부 → 미지정 안내', n.modal()['msg'] == '견적서가 지정되지 않았습니다.', n.modal())
    n.page.click('#pm-cancel-btn')
    check('F11 [지정하러 가기] → 오류 없음 · 제출 0건', not n.errors and not n.calls('submitRequisitionForClient'), n.errors)
    n.close()

    # F12 초기화 확인창 — 이전과 같은 [취소][초기화], 기본 카드
    h, n = open_form(browser, 'HEAD'), open_form(browser, 'NEW')
    for p in (h, n):
        p.ev("showConfirmModal('양식을 초기화하시겠습니까?', '입력한 내용과 첨부 파일이 모두 지워집니다.', doResetAll, '초기화')")
    mh, mn = h.modal(), n.modal()
    check('F12 다른 확인창(초기화) 문구·버튼 이전과 동일 · 기본 카드', [mh['msg'], mh['sub'], mh['btns']] == [mn['msg'], mn['sub'], mn['btns']]
          and not mn['wide'], [mh, mn])
    n.page.click('#pm-cancel-btn')
    check('F12 [취소] → 닫힘', not n.modal()['show'])
    for p in (h, n):
        check('F 대화상자 0건 · 스크립트 오류 0건', not p.dialogs and not p.errors, [p.dialogs, p.errors])
        p.close()

    # F13 좁은 화면(모바일) — 버튼이 카드 밖으로 넘치지 않는다
    n = Page(browser, 'Procurement_Form.html', 'NEW', {'webappUrl': WEBAPP}, FORM_RESP, width=380)
    n.page.wait_for_selector('#items-body tr')
    n.page.wait_for_function("document.getElementById('drafter').value === '김기안'")
    n.ev(FILL_VALID, [False])
    n.page.set_input_files('#file-in', [FILES['quote']])
    n.ev('submitForm()')
    m = n.modal()
    check('F13 폭 380px — 버튼이 카드 안(줄바꿈 허용)', m['show'] and m['fit'], m)
    n.close()


# ═════════════════════════════════════════════════════════════
REJ_DATA = {
    'ok': True, 'data': {
        'docNo': 'TQ-AP-26-004', 'issueDate': '2026-11-23', 'drafter': '김기안', 'dept': '품질팀', 'subject': 'LC공정 계측기 검교정',
        'purpose': '정기 검교정', 'vendorName': '(주)검교정', 'vendorEmail': 'sales@example.com', 'vendorContact': '박담당',
        'vendorPhone': '042-000-0000', 'deliveryDate': '2026-12-01', 'deliveryAddr': '대전', 'remarks': '',
        'items': [{'name': '정전기 측정기', 'spec': 'EA', 'qty': 2, 'price': 30500, 'currency': 'KRW'}], 'vatMode': 'excl',
        'rejectApprover': '이결재', 'rejectTime': '2026-11-23 10:00', 'rejectReason': '견적서 확인', 'rejectHistory': '',
    }}
REJ_RESP = {
    'getRejectDataForClient': {'mode': 'ok', 'value': REJ_DATA},
    'getRequisitionForViewer': {'mode': 'ok', 'value': {'ok': True, 'data': {'approvers': [
        {'label': '기안자', 'name': '김기안', 'email': 'kim@inlct.com'}, {'label': '결재자 1', 'name': '이결재', 'email': 'lee@inlct.com'}]}}},
}


def open_reject(browser, rev, files=(), pick=None):
    p = Page(browser, 'Procurement_Reject.html', rev, {'webappUrl': WEBAPP, 'token': 'tok-1', 'docNo': 'TQ-AP-26-004'}, REJ_RESP)
    p.page.wait_for_function("document.getElementById('subject').value === 'LC공정 계측기 검교정'")
    p.wait_calls('getRequisitionForViewer', 1)
    p.page.wait_for_function('cachedApprovers.length === 2')
    if files:
        p.page.set_input_files('#file-in', [FILES[k] for k in files])
        p.page.wait_for_function('(n) => document.querySelectorAll("#attach-list .attach-item").length >= n', arg=len(files))
    if pick is not None:
        p.page.check('input[name="quote-pick"][value="%d"]' % pick)
    return p


def reject(browser):
    print('\n[J] 반려 후 재상신 화면 (Procurement_Reject.html)')
    n = open_reject(browser, 'NEW')
    n.ev('doResubmit()')
    m = n.modal()
    check('J1 첨부 0건 → "첨부된 파일이 없습니다." · "견적서 없이 재상신하시겠습니까?" · [첨부하러 가기][견적서 없이 재상신]',
          m['msg'] == '첨부된 파일이 없습니다.' and m['sub'] == '견적서 없이 재상신하시겠습니까?'
          and m['btns'] == ['첨부하러 가기', '견적서 없이 재상신'] and m['wide'] and (max(m['tops']) - min(m['tops'])) <= 2 and m['fit'], m)
    n.page.click('#pm-cancel-btn')
    check('J2 [첨부하러 가기] → 재상신 0건 · 첨부 영역 강조', not n.calls('resubmitRequisitionForClient')
          and 'hint' in n.ev("document.getElementById('attach-zone').className"), n.ev("document.getElementById('attach-zone').className"))
    h = open_reject(browser, 'HEAD')
    h.ev('doResubmit()')
    n.ev('doResubmit()')
    ph, pn = go_on(h, 'resubmitRequisitionForClient'), go_on(n, 'resubmitRequisitionForClient')
    check('J3 [견적서 없이 재상신] → 서버 값 이전과 동일', ph == pn and pn['data']['attachments'] == [], [ph, pn])
    for p in (h, n):
        p.close()

    n = open_reject(browser, 'NEW', files=['quote', 'spec'])
    n.ev('doResubmit()')
    m = n.modal()
    check('J4 미지정 → "견적서가 지정되지 않았습니다." · [지정하러 가기][견적서 아님 — 계속 재상신] · 한 줄 · 카드 안',
          m['msg'] == '견적서가 지정되지 않았습니다.' and m['sub'] == "첨부 중 견적서가 있다면 파일 옆 '견적서'를 선택해 주세요."
          and m['btns'] == ['지정하러 가기', '견적서 아님 — 계속 재상신'] and (max(m['tops']) - min(m['tops'])) <= 2 and m['fit'], m)
    n.page.click('#pm-cancel-btn')
    hint = n.ev("Array.from(document.querySelectorAll('#attach-list .quote-pick')).map(e => e.classList.contains('hint'))")
    check('J5 [지정하러 가기] → 재상신 0건 · 라디오 강조', not n.calls('resubmitRequisitionForClient') and hint == [True, True], hint)
    n.close()

    h = open_reject(browser, 'HEAD', files=['quote', 'spec'])
    n = open_reject(browser, 'NEW', files=['quote', 'spec'])
    h.ev('doResubmit()')
    n.ev('doResubmit()')
    ph, pn = go_on(h, 'resubmitRequisitionForClient'), go_on(n, 'resubmitRequisitionForClient')
    check('J6 [견적서 아님 — 계속 재상신] → 서버 값 이전과 동일', ph == pn, [ph, pn])
    for p in (h, n):
        p.close()

    h = open_reject(browser, 'HEAD', files=['quote', 'spec'], pick=0)
    n = open_reject(browser, 'NEW', files=['quote', 'spec'], pick=0)
    for p in (h, n):
        p.ev('doResubmit()')
        p.wait_calls('resubmitRequisitionForClient', 1)
    ph, pn = h.calls('resubmitRequisitionForClient')[0]['args'][0], n.calls('resubmitRequisitionForClient')[0]['args'][0]
    check('J7 지정 → 모달 없이 재상신 · 서버 값 이전과 동일', ph == pn and pn['data']['attachments'][0].get('isQuote') is True, [ph, pn])
    for p in (h, n):
        check('J 대화상자 0건 · 스크립트 오류 0건', not p.dialogs and not p.errors, [p.dialogs, p.errors])
        p.close()

    # 폐기 확인창(showDiscardConfirm)은 그대로
    h, n = open_reject(browser, 'HEAD'), open_reject(browser, 'NEW')
    for p in (h, n):
        p.ev('showDiscardConfirm()')
    mh, mn = h.modal(), n.modal()
    check('J8 폐기 확인창 이전과 동일', [mh['msg'], mh['sub'], mh['btns']] == [mn['msg'], mn['sub'], mn['btns']] and not mn['wide'], [mh, mn])
    for p in (h, n):
        p.close()


# ═════════════════════════════════════════════════════════════
def home(browser):
    print('\n[A] 관리자 홈 배너 (Procurement_Home.html)')
    data = FIX['homeAdmin']
    tok = FIX['dismissToken']

    def open_home(rev, dismiss_resp=None):
        resp = {'getHomeDataForClient': {'mode': 'ok', 'value': data}}
        if dismiss_resp:
            resp['dismissQuoteUnpickedForClient'] = dismiss_resp
        p = Page(browser, 'Procurement_Home.html', rev, {'webappUrl': WEBAPP, 'initialMenu': 'mydocs'}, resp)
        p.page.wait_for_selector('.admin-alert')
        return p

    def rows(p):
        return p.ev("Array.from(document.querySelectorAll('#aa-panel-quote .aa-row')).map(r => r.innerText.replace(/\\s+/g, ' ').trim())")

    h = open_home('HEAD')
    n = open_home('NEW', {'mode': 'ok', 'value': {'ok': True, 'docNo': 'TQ-AP-26-004'}})
    check('A1 배지 2건(미지정 1 + 생성 요청 1)', n.text('.aa-badge') == '2' and h.text('.aa-badge') == '2', [h.text('.aa-badge'), n.text('.aa-badge')])
    n.page.click('[data-admin-task="quote"]')
    h.page.click('[data-admin-task="quote"]')
    rn, rh = rows(n), rows(h)
    un = [r for r in rn if 'TQ-AP-26-004' in r]
    check('A2 미지정 행: 상태 "견적서 미지정" · 첨부 링크 · 폴더 · QUO_ 이름 · [견적서 아님]', len(un) == 1 and '견적서 미지정' in un[0]
          and '20261123-2 인엘씨테크놀러지 견적서' in un[0] and '폴더 열기' in un[0] and 'QUO_TQ-AP-26-004.pdf' in un[0]
          and un[0].endswith('견적서 아님'), rn)
    pend_n = [r for r in rn if 'TQ-AP-26-030' in r]
    pend_h = [r for r in rh if 'TQ-AP-26-030' in r]
    check('A3 기존 생성 요청(PENDING) 행은 이전과 글자까지 같음', pend_n == pend_h and len(pend_n) == 1, [pend_h, pend_n])
    links = n.ev("Array.from(document.querySelectorAll('#aa-panel-quote .aa-row')[0].querySelectorAll('a')).map(a => a.href)")
    check('A4 첨부 링크 = Drive 파일 보기 주소', any('/file/d/' in u and u.endswith('/view') for u in links), links)

    n.page.click('.aa-btn[data-token="%s"]' % tok)
    check('A5 [견적서 아님] → 확인 모달(품의번호 · 영향 없음 안내) · 서버 호출 아직 0건',
          n.page.is_visible('#pkmodal-backdrop.show') and 'TQ-AP-26-004' in n.text('#pkmodal-body')
          and '결재 진행에는 영향이 없습니다' in n.text('#pkmodal-body') and not n.calls('dismissQuoteUnpickedForClient'), n.text('#pkmodal-body'))
    n.page.click('#pkmodal-footer .pkmodal-btn.secondary')
    check('A6 [취소] → 닫힘 · 호출 0건', not n.page.is_visible('#pkmodal-backdrop.show') and not n.calls('dismissQuoteUnpickedForClient'))
    n.page.click('.aa-btn[data-token="%s"]' % tok)
    n.page.click('#pkmodal-footer .pkmodal-btn.primary')
    n.wait_calls('dismissQuoteUnpickedForClient', 1)
    n.page.wait_for_selector('#pkmodal-body .pkmodal-success')
    check('A7 [견적서 아님] 확정 → 서버 호출 {token} 1건 · 완료 모달', n.calls('dismissQuoteUnpickedForClient')[0]['args'] == [{'token': tok}]
          and '견적서 없음으로 정리했습니다' in n.text('#pkmodal-body'), [n.calls('dismissQuoteUnpickedForClient'), n.text('#pkmodal-body')])
    n.page.click('#pkmodal-footer .pkmodal-btn.primary')
    rn2 = rows(n)
    check('A8 닫기 → 배지 1건 · 견적서 목록 다시 펼침 · 미지정 행 사라짐', n.text('.aa-badge') == '1'
          and n.page.is_visible('#aa-panel-quote') and len(rn2) == 1 and 'TQ-AP-26-030' in rn2[0], [n.text('.aa-badge'), rn2])
    for p in (h, n):
        check('A 대화상자 0건 · 스크립트 오류 0건', not p.dialogs and not p.errors, [p.dialogs, p.errors])
        p.close()

    # 실패 응답 — 오류 모달, 버튼 다시 활성, 배너 그대로
    f = open_home('NEW', {'mode': 'ok', 'value': {'ok': False, 'message': '관리자만 처리할 수 있습니다.'}})
    f.page.click('[data-admin-task="quote"]')
    f.page.click('.aa-btn[data-token="%s"]' % tok)
    f.page.click('#pkmodal-footer .pkmodal-btn.primary')
    f.page.wait_for_selector('#pkmodal-body .pkmodal-error')
    check('A9 서버 거부 → 오류 모달(사유) · 버튼 다시 활성 · 배지 그대로', '관리자만 처리할 수 있습니다.' in f.text('#pkmodal-body')
          and f.ev("document.querySelector('.aa-btn').disabled") is False and f.text('.aa-badge') == '2', f.text('#pkmodal-body'))
    f.page.click('#pkmodal-footer .pkmodal-btn.secondary')
    check('A9 닫기 → 모달 닫힘 · 오류 0건', not f.page.is_visible('#pkmodal-backdrop.show') and not f.errors and not f.dialogs, f.errors)
    f.close()


with sync_playwright() as pw:
    browser = pw.chromium.launch()
    try:
        form(browser)
        reject(browser)
        home(browser)
    finally:
        browser.close()

print('\n─────────────────────────────')
print(f'통과 {passed} / 실패 {failed}   (기준 {BASE_REF})')
if failed:
    print('실패 항목:\n  - ' + '\n  - '.join(failures))
    sys.exit(1)
