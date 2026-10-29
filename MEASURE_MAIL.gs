// [임시 측정] 결재 메일 지연 발송 검토용 (2026-10-29)
//  결재 응답 안에서 메일 1통이 차지하는 시간과, 지연 발송으로 바꿀 때 새로 붙는 비용(보관함 저장)을 잰다.
//  측정이 끝나면 이 파일을 통째로 지운다. 운영 코드는 이 파일을 호출하지 않는다.
//  공개 함수라 화면에서도 호출될 수 있으므로(SEC-07) 관리자만 실행, 수신자는 실행자 본인으로 고정한다.

var MEASURE_MAIL_ROUNDS = 5;

/**
 * GAS 편집기에서 수동 실행. 실행자 본인에게 결재요청 형식의 테스트 메일 5통을 보낸다.
 *  메일은 운영과 같은 sendApprovalEmailWithRetry(별칭·HTML·재시도 포함)로 보낸다.
 *  ※ 테스트 메일의 승인/반려 버튼은 누르지 말 것 — 없는 문서라 거부되고 시스템로그에 거부 1행이 남는다.
 *  운영에서 메일은 실행당 첫 GmailApp 호출이므로 '1회차' 값이 운영과 가장 가깝다.
 *  실행할 때마다 새 실행이 되므로 시간대를 바꿔 2~3번 돌리면 1회차 표본이 쌓인다.
 */
function measureMailLatency() {
  var me = getRequestUserEmail_();
  if (!isAdminUser(me)) throw new Error('관리자만 실행할 수 있습니다.');

  var props = PropertiesService.getScriptProperties();
  var items = [
    { name: '측정용 품목 A', qty: 2, price: 150000, currency: 'KRW' },
    { name: '측정용 품목 B', qty: 1, price: 320000, currency: 'KRW' },
  ];
  var approver = { label: '결재자', name: '측정용', email: me };
  var mailMs = [], getMs = [], setMs = [];

  for (var i = 1; i <= MEASURE_MAIL_ROUNDS; i++) {
    var data = {
      docNo:     '측정용-TEST',
      subject:   '[측정용] 메일 발송 시간 측정 ' + i + '/' + MEASURE_MAIL_ROUNDS + ' — 결재하지 마세요',
      drafter:   '측정용',
      issueDate: toDateStr(new Date()),
      extra:     buildMailExtraInfo('메일 발송 시간 측정 (결재 메일 지연 발송 검토)', '측정용 업체',
                                    items, calcTotal(items), '테스트 메일입니다. 승인/반려 버튼을 누르지 마세요.'),
    };

    // ① 운영 결재요청 메일과 같은 함수로 발송
    var t0 = Date.now();
    sendApprovalEmailWithRetry(approver, data, 'MEASURE-NOT-A-DOC', 0, 1);
    mailMs.push(Date.now() - t0);

    // ② 지연 발송 시 결재 응답에 새로 붙는 비용 — 하트비트 읽기 + 메일 1건을 보관함에 저장
    var t1 = Date.now();
    props.getProperty(AUDIT_FLUSH_HEARTBEAT);
    getMs.push(Date.now() - t1);

    var key = 'MEASURE_TMP_' + i;
    var t2 = Date.now();
    props.setProperty(key, JSON.stringify({ type: 'next_approval', approver: approver, data: data,
                                            token: 'MEASURE-NOT-A-DOC', rowNum: 0, idx: 1 }));
    setMs.push(Date.now() - t2);
    props.deleteProperty(key);
  }

  function stat(arr) {
    var s = arr.slice().sort(function(a, b) { return a - b; });
    return '1회차 ' + arr[0] + 'ms / 중앙값 ' + s[Math.floor(s.length / 2)] + 'ms / 최소 ' + s[0] +
           'ms / 최대 ' + s[s.length - 1] + 'ms  [' + arr.join(', ') + ']';
  }
  var summary = [
    '▶ 메일 1통 발송 : ' + stat(mailMs),
    '▶ 하트비트 읽기 : ' + stat(getMs),
    '▶ 보관함 저장   : ' + stat(setMs),
    '▶ 측정 시각 ' + toDateTimeStr(new Date()) + ' / 수신 ' + me,
  ].join('\n');
  console.log(summary);
  return { mailMs: mailMs, heartbeatGetMs: getMs, bufferSetMs: setMs };
}
