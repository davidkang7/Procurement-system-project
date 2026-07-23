# 구매결재시스템 — AI Context Document

**InLC Technology 구매결재시스템 (Procurement Approval System)**
**후속 개발 시 AI 어시스턴트가 참고하는 컨텍스트 문서**

| 항목 | 내용 |
| --- | --- |
| **문서 목적** | 향후 개발·디버깅·기능 확장 시 AI(또는 신규 개발자)가 이 시스템의 아키텍처, **지금까지 실제로 겪은 오류**, 반드시 지켜야 할 주의사항을 빠르게 파악하도록 돕는다 |
| **대상 시스템** | 구매결재시스템 v2.1 + INSP(검수보고서) Phase 2 모듈, Google Workspace 기반 |
| **런타임** | Google Apps Script (GAS) Web App, **`executeAs: USER_DEPLOYING` / `access: DOMAIN`** 배포 (모든 GAS 쿼터가 배포자=David 계정에 귀속) |
| **작성일 / 개정** | 2026-05-29 최초 → **2026-06-20 개정 (v2.2)** |
| **연계 문서** | 구매결재시스템_요구서_Phase1_v2.2, 구매결재시스템_셋팅매뉴얼_v1.3, INSP_구현설계서_v1.0 |

> **📌 이 문서를 읽는 AI에게**
> 이 시스템은 1인 개발자(David, davidkang@inlct.com)가 운영하는 사내 시스템이다. ERP 없이 Google Workspace(Apps Script + Sheets + Drive + Gmail)만으로 동작한다. 화려한 재작성보다 **기존 패턴 준수, 작은 단위 확인, 안전한 점진 변경**이 항상 우선이다. 코드는 자동 반영되지 않고 **개발자가 Apps Script 편집기에 직접 복붙**한다. 아래 §2의 "치명적 함정"과 §3의 "실제 운영 사고 사례"를 위반/반복하면 조용한 데이터 손실이나 전 사용자 장애로 이어진다. 코드를 수정하기 전에 반드시 §2·§3·§6을 먼저 확인하라.

> **⚠ v2.2 개정 핵심 — 무엇이 바뀌었나**
> ① **INSP(검수보고서) Phase 2 모듈 신설** — 별도 `INSP.gs`(~1,450줄) + 전용 HTML 3종. 별도 시트(검수보고서목록, 51컬럼), 회차 채번, 사진 압축·삭제, 종결 차단. (§9 신설)
> ② **실제 운영 사고 4건을 §3에 정식 기록** — (가) Gmail 스코프 미스매치로 결재 메일 전 건 미발송, (나) "새 배포" 반복으로 활성 배포 2개 갈라짐 → 검수 버튼 미표시, (다) Date 직렬화 null, (라) drafter 컬럼이 이메일이 아닌 성명. **이 네 가지는 반드시 숙지할 것.**
> ③ **스키마 변경 반영** — 품의서목록 v2.1: 구매방법(30)·구매조건/지급정보(31) 컬럼 추가로 `APPR_START`가 30→32로 이동.
> ④ **이메일 이중 경로 아키텍처**(의도적 분리, 통합 금지) 정식 문서화. (§10)
> ⑤ **UI 디자인 시스템**(와인레드/딥블루/청록) 및 **ERP 통합 로드맵** 추가. (§11, §12)

---

# 1. 시스템 한눈에 보기

## 1.1 무엇을 하는 시스템인가

기안자가 구매품의서(REQ)를 작성→결재받고, 1차 승인된 REQ를 구매팀이 픽업하여 구매팀 품의서(PRC)로 재기안→최종 결재받으면, PO번호 기준 폴더에 PDF·첨부가 자동 보관되는 결재 워크플로우 시스템이다. **Phase 2에서 검수보고서(INSP)가 추가되어, 최종승인(PRC)된 발주 건에 대해 입고/검수 결과를 회차별로 제출·결재하고 PO 폴더에 PDF로 보관**한다. 약 25명(→50명 확장) 사용.

## 1.2 핵심 자산 (v2.2 시점)

| 자산 | 내용 |
| --- | --- |
| Code.gs | 약 4,381줄, 22개 섹션 (단일 파일 백엔드) |
| INSP.gs | 약 1,450줄 — 검수보고서 모듈 (Step 1~6 누적본). Code.gs와 전역 스코프 공유 |
| INSP_DIAG.gs | 검수 결재 대기 미표시 원인 추적용 진단 함수(`debugInspPending`). 임시/디버깅용 |
| step1_audit_log.gs | 시스템로그 인프라 + AUDIT_EVENT 상수 (Code.gs와 병합 운영) |
| HTML 13종 | Procurement_Home, _Form, _PRC_Form, _Viewer, _Approval, _Reject, PDF_Template, _Admin_Discard, _Admin_ForceReleaseLock, **_INSP_Form, _INSP_Viewer, INSP_PDF_Template** (+ 디자인 시뮬레이터는 비배포) |
| Google Sheets (1파일·4시트) | 품의서목록(**32컬럼**+결재자블록), 결재자목록, 시스템로그(12컬럼·자동생성), **검수보고서목록(51컬럼·자동생성)** |
| Google Drive | STAGING_ROOT(진행중 임시), FINAL_ROOT(최종 보관) |
| 고급 서비스 | Drive API v3 (식별자 'Drive') — 필수 |
| 자동 트리거 | processQueueTrigger (1분 간격, 필수). INSP `insp_pdf` job도 동일 큐로 처리 |

**운영 CONFIG 식별자 (코드 수정 시 항상 이 값을 사용):**

| 키 | 값 |
| --- | --- |
| SHEET_ID | `1Jg0BfLiMrI9MUDMqwU7DWIAA8loVU5XwfumiWHMqMTA` |
| STAGING_ROOT_ID | `17MouMkuLeAcIkfml9kRwIs6OQVdYKKRc` |
| FINAL_ROOT_ID | `1_dXx1O-jyKHa1mnJ48eLzTEhzGB1t9VE` |
| WEBAPP_URL | `https://script.google.com/a/macros/inlct.com/s/AKfycbw-iL2l4m59rcZLuiBSgWkumFjBhg80QRm3Y2Kdz1_l9YJYqnWPDw_lhuTurXiCJ8Ip/exec` |
| INSP_SHEET_NAME | `검수보고서목록` |
| AUDIT_SHEET_NAME | `시스템로그` |
| scriptId (clasp) | `1QjMr3MXiFlon24bkCIP9NgLZ_oTrWqQ-eQ8f48R-spBVzRNCU6rRRyDQ` |

> **WEBAPP_URL은 `…iL2l4m…/exec`가 정답이다.** 과거에 "새 배포"를 잘못 만들어 `…ATwWeh…/exec` 같은 다른 URL이 생긴 적이 있다(§3.2 참조). 어떤 URL이 맞는지 헷갈리면 이 값을 기준으로 삼는다.

## 1.3 데이터 흐름 (요약)

**REQ/PRC 결재:** REQ 작성(검토중) → 순차결재(결재중) → 1차 최종승인(최종승인) → 구매팀 인박스 노출 → 선착순 락 점유(ClaimedBy 기록) → PRC 프리필·작성 → PRC 제출(부모 REQ: PRC생성됨) → 구매팀 결재 → PRC 최종승인(최종승인(PRC)) → **작업 큐 등록** → 1분 트리거가 PDF 생성 + FINAL 폴더 통합 이동.

**INSP 검수:** 최종승인(PRC)된 PRC를 원본 REQ 기안자가 픽업 → 검수보고서 작성(회차 자동 채번 `{REQ docNo}-NN`, 판정 합격/불합격, 사진 압축 업로드) → 제출(검토중) → 순차결재(기안자 자기결재 포함 최대 4블록) → 최종승인(INSP) → **insp_pdf 큐 등록** → 1분 트리거가 검수 PDF 생성 + FINAL/PO폴더 저장 + **사진 원본·임시폴더 삭제**. `IS_FINAL='Y'` 회차가 최종승인되면 해당 PRC는 검수 종결(추가 제출 차단).

## 1.4 권한 모델 (누적형)

전 사용자 공통 메뉴 → isProcurement(구매팀) 추가 메뉴 → isAdmin 관리자 섹션. 권한은 코드 내 CONFIG 배열(`PROCUREMENT_TEAM_EMAILS`, `ADMIN_EMAILS`)로 판정한다. 사용자 식별은 항상 서버에서 `getActiveUserEmail()`(=`Session.getActiveUser().getEmail()`)로 해결한다. 검수 작성 권한은 **원본 REQ 기안자 본인만**(이메일 소문자 정규화 비교).

---

# 2. 치명적 함정 (Critical Pitfalls) — 위반 금지

> 📌 *아래 항목들은 과거에 실제로 장애를 일으켰거나 일으킬 수 있는 것들이다. 코드 작성·수정 시 매번 확인하라. 특히 **2.9·2.10·2.11은 INSP 개발 중 실제로 터진 버그**이며, §3에 사고 기록이 있다.*

## 2.1 google.script.run + Date 객체 = 조용한 null

Sheets에서 읽은 Date 객체를 `google.script.run`으로 클라이언트에 반환하면 **에러 없이 null로 직렬화**되어, 클라이언트는 원인 모를 빈 값을 받는다. (실제 사고: §3.3)

- **반드시** 서버 반환 전에 `toDateStr()` / `toDateTimeStr()` 유틸로 문자열 변환할 것.
- 클라이언트에서도 `String()` 캐스팅으로 방어.
- 모든 서버 함수는 try/catch로 감싸고, catch에서 `err.toString() + err.stack`을 메시지에 담아 반환할 것. 그렇지 않으면 실패가 조용한 null로 나타나 디버깅이 불가능해진다.

## 2.2 앱 내부에서 fetch(WEBAPP_URL) 절대 금지

조직 도메인 배포 환경에서 클라이언트가 `fetch(WEBAPP_URL)`을 호출하면 CORS/401로 차단된다.

- 모든 클라이언트→서버 통신은 **google.script.run**을 사용한다.
- `google.script.run`은 ContentService 반환 객체를 받지 못한다. 서버 함수는 **plain object**(`{ok:true, message:'…', data:…}`)를 반환해야 한다.
- (doPost 경로는 별도로 `jsonResponse()` 래퍼를 통해 ContentService를 쓰지만, 이는 외부 POST 호환용이다. 신규 클라이언트 호출은 `google.script.run` + `*ForClient` 함수 패턴을 따른다.)

## 2.3 상태값은 한글이다 (영문 아님)

요구서 §8.2는 영문 논리 상태(SUBMITTED, APPROVED_1ST 등)를 쓰지만, **실제 코드는 한글 문자열**을 결재상태 컬럼에 저장한다. 영문 상태로 비교하면 어떤 필터도 매칭되지 않는다.

| 논리 상태 | 실제 코드 값 (REQ/PRC) |
| --- | --- |
| SUBMITTED | 검토중 |
| IN_APPROVAL | 결재중 (다음결재자명) |
| 업로드 실패 | 업로드오류 |
| REJECTED / 재제출 | 반려 / 재상신 |
| APPROVED_1ST | 최종승인 |
| PRC 생성됨(부모 REQ) | PRC생성됨 |
| FINAL_APPROVED | 최종승인(PRC) |
| 폐기 | 폐기 |

**INSP 상태값(검수보고서목록)도 한글이다:** `검토중` / `결재중 (이름)` / `반려` / `최종승인(INSP)`. 판정값은 `합격` / `불합격`. 영문 금지.

- 인박스 필터 조건은 `status === '최종승인' && docType === 'REQ'`.
- 진행 중 상태는 `'결재중 (' + 다음결재자 + ')'` 형태로 동적이므로 `indexOf('결재')` 부분 매칭을 쓰는 코드가 있다.
- 최종승인 여부 판정은 `status.indexOf('최종승인') >= 0` 패턴을 자주 쓴다('최종승인'과 '최종승인(PRC)' 모두 포함). **단, INSP의 '최종승인(INSP)'도 이 패턴에 걸리므로 docType/시트로 구분할 것.**

## 2.4 LockService 범위 최소화 — 첨부·이메일은 락 밖에서

withLock 안에는 **행 추가/토큰 발급/폴더 생성/순번 채번처럼 짧고 경합 위험이 큰 작업만** 넣는다.

- 첨부 업로드(가장 느림), 사진 업로드, 이메일 발송(재시도 포함)은 **반드시 락 밖**에서 처리한다. 락 안에 두면 LOCK_WAIT_MS(30초) 경합·타임아웃이 발생한다.
- 표준 단계 구성: ①락 밖 사전검증(용량 등) → ②최소 락(행/토큰/폴더/회차채번) → ③락 밖 첨부·사진 업로드 → ④메타 기록 → ⑤락 밖 이메일.
- INSP도 동일: **회차 채번(SEQ)은 락 안**(동시 제출 경합 방지), 사진 저장은 락 밖.

## 2.5 PDF 페이로드 화이트리스트 — 내부 메모 노출 금지

PDF에는 결재 시스템 내부에서만 보여야 하는 내용이 **절대** 들어가면 안 된다.

- `buildPdfPayload()`(REQ/PRC) 및 `generateInspPdf()`(INSP)는 화이트리스트 방식으로 필요한 필드만 담는다.
- 금지 필드: remarks, prcMemo, rejectLog, rejectHistory, claimedBy, claimedAt, fieldChanges, **INSP의 REJECT_LOG·RESUB_COUNT·PHOTO_LIST(파일 ID)** 등.
- PDF 관련 코드를 수정할 때 이 화이트리스트를 우회하거나 payload에 원본 row를 통째로 넣지 말 것.

## 2.6 Drive API는 supportsAllDrives: true 필수

공유 드라이브에서 파일을 다루므로 Drive Advanced Service 호출 시 `supportsAllDrives` 옵션을 빠뜨리면 파일 이동/생성이 실패한다. 또한 Apps Script 편집기에서 Drive 고급 서비스(식별자 'Drive', v3)가 추가되어 있어야 한다.

## 2.7 FINAL 통합 이동은 PRC 최종 승인 시점에만

PO번호는 PRC 단계에서 결정되므로, FINAL 폴더 통합 이동은 **REQ 1차 승인이 아니라 PRC 최종 승인 시점**에 일어난다. REQ 단계 PDF는 STAGING에만 생성된다(pdf_only job). INSP PDF는 PRC가 이미 만든 **FINAL/{PO번호} 폴더**에 저장된다(`getOrCreateFolder(poNo, issueDate, 'final')`).

## 2.8 코드 수정 후 반드시 [새 버전] 재배포 — 단, "새 배포"가 아니라 "기존 배포 편집"

CONFIG나 코드를 고치고 저장만 하면 사용자에게는 이전 버전이 노출된다. **반드시 다음 절차로만 배포한다.** (이 절차를 어겨서 실제 장애가 났다 — §3.2)

> **[배포] → [배포 관리] → 기존 운영 배포의 [✏️ 편집(연필)] → 버전 드롭다운에서 "새 버전" 선택 → [배포]**

- 절대 **"새 배포"를 새로 만들지 말 것.** "새 배포"는 매번 새 URL을 발급해 활성 배포가 갈라진다.
- 재배포 후 URL은 동일 유지된다(`…iL2l4m…/exec`).
- `executeAs: USER_DEPLOYING`이므로, **매니페스트(스코프) 변경 후에도 새 버전 재배포가 필수**다(권한 재승인이 배포에 묶임).

## 2.9 [INSP 실제 버그] DRAFTER 컬럼은 성명이지 이메일이 아니다

품의서목록 `COL.DRAFTER`(D열)에는 **기안자 성명(예: '강기종')**이 저장된다. **이메일이 아니다.** 따라서 기안자 본인 여부 같은 **신원 비교에는 절대 `COL.DRAFTER`를 쓰면 안 된다.**

- 기안자 이메일은 **결재자 블록 0번(기안자 자기결재)의 email 컬럼 = `COL.APPR_START + 2`**에서 읽고, **소문자 정규화 후 비교**한다.
- 같은 함정이 과거 홈 대시보드 "내 문서" 필터에서도 있었고(이메일 대신 성명 비교로 결재 블록 없는 행이 사라짐), INSP `completedDocs[0].drafter`에서도 재현됐다(§3.4).
- INSP 코드는 이를 학습해 `getInspFormDataForClient`·`_submitInspCore`에서 `reqInfo.row[COL.APPR_START + 2]`로 기안자 이메일을 얻고 `actor`와 소문자 비교한다.

## 2.10 [INSP 실제 버그] 동일인 다단계 결재 — 본인 단계 판정은 curIdx 우선

기안자가 결재선의 결재자로도 들어가는 등 **한 사람이 여러 결재 단계에 배정**될 수 있다. 이때 "본인 단계"를 단순히 "본인이 배정된 첫 단계"로 찾으면 이미 승인한 앞 단계가 잡혀 결재가 막힌다.

- `_findMyInspIdx(row, apprCount, curIdx, actor)`는 **①현재 차례(curIdx)가 본인이면 그 단계를 우선** 반환하고, ②아니면 본인이 배정된 첫 단계를 반환한다.
- 결재 처리(`processInspDecisionFromClient`)·뷰어(`getInspViewerDataForClient`)·대기목록(`_getInspMenusForClient`) 전부 이 헬퍼를 통한다.
- 일반적으로 **URL `idx` 파라미터를 신뢰하지 말고 서버가 actor로 본인 단계를 재판정**한다(아래 2.11과 동일 원칙).

## 2.11 결재 단계(idx)는 URL 파라미터를 신뢰하지 말고 서버가 판정

과거 REQ/PRC 결재에서 **approver `idx` URL 파라미터가 없으면 0으로 기본값 처리**되어, 홈에서 진입 시 엉뚱한 단계로 결재되는 버그가 있었다.

- 해결: doGet에서 idx가 있을 때만 넘기고(없으면 빈 문자열), 서버가 `myApproverIdx`를 **현재 사용자 이메일로 재계산**한다. URL `idx`는 힌트로만 쓰고 현재 사용자와 일치 검증 후 사용, 불일치면 가장 빠른 대기 단계로 폴백.
- INSP도 동일 패턴(`getInspViewerDataForClient`의 myApproverIdx 3단계 판정: 현재차례 본인 → URL힌트(검증) → 본인 첫 단계).

## 2.12 INSP_COL과 COL을 절대 혼용하지 말 것

검수보고서목록은 품의서목록과 **행 구조가 완전히 다르다.** 검수보고서목록 접근은 반드시 `INSP_COL` 상수로, 품의서목록은 `COL` 상수로 한다. 결재 블록 접근은 `inspApprCol(stageIdx, offset)` 헬퍼만 사용(하드코딩 금지). 숫자 리터럴로 컬럼 접근 금지.

---

# 3. 실제 운영 사고 사례 (Incident Log) — ★반드시 숙지★

> 📌 *아래는 운영/개발 중 **실제로 발생해 시간을 크게 잡아먹은 사고**들이다. 증상·근본원인·복구·재발방지를 정리했다. 같은 증상이 보이면 즉시 이 절을 참조하라. (사용자 요청에 따라 특별 강조 섹션으로 분리)*

## 3.1 ★Gmail 스코프 미스매치 → 결재 메일 전 건 미발송 (2026-06 해결)

**증상:** 승인요청·반려·완료 등 모든 결재 메일이 davidkang·mujung 가리지 않고 **조용히 전 건 미발송**. 예외 메시지: `Specified permissions are not sufficient … Required: gmail.send / gmail.compose / gmail.modify / mail.google.com`.

**근본 원인:** 코드는 `GmailApp.sendEmail`을 쓰는데, `appsscript.json`의 `oauthScopes`에는 `script.send_mail`(=MailApp 전용)만 있었다. **`GmailApp`은 `gmail.send`(또는 더 넓은 Gmail 스코프)를 요구**한다. **새 Google 계정에 재배포하면서 OAuth 재승인이 누락**되어 드러났다.

**복구:**
1. `appsscript.json` `oauthScopes`에 `https://www.googleapis.com/auth/gmail.send` 한 줄 추가(기존 줄 유지).
2. 편집기에서 `testMailSelf` 재실행 → 권한 동의 팝업 통과(고급 → 계속).
3. 자가 테스트 메일 도착 확인.
4. **웹앱 [새 버전] 재배포**(executeAs: USER_DEPLOYING이라 필수).

**재발방지/주의:**
- 현재 매니페스트에는 `gmail.send` · `gmail.settings.basic` · `script.send_mail`이 들어 있다. `gmail.settings.basic`은 **From 별칭(§10.1) 발송에 필요**하고, `script.send_mail`은 MailApp 전용이라 **지금은 미사용(레거시)** 이다 — MailApp 복귀 여지를 남겨 둔 것이니 지우지 말 것.
- 발송 쿼터 조회는 **현재 구현돼 있지 않다.** 도입한다면 `MailApp.getRemainingDailyQuota()`가 유일한 수단이다(`GmailApp`에 대응 메서드가 없음). 이때 `script.send_mail` 스코프가 실제로 쓰이게 된다.
- 코드가 GmailApp으로 라벨/스레드를 읽거나 수정하게 되면 `gmail.modify`/`mail.google.com`이 추가로 필요할 수 있다(현재는 발송만 하므로 `gmail.send`로 충분).

## 3.2 ★"새 배포" 반복 → 활성 배포 2개 갈라짐 → 검수 버튼 미표시 (2026-06 해결)

**증상:** Code.gs·HTML을 최신으로 붙여넣고 "새 버전 배포"를 했는데도, 웹앱 화면(iframe)이 받는 `STATE.data.completedDocs[0].drafter`가 옛 값(성명 '강기종')으로 와서 **검수보고서 제출 버튼이 안 보임.** 반면 편집기에서 `getHomeDataForClient()`를 직접 실행하면 최신 값(이메일)이 나옴. **"편집기는 최신, 화면은 옛 코드"**라는 결정적 증상.

**근본 원인:** 지난 며칠간 배포할 때 "기존 배포 편집"이 아니라 **매번 "새 배포"를 만들어** 활성(Active) 배포가 **2개**가 됨. 최신 코드는 새 배포(`…ATwWeh…/exec`)에 들어갔는데, 사용자/화면은 기존 운영 URL(`…iL2l4m…/exec` = WEBAPP_URL)을 호출 → 옛 코드로 응답.

**진단법(매우 유용 — 코드 vs 배포 불일치 구분):**
- 편집기에서 함수를 직접 실행한 결과와, 웹앱 iframe 콘솔(`DevTools 컨텍스트를 userHtmlFrame으로 전환`, top 아님)에서 본 `STATE.data` 결과를 비교한다. **둘이 다르면 100% 배포 반영 문제**(코드 문제 아님).
- [배포 관리]에서 **활성 배포가 몇 개인지** 확인. 2개 이상이면 갈라진 것.

**복구:**
1. **기존 운영 배포(`iL2l4m…` = WEBAPP_URL)를 편집(연필) → 버전 "새 버전" → 배포** ⇒ 기존 URL 유지한 채 최신 코드 반영(사용자 북마크 안 깨짐).
2. 잘못 만든 새 배포(`ATwWeh…`)는 **보관처리(Archive)**.

**재발방지:** §2.8의 배포 절차를 **항상** 따른다. 절대 "새 배포"를 만들지 않는다.

## 3.3 Date 직렬화 null → "품의서를 불러올 수 없습니다: 오류" (초기 사고)

**증상:** 결재 이메일의 Viewer 링크에서 "품의서를 불러올 수 없습니다". 실행 로그는 `getRequisitionForViewer`가 "Completed"인데 클라이언트는 `null` 수신.

**근본 원인:** Sheets가 날짜 셀(발행일자/납기일)을 **JavaScript Date 객체**로 반환 → `google.script.run` 직렬화기가 처리하지 못해 **에러 없이 null 반환**.

**복구:** `toDateStr()` 유틸 신설(Date·문자열을 안전히 포맷) + 함수 전체 try/catch로 실제 오류 노출 + 반환 필드 `String()` 캐스팅. (이후 모든 서버 함수의 표준 패턴이 됨 — §2.1)

## 3.4 DRAFTER 컬럼이 이메일이 아니라 성명 (INSP 개발 중)

**증상:** INSP 제출 버튼 표시 판정에서 기안자 본인 비교가 어긋남. `completedDocs[0].drafter`가 이메일이 아닌 성명으로 옴.

**근본 원인:** `COL.DRAFTER`(D열)에는 성명이 저장됨. 기안자 이메일은 결재자 블록 0번(`COL.APPR_START + 2`)에 있음. (§2.9)

**복구:** drafter 신원 비교를 `COL.APPR_START + 2`(이메일) + 소문자 정규화로 통일. 이 사고가 §3.2의 배포 사고와 겹쳐 진단이 길어졌으므로, **둘을 함께 의심**할 것.

---

# 4. GAS 운영 제약 (반드시 설계에 반영)

| 제약 | 수치 | 대응 |
| --- | --- | --- |
| 단일 실행 시간 | 6분/실행 | 다첨부+PDF+이동이 겹치는 최종 트리거는 비동기 작업 큐로 분할 처리 |
| 일일 트리거 시간 | **90분/계정/일** (※6시간 아님) | 트리거 누적 모니터링. 빈 큐는 즉시 종료되도록 유지. 큐 트리거 1분 간격이라 누적 ~12분/일 |
| Gmail 발송 한도 | ~1,500통/일 (Workspace 계정) | `USER_DEPLOYING`이라 전 사용자가 배포자(David) 1풀 공유. 50명×다단계 알림이면 피크일 근접 가능. **잔여 쿼터 모니터링은 아직 미구현** — 도입 시 `MailApp.getRemainingDailyQuota()` 사용(§3.1) |
| 동시 사용자 | ~30명 (Business 기준) | LockService 30초. 50명 확장 시 락 경합 모니터링 |
| IP 주소 취득 | 불가 | 시스템로그 ipAddress는 향후 확장용 빈값 유지 |

> **⚠ 단일 계정 의존 (단일 장애점)**
> `USE_DEPLOYING` 배포라 모든 GAS 쿼터(Gmail·트리거·실행시간)가 배포자(David) 1계정에 귀속된다. 소유자 퇴사 = 시스템 정지 위험. **백업 운영자 지정·일일 자동 백업이 아직 미구현(Q-20)** — 최우선 운영 과제. (ERP 확장 시 모듈별 배포자 분산이 자연 해법 — §12)

---

# 5. 비동기 작업 큐 (Job Queue)

최종 승인 시 즉시 PDF·이동을 처리하면 응답이 느려지므로 큐로 분리했다.

| 요소 | 값/역할 |
| --- | --- |
| 저장소 | Script Properties의 'PENDING_JOB_QUEUE' (JSON 배열) |
| enqueueJob() | 최종 승인 시 작업 적재. **완료 메일 수신자 유무 조건 밖에 배치**(없어도 PDF는 생성) |
| processQueueTrigger() | 1분 간격 트리거 진입점 (installQueueTrigger로 1회 등록 필수). while 루프로 시간 한도 내 다건 처리 |
| QUEUE_MAX_ATTEMPTS | 3회 재시도 후 'failed' + 관리자 알림 |
| QUEUE_MAX_RUNTIME_MS | 4분 (트리거당 안전 한도) |
| QUEUE_SAFETY_MARGIN_MS | 30초 (다음 job 진입 여유 부족 시 중단) |
| Job 타입 | pdf_only(REQ→STAGING PDF), pdf_and_consolidate(PRC→PDF+FINAL 통합), **insp_pdf(INSP→PDF+FINAL/PO폴더+사진삭제)** |
| 운영 함수 | getQueueStatus(조회), retryFailedJobs(재시도), clearQueue(초기화, ADMIN), uninstallQueueTrigger(제거) |

**주의:** `installQueueTrigger()`를 실행하지 않으면 결재는 정상 처리되지만 PDF·FINAL 이동이 전혀 일어나지 않는다. 셋팅 시 가장 흔한 누락 지점. **또한, 관리자/수동 경로로 상태를 최종승인으로 만들면 enqueueJob이 호출되지 않아 PDF가 생성되지 않는다**(실제 사례 있음 — 이 경우 `generatePdfForDocument`+`consolidateToFinalByPrc`를 직접 호출해 수동 복구).

---

# 6. 시스템로그(AuditLog) 인프라 — 현재 상태와 확장 방법

## 6.1 현재 상태 (v2.2)

- '시스템로그' 시트(12컬럼)에 **append-only** 기록. `ensureAuditLogSheet()`가 시트를 자동 보장.
- **fail-safe**: `writeAuditLog`는 try/catch로 감싸 로그 실패가 본 업무를 중단시키지 않는다(throw 안 함, false 반환).
- 기록되는 이벤트: **관리자·보안 이벤트** — ADMIN_FORCE_DISCARD, ADMIN_FORCE_RELEASE_LOCK, ADMIN_ACCESS_DENIED.
- **INSP 라이프사이클 이벤트는 처음부터 연결됨** — INSP_SUBMIT, INSP_APPROVE, INSP_REJECT, INSP_RESUBMIT, INSP_FINALIZED. (FR-41을 INSP에서 선행 적용)
- REQ/PRC 문서 라이프사이클 이벤트(DOC_SUBMIT, DOC_APPROVE 등)는 **상수 자리만 있고 미연결** — 다음 개발 1순위.

## 6.2 12개 컬럼 (AUDIT_COL)

logId(UUID), timestamp(Date), actor(이메일), actorRole(admin/procurement/approver/requester/unknown), eventType, docNo, docToken, docType(REQ/PRC/INSP), targetUser, reason, payload(JSON string), ipAddress(빈값).

## 6.3 신규 이벤트를 추가하는 표준 절차

- AUDIT_EVENT 상수에 이벤트 타입 추가.
- 해당 비즈니스 로직 함수의 적절한 지점에서 `writeAuditLog({ eventType, docNo, docToken, docType, payload })` 호출.
- payload는 객체로 넘기면 자동 `JSON.stringify`된다.
- 관리자 액션이면 진입 직후 `assertAdminWithLog(eventType, payload)`로 권한 검증(거부 시 자동 로그+throw), `requireAdminReason(reason)`으로 사유 강제.

> **확장 권장:** INSP가 이미 문서 라이프사이클 로그를 연결해 둔 좋은 선례다. REQ/PRC 제출/승인/반려/PRC 픽업 함수에 `writeAuditLog` 호출 한 줄씩만 추가하면 FR-41(문서 라이프사이클 로그)을 완성할 수 있다.

---

# 7. 관리자 기능 (A 그룹)

| 코드 | 기능 | 상태 | 진입 경로 / 함수 |
| --- | --- | --- | --- |
| A-1 | 강제 폐기 | ✅ 구현 | `?action=admin&fn=discard` / adminListDiscardableForClient, adminForceDiscardForClient |
| A-4 | PRC 락 강제 해제 | ✅ 구현 | `?action=admin&fn=force_release_lock` / adminListClaimedLocksForClient, adminForceReleaseLockForClient |
| A-2 | 대리 승인/반려 | ✗ 예정 | 라우팅·이벤트 상수만 |
| A-3 | 결재자 변경 | ✗ 예정 | — |
| A-5 | 상태 강제 변경 | ✗ 예정 | 현재는 시트 직접 편집으로 복구 |

**구현된 관리자 함수의 공통 패턴 (신규 관리자 기능 추가 시 그대로 따를 것):**

- `*ForClient(payload)` 래퍼 → `_*Core(payload)` 본체 구조.
- `_Core` 진입 직후: `assertAdminWithLog(이벤트, 컨텍스트)` → `requireAdminReason(payload.reason)`.
- 락이 필요한 변경은 withLock 안에서 `getActiveUserEmail()`을 **다시 확인**(락 대기 중 컨텍스트 변화 방어).
- 보호 규칙: 최종 완료 문서(`docType==='PRC' && status==='최종승인(PRC)'`)는 강제 폐기 대상에서 제외.
- 처리 후: `writeAuditLog` 기록 + 관계자(`_send*Notifications`)에게 알림.

---

# 8. 데이터 모델 핵심

## 8.1 품의서목록 (32컬럼 + 결재자 블록) — v2.1 스키마

| 범위 | 컬럼 |
| --- | --- |
| A~Q (0~16) | 제출일시, 품의번호, 발행일자, 기안자(**성명**), 부서, 품의제목, 용도, 업체명, 사업자번호(VENDOR_EMAIL), 업체담당자, 연락처, 납기일, 납품장소, 품목내역, 합계금액, 부가세포함합계, 비고 |
| R~Z (17~25) | 결재상태, 토큰, Drive폴더ID, 결재자수, 현재결재순번, 재상신횟수, 반려이력누적, 첨부파일ID목록, 이동상태 |
| AA~AD (26~29) | DocType(REQ/PRC), ParentDocId(PRC→REQ 토큰), ClaimedBy(락 점유자), ClaimedAt |
| **AE~AF (30~31)** | **PURCHASE_METHOD(구매방법), PAYMENT_INFO(구매조건/지급정보)** — v2.1 추가 |
| AG~ (32~) | 결재자 블록 × N — 블록당 6컬럼(label, name, email, status, processedAt, comment). **`APPR_START=32`, `APPR_COLS=6`** |

> **⚠ APPR_START가 30→32로 이동했다.** v2.1 마이그레이션(`runV21PurchaseConditionMigration`)이 ClaimedAt 뒤에 구매방법·구매조건 2칸을 삽입했다. 모든 결재 블록 참조는 **APPR_START 기준 상대 오프셋**이라 이동이 안전하다. 절대 숫자 리터럴로 결재 블록을 찾지 말 것. **기안자 이메일 = `COL.APPR_START + 2`**(§2.9).

## 8.2 검수보고서목록 (51컬럼) — INSP_COL (insp-v1.0)

`INSP_TOTAL_COLS = APPR_START(27) + MAX_APPROVERS(4) × APPR_COLS(6) = 51`.

| 범위 | 컬럼 |
| --- | --- |
| 0~14 | SUBMIT_AT, DOC_NO(`{REQ docNo}-NN`), SEQ(회차 1~99), PRC_TOKEN(조인키), REQ_TOKEN, PO_NO, REQ_NO, SUBJECT, VENDOR_NAME, DRAFTER(이메일), DRAFTER_NAME(성명), DEPT, ISSUE_DATE, RECEIVED_DATE, RECEIVED_NOTE |
| 15~26 | VERDICT(합격/불합격), IS_FINAL(Y/''), COMMENT, STATUS(한글), TOKEN(UUID), DRIVE_ID(사진 임시폴더), APPR_COUNT, APPR_IDX, RESUB_COUNT, REJECT_LOG, PHOTO_LIST(JSON), MOVE_STATUS(STAGING/FINAL) |
| 27~50 | 결재자 블록 ×4 — **블록0=기안자(자기결재), 블록1~3=결재자**. `APPR_START=27`, `MAX_APPROVERS=4`. 접근은 `inspApprCol(stageIdx, offset)` 헬퍼만 |

설계 원칙: PO_NO·REQ_NO·SUBJECT·VENDOR_NAME·DRAFTER_NAME·DEPT·ISSUE_DATE는 **제출 시점 스냅샷(비정규화)** — 원본 변경과 동기화하지 않음(검수 시점 고정, PDF 화이트리스트와 동일 사상). `_inspSchemaSelfCheck()`가 인덱스 연속성을 배포 전 검증.

> **주의:** INSP.gs 코드 주석 일부에 "27 + 3×6 = 45"라는 **옛 설계 흔적**이 남아 있으나, 실제는 기안자 블록을 포함해 **4블록·51컬럼**이다(Step 4에서 기안자 포함으로 변경). `testInspStep1`은 `INSP_TOTAL_COLS`(=51)로 검증한다.

## 8.3 공통 규약

- 행 조회는 토큰 기반 TextFinder/순회 사용. 토큰은 1회성 보안 + 행 식별 겸용.
- REQ↔PRC는 1:1 매핑. PRC는 자체 토큰 + ParentDocId로 원본 REQ 추적. INSP는 PRC_TOKEN으로 PRC를, REQ_TOKEN으로 원본 REQ를 추적.
- 락 점유는 status를 바꾸지 않고 ClaimedBy/ClaimedAt만 채운다(부모 REQ status는 '최종승인' 유지).
- 컬럼 인덱스는 0-based 상수로 접근하고, getRange에는 +1 한다.
- 일괄 쓰기는 `batchUpdate(sheet, rowNum, updates)`로 묶어 setValues 호출 수를 줄인다.

---

# 9. INSP(검수보고서) 모듈 — Phase 2 (신설)

## 9.1 개요

최종승인(PRC)된 발주 건에 대해 입고/검수 결과를 회차별로 제출·결재하고, 검수 PDF를 PO 폴더에 보관한다. 별도 파일 `INSP.gs`(~1,450줄, Step 1~6 누적)로 구현. Code.gs와 전역 스코프를 공유하므로 CONFIG·getOrCreateSheet·writeAuditLog·enqueueJob 등을 그대로 사용.

## 9.2 확정 사양 (Q-INSP-01~12)

- 문서번호: `{REQ docNo}-NN` 2자리 자동 채번(예 `TH-AP-26-029-01`).
- 결재선: 매번 신규 선택, 기안자(고정, 블록0=자기결재) + 결재자 최대 3인.
- 복수 제출 허용(분할 납품 회차별), 아코디언 UX. 앞 회차 결재중에도 다음 회차 제출 가능.
- 판정: 합격/불합격 2단계. 입고 내역 1줄(필수).
- 사진: 최대 6장, 장당 5MB(업로드 한도). **클라이언트 canvas 압축(긴 변 1600px, JPEG 0.8)** 후 base64 업로드. PDF 삽입 후 **원본·임시폴더 삭제(미보관)**.
- 종결: 최종 검수(IS_FINAL=Y) 회차가 최종승인(INSP)되면 품의 종결 + 추가 제출 차단. **차단은 서버 검증(`getInspFormDataForClient`/`_submitInspCore`)으로 강제**(클라이언트 버튼 숨김만으로 불충분).
- 열람: 일반은 본인 기안 건, 구매팀은 전체.
- 데이터: 기존 스프레드시트에 검수보고서목록 탭 추가. 사진은 STAGING에 임시 저장.

## 9.3 주요 서버 함수 (INSP.gs)

| 함수 | 역할 |
| --- | --- |
| ensureInspSheet / _inspSchemaSelfCheck / inspApprCol | 시트 보장, 스키마 자가점검, 결재 블록 셀 계산 |
| _getInspGroupsByPrc(ss) | PRC token별 그룹핑 (홈 결재완료 아코디언용). 동일 스프레드시트 read 1회 |
| getInspFormDataForClient(prcToken, resubToken) | 작성 폼 초기 데이터(스냅샷 프리필, 회차, 종결 차단, 재상신) |
| submitInspForClient → _submitInspCore | 제출/재상신. 락 안=검증+회차채번+행기록, 락 밖=사진저장+1차결재메일+감사로그 |
| getInspDecisionMetaForClient / getInspViewerDataForClient | 경량 결재 메타 / 전용 뷰어(본문·사진 base64·myApproverIdx) |
| processInspDecisionFromClient | 승인/반려. 본인 단계 `_findMyInspIdx`로 재판정. 최종승인 시 status='최종승인(INSP)' + insp_pdf enqueue |
| _processInspPdfJob / generateInspPdf / _cleanupInspPhotos | 큐 처리: PDF 생성(사진 base64) → FINAL/PO폴더 저장 → 사진·임시폴더 삭제. 멱등(MOVE_STATUS=FINAL이면 스킵) |
| _getInspMenusForClient(ss, actor, isProc, isAdmin) | 홈: 내 검수 결재 대기 / 최종 완료(INSP) / 관리자 전사 대기 목록 |

## 9.4 화면·라우팅

- 신규 HTML 3종: `Procurement_INSP_Form`(작성/재상신), `Procurement_INSP_Viewer`(전용 뷰어·결재, 청록 #0e7d72 테마), `INSP_PDF_Template`(PDF).
- doGet 추가: `?action=insp_form&token={PRC토큰}[&resub={INSP토큰}]`, `?action=insp_view&token=...&idx=...`.
- 결재 메일: 확인=전용 뷰어(insp_view), 승인/반려=경량 `Procurement_Approval`(템플릿 변수 `docKind='INSP'`로 분기 — REQ/PRC와 동일 3버튼 레이아웃). 제목 prefix `[검수보고서 결재요청]`.
- 홈: 기본 메뉴를 3그룹(기본/결재/조회)·5항목으로 재구성, "결재 완료" 뷰에 아코디언으로 검수 제출/결재 진입.

## 9.5 INSP 특이 패턴 (개발 시 준수)

- `enqueueJob`은 완료 메일 조건 밖에 배치(완료 메일 수신자 없어도 PDF 생성).
- 홈 대시보드는 INSP 함수에 `typeof` 가드를 둬 INSP 파일이 없어도 graceful degrade.
- 사진 base64 임베드: Drive 썸네일 URL은 웹앱 iframe에서 미인증으로 깨지므로, PDF·뷰어 모두 **STAGING Drive에서 바이트를 읽어 data URL로 직접 임베드**.
- 진단: `INSP_DIAG.gs`의 `debugInspPending()`이 행별 결재 블록·본인 단계·대기 표시 판정을 덤프(검수 대기 미표시 원인 추적).

---

# 10. 이메일 아키텍처 — 이중 경로 (의도적 분리, 통합 금지)

> 📌 *이 분리는 §3.1 스코프 사고를 겪은 뒤 David가 명시적으로 내린 설계 결정이다. **절대 단일 함수로 통합하지 말 것.***
>
> ⚠️ **2026-07-19 변경(d0b63e1)**: 발신 주소를 `admin-inlct@inlct.com`으로 통일하면서 관리자 알림 2종(A-1/A-4)도 **MailApp → GmailApp으로 교체**했다. `MailApp`이 From 별칭을 지원하지 않기 때문이다. **현재 코드에 `MailApp` 호출은 한 건도 없다**(주석에만 등장). 경로 분리 자체는 유지되며, 분리 축이 *라이브러리*에서 **호출 방식(재시도 유무)** 으로 바뀌었을 뿐이다.

| 경로 | 함수 / 위치 | 발송 방식 | 용도 |
| --- | --- | --- | --- |
| 일반 결재 알림 | `sendEmailWithRetry` (Code.gs §14) | GmailApp + **3회 재시도**, 최종 실패 시 throw | 승인요청·반려·완료. INSP도 재사용(`_sendInspApprovalEmail` 등) |
| 관리자 강제폐기 알림 | `GmailApp.sendEmail` 직접 (`[A-1]` 태그, Code.gs:1880) | GmailApp 직접 — **재시도·throw 없음**, 별칭 실패 시 `_sendWithoutAlias` **1회 폴백** | 강제 폐기 관계자 통지 |
| 관리자 락해제 알림 | `GmailApp.sendEmail` 직접 (`[A-4]` 태그, Code.gs:2195) | 〃 | 락 강제 해제 점유자 통지 |
| 관리자 오류 알림 | `notifyAdminError` (Code.gs §14) | GmailApp **직접**(재시도/throw 없음) | 오류 통지 — sendEmailWithRetry로 라우팅 금지 |

**분리 이유(fault isolation):** 단일 깔때기는 한 곳(예: 스코프)이 막히면 전 경로 메일이 동시에 죽고 "어느 메일이 죽었는지" 진단 단서가 사라진다. 특히 오류 알림이 재시도 경로를 타면 **발송 실패 → 오류 알림 발송 → 또 실패**의 재귀에 빠진다. 분리하면 한쪽이 실패해도 다른 쪽이 살고, 로그 태그(`[A-1]`/`[A-4]`)로 즉시 구분된다.

## 10.1 From 별칭 (CONFIG.MAIL_FROM) — 2026-07 도입

모든 발송은 `_mailOptions()`(Code.gs:3172)를 거쳐 `from: admin-inlct@inlct.com` + `name: '구매품의 시스템'`을 붙인다. 스크립트 실행 계정(David)과 무관하게 **대표 주소로 나가게** 하는 것이 목적이다.

- **사전 조건:** David Gmail → 설정 → 계정 → '다른 주소에서 메일 보내기'에 별칭이 **등록·인증**돼 있어야 하고, 매니페스트에 `gmail.settings.basic` 스코프가 있어야 한다(누락 시 `Specified permissions are not sufficient`).
- **폴백:** 별칭 발송이 실패하면 `_sendWithoutAlias()`(Code.gs:3188)가 **원인을 따지지 않고 1회는 별칭 없이 재발송**한다. 예외 문구로 원인을 분기하면 문구가 바뀔 때 폴백이 조용히 죽기 때문. 최악의 경우에도 발신자만 David로 표시될 뿐 알림은 유실되지 않는다.
- **점검:** GAS 에디터에서 `sendTestMailFromAlias()` 수동 실행. 운영 중 **수신 메일의 발신자가 `davidkang@inlct.com`으로 보이면 폴백이 돌고 있다는 신호**다(별칭 인증·스코프 점검 필요). 별칭 장애를 **자동으로 알려주는 경보는 아직 없다** — 발신자 육안 확인이 사실상 유일한 조기 신호다.
- **폴백 적용 범위:** 4개 발송 경로 **전부**. 2026-07-22에 A-1/A-4에 누락돼 있던 폴백을 보강했다(그 전까지는 별칭이 죽으면 관리자 통지 2종만 조용히 전멸했다).
- **경보 순환 의존 차단:** `ADMIN_NOTIFY_EMAILS`에 `admin-inlct` 외에 **`davidkang@inlct.com`을 함께 둔다**(Code.gs:61). 두 주소가 같으면 "admin-inlct 계정 장애"가 원인일 때 경보 메일까지 함께 죽는다. **admin-inlct를 배열 첫 번째로 유지**할 것 — `sendTestMailFromAlias()`가 `[0]`을 쓴다.

## 10.2 발송 기록(Sent)이 남는 위치 — 자주 오해하는 지점

- `executeAs: USER_DEPLOYING`이라 GmailApp은 **항상 배포자(David) 사서함**으로 발송한다. 따라서 보낸편지함 기록도 `davidkang@inlct.com`에만 쌓인다.
- `admin-inlct@inlct.com`은 **표시용 From 별칭일 뿐**이라 그 계정 보낸편지함은 **비어 있는 것이 정상**이다. 발송 감사 추적은 David 사서함에서 한다.
- 관리자 통지(A-1/A-4)가 **폴백까지 실패**하면 `_recordAdminNotifyFailure()`가 ① 시스템로그 시트에 `EMAIL_SEND_FAIL` 행을 남기고 ② `notifyAdminError()`로 관리자 메일을 시도한다. **시트 기록이 먼저**인 이유는 메일 경로가 통째로 죽은 상황에서도 살아남는 유일한 흔적이기 때문이다. A-1은 수신자별로 부르지 않고 **루프 밖에서 1회 집계 통보**한다(메일 폭풍 방지).
- "MailApp은 보낸편지함에 기록이 안 남는다"는 흔한 통설은 **사실이 아니다.** MailApp 시절(~2026-07-19)에 나간 A-1/A-4 메일도 David 사서함에 `SENT` 라벨로 남아 있다(2026-05~06 실물 확인). MailApp과 GmailApp의 실질적 차이는 **Sent 기록 여부가 아니라 From 별칭 지원 여부**다.

---

# 11. UI 디자인 시스템 (inlc-ppt 디자인 언어 적용)

홈 화면(`Procurement_Home.html`)에 사내 PPT 디자인 시스템을 적용했다. **JavaScript 렌더 로직이 참조하는 CSS 변수명·클래스명은 반드시 보존**하고 표현(색/모서리/그림자)만 변경하는 것이 철칙(google.script.run 호출·픽업 모달·INSP 아코디언이 깨지지 않도록).

| 토큰 | 값 / 용도 |
| --- | --- |
| 브랜드(REQ/일반) | 와인 레드 `#83161A` (그라데이션). 버튼·아바타·활성 네비·KPI 강조 |
| 구매팀(PRC) | 딥 블루 `#1E5F9E` (그라데이션). PRC 문서번호·구매팀 역할 pill·인박스·픽업 버튼/모달 |
| 검수(INSP) | 청록 `#0e7d72` / 시트 헤더 `#e4f4f2`. 검수 관련 요소 |
| 상태 배지 | **기존 의미색 그대로 보존** — 검토중(파랑), 결재중/재상신(주황), 최종승인/PRC생성됨(보라), 반려(빨강) |
| 폰트 | Pretendard (오프라인 시 시스템 폰트 폴백) |
| 캔버스/모서리 | 회색 캔버스 `#F4F4F4`. 카드 12px / 버튼 9px / pill 999px / 배지 6px |
| 로고 | 회사 로고 base64 임베드 ("IL" 텍스트 박스 대체) |

작업 방식: **production 파일을 건드리기 전 vanilla HTML 토글 시뮬레이터로 디자인을 먼저 검증**(예: dashboard_simulator_v*.html)한 뒤 적용. 적용 시 `<script>` 블록은 최소 변경(신규작성 버튼 색, nav-count zero 클래스 조건 등만).

---

# 12. ERP 통합 로드맵 (팀 웹앱 통합)

팀(경영·회계·SCM)이 각자 GAS 웹앱을 만들고 있으며, 구매결재시스템과 동일 구조를 따른다. 핵심 결론·방향:

- **URL 링크 연결은 "포털(런처)"이지 "ERP 통합"이 아니다.** 빠른 1단계지만 데이터는 사일로로 남는다.
- **진짜 통합의 핵심은 공유 데이터 계층** — ① 공유 마스터데이터 Sheet(거래처·직원·부서·계정과목, 공통 ID 참조; Sheet 읽기는 일일 쿼터 없음 → 저비용), ② 문서번호·상태값·키 컬럼 명명 규칙 통일, ③ 공유 GAS 라이브러리(인증·로깅·이메일·PDF 공통 유틸; **라이브러리는 호출 프로젝트 컨텍스트에서 실행되어 쿼터가 각 배포자에 귀속** → 중복 제거 + 쿼터 분산 동시 달성).
- **배포 주체:** 각 모듈은 담당자가 `USER_DEPLOYING`으로 배포하고, 공유 Sheet에는 팀 계정에 편집 권한 부여 → 알림 메일이 모듈별 배포자에서 나가 **쿼터 자연 분산**(단일 장애점·병목 완화).
- 단계: (1)URL 포털 → (2)공유 마스터+라이브러리 → (3)통합 리포팅(데이터 커지면 BigQuery/Connected Sheets) → (4)필요 시 AppSheet/정식 백엔드 전환. (AppSheet는 복잡 워크플로우엔 보류, 단순 신규 모듈엔 적합 — HR 학습/운동지원 등.)

---

# 13. 코딩 규약 및 권장 작업 방식

## 13.1 함수 네이밍 패턴

- 클라이언트 호출용: `xxxForClient(payload)` (plain object 반환) 또는 doPost 경로의 `_xxxCore(payload)`.
- 내부 헬퍼: 언더스코어 접두사(`_renderTemplate`, `_findMyInspIdx`, `inspApprCol` 등).
- 날짜 변환: `toDateStr` / `toDateTimeStr`. 문서번호 정규화: `normalizeDocNo`. 회차 패딩: `_pad2`.

## 13.2 변경 시 권장 절차 (이 시스템 개발자의 작업 스타일)

- 단일 파일(Code.gs)이므로 섹션 주석(`// N. …`)으로 구획을 찾는다. 22개 섹션 헤더가 목차 역할.
- **UX가 걸린 변경은 단일 HTML 시뮬레이터(외부 의존성 없음)로 먼저 검증**한 뒤 구현한다.
- 복잡 기능은 구조화된 Q&A(Q-XX-NN 형식)로 사양을 먼저 확정한다.
- 단계별로 확인받으며 진행한다. 큰 일괄 재작성보다 작은 검증 가능한 변경을 선호한다. 각 Step은 독립 배포·테스트 가능하게 나눈다.
- 수치(쿼터, 한도)·API 식별자·동작 설명은 **정확히** 적는다(개발자가 부정확한 수치 — 예: 트리거 90분을 6시간으로 — 를 적극 교정한다).
- **production HTML 수정 시 CSS 변수명·클래스명 보존**(JS 렌더가 참조).
- 코드 출력은 편집기에 **수동 복붙**되며 자동 반영되지 않는다. 모든 시스템 UI·문서·코드 주석은 한국어, 응답은 존댓말.

## 13.3 테스트 함수 활용

편집기 직접 실행: testCheckDocNo, testApproverList, testHomeData, testAuditLogInfra, testA4ListLocks/testA4ForceReleaseDryRun, **testInspStep1~6(INSP 단계별), debugInspPending(검수 대기 진단), debugWhichHome/debugCompletedPayload(배포 반영 진단)**.

---

# 14. 알려진 위험·미해결 과제

| 항목 | 내용 / 위험 |
| --- | --- |
| 단일 장애점 (최우선) | 배포자(David) 1계정 의존. 일일 자동 백업·백업 운영자 **미구현**(Q-20). 데이터 손상 시 복구 수단 부족 |
| REQ/PRC 문서 라이프사이클 로그 미연결 | 인프라·INSP 선례는 있으나 REQ/PRC 제출/승인 등 이벤트 미기록 |
| PRC 락 자동 해제 트리거 미등록 | TTL(24h) 설정은 있으나 releaseExpiredPrcLocks 시간 트리거 등록 필요(Q-19) |
| 상태 강제 변경 화면 부재 | 오류 복구를 시트 직접 편집에 의존(A-5 미구현) |
| REQ vs PRC 비교/필드변경 추적 미구현 | FR-46, FR-57, Q-18 |
| 서명 이미지·프리셋 미구현 | Q-04, Q-05 미결 |
| 50명 확장 시 락 경합·Gmail 한도 | 동시성·쿼터 모니터링. 임계 도달 시 모듈 분산 배포/별도 서버 재평가(Q-08) |
| INSP 주석의 컬럼 수 표기 불일치 | "45"로 적힌 옛 흔적 존재 — 실제 51. 혼동 주의(§8.2) |

---

# 15. 향후 확장 (Phase 2 / HR 문서)

- INSP 이후 Phase 2 로드맵: PO 자동생성/거래처 메일 발송, 거래명세서·세금계산서 관리, 거래처 마스터데이터 관리, 통합 리포팅.
- HR 워크플로우 모듈(학습 프로그램 신청, 운동/웰니스 지원 요청 등) — 경량 단일 단계 결재, 구매시스템과 별도 GAS 프로젝트. AppSheet 적합성 검토 대상.
- 결재 워크플로우·시스템로그·작업 큐 인프라는 재사용 가능하도록 설계되어 있어, 양식 메타와 PDF 템플릿 중심으로 확장하는 방향이 자연스럽다.

---

# 부록. 빠른 체크리스트 (코드 수정 전 확인)

| ☐ | 확인 항목 |
| --- | --- |
| ☐ | Date 반환값을 toDateStr/toDateTimeStr로 변환했는가 |
| ☐ | 서버 함수를 try/catch로 감싸고 stack을 메시지에 담았는가 |
| ☐ | 클라이언트 통신을 google.script.run으로 했는가 (fetch 금지) |
| ☐ | 상태값을 한글로 비교했는가 (영문 아님 / INSP는 '최종승인(INSP)') |
| ☐ | 첨부·사진·이메일을 락 밖에서 처리했는가 (회차 채번은 락 안) |
| ☐ | PDF payload에 금지 필드(remarks/REJECT_LOG 등)가 없는가 |
| ☐ | Drive 호출에 supportsAllDrives: true가 있는가 |
| ☐ | 컬럼 접근을 COL/INSP_COL 상수로 했는가 (숫자 리터럴·혼용 금지) |
| ☐ | 기안자 신원 비교를 APPR_START+2(이메일)+소문자로 했는가 (DRAFTER 성명 금지) |
| ☐ | 결재 단계 idx를 서버가 actor로 재판정했는가 (동일인 다단계 = curIdx 우선) |
| ☐ | enqueueJob을 완료 메일 조건 밖에 배치했는가 |
| ☐ | GmailApp 신규 사용 시 gmail.send + gmail.settings.basic(별칭) 스코프가 매니페스트에 있는가 |
| ☐ | 신규 발송 코드가 `_mailOptions()`를 거쳐 From 별칭을 적용했는가 (GmailApp.sendEmail 생짜 호출 금지) |
| ☐ | 이메일 이중 경로(일반=재시도 / 관리자=직접)를 통합하지 않았는가 |
| ☐ | 수정 후 [배포 관리→기존 배포 편집→새 버전]으로 재배포했는가 (새 배포 금지) |

*── 구매결재시스템 AI Context Document v2.3 끝 (2026-07-22) ──*
*v2.3 변경: §10 이메일 아키텍처를 현재 코드에 맞게 정정(MailApp→GmailApp 전환 반영), §10.1 From 별칭 / §10.2 Sent 기록 위치 신설, §3.1·§4의 MailApp 관련 서술 정정.*