# 구매결재시스템 — AI Context Document

**InLC Technology 구매결재시스템 (Procurement Approval System)**
**후속 개발 시 AI 어시스턴트가 참고하는 컨텍스트 문서**

| 항목 | 내용 |
| --- | --- |
| **문서 목적** | 향후 개발·디버깅·기능 확장 시 AI(또는 신규 개발자)가 이 시스템의 아키텍처, 지금까지 겪은 오류, 반드시 지켜야 할 주의사항을 빠르게 파악하도록 돕는다 |
| **대상 시스템** | 구매결재시스템 v2.1 (Phase 1, Google Workspace 기반) |
| **런타임** | Google Apps Script (GAS) Web App, "Execute as: Me" 배포 |
| **작성일** | 2026-05-29 |
| **연계 문서** | 구매결재시스템_요구서_Phase1_v2.1, 구매결재시스템_셋팅매뉴얼_v1.2 |

> **📌 이 문서를 읽는 AI에게**
> 이 시스템은 1인 개발자가 운영하는 사내 시스템이다. ERP 없이 Google Workspace(Apps Script + Sheets + Drive + Gmail)만으로 동작한다. 화려한 재작성보다 **기존 패턴 준수, 작은 단위 확인, 안전한 점진 변경**이 항상 우선이다. 아래 §2의 "치명적 함정"을 위반하면 조용한 데이터 손실이나 전 사용자 장애로 이어진다. 코드를 수정하기 전에 반드시 §2와 §5를 먼저 확인하라.

---

# 1. 시스템 한눈에 보기

## 1.1 무엇을 하는 시스템인가

기안자가 구매품의서(REQ)를 작성→결재받고, 1차 승인된 REQ를 구매팀이 픽업하여 구매팀 품의서(PRC)로 재기안→최종 결재받으면, PO번호 기준 폴더에 PDF·첨부가 자동 보관되는 결재 워크플로우 시스템이다. 약 25명(→50명 확장) 사용.

## 1.2 핵심 자산 (v2.1 시점)

| 자산 | 내용 |
| --- | --- |
| Code.gs | 약 4,000줄, 22개 섹션 (단일 파일 백엔드) |
| HTML 9종 | Procurement_Home, Procurement_Form, Procurement_PRC_Form, Procurement_Viewer, Procurement_Approval, Procurement_Reject, PDF_Template, Procurement_Admin_Discard, Procurement_Admin_ForceReleaseLock |
| Google Sheets (1파일·3시트) | 품의서목록(30컬럼+결재자블록), 결재자목록, 시스템로그(12컬럼, 자동생성) |
| Google Drive | STAGING_ROOT(진행중 임시), FINAL_ROOT(최종 보관) |
| 고급 서비스 | Drive API v3 (식별자 'Drive') — 필수 |
| 자동 트리거 | processQueueTrigger (1분 간격, 필수) |

## 1.3 데이터 흐름 (요약)

REQ 작성(검토중) → 순차결재(결재중) → 1차 최종승인(최종승인) → 구매팀 인박스 노출 → 선착순 락 점유(ClaimedBy 기록) → PRC 프리필·작성 → PRC 제출(부모 REQ: PRC생성됨) → 구매팀 결재 → PRC 최종승인(최종승인(PRC)) → **작업 큐 등록** → 1분 트리거가 PDF 생성 + FINAL 폴더 통합 이동.

## 1.4 권한 모델 (누적형)

전 사용자 공통 메뉴(5) → isProcurement(구매팀) 추가 메뉴(6) → isAdmin 관리자 섹션. 권한은 코드 내 CONFIG 배열(`PROCUREMENT_TEAM_EMAILS`, `ADMIN_EMAILS`)로 판정한다. 사용자 식별은 항상 서버에서 `Session.getActiveUser().getEmail()`로 해결한다.

---

# 2. 치명적 함정 (Critical Pitfalls) — 위반 금지

> 📌 *아래 항목들은 과거에 실제로 장애를 일으켰거나 일으킬 수 있는 것들이다. 코드 작성·수정 시 매번 확인하라.*

## 2.1 google.script.run + Date 객체 = 조용한 null

Sheets에서 읽은 Date 객체를 `google.script.run`으로 클라이언트에 반환하면 **에러 없이 null로 직렬화**되어, 클라이언트는 원인 모를 빈 값을 받는다.

- **반드시** 서버 반환 전에 `toDateStr()` / `toDateTimeStr()` 유틸로 문자열 변환할 것.
- 클라이언트에서도 `String()` 캐스팅으로 방어.
- 모든 서버 함수는 try/catch로 감싸고, catch에서 `err.toString() + err.stack`을 메시지에 담아 반환할 것. 그렇지 않으면 실패가 조용한 null로 나타나 디버깅이 불가능해진다.

## 2.2 앱 내부에서 fetch(WEBAPP_URL) 절대 금지

조직 도메인 배포 환경에서 클라이언트가 `fetch(WEBAPP_URL)`을 호출하면 CORS/401로 차단된다.

- 모든 클라이언트→서버 통신은 **google.script.run**을 사용한다.
- `google.script.run`은 ContentService 반환 객체를 받지 못한다. 서버 함수는 **plain object**(`{ok:true, message:'…', data:…}`)를 반환해야 한다.
- (doPost 경로는 별도로 `jsonResponse()` 래퍼를 통해 ContentService를 쓰지만, 이는 외부 POST 호환용이다. 신규 클라이언트 호출은 `google.script.run` + `*ForClient` 함수 패턴을 따른다.)

## 2.3 상태값은 한글이다 (영문 아님)

요구서 §8.2는 영문 논리 상태(SUBMITTED, APPROVED_1ST 등)를 쓰지만, **실제 코드는 한글 문자열**을 결재상태(R열)에 저장한다. 영문 상태로 비교하면 어떤 필터도 매칭되지 않는다.

| 논리 상태 | 실제 코드 값 |
| --- | --- |
| SUBMITTED | 검토중 |
| IN_APPROVAL | 결재중 (다음결재자명) |
| 업로드 실패 | 업로드오류 |
| REJECTED / 재제출 | 반려 / 재상신 |
| APPROVED_1ST | 최종승인 |
| PRC 생성됨(부모 REQ) | PRC생성됨 |
| FINAL_APPROVED | 최종승인(PRC) |
| 폐기 | 폐기 |

- 인박스 필터 조건은 `status === '최종승인' && docType === 'REQ'`.
- 진행 중 상태는 `'결재중 (' + 다음결재자 + ')'` 형태로 동적이므로 `indexOf('결재')` 같은 부분 매칭을 쓰는 코드가 있다. 정확 일치 비교 시 주의.
- 최종승인 여부 판정은 `status.indexOf('최종승인') >= 0` 패턴을 자주 쓴다('최종승인'과 '최종승인(PRC)' 모두 포함).

## 2.4 LockService 범위 최소화 — 첨부·이메일은 락 밖에서

withLock 안에는 **행 추가/토큰 발급/폴더 생성처럼 짧고 경합 위험이 큰 작업만** 넣는다.

- 첨부 업로드(가장 느림), 이메일 발송(재시도 포함)은 **반드시 락 밖**에서 처리한다. 락 안에 두면 LOCK_WAIT_MS(30초) 경합·타임아웃이 발생한다.
- 표준 단계 구성: ①락 밖 사전검증(용량 등) → ②최소 락(행/토큰/폴더) → ③락 밖 첨부 업로드 → ④첨부 메타 기록 → ⑤락 밖 이메일.
- withLock은 `LockService.getScriptLock()`을 쓰며 waitLock 실패 시 "시스템이 바쁩니다" 에러를 던진다.

## 2.5 PDF 페이로드 화이트리스트 — 내부 메모 노출 금지

PDF에는 결재 시스템 내부에서만 보여야 하는 내용이 **절대** 들어가면 안 된다.

- `buildPdfPayload()`는 화이트리스트 방식으로 필요한 필드만 담는다.
- 금지 필드(FORBIDDEN): remarks, prcMemo, rejectLog, rejectHistory, claimedBy, claimedAt, fieldChanges 등. 자체 점검 로직이 금지 필드 발견 시 throw하도록 되어 있다.
- PDF 관련 코드를 수정할 때 이 화이트리스트를 우회하거나 payload에 원본 row를 통째로 넣지 말 것.

## 2.6 Drive API는 supportsAllDrives: true 필수

공유 드라이브에서 파일을 다루므로 Drive Advanced Service 호출 시 `supportsAllDrives` 옵션을 빠뜨리면 파일 이동/생성이 실패한다. 또한 Apps Script 편집기에서 Drive 고급 서비스(식별자 'Drive', v3)가 추가되어 있어야 한다(누락 시 `consolidateToFinalByPrc`/`moveAttachmentsToFinal` 실패).

## 2.7 FINAL 통합 이동은 PRC 최종 승인 시점에만

PO번호는 PRC 단계에서 결정되므로, FINAL 폴더 통합 이동은 **REQ 1차 승인이 아니라 PRC 최종 승인 시점**에 일어난다. REQ 단계 PDF는 STAGING에만 생성된다(pdf_only job). 이 순서를 바꾸면 폴더명(PO번호)을 결정할 수 없다.

## 2.8 코드 수정 후 반드시 [새 버전] 재배포

CONFIG나 코드를 고치고 저장만 하면 사용자에게는 이전 버전이 노출된다. [배포 관리] → [새 버전]으로 재배포해야 반영된다. URL은 동일 유지된다.

---

# 3. GAS 운영 제약 (반드시 설계에 반영)

| 제약 | 수치 | 대응 |
| --- | --- | --- |
| 단일 실행 시간 | 6분/실행 | 다첨부+PDF+이동이 겹치는 최종 트리거는 비동기 작업 큐로 분할 처리 |
| 일일 트리거 시간 | **90분/계정/일** (※6시간 아님) | 트리거 누적 모니터링. 빈 큐는 즉시 종료되도록 유지 |
| Gmail 발송 한도 | ~1,500통/일 | 'Execute as: Me'로 전 사용자가 소유 계정 1풀 공유. 결재선 길이·리마인더 빈도 고려 |
| 동시 사용자 | ~30명 (Business 기준) | LockService 30초. 50명 확장 시 락 경합 모니터링 |
| IP 주소 취득 | 불가 | 시스템로그 ipAddress는 향후 확장용 빈값 유지 |

> **⚠ 단일 계정 의존**
> 소유 계정의 Gmail 한도·트리거 시간이 시스템 전체에 공유된다. 개인 계정이 아닌 **조직 운영용 계정**으로 배포할 것. 소유자 퇴사 = 시스템 정지 위험(단일 장애점).

---

# 4. 비동기 작업 큐 (Job Queue)

최종 승인 시 즉시 PDF·이동을 처리하면 응답이 느려지므로 큐로 분리했다.

| 요소 | 값/역할 |
| --- | --- |
| 저장소 | Script Properties의 'PENDING_JOB_QUEUE' (JSON 배열) |
| enqueueJob() | 최종 승인 시 작업 적재 |
| processQueueTrigger() | 1분 간격 트리거 진입점 (installQueueTrigger로 1회 등록 필수) |
| QUEUE_MAX_ATTEMPTS | 3회 재시도 후 'failed' + 관리자 알림 |
| QUEUE_MAX_RUNTIME_MS | 4분 (트리거당 안전 한도) |
| QUEUE_SAFETY_MARGIN_MS | 30초 (다음 job 진입 여유 부족 시 중단, 다음 트리거에서 계속) |
| Job 타입 | pdf_only (REQ 최종승인 → STAGING PDF), pdf_and_consolidate (PRC 최종승인 → PDF + FINAL 통합 이동) |
| 운영 함수 | getQueueStatus(조회), retryFailedJobs(재시도), clearQueue(초기화, ADMIN), uninstallQueueTrigger(제거) |

**주의:** `installQueueTrigger()`를 실행하지 않으면 결재는 정상 처리되지만 PDF·FINAL 이동이 전혀 일어나지 않는다. 셋팅 시 가장 흔한 누락 지점.

---

# 5. 시스템로그(AuditLog) 인프라 — 현재 상태와 확장 방법

## 5.1 현재 상태 (v2.1)

- '시스템로그' 시트(12컬럼)에 **append-only** 기록. `ensureAuditLogSheet()`가 시트를 자동 보장.
- **fail-safe**: `writeAuditLog`는 try/catch로 감싸 로그 실패가 본 업무를 중단시키지 않는다(throw 안 함, false 반환).
- 현재 기록되는 이벤트: **관리자·보안 이벤트만** — ADMIN_FORCE_DISCARD, ADMIN_FORCE_RELEASE_LOCK, ADMIN_ACCESS_DENIED.
- 문서 라이프사이클 이벤트(DOC_SUBMIT, DOC_APPROVE, DOC_REJECT, PRC_CLAIM 등)는 **상수 자리만 있고 미연결**.

## 5.2 12개 컬럼 (AUDIT_COL)

logId(UUID), timestamp(Date), actor(이메일), actorRole(admin/procurement/approver/requester/unknown), eventType, docNo, docToken, docType(REQ/PRC), targetUser, reason, payload(JSON string), ipAddress(빈값).

## 5.3 신규 이벤트를 추가하는 표준 절차

- AUDIT_EVENT 상수에 이벤트 타입 추가.
- 해당 비즈니스 로직 함수의 적절한 지점에서 `writeAuditLog({ eventType, docNo, docToken, docType, payload })` 호출.
- payload는 객체로 넘기면 자동 `JSON.stringify`된다(직렬화 불가 시 '[unserializable]').
- 관리자 액션이면 진입 직후 `assertAdminWithLog(eventType, payload)`로 권한 검증(거부 시 자동 로그+throw), `requireAdminReason(reason)`으로 사유 강제.

> **확장 권장:** 이미 인프라가 완성되어 있으므로, 문서 제출/승인/반려/PRC 픽업 함수에 `writeAuditLog` 호출 한 줄씩만 추가하면 FR-41(문서 라이프사이클 로그)을 손쉽게 완성할 수 있다. 이것이 다음 개발의 1순위 후보 중 하나다.

---

# 6. 관리자 기능 (A 그룹)

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
- doGet의 admin 라우팅은 isAdminUser 미통과 시 거부 페이지 + ADMIN_ACCESS_DENIED 로그.

---

# 7. 데이터 모델 핵심 (품의서목록 30컬럼)

| 범위 | 컬럼 |
| --- | --- |
| A~Q (0~16) | 제출일시, 품의번호, 발행일자, 기안자, 부서, 품의제목, 용도, 업체명, 사업자번호, 업체담당자, 연락처, 납기일, 납품장소, 품목내역, 합계금액, 부가세포함합계, 비고 |
| R~Z (17~25) | 결재상태, 토큰, Drive폴더ID, 결재자수, 현재결재순번, 재상신횟수, 반려이력누적, 첨부파일ID목록, 이동상태 |
| AA~AD (26~29) | DocType(REQ/PRC), ParentDocId(PRC→REQ 토큰), ClaimedBy(락 점유자), ClaimedAt |
| AE~ (30~) | 결재자 블록 × N — 블록당 6컬럼(label, name, email, status, processedAt, comment), 상수 APPR_START=30, APPR_COLS=6 |

**핵심 규약:**

- 행 조회는 토큰 기반 TextFinder(`findRowNumByToken`) 사용. 토큰은 1회성 보안 + 행 식별 겸용.
- REQ↔PRC는 1:1 매핑. PRC는 자체 토큰 + ParentDocId로 원본 REQ 추적.
- 락 점유는 status를 바꾸지 않고 ClaimedBy/ClaimedAt만 채운다(부모 REQ status는 '최종승인' 유지). 폐기·환원 시 status를 정확히 되돌려야 인박스 필터에 다시 걸린다.
- 컬럼 인덱스는 0-based COL 상수로 접근하고, getRange에는 +1 한다. 절대 숫자 리터럴로 컬럼 접근하지 말 것.
- 일괄 쓰기는 `batchUpdate(sheet, rowNum, updates)`로 묶어 setValues 호출 수를 줄인다.

---

# 8. 코딩 규약 및 권장 작업 방식

## 8.1 함수 네이밍 패턴

- 클라이언트 호출용: `xxxForClient(payload)` (plain object 반환) 또는 doPost 경로의 `_xxxCore(payload)`.
- 내부 헬퍼: 언더스코어 접두사(`_renderTemplate`, `_deriveActorRole` 등).
- 날짜 변환: `toDateStr` / `toDateTimeStr`. 문서번호 정규화: `normalizeDocNo`.

## 8.2 변경 시 권장 절차 (이 시스템 개발자의 작업 스타일)

- 단일 파일(Code.gs)이므로 섹션 주석(`// N. …`)으로 구획을 찾는다. 22개 섹션 헤더가 목차 역할을 한다.
- **UX가 걸린 변경은 단일 HTML 시뮬레이터(vanilla HTML/CSS/JS, 외부 의존성 없음)로 먼저 검증**한 뒤 구현한다(예: dashboard_simulator_v4.html). 개발자가 선호하는 방식이다.
- 단계별로 확인받으며 진행한다. 큰 일괄 재작성보다 작은 검증 가능한 변경을 선호한다.
- 수치(쿼터, 한도)·API 식별자·동작 설명은 **정확히** 적는다. 개발자가 부정확한 수치(예: 트리거 90분을 6시간으로 오기)를 적극적으로 교정한다.
- 모든 시스템 UI·문서·커뮤니케이션은 한국어.

## 8.3 테스트 함수 활용

Apps Script 편집기에서 직접 실행: testCheckDocNo(품의번호 중복), testApproverList(결재자 로드), testHomeData(홈 데이터), testAuditLogInfra(시스템로그), testA4ListLocks / testA4ForceReleaseDryRun(락 해제).

---

# 9. 알려진 위험·미해결 과제 (개발 시 인지할 것)

| 항목 | 내용 / 위험 |
| --- | --- |
| 단일 장애점 (최우선) | 관리자·소유 계정 1인 의존. 일일 자동 백업·백업 운영자 **미구현**(Q-20). 데이터 손상 시 복구 수단 부족 |
| 문서 라이프사이클 로그 미연결 | 인프라는 있으나 제출/승인 등 이벤트 미기록. 감사 추적 불완전 |
| PRC 락 자동 해제 트리거 미등록 | TTL(24h) 설정은 있으나 releaseExpiredPrcLocks 시간 트리거 등록 필요(Q-19) |
| 상태 강제 변경 화면 부재 | 오류 복구를 시트 직접 편집에 의존(A-5 미구현) → 수동 편집 시 상태 깨짐 위험 |
| REQ vs PRC 비교/필드변경 추적 미구현 | FR-46, FR-57, Q-18 |
| 서명 이미지·프리셋 미구현 | Q-04, Q-05 미결 |
| 50명 확장 시 락 경합 | 동시성 모니터링 필요. 임계 도달 시 GAS→별도 서버 재평가(Q-08) |

---

# 10. 향후 확장 (Phase 2 / HR 문서)

운동지원금·영어지원금·지출결의서·구매품의서 등 추가 양식으로 확장 예정. 확장 시:

- 현재 단일 docType(REQ/PRC) 구조를 일반화할지, 양식별 시트/플로우를 분리할지 설계 결정 필요.
- 결재 워크플로우·시스템로그·작업 큐 인프라는 재사용 가능하도록 설계되어 있으므로, 양식 메타와 PDF 템플릿 중심으로 확장하는 방향이 자연스럽다.

---

# 부록. 빠른 체크리스트 (코드 수정 전 확인)

| ☐ | 확인 항목 |
| --- | --- |
| ☐ | Date 반환값을 toDateStr/toDateTimeStr로 변환했는가 |
| ☐ | 서버 함수를 try/catch로 감싸고 stack을 메시지에 담았는가 |
| ☐ | 클라이언트 통신을 google.script.run으로 했는가 (fetch 금지) |
| ☐ | 상태값을 한글로 비교했는가 (영문 아님) |
| ☐ | 첨부·이메일을 락 밖에서 처리했는가 |
| ☐ | PDF payload에 금지 필드(remarks 등)가 없는가 |
| ☐ | Drive 호출에 supportsAllDrives: true가 있는가 |
| ☐ | 컬럼 접근을 COL 상수로 했는가 (숫자 리터럴 금지) |
| ☐ | 관리자 액션에 assertAdminWithLog + requireAdminReason을 넣었는가 |
| ☐ | 수정 후 [새 버전]으로 재배포했는가 |
