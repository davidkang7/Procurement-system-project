# 검수보고서(INSP) 기능 구현 설계서 v1.1 (구현 완료본)

**InLC Technology 구매결재시스템 — Phase 2 첫 모듈**
작성일: 2026-06-10 (v1.0 설계) / **개정: 2026-06-20 (v1.1 — 구현 완료 반영, 실제 스키마·버그·진단 도구 추가)**
기준 코드: Code.gs v2.1 (APPR_START=32) + **INSP.gs (Step 1~6 누적, ~1,450줄)**

> ✅ **구현 상태: Step 1~6 전부 구현·배포 완료.** 본 설계서는 v1.0의 설계 의도를 유지하되, **실제 구현과 달라진 부분(특히 결재 블록 수·총 컬럼 수)과 구현 중 발견·수정한 버그**를 v1.1에서 반영했다. 코드 수정 시 **반드시 Apps Script 편집기의 실코드(INSP.gs)를 기준**으로 하고, 본 문서는 의도·맥락 파악용으로 본다.

> ⚠️ **v1.0 → v1.1 핵심 차이 (먼저 읽을 것)**
> ① **결재 블록이 "기안자 포함 4블록"으로 변경됨.** v1.0 설계는 결재자 최대 3인(기안자 미포함)으로 `APPR_START=27 + 3×6 = 45`컬럼을 가정했으나, 구현 단계(Step 4)에서 **기안자 자기결재 블록(블록0)을 포함**하기로 바뀌어 `MAX_APPROVERS=4`, **총 51컬럼**(`APPR_START=27 + 4×6 = 51`)이 되었다. REQ/PRC와 동일하게 approvers[0]=기안자.
> ② **구현 중 버그 3건 발견·수정** — (가) DRAFTER 컬럼이 이메일이 아닌 성명(§9.1), (나) 동일인 다단계 결재 시 본인 단계 오판정 → `_findMyInspIdx`(§9.2), (다) "새 배포" 반복으로 검수 버튼 미표시(배포 사고, §9.3). **§9를 반드시 숙지.**
> ③ **진단 도구 `INSP_DIAG.gs`** 추가(검수 결재 대기 미표시 추적).

---

## 1. 확정 사양 요약 (Q-INSP-01~12 + 추가 결정)

| # | 항목 | 확정안 | 구현 |
|---|---|---|---|
| 01 | 리스트 상태 칼럼 | 검수 진행상태 (미제출 / 결재중 / 회차 완료 / 반려 / 종결) | ✅ 계산값(시뮬레이터 `poInspState` 사상) |
| 02 | 문서번호 | 품의번호 + 회차 2자리 자동 채번 (`TH-AP-26-029-01`) | ✅ `reqNo + '-' + _pad2(seq)` |
| 03 | 결재선 | 매번 신규 선택, **기안자(블록0, 고정) + 결재자 최대 3인** | ✅ 4블록 (기안자 포함) |
| 04 | 복수 제출 | 허용 (분할 납품 회차별), 아코디언 UX | ✅ |
| 05 | 열람 범위 | 일반: 본인 기안 건 / 구매팀: 전체 | ✅ `isProcurementUser` 분기 |
| 06 | 판정 | 합격 / 불합격 2단계 (한글) | ✅ |
| 07 | 사진 | 최대 6장, 장당 5MB(업로드 한도), **클라이언트 canvas 압축 후 업로드** | ✅ 긴 변 1600px / JPEG 0.8 |
| 08 | 반려 | 기존과 동일 (수정 후 재상신) | ✅ RESUB_COUNT 누적 |
| 09 | 종결 | 최종 검수 체크 회차 승인 완료 시 종결 + 추가 제출 차단 | ✅ **서버 검증** `_isInspLocked` |
| 10 | 입고 내역 | 텍스트 1줄 (필수) | ✅ |
| 11 | 동시 진행 | 앞 회차 결재중에도 다음 회차 제출 허용 | ✅ |
| 12 | 회차 자릿수 | 2자리 (-01 ~ -99) | ✅ 99 초과 거부 |
| 추가 | 화면 구성 | INSP 작성·뷰어는 전용 HTML, 승인/반려는 기존 재사용(`docKind`) | ✅ |
| 추가 | 데이터 저장 | 기존 스프레드시트에 `검수보고서목록` 탭 추가 | ✅ 자동 생성 |
| 추가 | 사진 보관 | PDF 삽입 후 원본·임시 폴더 삭제 (미보관) | ✅ `_cleanupInspPhotos` |

---

## 2. 데이터 스키마 — `검수보고서목록` 탭 (INSP_COL) — ★실제 구현본★

기존 `품의서목록`과 행 구조가 다르므로 **별도 상수 `INSP_COL`을 정의**하고, 품의서목록의 `COL`과 절대 혼용하지 않는다. 결재자 블록은 기존과 동일한 6컬럼 패턴(label, name, email, status, processedAt, comment)을 재사용한다.

```javascript
// ================================================================
// 검수보고서목록 열 인덱스 상수 (0-based) — insp-v1.0 (실제 구현)
// ================================================================
var INSP_COL = {
  SUBMIT_AT:     0,   DOC_NO:        1,   // {REQ docNo}-NN
  SEQ:           2,   PRC_TOKEN:     3,   REQ_TOKEN:     4,
  PO_NO:         5,   REQ_NO:        6,   SUBJECT:       7,   VENDOR_NAME:   8,
  DRAFTER:       9,   // ★작성자 이메일 (성명 아님)
  DRAFTER_NAME:  10,  DEPT:          11,  ISSUE_DATE:    12,
  RECEIVED_DATE: 13,  RECEIVED_NOTE: 14,
  VERDICT:       15,  // '합격' | '불합격' (한글)
  IS_FINAL:      16,  // 'Y' | ''  최종 검수 여부
  COMMENT:       17,  STATUS:        18,  // 한글 상태
  TOKEN:         19,  DRIVE_ID:      20,  // 사진 임시 폴더(STAGING)
  APPR_COUNT:    21,  APPR_IDX:      22,  RESUB_COUNT:   23,  REJECT_LOG:    24,
  PHOTO_LIST:    25,  // JSON [{id,name,size}]
  MOVE_STATUS:   26,  // 'STAGING' | 'FINAL'
  APPR_START:    27,  // 결재 블록 시작 (블록0=기안자, 블록1~3=결재자)
  APPR_COLS:     6,
  MAX_APPROVERS: 4,   // ★기안자 1 + 결재자 최대 3. approvers[0]=기안자
};
INSP_COL._VERSION = 'insp-v1.0';

// 시트 총 컬럼 수 = 27 + 4×6 = 51  (※v1.0 설계의 45가 아님)
var INSP_TOTAL_COLS = INSP_COL.APPR_START + (INSP_COL.MAX_APPROVERS * INSP_COL.APPR_COLS); // = 51
```

> ⚠️ **결재 블록 수 변경 (45 → 51).** v1.0 설계는 결재자 3인(기안자 미포함)을 가정했으나, 구현은 REQ/PRC와 동일하게 **기안자 자기결재 블록(블록0)을 포함해 4블록·51컬럼**으로 확정됐다. INSP.gs 일부 주석에 "27 + 3×6 = 45"라는 옛 흔적이 남아 있으나 **실제는 51**이며, `testInspStep1`·`_inspSchemaSelfCheck`는 51 기준으로 검증한다.

**비정규화 스냅샷 필드:** PO_NO·REQ_NO·SUBJECT·VENDOR_NAME·DRAFTER_NAME·DEPT·ISSUE_DATE는 PRC/REQ에서 **제출 시점에 복사한 스냅샷**이다. 검수 목록 조회 시 품의서목록을 재스캔하지 않기 위한 것으로, 원본 변경 시 동기화하지 않는다(검수 시점 고정 — PDF 화이트리스트와 동일 사상).

**시트 보장:** `ensureInspSheet()`가 탭 부재 시 헤더와 함께 자동 생성(`ensureAuditLogSheet()` 패턴 복제). 헤더 결재 블록 prefix는 `drafter_appr / appr1 / appr2 / appr3`. 시트 헤더 배경 청록 `#e4f4f2`로 품의서목록과 시각 구분.

**스키마 자가 점검:** `_inspSchemaSelfCheck()`가 INSP_COL 본체 인덱스가 0부터 빈틈없이 연속하는지 + 키 수가 APPR_START+1과 일치하는지 검증(배포 전 인덱스 누락·중복 차단). `inspApprCol(stageIdx, offset)` 헬퍼로만 결재 블록 접근(하드코딩 금지).

---

## 3. 상태값 정의 (한글 — 영문 금지)

| 상태 | 의미 | 전이 |
|---|---|---|
| `검토중` | 제출 직후, 1차 결재자 대기 | 제출 시 |
| `결재중 (이름)` | N차 결재 진행 | 기존 패턴 동일 |
| `반려` | 결재자 반려 | → 재상신 시 `검토중`, RESUB_COUNT+1 |
| `최종승인(INSP)` | 전체 결재 완료 | 큐에 insp_pdf job 등록 |

**품의(PO) 단위 집계 상태는 저장하지 않고 계산**한다. 종결 차단 판정은 `_isInspLocked(inspList)` — **최종 검수(isFinal) 회차가 '반려'가 아닌 상태로 존재하면 차단**(결재 대기/진행중, 또는 이미 최종승인(INSP)). 반려된 최종 회차는 재상신 대상이라 차단에서 제외. 신규 제출 차단은 **`getInspFormDataForClient`(폼 진입)와 `_submitInspCore`(제출, 락 안) 양쪽 서버 검증**으로 강제(클라이언트 버튼 숨김만으로는 불충분).

---

## 4. 화면·라우팅 구성 (구현본)

### 4.1 신규 HTML 파일 (3개) — 구현 완료

| 파일 | 역할 | 비고 |
|---|---|---|
| `Procurement_INSP_Form.html` | 검수보고서 작성/재상신 | 결재란·진행바·드롭다운 + 사진 canvas 압축 업로더 |
| `Procurement_INSP_Viewer.html` | 결재자·열람용 전용 뷰어 (사진·판정·결재) | 청록 #0e7d72 테마, 사진 base64 임베드 |
| `INSP_PDF_Template.html` | PDF 생성 템플릿 | 청록 테마, 사진·결재 블록 삽입 |

### 4.2 기존 파일 수정 — 구현 완료

| 파일 | 수정 내용 |
|---|---|
| `Code.gs` | doGet 라우팅 2건(`insp_form`/`insp_view`), CONFIG `INSP_SHEET_NAME`, AUDIT_EVENT INSP 5종, processQueueTrigger에 `insp_pdf` job 분기(`typeof _processInspPdfJob` 가드) |
| `Procurement_Home.html` | 기본 메뉴 3그룹·5항목 재구성, "결재 완료" 아코디언 렌더(검수 제출/결재 진입), btn-insp CSS |
| `Procurement_Approval.html` | **재사용** — 템플릿 변수 `docKind='INSP'`면 INSP 서버 함수 호출·뷰어 링크로 분기. 결재자 라벨은 서버 `getInspDecisionMetaForClient`에서 받아 표시 |

### 4.3 doGet 라우팅 (구현본)

```javascript
// 검수보고서 작성 (PRC 픽업 또는 반려 재상신)
//   ?action=insp_form&token={PRC토큰}[&resub={INSP토큰}]
// 검수보고서 전용 뷰어 (열람 + 결재)
//   ?action=insp_view&token={INSP토큰}&idx={힌트}
// 결재 메일 승인/반려는 기존 approve 라우팅에 docKind=INSP 통과
//   ?action=approve&docKind=INSP&token=...&decision=...&idx=...
```

홈 "검수보고서 제출" 버튼: `WEBAPP_URL + '?action=insp_form&token={PRC토큰}'` (PRC 픽업 패턴과 동일).

---

## 5. 서버 함수 설계 (INSP.gs — 구현본)

### 5.1 조회

```
_getInspGroupsByPrc(ss)            [홈 아코디언용]
  - 검수보고서목록 1회 getValues() → PRC_TOKEN 기준 그룹핑, 회차 오름차순
  - 동일 스프레드시트라 openById 재비용은 탭 read 1회뿐. fail-safe(실패해도 홈 안 막음)
  - Date 필드 전부 toDateStr/toDateTimeStr 변환

getInspFormDataForClient(prcToken, resubToken)   [작성 폼]
  - 대상 PRC(최종승인(PRC)) + 원본 REQ 스냅샷을 모아 프리필
  - 권한: 원본 REQ 기안자 본인만 (★기안자 이메일 = COL.APPR_START + 2, 소문자 비교)
  - 종결 차단(_isInspLocked) + 회차 채번(max(SEQ)+1) + 재상신 prefill(반려 건만)

getInspViewerDataForClient(token, hintIdx)       [전용 뷰어]
  - 본문 스냅샷 + 결재 블록 + 사진 + myApproverIdx
  - 사진은 STAGING Drive 바이트를 base64 data URL로 직접 임베드
    (Drive 썸네일 URL은 웹앱 iframe에서 미인증으로 깨지므로)

getInspDecisionMetaForClient(token)              [경량 결재 메타]
  - Approval(docKind=INSP) 화면용. curLabel/myLabel/isMyTurn 등
```

### 5.2 제출 (`submitInspForClient` → `_submitInspCore`)

```
검증 (락 밖): 판정/입고내역 필수, approvers[0]=기안자 본인(소문자), 사진 ≤6장·≤5MB
락 안 (동시 제출 경합 방지):
  1) 대상 PRC 재검증(status==='최종승인(PRC)') — 폼 데이터 불신
  2) 원본 REQ 기안자 본인 (★reqInfo.row[COL.APPR_START + 2] 비교)
  3) 종결 차단(_isInspLocked) / 재상신이면 반려 상태만 허용
  4) 회차 채번 max(SEQ)+1 (99 초과 거부), DOC_NO = reqNo-NN
  5) INSP STAGING 폴더 생성('INSP_'+docNo), 행 기록(상태 '검토중', 4블록, MOVE_STATUS 'STAGING')
락 밖: 사진 저장(_saveInspPhotos → PHOTO_LIST 갱신), 1차 결재자 메일, 감사 로그(INSP_SUBMIT/RESUBMIT)
```

### 5.3 사진 처리 (확정 플로우)

```
[클라이언트] 선택 → canvas 리사이즈(긴 변 1600px, JPEG 0.8, 장당 ~300KB-1MB)
            → 미리보기 dataURL = 업로드 데이터와 동일(미리보기=최종 PDF 일치)
[제출]      → base64 서버 전송 → STAGING/INSP_{docNo}/ 저장(supportsAllDrives:true)
[결재 완료] → insp_pdf job → 템플릿에 base64 삽입 → PDF 생성
[PDF 성공]  → FINAL/PO폴더로 저장 → 사진 파일+임시 폴더 삭제, PHOTO_LIST='', DRIVE_ID 갱신, MOVE_STATUS='FINAL'
[PDF 실패]  → 사진 보존(재시도 3회 — QUEUE_MAX_ATTEMPTS — 모두 실패해도 수동 복구용 유지)
```

압축 근거: 원본 6장×5MB를 base64 삽입하면 HTML이 ~40MB가 되어 `getAs(PDF)` 실패 위험. 압축 후 6장×~1MB ≈ base64 8MB로 안전권. 5MB는 **업로드 허용 한도**, 실제 저장·삽입은 압축본.

### 5.4 결재 (`processInspDecisionFromClient`)

```
withLock:
  - 상태 가드(반려/최종승인(INSP)면 거부)
  - ★본인 단계 자동 판정 _findMyInspIdx(row, apprCount, curIdx, actor)
    → myIdx<0: 권한 없음 / myIdx≠curIdx: 순서 아님
  - 현재 블록 status/processedAt/comment 기록
  - reject: status='반려', REJECT_LOG 누적
  - approve & nextIdx<apprCount: status='결재중 (다음이름)', APPR_IDX+1
  - approve & 최종: status='최종승인(INSP)', finalized=true
락 밖:
  - 메일(reject→기안자 / next→다음결재자 / completion→기안자), 감사 로그
  - finalized면 enqueueJob({type:'insp_pdf', token}) — ★완료 메일 조건 밖 배치
```

`Procurement_Approval.html` 재사용: 템플릿 변수 `docKind`로 호출 대상 함수만 분기. 뷰어 링크는 `?action=insp_view`. **idx URL 파라미터를 신뢰하지 않고 서버가 actor로 본인 단계 재판정**(기존 버그 학습).

### 5.5 큐 job 타입 (`insp_pdf`)

```
processQueueTrigger() 내 'insp_pdf' → _processInspPdfJob(job):
  - 상태 가드(최종승인(INSP)), 멱등(MOVE_STATUS==='FINAL'이면 스킵)
  - generateInspPdf(token): INSP_PDF_Template 렌더(사진 base64) → getAs(PDF)
    → FINAL/{PO번호} 폴더 저장(getOrCreateFolder(poNo, issueDate, 'final'))
    → 동일 docNo 이전 PDF는 휴지통 처리(중복 방지)
  - 성공 시 _cleanupInspPhotos(row): 사진 원본 삭제 + INSP_ STAGING 임시폴더 삭제
  - 시트 갱신: DRIVE_ID=FINAL폴더, MOVE_STATUS='FINAL', PHOTO_LIST='[]'
  ※ PDF는 화이트리스트 방식 — REJECT_LOG/RESUB_COUNT/PHOTO_LIST(파일 ID) 등 내부 필드 제외
```

### 5.6 감사 로그 (FR-41 선행 적용)

```
AUDIT_EVENT: INSP_SUBMIT, INSP_APPROVE, INSP_REJECT, INSP_RESUBMIT, INSP_FINALIZED(종결)
각 코어 함수에서 writeAuditLog({eventType, docNo, docToken, docType:'INSP', reason}) 1줄씩.
※ 검수보고서는 문서 라이프사이클 로그를 처음부터 연결(기존 REQ/PRC는 추후 — AI Context §6.1)
```

---

## 6. 이메일 (구현본)

- 결재 요청/반려/완료: 기존 `sendEmailWithRetry`(GmailApp) 재사용. 전용 래퍼 `_sendInspApprovalEmail`/`_sendInspRejectEmail`/`_sendInspCompletionEmail`. 제목 prefix `[검수보고서 …]`.
- 결재 메일 3버튼: 확인=전용 뷰어(`insp_view`), 승인/반려=경량 `Procurement_Approval`(`docKind=INSP`). REQ/PRC와 동일 레이아웃, 청록 테마.
- 결재 링크: `?action=approve&docKind=INSP&token=...&idx=...&decision=...`
- 관리자 오류 알림: 기존 `notifyAdminError` 그대로(이중 경로 분리 원칙 — sendEmailWithRetry로 라우팅 금지).

---

## 7. 구현 순서 (Step 1~6 — 전부 완료)

| Step | 내용 | 산출물 | 상태 |
|---|---|---|---|
| 1 | INSP_COL + ensureInspSheet + AUDIT_EVENT + 자가점검 | INSP.gs | ✅ (testInspStep1 ALL PASS) |
| 2 | getHomeData 확장 + 홈 "결재 완료" 아코디언 | Code.gs, Home.html | ✅ |
| 3 | 작성 폼 + submitInspForClient + 사진 업로드 | INSP_Form.html, INSP.gs | ✅ |
| 4 | 결재 승인/반려 + 종결 차단 + **결재란 기안자 포함(4블록)** | INSP.gs, Approval.html(docKind) | ✅ |
| 5 | 전용 뷰어 + 내 검수 결재 대기 + 최종 완료(INSP) | INSP_Viewer.html | ✅ |
| 6 | 최종승인 시 PDF 생성 → FINAL/PO폴더 → 사진 삭제(insp_pdf 큐) | INSP_PDF_Template.html, INSP.gs | ✅ |

각 Step 완료 시 **[배포 관리→기존 배포 편집→새 버전]** 재배포(절대 "새 배포" 아님 — §9.3). Step 3부터 신규 서버 함수가 추가되므로 재배포 누락 시 `google.script.run` 호출 실패.

---

## 8. 점검 체크리스트

- ☐ Date 반환값 toDateStr/toDateTimeStr 변환
- ☐ 클라이언트 통신 google.script.run (fetch 금지)
- ☐ 상태값·판정값 한글 비교 (영문 금지)
- ☐ 컬럼 접근 INSP_COL 상수 + inspApprCol 헬퍼 (숫자 리터럴·COL 혼용 금지)
- ☐ Drive 호출 supportsAllDrives: true
- ☐ 사진 업로드는 락 밖 / 회차 채번은 락 안
- ☐ enqueueJob은 완료 메일 조건 밖 배치
- ☐ PDF payload 화이트리스트 + 자가점검 throw
- ☐ 종결 차단은 서버 검증 (클라이언트 버튼 숨김 의존 금지)
- ☐ PDF 실패 시 사진 보존 (삭제는 PDF 성공 후에만)
- ☐ ★기안자 신원 비교는 COL.APPR_START+2(이메일)+소문자 (DRAFTER 성명 금지 — §9.1)
- ☐ ★본인 단계 판정은 _findMyInspIdx(curIdx 우선 — 동일인 다단계 — §9.2)
- ☐ 수정 후 [기존 배포 편집→새 버전] 재배포 (새 배포 금지 — §9.3)

---

## 9. ★구현 중 발견·수정한 버그 (반드시 숙지)★

> 📌 *INSP 구현·테스트 단계에서 실제로 시간을 크게 잡아먹은 이슈들이다. AI Context Document §3과 교차 참조. 사용자 요청에 따라 특별 강조.*

### 9.1 DRAFTER 컬럼은 이메일이 아니라 성명

**증상:** 검수 제출 권한 판정(기안자 본인 비교)이 어긋나고, 홈 `completedDocs[0].drafter`가 이메일이 아닌 성명('강기종')으로 옴.

**근본 원인:** 품의서목록 `COL.DRAFTER`(D열)에는 **성명**이 저장된다. 기안자 이메일은 결재자 블록 0번(기안자 자기결재)의 email = **`COL.APPR_START + 2`**에 있다.

**해결:** 모든 기안자 신원 비교를 `reqInfo.row[COL.APPR_START + 2]`(이메일) + **소문자 정규화**로 통일. `getInspFormDataForClient`·`_submitInspCore`에 적용. 같은 함정이 과거 홈 "내 문서" 필터에서도 있었다(결재 블록 없는 행이 사라짐).

### 9.2 동일인 다단계 결재 — 본인 단계 오판정

**증상:** 기안자가 결재선의 결재자로도 들어간 경우(자기결재 블록0 + 결재자 블록 중복 배정), "본인이 배정된 첫 단계"로 단순 검색하면 이미 승인한 앞 단계가 잡혀 **검수 결재 대기에 안 뜨거나 결재가 막힘**.

**근본 원인:** "본인 단계"를 first-match로 찾으면 현재 차례가 아닌 과거 단계가 잡힘.

**해결:** `_findMyInspIdx(row, apprCount, curIdx, actor)` 도입 — **①현재 차례(curIdx)가 본인이면 그 단계를 우선** 반환, ②아니면 본인 첫 단계. 결재 처리·뷰어·대기목록 전부 이 헬퍼를 통한다. 뷰어 myApproverIdx는 3단계(현재차례 본인 → URL힌트 검증 → 본인 첫 단계).

**진단 도구:** `INSP_DIAG.gs`의 `debugInspPending()` — 행별로 결재 블록(이름/이메일/status/현재차례/=나 표시)과 `myIdx/statusOk/myStatus/대기표시` 판정을 콘솔에 덤프해 미표시 원인을 즉시 식별.

### 9.3 "새 배포" 반복 → 활성 배포 2개 → 검수 버튼 미표시 (배포 사고)

**증상:** Code.gs·Home.html을 최신으로 붙여넣고 "새 버전 배포"를 했는데도 화면(iframe)의 `STATE.data.completedDocs[0].drafter`가 옛 성명으로 와서 **검수 제출 버튼이 안 보임.** 편집기에서 `getHomeDataForClient()` 직접 실행은 최신(이메일). "편집기는 최신, 화면은 옛 코드"라는 결정적 증상.

**근본 원인:** 며칠간 배포 시 "기존 배포 편집"이 아니라 **"새 배포"를 반복 생성**해 활성 배포가 2개가 됨. 최신 코드는 새 배포(`…ATwWeh…/exec`)에, 사용자/화면은 기존 WEBAPP_URL(`…iL2l4m…/exec`)을 호출 → 옛 코드 응답.

**진단법:** 편집기 직접 실행 결과 vs 웹앱 iframe 콘솔(`DevTools 컨텍스트 userHtmlFrame`)의 `STATE.data` 비교. 다르면 100% 배포 반영 문제. [배포 관리]에서 활성 배포가 2개 이상이면 갈라진 것. (보조: `debugWhichHome`, `debugCompletedPayload`.)

**복구:** 기존 운영 배포(`iL2l4m…`)를 편집(연필) → 버전 "새 버전" → 배포(기존 URL 유지) ⇒ 잘못 만든 새 배포(`ATwWeh…`)는 Archive.

**재발방지:** 배포는 **항상** "[배포 관리]→기존 배포 편집(연필)→버전 '새 버전'→배포". 절대 "새 배포" 금지. (이 사고가 §9.1 DRAFTER 버그와 겹쳐 진단이 길어졌으므로, 검수 버튼 미표시 시 **두 가지를 함께 의심**.)

---

*── 검수보고서(INSP) 구현 설계서 v1.1 (구현 완료본) 끝 (2026-06-20) ──*