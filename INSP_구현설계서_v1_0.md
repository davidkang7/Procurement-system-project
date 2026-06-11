# 검수보고서(INSP) 기능 구현 설계서 v1.0

**InLC Technology 구매결재시스템 — Phase 2 첫 모듈**
작성일: 2026-06-10 / 기준 코드: Code.gs v2.1 (APPR_START=32, PURCHASE_METHOD·PAYMENT_INFO 반영본)

> ⚠️ 본 설계서의 컬럼 인덱스는 **운영본 v2.1 스키마 기준**입니다. 프로젝트 파일의 Code.gs 사본(v2.0)과 다르므로, 구현 시 반드시 Apps Script 편집기의 실코드를 기준으로 합니다.

---

## 1. 확정 사양 요약 (Q-INSP-01~12 + 추가 결정)

| # | 항목 | 확정안 |
|---|---|---|
| 01 | 리스트 상태 칼럼 | 검수 진행상태 (미제출 / 결재중 N건 / N회 완료 / 반려 있음 / 종결) |
| 02 | 문서번호 | 품의번호 + 회차 2자리 자동 채번 (`TH-AP-26-029-01`) |
| 03 | 결재선 | 매번 신규 선택, 기안자(고정) + 결재자 최대 3인, 기존 결재란 UI 재사용 |
| 04 | 복수 제출 | 허용 (분할 납품 회차별), 아코디언 UX |
| 05 | 열람 범위 | 일반: 본인 기안 건 / 구매팀: 전체 |
| 06 | 판정 | 합격 / 불합격 2단계 |
| 07 | 사진 | 최대 6장, 장당 5MB(원본), **클라이언트 측 자동 압축 후 업로드(§5.3)** |
| 08 | 반려 | 기존과 동일 (수정 후 재상신) |
| 09 | 종결 | "최종 검수" 체크 회차 승인 완료 시 품의 종결 + 추가 제출 차단 |
| 10 | 입고 내역 | 텍스트 1줄 |
| 11 | 동시 진행 | 앞 회차 결재중에도 다음 회차 제출 허용 |
| 12 | 회차 자릿수 | 2자리 (-01 ~ -99) |
| 추가 | 화면 구성 | INSP 작성·뷰어는 전용 HTML, **승인/반려 처리 페이지는 기존 재사용** |
| 추가 | 데이터 저장 | 기존 스프레드시트(`CONFIG.SHEET_ID`)에 `검수보고서목록` 탭 추가 |
| 추가 | 사진 보관 | PDF 삽입 후 원본·임시 폴더 삭제 (미보관) |

---

## 2. 데이터 스키마 — `검수보고서목록` 탭 (INSP_COL)

기존 `품의서목록`과 행 구조가 다르므로 **별도 상수 `INSP_COL`을 정의**하고, 품의서목록의 `COL`과 절대 혼용하지 않습니다. 결재자 블록은 기존과 동일한 6컬럼 패턴(label, name, email, status, processedAt, comment)을 재사용합니다.

```javascript
// ================================================================
// 검수보고서목록 열 인덱스 상수 (0-based) — v1.0
// ================================================================
var INSP_COL = {
  SUBMIT_AT:     0,   // 제출 시각 (Date)
  DOC_NO:        1,   // 검수보고서 번호: {REQ docNo}-{NN}  ex) TH-AP-26-029-01
  SEQ:           2,   // 회차 (number)
  PRC_TOKEN:     3,   // 대상 PRC의 token (품의서목록 조인 키)
  REQ_TOKEN:     4,   // 원본 REQ의 token
  PO_NO:         5,   // PO 번호 (PRC docNo 복사 — 조회 성능용 비정규화)
  REQ_NO:        6,   // 품의 번호 (REQ docNo 복사)
  SUBJECT:       7,   // 품명 (REQ subject 복사)
  VENDOR_NAME:   8,   // 업체명
  DRAFTER:       9,   // 작성자 이메일
  DRAFTER_NAME:  10,  // 작성자 성명
  DEPT:          11,  // 부서
  ISSUE_DATE:    12,  // 품의 기안일자 (REQ issueDate)
  RECEIVED_DATE: 13,  // 입고일자
  RECEIVED_NOTE: 14,  // 입고 내역 (텍스트 1줄)
  VERDICT:       15,  // '합격' | '불합격'  (한글 저장 — 영문 금지)
  IS_FINAL:      16,  // 'Y' | ''  최종 검수 여부
  COMMENT:       17,  // 검수 의견
  STATUS:        18,  // 상태 (한글, §4 참조)
  TOKEN:         19,  // INSP 고유 토큰 (Utilities.getUuid)
  DRIVE_ID:      20,  // 사진 임시 폴더 ID (STAGING) → PDF 생성 후 비움
  APPR_COUNT:    21,
  APPR_IDX:      22,
  RESUB_COUNT:   23,
  REJECT_LOG:    24,
  PHOTO_LIST:    25,  // JSON: [{id, name, size}] — 임시 사진 파일 ID 목록
  MOVE_STATUS:   26,  // 'STAGING' | 'FINAL'
  APPR_START:    27,  // 결재자 블록 시작 (기안자 미포함, 결재자 1~3)
  APPR_COLS:     6,   // 기존과 동일: label, name, email, status, processedAt, comment
};
INSP_COL._VERSION = 'insp-v1.0';
```

설계 원칙: PO_NO·REQ_NO·SUBJECT·VENDOR_NAME은 PRC/REQ에서 복사한 **비정규화 필드**입니다. 검수보고서 목록 조회 시 품의서목록을 재스캔하지 않기 위한 것으로, 원본 변경 시 동기화하지 않습니다(검수 시점 스냅샷 — PDF 화이트리스트 원칙과 동일한 사상).

시트 보장: `ensureInspSheet()` 함수가 탭 부재 시 헤더와 함께 자동 생성 (`ensureAuditLogSheet()` 패턴 복제).

---

## 3. 상태값 정의 (한글 — 영문 금지)

| 상태 | 의미 | 전이 |
|---|---|---|
| `검토중` | 제출 직후, 1차 결재자 대기 | 제출 시 |
| `결재중 (이름)` | N차 결재 진행 | 기존 패턴 동일 |
| `반려` | 결재자 반려 | → 수정 재상신 시 `검토중`, RESUB_COUNT+1 |
| `최종승인(INSP)` | 전체 결재 완료 | 큐에 insp_pdf job 등록 |

**품의(PO) 단위 집계 상태는 저장하지 않고 계산**합니다(시뮬레이터의 `poInspState()` 로직과 동일). `IS_FINAL='Y'` 행이 `최종승인(INSP)`이면 해당 PRC는 "검수 완료(종결)" — 신규 제출 차단은 `submitInspForClient` 서버 검증으로 강제합니다(클라이언트 버튼 숨김만으로는 불충분).

---

## 4. 화면·라우팅 구성

### 4.1 신규 HTML 파일 (3개)

| 파일 | 역할 | 기반 |
|---|---|---|
| `Procurement_INSP_Form.html` | 검수보고서 작성/재상신 | 시뮬레이터 v3 폼 + `Procurement_Form.html`의 결재란·진행바·드롭다운 코드 이식 |
| `Procurement_INSP_Viewer.html` | 결재자·열람용 INSP 전용 뷰어 (사진·판정 표시) | `Procurement_Viewer.html` 레이아웃 차용 |
| `INSP_PDF_Template.html` | PDF 생성 템플릿 (청록 테마, 샘플 양식 구조) | `PDF_Template.html` 패턴 |

### 4.2 기존 파일 수정

| 파일 | 수정 내용 |
|---|---|
| `Code.gs` | doGet 라우팅 2건, INSP 섹션 신설(§5), getHomeDataForClient 확장(§5.1), 큐 job 타입 추가(§6) |
| `Procurement_Home.html` | 기본 메뉴에 "결재 완료" 추가, 결재 완료 뷰(아코디언) 렌더 함수 추가 — 시뮬레이터 v3 코드 이식 |
| `Procurement_Approval.html` | **재사용** — 템플릿 변수 `docKind` 추가('INSP'면 INSP 뷰어 링크·INSP 서버 함수 호출로 분기) |
| `Procurement_Reject.html` | **재사용** — 동일하게 `docKind` 분기 |

### 4.3 doGet 라우팅 추가

```javascript
// 검수보고서 작성 화면
if (action === 'insp_form') {
  return _renderTemplate('Procurement_INSP_Form', {
    prcToken: e.parameter.token || '',
    webappUrl: CONFIG.WEBAPP_URL,
  }, '검수보고서 작성');
}
// 검수보고서 뷰어
if (action === 'insp_view') {
  return _renderTemplate('Procurement_INSP_Viewer', {
    token: e.parameter.token || '',
    webappUrl: CONFIG.WEBAPP_URL,
  }, '검수보고서');
}
// 기존 action === 'approve' 분기에 docKind 파라미터 통과 추가
//   (이메일 링크: ...?action=approve&docKind=INSP&token=...)
```

홈 "검수보고서 제출" 버튼: `window.top.location.href = WEBAPP_URL + '?action=insp_form&token={PRC토큰}'` (PRC 픽업 → PRC 작성 전환과 동일 패턴).

---

## 5. 서버 함수 설계 (Code.gs 신설 섹션)

### 5.1 조회

```
getHomeDataForClient()          [기존 수정]
  - 기존 루프에서 docType==='PRC' && status==='최종승인(PRC)' 행 수집 → completedDocs
  - 검수보고서목록 탭 1회 getValues() → PRC_TOKEN 기준 그룹핑하여 insps[] 부착
  - 열람 범위: isProcurementUser(actor) ? 전체 : 부모 REQ drafter === actor
  - openById는 기존 1회 그대로 (동일 스프레드시트 — 추가 비용은 탭 read 1회뿐)
  - Date 필드는 전부 toDateStr/toDateTimeStr 변환 후 반환

getInspForViewer(token)        [신규]
  - INSP 행 + 결재자 블록 + 사진 임시 파일 목록(결재 진행 중일 때만) 반환
  - 사진은 Drive 파일을 base64 dataURL로 반환하지 않고,
    뷰어에서 <img src="https://drive.google.com/thumbnail?id=...&sz=w800"> 사용
    (도메인 내 열람 — 폴더가 공유드라이브 내부이므로 조직 구성원 접근 가능)
```

### 5.2 제출

```
submitInspForClient(data) → _submitInspCore(data)
  검증 (락 안):
   1) assertDrafterIsSelf 패턴 — 부모 REQ의 drafter === actor 인지 검증
   2) 대상 PRC status === '최종승인(PRC)' 인지 검증
   3) 종결 차단: 동일 PRC_TOKEN에 IS_FINAL='Y' && STATUS='최종승인(INSP)' 행 존재 시 거부 (Q-INSP-09)
   4) 회차 채번: 동일 PRC_TOKEN의 max(SEQ)+1 — 락 안에서 계산 (동시 제출 경합 방지)
   5) DOC_NO = reqNo + '-' + pad2(seq), 99회 초과 시 거부
  처리 (락 안): 행 append (상태 '검토중', MOVE_STATUS 'STAGING')
  처리 (락 밖 — 기존 원칙): 사진 업로드 → STAGING/{INSP docNo}/ 폴더 생성,
   PHOTO_LIST·DRIVE_ID 갱신, 1차 결재자 이메일 발송(sendEmailWithRetry)
```

### 5.3 사진 처리 (확정 플로우)

```
[클라이언트] 선택 → canvas 리사이즈 (긴 변 1600px, JPEG 품질 0.8, 장당 ~300KB-1MB)
            → 미리보기 dataURL = 업로드 데이터와 동일 (미리보기 = 최종 PDF 일치 보장)
[제출]      → base64로 서버 전송 → STAGING/INSP_{docNo}/ 저장 (supportsAllDrives:true)
[결재 완료] → 큐 insp_pdf job → 템플릿에 base64 삽입 → PDF 생성
[PDF 성공]  → FINAL/PO폴더로 PDF 이동 → 사진 파일+임시 폴더 삭제
            → PHOTO_LIST='', DRIVE_ID='', MOVE_STATUS='FINAL'
[PDF 실패]  → 사진 보존 (재시도 3회 — QUEUE_MAX_ATTEMPTS — 모두 실패 시 관리자 알림,
              사진은 수동 복구를 위해 유지)
```

압축 근거: 원본 6장 × 5MB를 base64 삽입하면 HTML이 약 40MB가 되어 `getAs(PDF)` 실패 위험. 압축 후 6장 × ~1MB ≈ base64 8MB로 안전권. 5MB는 **업로드 허용 한도**, 실제 저장·삽입은 압축본.

### 5.4 결재

```
approveInspForClient(payload) → _approveInspCore(payload)
  - 기존 _approveCore 로직 복제 후 INSP_COL 기준으로 치환 (시트: 검수보고서목록)
  - 서버 측 myApproverIdx 자동 판정 로직 동일 적용 (idx URL 파라미터 의존 금지 — 기존 버그 학습)
  - 최종 승인 시: status='최종승인(INSP)' + enqueueJob({type:'insp_pdf', token})
    ※ enqueueJob은 완료 메일 수신자 유무 조건 밖에 배치 (기존 학습 사항)
  - 반려 시: 기존 반려 패턴 동일 (REJECT_LOG 누적, 기안자 메일)
rejectInspForClient / resubmitInspForClient — 동일 패턴
```

`Procurement_Approval.html` 재사용 방법: 템플릿 변수 `docKind`를 받아 `google.script.run` 호출 대상 함수명만 분기. 뷰어 링크는 `?action=insp_view&token=...`으로 분기.

### 5.5 큐 job 타입 추가

```
processQueueTrigger() 내 job 분기에 'insp_pdf' 추가:
  1) generateInspPdf(token): INSP_PDF_Template 렌더(사진 base64 삽입) → Blob → PDF
  2) FINAL 위치 결정: PRC_TOKEN으로 품의서목록 조회 → PRC 행 MOVE_STATUS==='FINAL'이면
     DRIVE_ID(=FINAL 폴더)에 저장. 'STAGING'이면(이론상 없음 — PRC 최종승인이 전제)
     PRC STAGING에 저장하고 경고 로그.
  3) 사진·임시 폴더 삭제, 시트 갱신 (§5.3)
  ※ PDF payload는 buildInspPdfPayload() 화이트리스트 방식 — REJECT_LOG, RESUB_COUNT,
    PHOTO_LIST(파일 ID) 등 내부 필드 제외. 자체 점검 throw 로직 동일 적용.
```

### 5.6 감사 로그

```
AUDIT_EVENT에 추가 (step1_audit_log.gs):
  INSP_SUBMIT, INSP_APPROVE, INSP_REJECT, INSP_RESUBMIT, INSP_FINALIZED(종결)
각 코어 함수에서 writeAuditLog({eventType, docNo, docToken, docType:'INSP', payload}) 1줄씩.
※ 검수보고서는 문서 라이프사이클 로그를 처음부터 연결 (FR-41 선행 적용 — 기존 REQ/PRC는 추후)
```

---

## 6. 이메일

- 결재 요청/반려/완료: 기존 `sendEmailWithRetry`(GmailApp) 재사용. 제목 prefix `[검수]` 추가.
- 결재 링크: `?action=approve&docKind=INSP&token=...&decision=...`
- 관리자 오류 알림: 기존 `notifyAdminError` 그대로 (분리 원칙 유지 — sendEmailWithRetry로 라우팅 금지).

---

## 7. 구현 순서 (단계별 — 각 단계 독립 배포·테스트 가능)

| Step | 내용 | 산출물 | 검증 |
|---|---|---|---|
| 1 | INSP_COL 상수 + ensureInspSheet + AUDIT_EVENT 추가 | Code.gs | testInspSheet()로 탭 자동 생성 확인 |
| 2 | getHomeDataForClient 확장 + 홈 "결재 완료" 메뉴/아코디언 | Code.gs, Procurement_Home.html | 최종승인(PRC) 건 리스트·열람 범위 확인 |
| 3 | INSP 작성 폼 + submitInspForClient + 사진 업로드 | Procurement_INSP_Form.html, Code.gs | 제출 → 시트 행·STAGING 사진·채번·종결 차단 확인 |
| 4 | 결재 플로우 (approve/reject/resubmit + Approval/Reject docKind 분기) | Code.gs, 기존 HTML 2개 | 메일 링크 → 승인/반려 → 상태 전이 확인 |
| 5 | INSP 뷰어 | Procurement_INSP_Viewer.html | 결재자가 사진·판정 열람 확인 |
| 6 | INSP_PDF_Template + insp_pdf 큐 job + 사진 삭제 | INSP_PDF_Template.html, Code.gs | 최종승인 → PDF가 FINAL/PO폴더 생성 + 사진 삭제 확인 |
| 7 | 감사 로그 연결 + 마무리 점검 | Code.gs | 시스템로그 기록 확인 |

각 Step 완료 시 **[새 버전] 재배포** 필수. Step 3부터 신규 서버 함수가 추가되므로 재배포 누락 시 `google.script.run` 호출이 실패합니다.

---

## 8. 점검 체크리스트 (AI Context Document 부록 + INSP 추가 항목)

- ☐ Date 반환값 toDateStr/toDateTimeStr 변환
- ☐ 클라이언트 통신 google.script.run (fetch 금지)
- ☐ 상태값·판정값 한글 비교 (영문 금지)
- ☐ 컬럼 접근 INSP_COL 상수 (숫자 리터럴 금지, COL과 혼용 금지)
- ☐ Drive 호출 supportsAllDrives: true
- ☐ 첨부(사진) 업로드는 락 밖에서 처리
- ☐ enqueueJob은 완료 메일 조건 밖에 배치
- ☐ PDF payload 화이트리스트 + 자체 점검 throw
- ☐ 회차 채번은 락 안에서 (동시 제출 경합)
- ☐ 종결 차단은 서버 검증 (클라이언트 버튼 숨김에 의존 금지)
- ☐ PDF 실패 시 사진 보존 (삭제는 PDF 성공 확인 후에만)
- ☐ 수정 후 [새 버전] 재배포

---

*── 검수보고서(INSP) 구현 설계서 v1.0 끝 ──*
