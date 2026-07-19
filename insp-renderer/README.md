# 검수보고서(INSP) PDF 렌더러

GAS가 결재 완료 시 STAGING 폴더에 기록한 `manifest.json` + 사진을 읽어
Playwright(headless Chromium)로 PDF를 만들고 `FINAL/{PO번호}` 폴더에 업로드한다.

구글 문서 변환 서버(불안정)를 쓰지 않으므로 `no servers are currently available` 류의
일시 장애가 없고, 실패해도 사진이 보존되어 있어 언제든 다시 생성할 수 있다.

## 전체 흐름

```
기안자/팀원이 검수보고서 제출 → 결재 진행 → 결재 완료(최종승인)
   └→ [GAS] manifest.json 기록 + 사진 보존 + David에게 작업요청 메일
        └→ [David → 클로드 에이전트] "이 건 PDF 만들어줘"
             └→ [render_insp.py] 사진 다운로드 → Playwright 렌더 → PO 폴더 업로드
                  └→ [GAS] markInspPdfDone(token, fileId) 로 마감
```

## 최초 1회 설정

### 1) 파이썬 패키지 — ✅ 완료됨

이 폴더의 `.venv/`에 이미 설치되어 있다 (playwright + chromium 포함).
`python`이 PATH에 없으므로(아나콘다만 설치됨) **venv의 파이썬을 직접 지정해서** 실행한다:

```powershell
cd "C:\Users\davidkang\OneDrive\Desktop\David Kang\procurement-system\Source-Code\insp-renderer"
.\.venv\Scripts\python.exe render_insp.py --help
```

> 환경을 다시 만들어야 할 경우:
> ```powershell
> & "C:\Users\davidkang\anaconda3\python.exe" -m venv .venv
> .\.venv\Scripts\python.exe -m pip install -r requirements.txt
> .\.venv\Scripts\python.exe -m playwright install chromium
> ```

### 2) 구글 OAuth 자격증명

1. [Google Cloud Console](https://console.cloud.google.com/) → 프로젝트 선택/생성
2. **API 및 서비스 → 라이브러리** → `Google Drive API` 사용 설정
3. **OAuth 동의 화면** → User Type을 **내부(Internal)** 로 설정
   - ⚠ "외부/테스트"로 두면 갱신 토큰이 **7일마다 만료**되어 재로그인을 반복하게 된다.
     inlct.com Workspace 계정이므로 반드시 **내부**로 설정할 것.
4. **사용자 인증 정보 → OAuth 클라이언트 ID → 애플리케이션 유형: 데스크톱 앱**
5. JSON 다운로드 → 이 폴더에 **`credentials.json`** 으로 저장

최초 실행 시 브라우저가 열려 동의를 요청하고, 이후 `token.json`이 만들어져 자동 갱신된다.

> `credentials.json`과 `token.json`은 `.gitignore`에 있으므로 커밋되지 않는다.
> `token.json`은 David 계정의 Drive 접근 권한이므로 외부에 공유하지 말 것.

## 사용법

```powershell
cd "C:\Users\davidkang\OneDrive\Desktop\David Kang\procurement-system\Source-Code\insp-renderer"

# 1) 미리보기 (업로드 없이 로컬 PDF만 생성 — 레이아웃 확인용)
.\.venv\Scripts\python.exe render_insp.py --folder <STAGING 폴더 링크> --no-upload

# 2) 실제 실행 (렌더 + PO 폴더 업로드)
.\.venv\Scripts\python.exe render_insp.py --folder <STAGING 폴더 링크>

# 3) 로컬 manifest.json으로 실행
.\.venv\Scripts\python.exe render_insp.py --manifest .\manifest.json
```

`--folder`에는 David가 받은 작업요청 메일의 **입력 폴더 링크**를 그대로 붙여넣으면 된다.
(폴더 ID만 넣어도 되고, `https://drive.google.com/drive/folders/...` 전체 URL도 인식한다)

실행이 끝나면 마지막에 GAS에서 실행할 마감 명령이 출력된다:

```
markInspPdfDone("<token>", "<업로드된 PDF 파일 id>")
```

이걸 Apps Script 편집기에서 실행하면 대기목록(`listInspAwaitingPdf()`)에서 빠진다.

## 설계 원칙

- **결정적(deterministic)**: 이 스크립트는 내용을 생성하거나 추론하지 않는다.
  `manifest.json`에 있는 값만 그대로 렌더한다. 판정·금액·결재자 이름을
  LLM이 만들어내면 문서 위조가 되므로, 에이전트는 이 스크립트를 **실행**만 한다.
- **매니페스트 계약**: 파이썬은 구글 시트의 45컬럼 스키마를 알지 못한다.
  `manifest.json`의 이름표(`docNo`, `verdict`, `photos[].id` …)만 계약으로 삼으므로,
  GAS 시트 구조가 바뀌어도 이 스크립트는 영향받지 않는다.
- **업로드 후 정리**: 같은 회차의 이전 PDF는 **새 업로드가 성공한 뒤에만** 휴지통으로 보낸다.
  (먼저 지우면 업로드 실패 시 멀쩡한 이전 PDF만 사라진다)

## manifest.json 스키마 (`insp-manifest-v1`)

GAS의 `_buildInspManifest()` (INSP.gs)가 생성한다.

| 키 | 설명 |
|---|---|
| `docNo` / `seq` / `isFinal` | 검수보고서 번호, 회차, 최종검수 여부 |
| `reqNo` / `poNo` | 품의번호 / PO번호 |
| `subject` / `vendorName` / `drafterName` / `dept` | 품명 / 업체 / 담당자 / 부서 |
| `issueDate` / `receivedDate` / `receivedNote` | 기안일 / 입고일 / 입고내역 |
| `verdict` / `comment` | 판정(합격·불합격) / 검수의견 |
| `approvers[]` | `label`, `name`, `status`, `processedAt` — 결재란에 찍힘 |
| `photos[]` | `id`(Drive 파일 id), `name` — id로 직접 다운로드 |
| `finalFolderId` | 업로드 대상 = REQ/PRC 파일이 있는 `FINAL/{연}/{월}/{PO번호}` 폴더 |
| `token` | 마감(`markInspPdfDone`)에 사용 |

## 문제 해결

| 증상 | 원인/조치 |
|---|---|
| `credentials.json 없음` | 위 "구글 OAuth 자격증명" 절차 수행 |
| 브라우저 동의가 자꾸 반복됨 | OAuth 동의 화면이 "테스트" 모드 → **내부(Internal)** 로 변경 |
| `manifest.json 이 없습니다` | 아직 결재 미완료. GAS에서 `rerunInspHandoff("token")` 실행 |
| 한글이 네모(□)로 나옴 | 맑은 고딕 부재. Windows에는 기본 설치돼 있음 |
| 사진이 일부만 나옴 | 경고 로그 확인 — 해당 사진이 휴지통에 있거나 권한 문제 |
