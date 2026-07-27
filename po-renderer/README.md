# po-renderer — 주문서(PURCHASE ORDER) 생성기

검수보고서(INSP) 파이썬 렌더러와 **같은 방식**으로, 구매 결재가 완료된 구매품의서(PRC)
정보를 바탕으로 **주문서(PO) Excel(.xlsx)** 을 자동 생성한다.

```
[GAS] PRC 최종승인(최종승인(PRC))
   → 큐 job 말미 _preparePoHandoff()  (PO.gs)
   → FINAL/{PO} 폴더에 po_manifest.json 기록 + David에게 작업요청 메일
[로컬] python render_po.py --folder <FINAL/{PO} 폴더 링크>
   → QF-741-2 템플릿(po_template.xlsx) 셀만 채워 xlsx 생성
   → 주문서 폴더(SCM_Innovation/02. Purchase/주문서)에 {품의제목}_{업체명}.xlsx 업로드
[GAS] markPoDone("prcToken", "파일ID")   (감사로그 마감)
```

## 왜 openpyxl로 저장하지 않는가 (중요)

주문서 템플릿(QF-741-2)의 **InLC 로고·David 서명 이미지**와 **표 괘선**(drawing1.xml의
Line 도형)은 openpyxl이 `load → save` 시 **전부 소실**시킨다(검증 완료).
그래서 이 렌더러는 xlsx를 zip으로 열어 `xl/worksheets/sheet2.xml`의 **대상 셀 값만**
치환하고, 나머지(media·drawings·styles·rels)는 **바이트 그대로 보존**한다.
런타임에 openpyxl이 필요 없다(표준 라이브러리 zipfile/xml만 사용).

## 설정 (최초 1회)

```powershell
cd "...\po-renderer"
# 가상환경은 이미 있음(.venv). 없으면:
#   & "C:\Users\davidkang\anaconda3\python.exe" -m venv .venv
#   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
# credentials.json 은 insp-renderer 것을 복사해 둠(공유 GCP OAuth 데스크톱 클라이언트).
```

`credentials.json` / `token.json` 은 **gitignore(커밋 금지)**. 최초 실행 시 브라우저 동의 1회.

## 실행

```powershell
# STAGING(=FINAL/{PO}) 폴더에서 po_manifest.json을 찾아 xlsx 생성 + 주문서 폴더 업로드
.\.venv\Scripts\python.exe render_po.py --folder <FINAL/{PO} 폴더 ID 또는 URL>

# 업로드 없이 로컬 미리보기만
.\.venv\Scripts\python.exe render_po.py --folder <ID> --no-upload

# 로컬 매니페스트로 생성(개발/디버깅)
.\.venv\Scripts\python.exe render_po.py --manifest .\manifest.json --no-upload
```

`python`이 PATH에 없으므로 venv 파이썬을 직접 지정한다. 완료 후 GAS 편집기에서
`markPoDone("prcToken", "생성된파일ID")` 실행해 마감한다.

## 매니페스트 스키마 (po-manifest-v1)

| 필드 | 주문서 위치 | 비고 |
|---|---|---|
| `poNo` | P/O-No (C9) | = PRC 품의번호 |
| `paymentTerms` | Payment Terms (C10) | 지급정보 → 구매방법 폴백 |
| `destination` | Destination (G10) | 빈값이면 "INLC Technology, Daejeon, Korea" |
| `reference` | Reference NO.(품의번호) (C11) | 원 REQ 문서번호 |
| `vendorName` | Shipper (G12) | |
| `items[]` | 품목표 (14,16,… 행) | `{name, spec, qty, price, currency}` |
| `orderFolderId` | 업로드 대상 | 주문서 폴더 Drive ID |

- 품목 설명 = `name / spec`, 단위는 기본 `ea`. 금액(G열)·합계(G37)는 수식 유지 + 캐시값 갱신.
- 품목 11개까지 자동 채움(14~34행). 초과분은 경고 후 Excel에서 수동 추가.

## 알려진 한계 (v1)

- **통화 서식**: 템플릿이 USD($) 기준. KRW 등 다른 통화는 값은 채워지나 셀 표시서식이
  $로 나올 수 있어 사용자가 Excel에서 서식을 조정한다. (통화별 템플릿 분리는 향후 과제)
- 사용자가 생성된 xlsx에서 영문 설명/납기/비고 등 필요한 부분을 최종 편집하는 것을 전제로 한다.

관련: `insp-renderer/`(같은 패턴, INSP PDF), GAS `PO.gs`.
