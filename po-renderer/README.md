# po-renderer — 주문서(PURCHASE ORDER) 생성기 v2

구매 결재가 완료된 구매품의서(PRC)로 **주문서 Excel + PDF**를 만든다.
운영 규칙(양식·업체 표기·검증 절차)은 [`주문서_생성_규칙.md`](주문서_생성_규칙.md)에 있다.

```
[GAS] PRC 최종승인 → PO.gs _preparePoHandoff()
        → FINAL/{PO} 폴더에 po_manifest.json + David에게 작업요청 메일
[로컬] python render_po.py --folder <FINAL/{PO} 폴더>
        → 2026 기준 양식(거래조건 2p 포함)에 값 기입
        → Excel COM 재계산 + PDF → 자체 검증
        → 주문서 폴더에 {PO번호}_{업체명}.xlsx + .pdf 업로드
[GAS] markPoDone("prcToken", "파일ID")   (감사로그 마감)
```

## 설치 (최초 1회)

```powershell
cd "...\po-renderer"
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

- **Pillow가 없으면 로고·서명 이미지가 사라진 주문서가 만들어진다**(렌더러가 막는다).
- **Excel이 설치된 PC에서만** 동작한다(수식 재계산·PDF 내보내기에 COM을 쓴다).
- `credentials.json`은 insp-renderer 것을 복사해 둔 공용 OAuth 데스크톱 클라이언트.
  `credentials.json` / `token.json` 은 **커밋 금지**. 최초 실행 시 브라우저 동의 1회.

## 실행

```powershell
# ① 결재 핸드오프 (권장)
.\.venv\Scripts\python.exe render_po.py --folder <FINAL/{PO} 폴더 ID 또는 URL>

# ② DB 백업에서 직접 (과거 건 재생성 / 핸드오프 실패 시)
.\.venv\Scripts\python.exe render_po.py --db "..\구매결재시스템_DB_20260915.xlsx" `
    --po TO-PO-26-278 --po TO-PO-26-279

# 로컬 확인만 (업로드 안 함)
... --no-upload            # out\ 폴더에 생성
... --copy-to "G:\Shared drives\SCM_Innovation\02. Purchase\주문서"
```

주요 옵션

| 옵션 | 용도 |
|---|---|
| `--kind krw\|fx` | 기준 양식 강제(기본은 품목 통화로 판단) |
| `--base <xlsx>` | 다른 양식을 기준으로 사용(옛 업체 양식 — 권장하지 않음) |
| `--scale 95` | 인쇄 배율 고정(기본은 2페이지가 되도록 자동 선택) |
| `--payment`, `--destination` | Payment Terms(C10) / Destination(G10) 덮어쓰기 |
| `--allow-unapproved` | `--db` 사용 시 결재 미완료 건도 허용 |
| `--force` | 검증 실패에도 계속 진행(원칙적으로 쓰지 않는다) |

## 파일 구성

| 파일 | 역할 |
|---|---|
| `render_po.py` | CLI · Drive 입출력 · Excel COM 호출 · 인쇄 설정 선택 |
| `po_form.py` | 양식 지식(셀 맵·통화 서식·레터헤드), 채우기, 검증 |
| `po_source.py` | 매니페스트 / DB 백업 → 공통 주문 데이터 변환 |
| `vendor_rules.json` | 업체별 Shipper 표기·REMARK 라벨·결제조건 — **신규 업체는 여기만 수정** |
| `excel_finalize.ps1` | Excel COM 재계산·저장·PDF 내보내기 (ASCII 전용) |
| `templates/po_base_krw.xlsx` | 내자 기준 양식 (TO-PO-26-261_태성테크, 국문 거래조건 2p) |
| `templates/po_base_fx.xlsx` | 외자 기준 양식 (TO-PO-26-264_Coherent, 영문 거래조건 2p) |

## 자체 검증과 사람의 몫

렌더러는 합계 대조, 기준 양식 원 업체 흔적, 로고·서명 유무, `###`(열 너비 부족),
PDF 페이지 구성을 확인한다. **글자 잘림·줄바꿈은 자동으로 잡히지 않으므로**
생성된 PDF를 반드시 눈으로 확인한다.

## 매니페스트 스키마 (po-manifest-v2)

| 필드 | 주문서 위치 | 비고 |
|---|---|---|
| `poNo` | P/O-No (C9) | = PRC 품의번호 |
| `issueDate` | Date (G9) | |
| `paymentTerms` | Payment Terms (C10) | 품의서 구매조건 원문 → 렌더러가 표기 변환 |
| `deliveryDate` | DEL'Y DATE (H열) | |
| `vendorName` | Shipper (G12) | 표기는 `vendor_rules.json` |
| `items[]` | 품목표 (14행~) | `{name, spec, qty, price, currency}` |
| `totalAmt` | 합계 대조값 | 불일치 시 생성 중단 |
| `orderFolderId` | 업로드 대상 | 주문서 폴더 Drive ID |

Destination(G10)은 관례값(내자 `KOREA` / 외자 `INLC Technology, Daejeon, Korea`)을 쓴다.
품의서 납품장소(`deliveryAddr`)는 참고용으로만 전달되며 주문서에 그대로 찍지 않는다.
Reference NO.(C11)는 관례상 공란이다.

관련: `PO.gs`(핸드오프) · `insp-renderer/`(같은 패턴의 검수보고서 렌더러)
