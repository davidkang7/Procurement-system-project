# -*- coding: utf-8 -*-
"""
주문서(QF-741-2) 양식 지식 — 기준 양식 선택 · 셀 채우기 · 자체 검증.

2026-09 운영 규칙을 코드로 옮긴 모듈이다. 근거는 `주문서_생성_규칙.md`.
  - 주문서는 **2페이지**다: 1p QF-741-2 폼 + 2p 표준거래조건(내자=국문 / 외자=영문).
    기준 양식(templates/po_base_krw.xlsx = TO-PO-26-261 태성테크,
              templates/po_base_fx.xlsx  = TO-PO-26-264 Coherent)에 2p가 이미 들어 있다.
    → 거래조건을 따로 삽입(add_terms)하지 않는다. 2025년 이전 양식을 복사해 쓰면 안 된다
      (구주소·전임 담당자 레터헤드, 거래조건 페이지 없음, 제목 잘림).
  - openpyxl로 load→save 해도 이 두 기준 양식은 로고·서명 이미지와 괘선이 보존된다(검증됨).
    괘선을 Line 도형으로 그린 옛 양식(예: 2025년 Coherent)은 보존이 깨질 수 있으므로
    --base 로 옛 파일을 지정하는 것은 권장하지 않는다.
  - openpyxl은 수식 캐시를 남기지 않는다 → 저장 후 반드시 Excel COM 재계산(excel_finalize.ps1).
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import date
from io import BytesIO
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "templates"
RULES_FILE = BASE_DIR / "vendor_rules.json"

TEMPLATES = {
    "krw": TEMPLATE_DIR / "po_base_krw.xlsx",   # 내자 기준: TO-PO-26-261_태성테크 (국문 거래조건 2p)
    "fx":  TEMPLATE_DIR / "po_base_fx.xlsx",    # 외자 기준: TO-PO-26-264_Coherent (영문 거래조건 2p)
}

# 기준 양식에 남아 있는 원 업체 흔적 — 렌더 후에도 남아 있으면 오염으로 보고 실패시킨다.
BASE_RESIDUE = {
    "krw": ["태성테크", "TO-PO-26-261"],
    "fx":  ["COHERENT", "Coherent", "TO-PO-26-264"],
}

SHEET_NAME = "주문서"

# 셀 맵 (QF-741-2 Rev 0)
CELL_PO_NO      = "C9"    # P/O-No
CELL_DATE       = "G9"    # Date(발행일자)
CELL_PAYMENT    = "C10"   # Payment Terms
CELL_DEST       = "G10"   # Destination
CELL_REFERENCE  = "C11"   # Reference NO.(품의번호) — 관례상 공란
CELL_AMOUNT     = "C12"   # Amount = =G{합계행}
CELL_SHIPPER    = "G12"   # Shipper(업체)
ITEM_FIRST_ROW  = 14      # 품목 첫 행(13행은 머리글)
COL_NO, COL_DESC, COL_QTY, COL_UNIT, COL_PRICE, COL_AMT, COL_DELY, COL_REMARK = 1, 2, 4, 5, 6, 7, 8, 9

# 통화별 표시서식 (실제 주문서에서 추출)
NUMBER_FORMATS = {
    "KRW": r'\₩#,##0_);[RED]"(₩"#,##0\)',
    "USD": '_-[$$-409]* #,##0.00_ ;_-[$$-409]* \\-#,##0.00\\ ;_-[$$-409]* \\-??_ ;_-@_ ',
    "JPY": '_-[$¥-411]* #,##0_-;\\-[$¥-411]* #,##0_-;_-[$¥-411]* "-"_-;_-@_-',
    "EUR": '_-[$€-2]* #,##0.00_ ;_-[$€-2]* \\-#,##0.00\\ ;_-[$€-2]* \\-??_ ;_-@_ ',
}
DATE_FORMAT = "yyyy-mm-dd"

# 행 높이 = 글꼴 크기 × 이 비율. Excel 기본 줄간격(약 1.2)에 여백을 더한 값으로,
# 18pt 제목이 19.5pt 행에 갇혀 위가 잘리던 문제를 막는다.
ROW_HEIGHT_RATIO = 1.45

# 인쇄 배율 후보 — 주문서는 2페이지(1p 폼 + 2p 거래조건)로 떨어져야 한다.
# 기준 양식 원본은 fitToPage=True·배율 100이라 거래조건 마지막 줄이 3페이지로 밀린다.
# 실제 발행분(2페이지짜리)은 모두 fitToPage=False + 배율 94~97 이었다 → 같은 방식으로 맞춘다.
# 품목이 많아 1페이지가 넘칠 때를 대비해 배율을 한 단계씩 낮추며 재시도한다.
#   내자(국문 조건문)는 97%면 2페이지에 붙는다. 외자(영문 조건문)는 분량이 많아
#   84%까지 내려가야 2페이지가 된다 — 열 폭이 넓어 89% 이상에서는 가로가 잘린다.
#   3페이지(글자 크게)로 발행하려면 --scale 로 배율을 고정한다.
PAGE_SCALES = {"krw": [97, 95, 93, 91], "fx": [95, 93, 91, 89, 87, 85, 84, 82]}

# 현행 레터헤드 (기준 양식엔 이미 반영돼 있으나, 옛 양식을 --base 로 쓸 때를 대비해 항상 재기입)
LETTERHEAD = {
    "krw": {
        "F4": "Migun Techno World 2cha #A 204. Techno 2-ro 187, ",
        "F5": "Yuseong-gu, Daejeon, Korea 34025",
        "F6": " TEL   :  82-42-721-0721",
        "F7": " FAX  :  82-42-721-0722",
        "F8": "Contact to : Kichong Kang",
        "H6": "C.P : 010-3188-3265",
        "H8": " E-mail : davidkang@inlct.com",
    },
    "fx": {
        "F4": "Migun Techno World 2cha #A 204. Techno 2-ro 187, ",
        "F5": "Yuseong-gu, Daejeon, Korea 34025",
        "F6": " TEL   :  82-42-721-0721",
        "F7": " FAX  :  82-42-721-0722",
        "F8": "Contact to : David Kang",
        "H6": "C.P : 82-10-3188-3265",
        "H8": " E-mail : davidkang@inlct.com",
    },
}

# 품목내역 둘째 필드가 이 목록에 들어맞으면 '단위'로, 아니면 '규격'으로 취급한다.
UNIT_WORDS = {
    "ea", "pc", "pcs", "set", "sets", "box", "roll", "lot", "kg", "g", "m", "mm", "cm",
    "hr", "hour", "md", "spl", "unit", "units", "pack",
    "개", "대", "매", "권", "본", "세트", "조", "식", "회", "건", "박스", "롤", "장", "병",
    "캔", "통", "쌍", "년", "월", "일",
}
DEFAULT_UNIT = "ea"


# ──────────────────────────────────────────────────────────────
# 업체 규칙
# ──────────────────────────────────────────────────────────────
def _norm_vendor(name: str) -> str:
    s = str(name or "")
    s = re.sub(r"㈜|\(주\)|\(株\)|주식회사", "", s)
    s = re.sub(r"\s+|[.,]", "", s)
    return s.strip().lower()


class VendorRules:
    def __init__(self, path: Path = RULES_FILE):
        data = json.loads(path.read_text(encoding="utf-8"))
        self.vendors = {_norm_vendor(k): v for k, v in data.get("vendors", {}).items()}
        self.payment_alias = data.get("payment_alias", {})
        self.defaults = data.get("defaults", {})

    def get(self, vendor_name: str) -> dict:
        """업체 규칙 조회. 품의서 업체명에는 법인격·수식어가 붙어 오므로
        (예: 'EverproX Technologies(舊 BDX)') 정확히 일치하지 않으면 가장 긴 부분일치 키를 쓴다."""
        norm = _norm_vendor(vendor_name)
        if not norm:
            return {}
        if norm in self.vendors:
            return self.vendors[norm]
        hits = [k for k in self.vendors if len(k) >= 3 and k in norm]
        if hits:
            return self.vendors[max(hits, key=len)]
        return {}

    def shipper(self, vendor_name: str) -> str:
        return self.get(vendor_name).get("shipper") or str(vendor_name or "").strip()

    def file_label(self, vendor_name: str) -> str:
        """파일명에 쓰는 업체 표기 — 주문서 폴더 관례는 짧은 상호다
        (TO-PO-26-264_Coherent.xlsx, TO-PO-26-274_신광정보통신.xlsx)."""
        rule = self.get(vendor_name)
        label = rule.get("file") or rule.get("shipper") or str(vendor_name or "").strip()
        return re.sub(r"㈜|\(주\)|\(株\)|주식회사", "", label).strip()

    def remark_label(self, vendor_name: str) -> str:
        return self.get(vendor_name).get("remark") or ""

    def base_override(self, vendor_name: str) -> str:
        return self.get(vendor_name).get("base") or ""

    def payment_terms(self, vendor_name: str, raw: str, kind: str) -> str:
        """품의서 구매조건/구매방법 → 주문서 Payment Terms 표기."""
        raw = str(raw or "").strip()
        alias = self.payment_alias.get(kind) or {}
        if raw and raw in alias:
            return alias[raw]
        if raw:
            return raw
        return (self.get(vendor_name).get("payment")
                or (self.defaults.get(kind) or {}).get("payment") or "")

    def destination(self, kind: str) -> str:
        return (self.defaults.get(kind) or {}).get("destination") or "KOREA"


# ──────────────────────────────────────────────────────────────
# 주문 데이터 모델
# ──────────────────────────────────────────────────────────────
@dataclass
class PoItem:
    name: str
    qty: float
    price: float
    unit: str = DEFAULT_UNIT
    spec: str = ""          # 규격/부가설명 — 있으면 품목 아래 행에 적는다
    currency: str = "KRW"

    @property
    def amount(self) -> float:
        return float(self.qty) * float(self.price)


@dataclass
class PoOrder:
    po_no: str
    vendor_name: str
    items: list
    issue_date: date | None = None
    delivery_date: date | None = None
    payment_terms_raw: str = ""
    destination: str = ""
    reference: str = ""              # Reference NO. — 관례상 공란
    total_amt: float | None = None   # 품의서 합계금액(대조용)
    currency: str = "KRW"
    subject: str = ""
    prc_token: str = ""

    @property
    def kind(self) -> str:
        """기준 양식 구분: 내자(krw) / 외자(fx)."""
        return "krw" if str(self.currency or "KRW").upper() == "KRW" else "fx"

    @property
    def calc_total(self) -> float:
        return sum(i.amount for i in self.items)


def split_unit_spec(second_field: str):
    """품목내역 둘째 필드 → (단위, 규격).

    실무상 이 자리에는 '1 ea', 'ea', '개'처럼 단위가 들어오지만 'M8x20', '1.8M' 같은
    실제 규격이 들어오는 경우도 있다. 단위로 읽히면 E열(단위)에, 아니면 규격으로 보고
    품목 아래 행에 적는다(값을 버리지 않는다).

    ⚠ 수량이 붙은 표기는 '1'일 때만 단위로 본다. '1.8M'은 길이 규격이지 단위가 아니다
      (2026-09-17 TO-PO-26-278: '그라운드 코드 1.8M 3개'가 '3 M'으로 찍힐 뻔했다)."""
    s = str(second_field or "").strip()
    if not s or re.fullmatch(r"\d+(?:\.\d+)?", s):   # 빈칸 또는 숫자만("1") → 단위 기본값
        return DEFAULT_UNIT, ""
    m = re.fullmatch(r"(?:(\d+(?:\.\d+)?)\s*)?([A-Za-z가-힣]{1,4})", s)
    if m and m.group(2).lower() in UNIT_WORDS and (m.group(1) is None or float(m.group(1)) == 1):
        return m.group(2), ""
    return DEFAULT_UNIT, s


# ──────────────────────────────────────────────────────────────
# 양식 열기 / 구조 파악
# ──────────────────────────────────────────────────────────────
@dataclass
class FormLayout:
    sheet: Worksheet
    kind: str
    item_first: int
    item_last: int
    sum_row: int
    template: Path
    wb: object = field(repr=False, default=None)
    images: list = field(repr=False, default_factory=list)   # 로고·서명 원본 바이트


def load_form(kind: str, base_path=None) -> FormLayout:
    """기준 양식을 열고 품목 구간·합계행을 실측한다(양식마다 달라 하드코딩하지 않는다)."""
    template = Path(base_path) if base_path else TEMPLATES[kind]
    if not template.exists():
        raise SystemExit(f"[오류] 기준 양식 없음: {template}")

    # ⚠ Pillow가 없으면 openpyxl이 load→save 때 그림을 조용히 버린다 →
    #   InLC 로고와 서명이 빠진 주문서가 발행된다(검증으로도 안 잡히는 결함).
    try:
        import PIL  # noqa: F401
    except ImportError:
        raise SystemExit("[오류] Pillow 미설치 — 로고·서명 이미지가 사라진 주문서가 만들어집니다.\n"
                         "  .\\.venv\\Scripts\\python.exe -m pip install pillow")

    wb = openpyxl.load_workbook(template)

    ws = None
    for s in wb.worksheets:
        if s["A3"].value and "PURCHASE" in str(s["A3"].value).upper():
            ws = s
            break
    if ws is None:
        raise SystemExit(f"[오류] 주문서 시트를 찾지 못함(A3='PURCHASE ORDER' 없음): {template}")
    ws.title = SHEET_NAME

    sum_row = 0
    for r in range(ITEM_FIRST_ROW + 1, 60):
        v = ws.cell(r, COL_AMT).value
        if isinstance(v, str) and v.upper().startswith(f"=SUM(G{ITEM_FIRST_ROW}"):
            sum_row = r
            break
    if not sum_row:
        raise SystemExit(f"[오류] 합계행(=SUM(G14:...))을 찾지 못함: {template}")

    # 그림 원본을 메모리에 떠 둔다 — openpyxl은 한 번 저장하면 이미지 파일 핸들을 닫아
    # 같은 워크북을 두 번째 저장할 때 'I/O operation on closed file'로 죽는다(인쇄배율 재시도).
    images = [img._data() for img in ws._images]

    return FormLayout(sheet=ws, kind=kind, item_first=ITEM_FIRST_ROW,
                      item_last=sum_row - 1, sum_row=sum_row, template=template,
                      wb=wb, images=images)


# ──────────────────────────────────────────────────────────────
# 채우기
# ──────────────────────────────────────────────────────────────
def _set(ws: Worksheet, ref: str, value):
    cell = ws[ref]
    if isinstance(cell, MergedCell):      # 병합 영역의 좌상단이 아니면 건드리지 않는다
        return
    cell.value = value


def _shrink(ws: Worksheet, row: int, col: int):
    """좁은 열(단위·납기)의 글자 잘림 방지 — 자동 검사로 못 잡는 결함이라 항상 켠다."""
    c = ws.cell(row, col)
    try:
        c.alignment = c.alignment.copy(shrinkToFit=True)
    except Exception:
        pass


def fix_row_heights(layout: FormLayout):
    """글꼴 크기에 비해 낮은 행의 높이를 키운다 — 글자 윗부분 잘림 방지.

    기준 양식의 제목행(A3 'PURCHASE ORDER', 18pt 굵게)은 행 높이가 19.5pt뿐이라
    Excel 화면에서 글자 위가 잘려 보인다(세로 정렬이 '아래쪽'이라 위를 깎는다).
    PDF로 내보내면 멀쩡해서 자동 검사로는 안 잡히던 결함이다(2026-09-17 확인).
    병합 셀은 Excel의 자동 맞춤이 동작하지 않으므로 필요한 높이를 계산해 직접 넣는다.
    """
    ws = layout.sheet
    for row in range(1, ITEM_FIRST_ROW):          # 머리글 블록(1~13행)
        max_pt = 0.0
        for col in range(1, COL_REMARK + 1):
            cell = ws.cell(row, col)
            if cell.value in (None, "") and not isinstance(cell, MergedCell):
                continue
            size = getattr(cell.font, "size", None)
            if size:
                max_pt = max(max_pt, float(size))
        if not max_pt:
            continue
        need = round(max_pt * ROW_HEIGHT_RATIO, 2)
        cur = ws.row_dimensions[row].height
        if cur is None or cur < need:
            ws.row_dimensions[row].height = need     # height 설정 자체가 customHeight를 켠다


def _clear_body(layout: FormLayout):
    """품목 구간만 비운다. 2페이지 거래조건 블록·푸터·서명은 절대 건드리지 않는다."""
    ws = layout.sheet
    for r in range(layout.item_first, layout.item_last + 1):
        for col in range(1, COL_REMARK + 1):
            c = ws.cell(row=r, column=col)
            if not isinstance(c, MergedCell):
                c.value = None


def fill(layout: FormLayout, order: PoOrder, rules: VendorRules):
    """주문서 셀 채우기. 반환값은 경고 목록."""
    ws, kind = layout.sheet, layout.kind
    warnings = []

    _clear_body(layout)
    fix_row_heights(layout)

    for ref, text in LETTERHEAD[kind].items():
        _set(ws, ref, text)

    _set(ws, CELL_PO_NO, order.po_no)
    if order.issue_date:
        _set(ws, CELL_DATE, order.issue_date)
        ws[CELL_DATE].number_format = DATE_FORMAT
    _set(ws, CELL_PAYMENT, rules.payment_terms(order.vendor_name, order.payment_terms_raw, kind))
    _set(ws, CELL_DEST, order.destination or rules.destination(kind))
    _set(ws, CELL_REFERENCE, order.reference or None)      # 관례상 공란
    _set(ws, CELL_SHIPPER, rules.shipper(order.vendor_name))
    _set(ws, CELL_AMOUNT, f"=G{layout.sum_row}")

    fmt = NUMBER_FORMATS.get(str(order.currency).upper(), NUMBER_FORMATS["KRW"])
    ws[CELL_AMOUNT].number_format = fmt      # 기준 양식의 통화($ 등)가 그대로 남지 않게
    remark_label = rules.remark_label(order.vendor_name)

    row = layout.item_first
    for idx, it in enumerate(order.items, start=1):
        need = 2 if it.spec else 1
        if row + need - 1 > layout.item_last:
            warnings.append(
                f"품목 {idx}번부터 양식 품목칸({layout.item_first}~{layout.item_last}행)을 넘어 "
                f"기재하지 못했습니다. 행을 추가해 수동 입력하거나 주문을 분할하세요.")
            break
        ws.cell(row, COL_NO, idx)
        ws.cell(row, COL_DESC, it.name)
        ws.cell(row, COL_QTY, it.qty)
        ws.cell(row, COL_UNIT, it.unit or DEFAULT_UNIT)
        _shrink(ws, row, COL_UNIT)
        ws.cell(row, COL_PRICE, it.price).number_format = fmt
        ws.cell(row, COL_AMT, f"=D{row}*F{row}").number_format = fmt
        if order.delivery_date:
            ws.cell(row, COL_DELY, order.delivery_date).number_format = DATE_FORMAT
            _shrink(ws, row, COL_DELY)
        if it.spec:
            ws.cell(row + 1, COL_DESC, it.spec)
        row += need

    if remark_label:
        ws.cell(layout.item_first, COL_REMARK, remark_label)   # 값은 비운다(임의 생성 금지)

    _set(ws, f"G{layout.sum_row}", f"=SUM(G{layout.item_first}:G{layout.item_last})")
    ws.cell(layout.sum_row, COL_AMT).number_format = fmt

    layout.wb.active = layout.wb.worksheets.index(ws)
    return warnings


def set_print_mode(layout: FormLayout, mode: str, scale: int = 0):
    """인쇄 설정. mode='scale'(배율 고정) | 'width'(가로 1페이지 맞춤).

    - 'width'는 기준 양식 원본 설정(fitToPage=True)이다. 가로는 반드시 한 페이지에 들어가지만
      배율이 무시돼 거래조건 마지막 줄이 3페이지로 밀리는 경우가 있다.
    - 'scale'은 배율을 직접 낮춰 2페이지로 붙이는 방식(실제 2페이지 발행분이 쓰던 방식).
      다만 열 폭이 넓은 양식(예: Coherent)은 배율을 낮춰도 가로가 잘려 페이지가 옆으로 쪼개진다.
      → 호출부에서 세로 분할(VPageBreaks)이 없는 설정만 채택한다.
    - fitToHeight=0(여러 페이지 허용), verticalCentered=False(내용이 페이지 중앙에 뜨는 현상 방지)는
      두 모드 공통으로 강제한다."""
    ws = layout.sheet
    ws.page_setup.fitToHeight = 0
    ws.print_options.verticalCentered = False
    pr = ws.sheet_properties.pageSetUpPr
    if mode == "scale":
        ws.page_setup.scale = int(scale)
        ws.page_setup.fitToWidth = None
        if pr is not None:
            pr.fitToPage = False
    else:
        ws.page_setup.scale = None
        ws.page_setup.fitToWidth = 1
        if pr is not None:
            pr.fitToPage = True


def verify(layout: FormLayout, order: PoOrder):
    """저장 전 자체 검증 — 반환값은 오류 목록(비어야 정상)."""
    ws = layout.sheet
    errors = []

    if order.total_amt is not None:
        diff = abs(order.calc_total - float(order.total_amt))
        if diff > 0.5:
            errors.append(f"합계 불일치: 품목 계산합 {order.calc_total:,.2f} "
                          f"≠ 품의서 합계금액 {float(order.total_amt):,.2f}")

    # 기준 양식의 원 업체 흔적이 남았는가 (오염된 주문서 발행 방지)
    own = f"{order.vendor_name} {order.po_no} {ws[CELL_SHIPPER].value or ''}".lower()
    for token in BASE_RESIDUE.get(layout.kind, []):
        if token.lower() in own:                 # 발주 업체가 곧 기준 양식 업체인 경우
            continue
        for r in range(1, layout.sum_row + 1):
            for col in range(1, COL_REMARK + 1):
                v = ws.cell(r, col).value
                if isinstance(v, str) and token in v:
                    errors.append(f"기준 양식 잔재: {ws.cell(r, col).coordinate}='{v}' (기준 업체 '{token}')")

    if len(getattr(ws, "_images", [])) < 2:      # InLC 로고 + 서명
        errors.append(f"로고·서명 이미지가 {len(getattr(ws, '_images', []))}개뿐입니다 "
                      "(정상 2개). Pillow 설치 여부와 기준 양식을 확인하세요.")
    if not str(ws[CELL_PO_NO].value or "").strip():
        errors.append("P/O-No(C9)가 비었습니다.")
    if not str(ws[CELL_SHIPPER].value or "").strip():
        errors.append("Shipper(G12)가 비었습니다.")
    return errors


def save(layout: FormLayout, out_path: Path) -> Path:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    for img, data in zip(layout.sheet._images, layout.images):
        img.ref = BytesIO(data)          # 저장할 때마다 새 핸들을 물려 준다(재저장 대비)
    layout.wb.save(out_path)
    return out_path
