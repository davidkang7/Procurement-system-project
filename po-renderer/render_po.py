#!/usr/bin/env python3
"""
주문서(PURCHASE ORDER) 생성기 — 로컬 실행 전용.

GAS가 PRC 최종승인 시 STAGING 폴더에 기록한 manifest.json을 입력으로 받아
QF-741-2 주문서 Excel 템플릿(po_template.xlsx)의 셀 값만 채워
주문서 폴더에 `{품의제목}_{업체명}.xlsx` 로 저장한다.

⚠ 이 스크립트는 결정적(deterministic)이다. 내용을 생성/추론하지 않는다.
   manifest.json의 값만 그대로 채운다. (문서 위조 방지)

⚠ 왜 openpyxl로 저장하지 않는가:
   템플릿의 InLC 로고·서명 이미지와 표 괘선(drawing1.xml의 Line 도형)은
   openpyxl이 load→save 시 전부 소실시킨다(검증 완료). 그래서 xlsx를 zip으로 열어
   sheet2.xml의 대상 셀 값만 치환하고 나머지(media/drawings/styles/rels)는
   바이트 그대로 보존한다.

사용법:
  # STAGING 폴더에서 manifest를 찾아 xlsx 생성 + 업로드
  python render_po.py --folder <STAGING 폴더 ID 또는 URL>

  # 업로드 없이 로컬 xlsx만 (미리보기)
  python render_po.py --folder <ID> --no-upload

  # 로컬 manifest.json으로 생성
  python render_po.py --manifest ./manifest.json
"""

from __future__ import annotations

import argparse
import io
import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape as xml_escape

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

# ──────────────────────────────────────────────────────────────
# 설정
# ──────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"   # GCP OAuth 데스크톱 클라이언트
TOKEN_FILE = BASE_DIR / "token.json"               # 최초 동의 후 자동 생성/갱신
TEMPLATE_FILE = BASE_DIR / "templates" / "po_template.xlsx"
SHEET_XML = "xl/worksheets/sheet2.xml"             # 주문서 본문 시트('photop')

SCOPES = ["https://www.googleapis.com/auth/drive"]

MANIFEST_NAME = "po_manifest.json"
SUPPORTED_SCHEMA = "po-manifest-v1"

# 주문서 양식 셀 좌표 (QF-741-2 Rev 0)
CELL_PO_NO       = "C9"    # P/O-No
CELL_PAY_TERMS   = "C10"   # Payment Terms
CELL_DEST        = "G10"   # Destination
CELL_REFERENCE   = "C11"   # Reference NO.(품의번호)
CELL_MAKER       = "G11"   # Maker
CELL_SHIPPER     = "G12"   # Shipper(업체명)

DEFAULT_DEST = "INLC Technology, Daejeon, Korea"
DEFAULT_UNIT = "ea"

# 품목 슬롯: 14,16,...,34 행 (2행 1블록, 총 11개). 넘치면 경고 후 잘림(사용자 수동 추가).
ITEM_START_ROW = 14
ITEM_ROW_STEP  = 2
ITEM_MAX_SLOTS = 11


# ──────────────────────────────────────────────────────────────
# 인증 / Drive
# ──────────────────────────────────────────────────────────────
def get_drive():
    """OAuth 데스크톱 흐름으로 Drive 서비스 생성. 최초 1회만 브라우저 동의."""
    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_FILE.exists():
                sys.exit(
                    f"[오류] {CREDENTIALS_FILE.name} 없음.\n"
                    "  insp-renderer/credentials.json 을 복사해 두거나, GCP 콘솔에서\n"
                    "  OAuth 클라이언트 ID(데스크톱 앱) JSON을 이 위치에 저장하세요."
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")

    return build("drive", "v3", credentials=creds, cache_discovery=False)


def extract_id(value: str) -> str:
    """폴더/파일 URL 또는 순수 ID에서 ID만 뽑는다."""
    if not value:
        return ""
    m = re.search(r"/(?:folders|d)/([A-Za-z0-9_-]+)", value)
    if m:
        return m.group(1)
    m = re.search(r"[?&]id=([A-Za-z0-9_-]+)", value)
    if m:
        return m.group(1)
    return value.strip()


def find_manifest(drive, folder_id: str) -> str:
    """STAGING 폴더에서 po_manifest.json 파일 ID를 찾는다."""
    resp = drive.files().list(
        q=f"'{folder_id}' in parents and name = '{MANIFEST_NAME}' and trashed = false",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = resp.get("files", [])
    if not files:
        sys.exit(
            f"[오류] 폴더에 {MANIFEST_NAME} 이 없습니다: {folder_id}\n"
            "  GAS에서 PRC 결재가 최종 완료되어야 매니페스트가 생성됩니다.\n"
            "  (필요 시 GAS에서 rerunPoHandoff(\"prcToken\") 실행)"
        )
    return files[0]["id"]


def download_bytes(drive, file_id: str) -> bytes:
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, drive.files().get_media(fileId=file_id, supportsAllDrives=True))
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()


def upload_xlsx(drive, xlsx_path: Path, folder_id: str, file_name: str) -> str:
    """xlsx를 주문서 폴더에 업로드. 같은 이름의 이전 파일은 업로드 성공 후에만 정리."""
    meta = {"name": file_name, "parents": [folder_id]}
    media = MediaFileUpload(
        str(xlsx_path),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False,
    )
    created = drive.files().create(
        body=meta, media_body=media, fields="id, name, webViewLink", supportsAllDrives=True
    ).execute()
    new_id = created["id"]

    # 동일 이름 이전 파일 정리 (안전: 새 업로드 성공 후에만)
    resp = drive.files().list(
        q=f"'{folder_id}' in parents and name = '{file_name}' and trashed = false",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    for f in resp.get("files", []):
        if f["id"] != new_id:
            try:
                drive.files().update(fileId=f["id"], body={"trashed": True},
                                     supportsAllDrives=True).execute()
                print(f"  · 이전 파일 정리: {f['name']}")
            except Exception as e:
                print(f"  · 이전 파일 정리 실패({f['name']}): {e}")

    return new_id


# ──────────────────────────────────────────────────────────────
# xlsx 셀 패치 (zip/XML 직접 수정 — 이미지·도형·서식 보존)
# ──────────────────────────────────────────────────────────────
def _cell_pattern(ref: str) -> str:
    """<c r="ref" .../> (self-closing) 또는 <c r="ref" ...>...</c> 를 매칭.
    self-closing의 '/'를 attr 부분이 삼키지 않도록 두 형태를 분리한다."""
    esc = re.escape(ref)
    return r'<c r="%s"(?: [^>]*?)?/>|<c r="%s"(?: [^>]*?)?>.*?</c>' % (esc, esc)


def _cell_style(sheet_xml: str, ref: str) -> str:
    """셀의 기존 스타일 인덱스(s=) 추출. 없으면 빈 문자열."""
    m = re.search(r'<c r="%s"((?: [^>]*?)?)/?>' % re.escape(ref), sheet_xml)
    if not m:
        return ""
    sm = re.search(r's="(\d+)"', m.group(1))
    return sm.group(1) if sm else ""


def _replace_cell(sheet_xml: str, ref: str, new_cell: str) -> str:
    """<c r="ref" .../> 또는 <c r="ref">...</c> 전체를 new_cell로 치환."""
    new_xml, n = re.subn(_cell_pattern(ref), lambda _m: new_cell, sheet_xml, count=1, flags=re.S)
    if n == 0:
        raise ValueError(f"셀을 찾지 못함: {ref}")
    return new_xml


def _style_attr(style: str) -> str:
    return f' s="{style}"' if style else ""


def set_string(sheet_xml: str, ref: str, text: str) -> str:
    style = _cell_style(sheet_xml, ref)
    text = "" if text is None else str(text)
    cell = (f'<c r="{ref}"{_style_attr(style)} t="inlineStr">'
            f'<is><t xml:space="preserve">{xml_escape(text)}</t></is></c>')
    return _replace_cell(sheet_xml, ref, cell)


def set_number(sheet_xml: str, ref: str, num) -> str:
    style = _cell_style(sheet_xml, ref)
    # 정수면 정수로, 아니면 실수로
    if isinstance(num, float) and num.is_integer():
        num = int(num)
    cell = f'<c r="{ref}"{_style_attr(style)}><v>{num}</v></c>'
    return _replace_cell(sheet_xml, ref, cell)


def set_formula(sheet_xml: str, ref: str, formula: str, cached) -> str:
    style = _cell_style(sheet_xml, ref)
    if isinstance(cached, float) and cached.is_integer():
        cached = int(cached)
    cell = (f'<c r="{ref}"{_style_attr(style)}>'
            f'<f>{xml_escape(formula)}</f><v>{cached}</v></c>')
    return _replace_cell(sheet_xml, ref, cell)


def clear_cell(sheet_xml: str, ref: str) -> str:
    """값/수식 제거, 스타일(괘선) 보존."""
    style = _cell_style(sheet_xml, ref)
    return _replace_cell(sheet_xml, ref, f'<c r="{ref}"{_style_attr(style)}/>')


def build_description(item: dict) -> str:
    name = str(item.get("name") or "").strip()
    spec = str(item.get("spec") or "").strip()
    if spec:
        return f"{name} / {spec}" if name else spec
    return name


def fill_sheet(sheet_xml: str, m: dict[str, Any]) -> str:
    """주문서 셀 채우기. 수식 셀(G열 금액, 합계, TODAY 등)은 건드리지 않는다."""
    xml = sheet_xml

    # 헤더/메타
    xml = set_string(xml, CELL_PO_NO,     m.get("poNo") or "")
    xml = set_string(xml, CELL_PAY_TERMS, m.get("paymentTerms") or "")
    xml = set_string(xml, CELL_DEST,      m.get("destination") or DEFAULT_DEST)
    xml = set_string(xml, CELL_REFERENCE, m.get("reference") or "")
    xml = set_string(xml, CELL_SHIPPER,   m.get("vendorName") or "")

    # 품목
    items = m.get("items") or []
    if len(items) > ITEM_MAX_SLOTS:
        print(f"  · [경고] 품목 {len(items)}개 중 {ITEM_MAX_SLOTS}개만 채웁니다. "
              f"나머지는 Excel에서 행을 추가해 수동 입력하세요.")
    total = 0.0
    for slot in range(ITEM_MAX_SLOTS):
        row = ITEM_START_ROW + slot * ITEM_ROW_STEP
        col_a, col_b, col_d, col_e, col_f, col_g = (
            f"A{row}", f"B{row}", f"D{row}", f"E{row}", f"F{row}", f"G{row}")
        if slot < len(items):
            it = items[slot]
            qty = float(it.get("qty") or 0)
            price = float(it.get("price") or 0)
            amount = qty * price
            total += amount
            unit = str(it.get("unit") or DEFAULT_UNIT)
            xml = set_number(xml, col_a, slot + 1)
            xml = set_string(xml, col_b, build_description(it))
            xml = set_number(xml, col_d, qty)
            xml = set_string(xml, col_e, unit)
            xml = set_number(xml, col_f, price)
            xml = set_formula(xml, col_g, f"D{row}*F{row}", amount)
        else:
            for ref in (col_a, col_b, col_d, col_e, col_f, col_g):
                xml = clear_cell(xml, ref)

    # 합계 캐시값 갱신 (Excel 재계산 전에도 올바른 총액 표시). 수식 자체는 유지.
    xml = set_formula(xml, "G37", "SUM(G14:G36)", total)  # Amount(합계)
    xml = set_formula(xml, "C12", "G37", total)           # 상단 Amount 표시
    return xml


def render_xlsx(m: dict[str, Any], out_path: Path) -> Path:
    """템플릿 zip을 열어 sheet2.xml만 치환하고 새 xlsx로 저장."""
    if not TEMPLATE_FILE.exists():
        sys.exit(f"[오류] 템플릿 없음: {TEMPLATE_FILE}")

    with zipfile.ZipFile(TEMPLATE_FILE, "r") as zin:
        names = zin.namelist()
        if SHEET_XML not in names:
            sys.exit(f"[오류] 템플릿에 {SHEET_XML} 없음 — 잘못된 템플릿")
        sheet_xml = zin.read(SHEET_XML).decode("utf-8")
        new_sheet = fill_sheet(sheet_xml, m)

        out_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == SHEET_XML:
                    data = new_sheet.encode("utf-8")
                zout.writestr(item, data)
    return out_path


# ──────────────────────────────────────────────────────────────
# 유틸
# ──────────────────────────────────────────────────────────────
def safe_name(name: str, fallback: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]', "_", (name or "").strip())
    return cleaned or fallback


def po_file_name(m: dict[str, Any]) -> str:
    """파일명 = {품의제목}_{업체명}.xlsx (사용자 지정 규칙)."""
    subject = safe_name(m.get("subject"), m.get("poNo") or "PO")
    vendor = safe_name(m.get("vendorName"), "vendor")
    return f"{subject}_{vendor}.xlsx"


# ──────────────────────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────────────────────
def main() -> None:
    ap = argparse.ArgumentParser(description="주문서(PO) Excel 생성기 (po_manifest.json 기반)")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--folder", help="STAGING 폴더 ID 또는 URL (po_manifest.json이 있는 폴더)")
    src.add_argument("--manifest", help="로컬 po_manifest.json 경로")
    ap.add_argument("--out", help="xlsx 저장 경로(로컬 사본). 생략 시 ./out/ 에 저장")
    ap.add_argument("--no-upload", action="store_true", help="Drive 업로드 없이 로컬 xlsx만 생성")
    args = ap.parse_args()

    need_drive = bool(args.folder) or (not args.no_upload)
    drive = get_drive() if need_drive else None

    # 매니페스트 확보
    if args.folder:
        folder_id = extract_id(args.folder)
        print(f"[1/3] 매니페스트 조회: 폴더 {folder_id}")
        manifest_id = find_manifest(drive, folder_id)
        manifest = json.loads(download_bytes(drive, manifest_id).decode("utf-8"))
    else:
        print(f"[1/3] 매니페스트 읽기: {args.manifest}")
        manifest = json.loads(Path(args.manifest).read_text(encoding="utf-8"))

    schema = manifest.get("schemaVersion")
    if schema != SUPPORTED_SCHEMA:
        print(f"  · [경고] 매니페스트 스키마 불일치: {schema} (지원: {SUPPORTED_SCHEMA})")

    print(f"      PO {manifest.get('poNo','')} / {manifest.get('subject','')} / "
          f"{manifest.get('vendorName','')} / 품목 {len(manifest.get('items') or [])}건")

    # 렌더
    file_name = po_file_name(manifest)
    out_path = Path(args.out) if args.out else (BASE_DIR / "out" / file_name)
    print(f"[2/3] 주문서 xlsx 생성: {file_name}")
    render_xlsx(manifest, out_path)
    print(f"      로컬 저장: {out_path}")

    # 업로드
    if args.no_upload:
        print("[3/3] 업로드 생략(--no-upload)")
        file_id = None
    else:
        order_folder = manifest.get("orderFolderId")
        if not order_folder:
            sys.exit("[오류] 매니페스트에 orderFolderId 없음 — 주문서 폴더를 알 수 없습니다.")
        print(f"[3/3] 주문서 폴더 업로드: {order_folder}")
        file_id = upload_xlsx(drive, out_path, order_folder, file_name)
        print(f"      업로드 완료: {file_name} (id={file_id})")

    token = manifest.get("prcToken", "")
    print("\n─────────────────────────────────────────────")
    if file_id:
        print("완료. GAS에서 아래를 실행해 마감하세요:")
        print(f'  markPoDone("{token}", "{file_id}")')
    else:
        print("로컬 미리보기만 생성했습니다. 확인 후 --no-upload 없이 다시 실행하세요.")
    print("─────────────────────────────────────────────")


if __name__ == "__main__":
    main()
