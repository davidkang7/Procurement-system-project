#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
주문서(PURCHASE ORDER) 생성기 v2 — 로컬 실행 전용.

  [입력] ① GAS 핸드오프 매니페스트(po_manifest.json)  ② 구매결재시스템 DB 백업 xlsx
  [양식] 2026 기준 양식 (내자=태성테크 261 / 외자=Coherent 264) — 표준거래조건 2p 포함
  [출력] {PO번호}_{업체명}.xlsx + 같은 이름 .pdf → 주문서 폴더(Drive 또는 G:\ 경로)

⚠ 결정적(deterministic)이다. 내용을 생성·추론하지 않는다. 입력에 없는 값은 비운다.

사용법:
  # ① 결재 핸드오프 (권장 경로)
  .\.venv\Scripts\python.exe render_po.py --folder <FINAL/{PO} 폴더 ID 또는 URL>

  # ② DB 백업에서 직접 (과거 건 재생성 / 핸드오프 실패 시)
  .\.venv\Scripts\python.exe render_po.py --db "..\구매결재시스템_DB_20260915.xlsx" ^
      --po TO-PO-26-278 --po TO-PO-26-279 --no-upload

  # 옛 업체 양식을 굳이 기준으로 쓸 때(권장하지 않음)
  ... --base "G:\...\주문서\TO-PO-26-270_빅터스.xlsx"

작업 후에는 반드시 PDF를 눈으로 확인한다(글자 잘림은 자동 검사로 못 잡는다).
완료 후 GAS에서 markPoDone("prcToken", "파일ID") 실행.
"""

from __future__ import annotations

import argparse
import io
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

import po_form
import po_source
from po_form import PoOrder, VendorRules

BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"
FINALIZE_PS1 = BASE_DIR / "excel_finalize.ps1"
OUT_DIR = BASE_DIR / "out"

SCOPES = ["https://www.googleapis.com/auth/drive"]
MANIFEST_NAME = "po_manifest.json"
ORDER_FOLDER_ID = "1Ot_KcJlCS8oZpu_IKCwAqz2CSPUTd81c"   # SCM_Innovation/02. Purchase/주문서

for _stream in (sys.stdout, sys.stderr):      # 콘솔 기본 코드페이지(cp949)에서 한글 깨짐 방지
    try:
        _stream.reconfigure(encoding="utf-8")
    except Exception:
        pass


# ──────────────────────────────────────────────────────────────
# Drive
# ──────────────────────────────────────────────────────────────
def get_drive():
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_FILE.exists():
                sys.exit(f"[오류] {CREDENTIALS_FILE.name} 없음. insp-renderer/credentials.json 을 복사하세요.")
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def extract_id(value: str) -> str:
    if not value:
        return ""
    m = re.search(r"/(?:folders|d)/([A-Za-z0-9_-]+)", value) or re.search(r"[?&]id=([A-Za-z0-9_-]+)", value)
    return m.group(1) if m else value.strip()


def fetch_manifest(drive, folder_id: str) -> dict:
    from googleapiclient.http import MediaIoBaseDownload

    resp = drive.files().list(
        q=f"'{folder_id}' in parents and name = '{MANIFEST_NAME}' and trashed = false",
        fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    files = resp.get("files", [])
    if not files:
        sys.exit(f"[오류] 폴더에 {MANIFEST_NAME} 이 없습니다: {folder_id}\n"
                 '  GAS에서 rerunPoHandoff("prcToken") 으로 재생성하거나 --db 경로를 쓰세요.')
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, drive.files().get_media(fileId=files[0]["id"], supportsAllDrives=True))
    done = False
    while not done:
        _, done = dl.next_chunk()
    return json.loads(buf.getvalue().decode("utf-8"))


def upload(drive, path: Path, folder_id: str) -> str:
    from googleapiclient.http import MediaFileUpload

    mime = ("application/pdf" if path.suffix.lower() == ".pdf"
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    created = drive.files().create(
        body={"name": path.name, "parents": [folder_id]},
        media_body=MediaFileUpload(str(path), mimetype=mime, resumable=False),
        fields="id, name", supportsAllDrives=True,
    ).execute()
    new_id = created["id"]

    resp = drive.files().list(                      # 같은 이름 이전 파일은 업로드 성공 후에만 정리
        q=f"'{folder_id}' in parents and name = '{path.name}' and trashed = false",
        fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    for f in resp.get("files", []):
        if f["id"] != new_id:
            try:
                drive.files().update(fileId=f["id"], body={"trashed": True}, supportsAllDrives=True).execute()
                print(f"      · 이전 파일 정리: {f['name']}")
            except Exception as e:
                print(f"      · 이전 파일 정리 실패({f['name']}): {e}")
    return new_id


# ──────────────────────────────────────────────────────────────
# Excel COM 마무리 (수식 재계산 + PDF)
# ──────────────────────────────────────────────────────────────
def finalize(xlsx: Path, pdf: Path | None):
    """Excel COM으로 수식 캐시 생성·저장하고 PDF를 내보낸다. 반환: 진단값 dict."""
    cmd = ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
           "-File", str(FINALIZE_PS1), "-Path", str(xlsx)]
    if pdf:
        cmd += ["-Pdf", str(pdf)]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    out = (proc.stdout or "") + (proc.stderr or "")
    info = dict(re.findall(r"^([A-Z_]+)=(.*)$", out, flags=re.M))
    if proc.returncode != 0 or "ERROR" in info:
        raise SystemExit(f"[오류] Excel 재계산/PDF 실패: {info.get('ERROR') or out.strip()[:500]}\n"
                         "  Excel이 설치된 PC에서 실행해야 합니다. 숨은 Excel 프로세스가 남았는지 확인하세요.")
    return info


def pdf_page_count(pdf: Path) -> int:
    data = pdf.read_bytes()
    n = len(re.findall(rb"/Type\s*/Page[^s]", data))
    return n or len(re.findall(rb"/Type\s*/Pages?", data))


# ──────────────────────────────────────────────────────────────
# 생성
# ──────────────────────────────────────────────────────────────
def choose_print_mode(layout, xlsx: Path, kind: str, args):
    """Excel로 페이지 나눔만 재보고 인쇄 설정을 고른다. 반환: (mode, scale).

    판정 규칙
      · 세로 나눔(VPageBreaks)이 하나라도 있으면 폼이 가로로 잘린다 → 그 설정은 쓰지 않는다.
      · 그 중 페이지 수가 가장 적은 설정을 쓰고, 2페이지가 나오면 즉시 채택한다.
      · 배율로 2페이지를 못 만드는 넓은 양식(Coherent 등)은 '가로 1페이지 맞춤'으로 돌아가
        3페이지(1p 폼 + 2p 거래조건 연장)로 발행한다 — 실제 발행분과 같은 모습이다.
    """
    if args.scale:
        return "scale", args.scale

    best = None                        # (pages, mode, scale)
    for scale in po_form.PAGE_SCALES[kind]:
        po_form.set_print_mode(layout, "scale", scale)
        po_form.save(layout, xlsx)
        info = finalize(xlsx, None)
        v = int(info.get("VPAGEBREAKS") or 0)
        pages = (int(info.get("HPAGEBREAKS") or 0) + 1) * (v + 1)
        if v == 0:
            if best is None or pages < best[0]:
                best = (pages, "scale", scale)
            if pages == 2:
                return "scale", scale
        print(f"      · 배율 {scale}% → {pages}페이지" + (" (가로 잘림)" if v else ""))

    po_form.set_print_mode(layout, "width")
    po_form.save(layout, xlsx)
    info = finalize(xlsx, None)
    pages = (int(info.get("HPAGEBREAKS") or 0) + 1) * (int(info.get("VPAGEBREAKS") or 0) + 1)
    print(f"      · 가로 1페이지 맞춤 → {pages}페이지")
    if best is None or pages <= best[0]:
        return "width", 0
    return best[1], best[2]


def safe_name(name: str, fallback: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]', "_", (name or "").strip())
    cleaned = cleaned.rstrip(" .")          # 윈도 파일명은 마침표·공백으로 끝날 수 없다
    return cleaned or fallback


def render_one(order: PoOrder, rules: VendorRules, args) -> dict:
    """주문 1건 → xlsx(+pdf) 생성. 반환: 결과 요약."""
    kind = args.kind or rules.base_override(order.vendor_name) or order.kind
    layout = po_form.load_form(kind, args.base)

    print(f"  · 기준 양식: {layout.template.name} ({'내자' if kind == 'krw' else '외자'}) "
          f"/ 품목칸 {layout.item_first}~{layout.item_last}행 / 합계 G{layout.sum_row}")

    if args.destination:
        order.destination = args.destination
    if args.payment:
        order.payment_terms_raw = args.payment

    warnings = po_form.fill(layout, order, rules)
    errors = po_form.verify(layout, order)
    if errors and not args.force:
        for e in errors:
            print(f"  ✗ {e}")
        raise SystemExit("[중단] 검증 실패 — 원인을 고친 뒤 다시 실행하세요(무시하려면 --force).")
    for e in errors:
        print(f"  ✗ (무시됨) {e}")
    for w in warnings:
        print(f"  · [경고] {w}")

    out_dir = Path(args.out_dir) if args.out_dir else OUT_DIR
    stem = f"{safe_name(order.po_no, 'PO')}_{safe_name(rules.file_label(order.vendor_name), 'vendor')}"
    xlsx = out_dir / f"{stem}.xlsx"
    pdf = None if args.no_pdf else out_dir / f"{stem}.pdf"

    # 인쇄 설정 선택: 1p 주문서 + 2p 거래조건 = 2페이지가 목표.
    #   · 배율을 낮추면 거래조건이 2페이지 안에 들어가지만, 열 폭이 넓은 양식은 가로로 쪼개진다.
    #   · Excel을 열어 페이지 나눔 수만 먼저 재보고(PDF 없이) 가장 적은 페이지 설정을 채택한다.
    po_form.save(layout, xlsx)
    mode, scale = choose_print_mode(layout, xlsx, kind, args)
    po_form.set_print_mode(layout, mode, scale)
    po_form.save(layout, xlsx)
    info = finalize(xlsx, pdf)
    print(f"  · 인쇄 설정: " + (f"배율 {scale}%" if mode == "scale" else "가로 1페이지 맞춤"))
    print(f"  · xlsx 저장: {xlsx}")
    pages = pdf_page_count(pdf) if pdf else None

    if info.get("HASH_CELLS"):
        print(f"  ✗ 열 너비 부족(### 표시): {info['HASH_CELLS']} — Excel에서 열 너비를 넓히세요.")
    amount = float(info.get("AMOUNT") or 0)
    if order.total_amt is not None and abs(amount - float(order.total_amt)) > 0.5:
        print(f"  ✗ 재계산 합계 불일치: 주문서 {amount:,.2f} ≠ 품의서 {float(order.total_amt):,.2f}")
    else:
        print(f"  · 합계 확인: {amount:,.2f} {order.currency}")

    if pdf:
        if pages == 2:
            print(f"  · PDF 2페이지(1p 주문서 + 2p 표준거래조건): {pdf}")
        elif pages == 3:
            print(f"  · PDF 3페이지 — 거래조건이 한 페이지를 넘었습니다(넓은 양식에서 정상 범위). "
                  f"2페이지로 붙이려면 --scale 로 배율을 낮추세요: {pdf}")
        else:
            print(f"  ✗ PDF {pages}페이지 — 1p 주문서 + 2p~ 거래조건 구성이 아닙니다. "
                  f"인쇄 영역/배율을 확인하세요: {pdf}")

    return {"order": order, "xlsx": xlsx, "pdf": pdf, "errors": errors, "warnings": warnings}


def deliver(results, args):
    """주문서 폴더로 배포 — Drive 업로드(기본) 또는 로컬/네트워크 경로 복사."""
    if args.copy_to:
        dest = Path(args.copy_to)
        dest.mkdir(parents=True, exist_ok=True)
        for r in results:
            for p in (r["xlsx"], r["pdf"]):
                if p:
                    shutil.copy(p, dest / p.name)
                    print(f"  · 복사: {dest / p.name}")
        return
    if args.no_upload:
        print("  · 업로드 생략(--no-upload) — out 폴더 파일을 확인 후 수동 배포하세요.")
        return

    drive = get_drive()
    folder = args.order_folder or ORDER_FOLDER_ID
    for r in results:
        for p in (r["xlsx"], r["pdf"]):
            if p:
                fid = upload(drive, p, folder)
                print(f"  · 업로드: {p.name} (id={fid})")


def main():
    ap = argparse.ArgumentParser(description="주문서(PO) 생성기 v2")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--folder", help="FINAL/{PO} 폴더 ID 또는 URL (po_manifest.json 위치)")
    src.add_argument("--manifest", help="로컬 po_manifest.json 경로")
    src.add_argument("--db", help="구매결재시스템 DB 백업 xlsx 경로 (--po 와 함께)")
    ap.add_argument("--po", action="append", default=[], help="--db 사용 시 대상 품의번호(반복 지정 가능)")
    ap.add_argument("--base", help="기준 양식 xlsx 직접 지정(옛 업체 양식 — 권장하지 않음)")
    ap.add_argument("--kind", choices=["krw", "fx"], help="기준 양식 강제(내자/외자). 기본은 통화로 판단")
    ap.add_argument("--out-dir", help="산출물 폴더(기본 ./out)")
    ap.add_argument("--destination", help="Destination(G10) 덮어쓰기")
    ap.add_argument("--payment", help="Payment Terms(C10) 덮어쓰기")
    ap.add_argument("--order-folder", help="업로드 대상 Drive 폴더 ID(기본 주문서 폴더)")
    ap.add_argument("--copy-to", help="Drive 업로드 대신 이 경로로 복사 (예: G:\\...\\주문서)")
    ap.add_argument("--no-upload", action="store_true", help="업로드/복사 없이 로컬 생성만")
    ap.add_argument("--no-pdf", action="store_true", help="PDF 생성 생략(검증 불가 — 권장하지 않음)")
    ap.add_argument("--scale", type=int, help="인쇄 배율(%%) 고정. 기본은 2페이지가 될 때까지 자동 조정")
    ap.add_argument("--allow-unapproved", action="store_true", help="--db 사용 시 최종승인(PRC) 아닌 건도 허용")
    ap.add_argument("--force", action="store_true", help="검증 실패에도 계속 진행")
    args = ap.parse_args()

    if args.db and not args.po:
        sys.exit("[오류] --db 는 --po TO-PO-26-XXX 와 함께 씁니다.")

    rules = VendorRules()
    orders = []
    manifests = []

    if args.folder:
        drive = get_drive()
        print(f"[1/3] 매니페스트 조회: 폴더 {extract_id(args.folder)}")
        m = fetch_manifest(drive, extract_id(args.folder))
        po_source.check_schema(m)
        manifests.append(m)
        orders.append(po_source.from_manifest(m))
    elif args.manifest:
        print(f"[1/3] 매니페스트 읽기: {args.manifest}")
        m = po_source.load_manifest_file(Path(args.manifest))
        po_source.check_schema(m)
        manifests.append(m)
        orders.append(po_source.from_manifest(m))
    else:
        print(f"[1/3] DB 읽기: {args.db}")
        for po_no in args.po:
            orders.append(po_source.from_db(Path(args.db), po_no.strip(),
                                            require_approved=not args.allow_unapproved))

    results = []
    for order in orders:
        print(f"[2/3] 주문서 생성: {order.po_no} / {order.vendor_name} / "
              f"품목 {len(order.items)}건 / {order.calc_total:,.2f} {order.currency}")
        results.append(render_one(order, rules, args))

    print("[3/3] 주문서 폴더 배포")
    deliver(results, args)

    print("\n─────────────────────────────────────────────")
    print("생성한 주문서:")
    for r in results:
        print(f"  {r['xlsx'].name}" + ("  + PDF" if r["pdf"] else ""))
    print("\n※ PDF를 반드시 육안 확인하세요 — 글자 잘림·줄바꿈은 자동 검사로 잡히지 않습니다.")
    for m in manifests:
        if m.get("prcToken"):
            print(f'※ 마감: GAS에서  markPoDone("{m["prcToken"]}", "업로드된 파일ID")')
    print("─────────────────────────────────────────────")


if __name__ == "__main__":
    main()
