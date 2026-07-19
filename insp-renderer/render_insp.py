#!/usr/bin/env python3
"""
검수보고서(INSP) PDF 렌더러 — 로컬 실행 전용.

GAS가 결재 완료 시 STAGING 폴더에 기록한 manifest.json과 사진을 입력으로 받아
Playwright(headless Chromium)로 PDF를 만들고 FINAL/{PO번호} 폴더에 업로드한다.

⚠ 이 스크립트는 결정적(deterministic)이다. 내용을 생성하거나 추론하지 않는다.
   manifest.json에 있는 값만 그대로 렌더한다. (문서 위조 방지)

사용법:
  # STAGING 폴더에서 manifest를 찾아 렌더 + 업로드
  python render_insp.py --folder <STAGING 폴더 ID 또는 URL>

  # 업로드 없이 미리보기만 (로컬 PDF 확인)
  python render_insp.py --folder <ID> --no-upload

  # 로컬에 내려둔 manifest.json으로 렌더 (사진은 Drive에서 id로 받음)
  python render_insp.py --manifest ./manifest.json
"""

from __future__ import annotations

import argparse
import io
import json
import re
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from jinja2 import Environment, FileSystemLoader, select_autoescape
from playwright.sync_api import sync_playwright

# ──────────────────────────────────────────────────────────────
# 설정
# ──────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"   # GCP OAuth 데스크톱 클라이언트
TOKEN_FILE = BASE_DIR / "token.json"               # 최초 동의 후 자동 생성/갱신
TEMPLATE_DIR = BASE_DIR / "templates"
TEMPLATE_NAME = "insp_report.html"

# 사진 다운로드 + PDF 업로드에 필요. 공유 드라이브의 기존 파일을 다뤄야 하므로 전체 drive 스코프.
SCOPES = ["https://www.googleapis.com/auth/drive"]

MANIFEST_NAME = "manifest.json"
SUPPORTED_SCHEMA = "insp-manifest-v1"


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
                    "  GCP 콘솔 → 사용자 인증 정보 → OAuth 클라이언트 ID(데스크톱 앱)를 만들어\n"
                    f"  JSON을 {CREDENTIALS_FILE} 로 저장하세요. (동의 화면은 '내부(Internal)'로 설정)"
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
    """STAGING 폴더에서 manifest.json 파일 ID를 찾는다."""
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
            "  GAS에서 결재가 최종 완료되어야 매니페스트가 생성됩니다.\n"
            "  (필요 시 GAS에서 rerunInspHandoff(\"token\") 실행)"
        )
    return files[0]["id"]


def download_file(drive, file_id: str, dest: Path) -> Path:
    """Drive 파일을 로컬로 내려받는다."""
    request = drive.files().get_media(fileId=file_id, supportsAllDrives=True)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    dest.write_bytes(buf.getvalue())
    return dest


def upload_pdf(drive, pdf_path: Path, folder_id: str, doc_no: str) -> str:
    """PDF를 FINAL/{PO} 폴더에 업로드하고, 같은 회차의 이전 PDF를 정리한다.

    ⚠ 순서 중요: 새 파일 업로드가 성공한 뒤에만 이전 버전을 휴지통으로 보낸다.
       (먼저 지우면 업로드 실패 시 멀쩡한 이전 PDF만 사라진다)
    """
    meta = {"name": pdf_path.name, "parents": [folder_id]}
    media = MediaFileUpload(str(pdf_path), mimetype="application/pdf", resumable=False)
    created = drive.files().create(
        body=meta, media_body=media, fields="id, name, webViewLink", supportsAllDrives=True
    ).execute()
    new_id = created["id"]

    # 동일 검수보고서의 이전 PDF 정리 (INSP_{docNo}_ 접두사)
    prefix = f"INSP_{doc_no}_"
    resp = drive.files().list(
        q=f"'{folder_id}' in parents and trashed = false",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    for f in resp.get("files", []):
        if f["id"] != new_id and f["name"].startswith(prefix) and f["name"].endswith(".pdf"):
            try:
                drive.files().update(fileId=f["id"], body={"trashed": True},
                                     supportsAllDrives=True).execute()
                print(f"  · 이전 PDF 정리: {f['name']}")
            except Exception as e:  # 정리 실패는 치명적이지 않음
                print(f"  · 이전 PDF 정리 실패({f['name']}): {e}")

    return new_id


# ──────────────────────────────────────────────────────────────
# 렌더
# ──────────────────────────────────────────────────────────────
def safe_name(name: str, fallback: str) -> str:
    """파일명으로 쓸 수 있게 정리."""
    cleaned = re.sub(r'[\\/:*?"<>|]', "_", (name or "").strip())
    return cleaned or fallback


def render_pdf(manifest: dict[str, Any], workdir: Path, drive) -> Path:
    """사진 다운로드 → HTML 렌더 → Playwright로 PDF 생성. 생성된 PDF 경로 반환."""
    # 1) 사진 다운로드 (매니페스트의 Drive id 기준)
    photos = []
    for idx, p in enumerate(manifest.get("photos") or [], start=1):
        pid = p.get("id")
        if not pid:
            continue
        fname = safe_name(p.get("name"), f"photo_{idx}.jpg")
        local = workdir / f"{idx:02d}_{fname}"
        try:
            download_file(drive, pid, local)
            # 템플릿과 같은 폴더에 두므로 상대경로로 참조
            photos.append({"src": local.name, "name": p.get("name") or local.name})
            print(f"  · 사진 {idx}: {fname}")
        except Exception as e:
            print(f"  · [경고] 사진 {idx} 다운로드 실패({pid}): {e}")

    expected = len(manifest.get("photos") or [])
    if expected and len(photos) != expected:
        print(f"  · [경고] 사진 {expected}장 중 {len(photos)}장만 확보됨")

    # 2) HTML 렌더
    env = Environment(
        loader=FileSystemLoader(str(TEMPLATE_DIR)),
        autoescape=select_autoescape(["html"]),
    )
    template = env.get_template(TEMPLATE_NAME)
    html = template.render(
        m=manifest,
        photos=photos,
        verdict_class="pass" if manifest.get("verdict") == "합격" else "fail",
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
    )
    html_path = workdir / "report.html"
    html_path.write_text(html, encoding="utf-8")

    # 3) Playwright → PDF
    doc_no = manifest.get("docNo") or "INSP"
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    pdf_path = workdir / f"INSP_{doc_no}_{ts}.pdf"

    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        try:
            page = browser.new_page()
            page.goto(html_path.as_uri(), wait_until="load")
            page.emulate_media(media="print")
            page.pdf(
                path=str(pdf_path),
                format="A4",
                print_background=True,
                prefer_css_page_size=True,
            )
        finally:
            browser.close()

    return pdf_path


# ──────────────────────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────────────────────
def main() -> None:
    ap = argparse.ArgumentParser(description="검수보고서 PDF 생성기 (manifest.json 기반)")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--folder", help="STAGING 폴더 ID 또는 URL (manifest.json이 있는 폴더)")
    src.add_argument("--manifest", help="로컬 manifest.json 경로")
    ap.add_argument("--out", help="PDF 저장 경로(로컬 사본). 생략 시 ./out/ 에 저장")
    ap.add_argument("--no-upload", action="store_true", help="Drive 업로드 없이 로컬 PDF만 생성")
    args = ap.parse_args()

    drive = get_drive()

    # 매니페스트 확보
    if args.folder:
        folder_id = extract_id(args.folder)
        print(f"[1/4] 매니페스트 조회: 폴더 {folder_id}")
        manifest_id = find_manifest(drive, folder_id)
        raw = io.BytesIO()
        downloader = MediaIoBaseDownload(
            raw, drive.files().get_media(fileId=manifest_id, supportsAllDrives=True)
        )
        done = False
        while not done:
            _, done = downloader.next_chunk()
        manifest = json.loads(raw.getvalue().decode("utf-8"))
    else:
        print(f"[1/4] 매니페스트 읽기: {args.manifest}")
        manifest = json.loads(Path(args.manifest).read_text(encoding="utf-8"))

    schema = manifest.get("schemaVersion")
    if schema != SUPPORTED_SCHEMA:
        print(f"  · [경고] 매니페스트 스키마 불일치: {schema} (지원: {SUPPORTED_SCHEMA})")

    doc_no = manifest.get("docNo", "")
    print(f"      문서번호 {doc_no} / {manifest.get('subject', '')} / 판정 {manifest.get('verdict', '')}")

    # 렌더
    with tempfile.TemporaryDirectory(prefix="insp_") as tmp:
        workdir = Path(tmp)
        print("[2/4] 사진 다운로드 + PDF 렌더")
        pdf_path = render_pdf(manifest, workdir, drive)

        # 로컬 사본 저장
        out_path = Path(args.out) if args.out else (BASE_DIR / "out" / pdf_path.name)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(pdf_path.read_bytes())
        print(f"[3/4] 로컬 PDF 저장: {out_path}")

        # 업로드
        if args.no_upload:
            print("[4/4] 업로드 생략(--no-upload)")
            file_id = None
        else:
            final_folder = manifest.get("finalFolderId")
            if not final_folder:
                sys.exit("[오류] 매니페스트에 finalFolderId 없음 — 업로드 대상 폴더를 알 수 없습니다.")
            print(f"[4/4] PO 폴더 업로드: {final_folder}")
            file_id = upload_pdf(drive, pdf_path, final_folder, doc_no)
            print(f"      업로드 완료: {pdf_path.name} (id={file_id})")

    # 마무리 안내
    token = manifest.get("token", "")
    print("\n─────────────────────────────────────────────")
    if file_id:
        print("완료. GAS에서 아래를 실행해 마감하세요:")
        print(f'  markInspPdfDone("{token}", "{file_id}")')
    else:
        print("로컬 미리보기만 생성했습니다. 확인 후 --no-upload 없이 다시 실행하세요.")
    print("─────────────────────────────────────────────")


if __name__ == "__main__":
    main()
