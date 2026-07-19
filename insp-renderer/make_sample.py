#!/usr/bin/env python3
"""샘플 검수보고서 PDF 생성 — 구글 인증 없이 레이아웃 확인용.

가짜 데이터와 가짜 사진으로 render_insp.render_pdf()를 그대로 돌린다.
템플릿(templates/insp_report.html)을 수정한 뒤 결과를 눈으로 확인할 때 사용한다.

사용법:
    .\\.venv\\Scripts\\python.exe make_sample.py            # 사진 3장 (기본)
    .\\.venv\\Scripts\\python.exe make_sample.py --photos 6  # 사진 6장 (2페이지 확인)
    .\\.venv\\Scripts\\python.exe make_sample.py --photos 0  # 사진 없음

결과: out/SAMPLE_검수보고서_사진{N}장.pdf  +  같은 이름의 .png (미리보기)
"""

from __future__ import annotations

import argparse
import shutil
import struct
import tempfile
import zlib
from pathlib import Path

import render_insp
from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).resolve().parent
OUT_DIR = BASE_DIR / "out"

# 가짜 사진 팔레트 (색·비율을 다르게 해서 object-fit 동작 확인)
PALETTE = [
    ("front.jpg", 800, 600, (70, 130, 180)),
    ("label.jpg", 600, 800, (180, 120, 70)),
    ("serial_number.jpg", 1000, 400, (110, 160, 110)),
    ("packing.jpg", 800, 800, (150, 110, 160)),
    ("dimension.jpg", 900, 500, (200, 160, 80)),
    ("overall.jpg", 700, 700, (90, 140, 150)),
]


def make_png(path: Path, w: int, h: int, rgb: tuple[int, int, int]) -> None:
    """의존성 없이 단색 PNG를 만든다 (테스트용 가짜 사진)."""
    def chunk(typ: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + typ + data
                + struct.pack(">I", zlib.crc32(typ + data) & 0xFFFFFFFF))

    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
    png = (b"\x89PNG\r\n\x1a\n"
           + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
           + chunk(b"IDAT", zlib.compress(raw))
           + chunk(b"IEND", b""))
    path.write_bytes(png)


def build_manifest(n_photos: int) -> dict:
    return {
        "schemaVersion": "insp-manifest-v1",
        "token": "SAMPLE-TOKEN-0000",
        "docNo": "TH-AP-26-029-01",
        "seq": 1,
        "isFinal": True,
        "reqNo": "TH-AP-26-029",
        "poNo": "PO-2026-0142",
        "subject": "유량계 교체 부품 (DN50 오리피스 플레이트 외 2종)",
        "vendorName": "○○계측기술",
        "drafterName": "홍길동",
        "dept": "기술팀",
        "issueDate": "2026-07-10",
        "receivedDate": "2026-07-15",
        "receivedNote": "3EA 입고, 외관 이상 없음. 납품서 대조 완료.",
        "verdict": "합격",
        "comment": ("치수 및 작동 상태 정상 확인.\n"
                    "- 오리피스 내경 측정값 규격 내\n"
                    "- 플랜지 면 손상 없음\n"
                    "- 시리얼 번호 납품서 일치"),
        "approvers": [
            {"label": "기안자", "name": "홍길동", "status": "승인", "processedAt": "2026-07-15 14:02"},
            {"label": "결재자 1", "name": "김철수", "status": "승인", "processedAt": "2026-07-15 16:30"},
            {"label": "결재자 2", "name": "박영희", "status": "승인", "processedAt": "2026-07-16 09:14"},
        ],
        "photos": [{"id": f"sample_{i}", "name": PALETTE[i][0]} for i in range(n_photos)],
        "finalFolderId": "SAMPLE_FOLDER",
        "generatedAt": "2026-07-19 16:31",
    }


def main() -> None:
    ap = argparse.ArgumentParser(description="샘플 검수보고서 PDF 생성 (인증 불필요)")
    ap.add_argument("--photos", type=int, default=3, choices=range(0, 7),
                    help="가짜 사진 장수 (0~6, 기본 3)")
    args = ap.parse_args()

    OUT_DIR.mkdir(exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="insp_sample_") as tmp:
        tmpdir = Path(tmp)

        # 가짜 사진 생성 + download_file 대체 (Drive 접근 없이 동작)
        fakes: dict[str, Path] = {}
        for i in range(args.photos):
            name, w, h, rgb = PALETTE[i]
            p = tmpdir / f"src_{name}.png"
            make_png(p, w, h, rgb)
            fakes[f"sample_{i}"] = p

        render_insp.download_file = (
            lambda drive, file_id, dest: (shutil.copy(fakes[file_id], dest), dest)[1]
        )

        workdir = tmpdir / "work"
        workdir.mkdir()
        pdf = render_insp.render_pdf(build_manifest(args.photos), workdir, drive=None)

        stem = f"SAMPLE_검수보고서_사진{args.photos}장"
        pdf_out = OUT_DIR / f"{stem}.pdf"
        shutil.copy(pdf, pdf_out)

        # 미리보기 PNG (A4 인쇄영역 기준)
        mm = 96 / 25.4
        png_out = OUT_DIR / f"{stem}.png"
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            try:
                page = browser.new_page(
                    viewport={"width": int((210 - 32) * mm), "height": int((297 - 28) * mm)}
                )
                page.emulate_media(media="print")
                page.goto((workdir / "report.html").as_uri(), wait_until="load")
                page.screenshot(path=str(png_out), full_page=True)
            finally:
                browser.close()

    print(f"\nPDF : {pdf_out}  ({pdf_out.stat().st_size // 1024} KB)")
    print(f"PNG : {png_out}")


if __name__ == "__main__":
    main()
