# GAS PDF 이미지 렌더링 한계 및 해결법

이 프로젝트의 PDF 생성에서 `Utilities.newBlob(html,'text/html').getAs(MimeType.PDF)` 변환기는 **`<img>` 이미지를 전혀 렌더링하지 못한다** — data URI(base64)든 외부 URL이든 모두 "깨진 이미지" 자리표시자로 나온다. CSS(배경색·표·텍스트)는 정상 렌더된다.

검수보고서(INSP) PDF에서 사진이 깨지던 원인이 이것이었다. 사진 크기를 줄여도(썸네일) 해결 안 됨 — 크기 문제가 아니라 변환기가 이미지를 아예 무시하는 구조적 한계.

## 해결법 (2026-06-29 적용)

HTML을 먼저 Google 문서로 변환(`Drive.Files.create`, mimeType=`GOOGLE_DOCS`)한 뒤 그 문서를 `getAs(MimeType.PDF)`로 내보낸다. 문서 변환기는 base64 data URI 이미지를 실제로 임베드한다. 무거운 변환은 Google 서버가 처리해 스크립트 부하가 작다. 임시 문서는 변환 후 `Drive.Files.remove`로 삭제.

해당 구현: `INSP.gs`의 `generateInspPdf()` (HTML → 임시 Google 문서 → PDF 추출 → 임시 문서 삭제).

## 주의

- Google 문서 변환기는 `<style>` 클래스 CSS를 **부분적으로만** 반영한다. flex / `@page` / 정밀 패딩 등은 무시되므로, **표 기반 레이아웃 + 인라인 스타일**을 써야 안정적으로 렌더된다.
- 검수 사진 그리드는 flex → 2열 `<table>` + 인라인 스타일로 변경함 (`INSP_PDF_Template.html`).
- 사진은 원본 대신 Drive 썸네일(약 1000px 축소본)을 base64로 임베드 → 임시 문서가 가벼워 변환도 빠름 (`_fetchInspPhotoBlobForPdf()`).

## REQ/PRC PDF는 무관

REQ/PRC PDF(`PDF_Template.html`)는 실사진 없이 CSS 결재란만 있어 기존 `getAs(MimeType.PDF)` 직접 변환으로도 문제없음. **INSP PDF만** 이 문서 변환 경로를 사용한다.
