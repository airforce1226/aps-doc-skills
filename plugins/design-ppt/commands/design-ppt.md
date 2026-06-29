---
description: Claude Design 스타일 HTML 덱을 편집 가능한 네이티브 .pptx로 빌드 (이미지본은 --mode image, IT전략팀 규칙 적용)
argument-hint: "[보고서 제목/유형] [기존 deck.html 경로(선택)]"
---

You are starting the design-ppt workflow.

**REQUIRED:** Use the `creating-design-ppt` skill and follow its workflow exactly:
1. Gather content (report type + reader 임원/실무); never invent unknowns → mark `(미정 — 추후 확정)`.
2. Assemble a `deck.html` from `assets/sections/` snippets; fill each section's content and `data-speaker-notes`, following `assets/design-tokens.md` and `assets/base.css`.
3. Build the deck — **default to editable native objects**:
   `python scripts/build_design_ppt.py deck.html "<제목> v1.0.pptx" --mode native`.
   Use `--mode image` only if the user asks for a pixel-perfect (non-editable) image deck.
4. Verify the deck (slide count, `author == "IT전략팀"`; in native mode confirm text/표/도형 are
   real editable objects, not a baked picture — check the build's `native=/raster=` report).

Team rules (enforced): author = **IT전략팀** (never a personal name); filename uses spaces, **no underscores**, with a version suffix like `v1.0` at the end.

User request / arguments: $ARGUMENTS

Guidance:
- If a report title/type is given, use it; otherwise ask. If a `deck.html` path is given, base the build on it.
- **보안 등급은 묻지 않는다.** APS 사내 보고서는 모두 **대외비**이므로 기본으로 전 슬라이드에 "대외비" 배지를 적용한다. 사용자가 명시적으로 다른 등급/공개를 요청할 때만 변경한다.
