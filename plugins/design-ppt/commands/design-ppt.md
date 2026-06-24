---
description: Claude Design 스타일 HTML 덱을 픽셀 그대로 .pptx로 빌드 (IT전략팀 규칙 적용)
argument-hint: "[보고서 제목/유형] [기존 deck.html 경로(선택)]"
---

You are starting the design-ppt workflow.

**REQUIRED:** Use the `creating-design-ppt` skill and follow its workflow exactly:
1. Gather content (report type + reader 임원/실무); never invent unknowns → mark `(미정 — 추후 확정)`.
2. Assemble a `deck.html` from `assets/sections/` snippets; fill each section's content and `data-speaker-notes`, following `assets/design-tokens.md` and `assets/base.css`.
3. Build the deck: `python scripts/build_design_ppt.py deck.html "<제목> v1.0.pptx"`.
4. Verify the deck (slide count, notes present, `author == "IT전략팀"`).

Team rules (enforced): author = **IT전략팀** (never a personal name); filename uses spaces, **no underscores**, with a version suffix like `v1.0` at the end.

User request / arguments: $ARGUMENTS

Guidance:
- If a report title/type is given, use it; otherwise ask. If a `deck.html` path is given, base the build on it.
