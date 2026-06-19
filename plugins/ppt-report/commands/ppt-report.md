---
description: PPT 보고서를 docs(.docx) 먼저 생성 후 그 문서를 토대로 .pptx 제작 (IT전략팀 규칙 적용)
argument-hint: "[보고서 제목/유형] [초안 파일 경로(선택)]"
---

You are starting the PPT report workflow.

**REQUIRED:** Use the `creating-ppt-reports` skill and follow its workflow exactly:
1. Gather content (report type + reader 임원/실무); never invent unknowns → mark `TBD`.
2. **Stage 1 — build the source `.docx`** with `scripts/build_doc.py`.
3. **Stage 2 — build the `.pptx` from that document** with `scripts/build_ppt.py`.
4. Verify the deck (slide count, headings, tables, `author == "IT전략팀"`).
5. Apply the final PPT proofing step — **검토 > 언어 > 교정 언어 설정 > 사용할 언어 "한국어" > "맞춤법 검사 안 함" 체크 > 확인** — and only then report that the report is finished.

Team rules (enforced): author = **IT전략팀** (never a personal name); filename uses spaces, **no underscores**, with a version suffix like `v1.0` at the end; no spell/grammar red squiggles.

User request / arguments: $ARGUMENTS

Guidance:
- If a report title/type is given, use it; otherwise ask. If a draft file path is given (or a draft clearly exists in the working folder), offer to base the document on it.
- Always do Stage 1 (document) before Stage 2 (slides). The document is the source of truth for the deck.
