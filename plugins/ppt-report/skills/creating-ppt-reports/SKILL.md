---
name: creating-ppt-reports
description: Use when asked to create a Korean PPT report / PPT 보고서 (착수 보고서, 진행/결과 보고 등) as a .pptx — runs the docs-first workflow (generate a source .docx, then build the slide deck from it) and applies the IT전략팀 team rules.
---

# Creating PPT Reports (docs → ppt)

## Overview

Produce a Korean PowerPoint report in **two stages**: first write the source
document as a styled `.docx`, then build the slide deck **from that document's
content**. Both stages are spec-driven (UTF-8 JSON → bundled renderers) so
styling stays consistent and the team rules are enforced in code.

**Core principle:** Document first, slides second. The `.docx` is the source of
truth; the `.pptx` is derived from it — never invent new facts at the slide
stage.

## When to Use

- "PPT 보고서 만들어줘", "착수 보고서 발표자료 만들어줘", ".pptx 로 보고서 작성"
- A slide deck report is the deliverable (not just a document → use
  `drafting-project-documents` for a doc-only request).
- NOT for free-form prose or a `.docx`-only deliverable.

## Workflow

1. **Gather content** (same discipline as document drafting): identify report
   type + reader (임원 vs 실무), interview the user or read an existing draft.
   Never invent unknowns — mark them `"TBD"` (renders as grey "(미정 — 추후 확정)").
2. **Stage 1 — build the source `.docx`.** Write a doc spec JSON (schema in the
   header of `scripts/build_doc.py`) and render:
   `python scripts/build_doc.py docspec.json "<제목> v1.0.docx"`
3. **Stage 2 — build the `.pptx` from the doc.** Derive a slide spec from the
   document's sections (one section → one slide; condense prose into bullets),
   schema in the header of `scripts/build_ppt.py`, then render:
   `python scripts/build_ppt.py pptspec.json "<제목> v1.0.pptx"`
4. **Verify.** Re-open the `.pptx` with `python-pptx` and read back slide count,
   headings, and tables. Confirm `core_properties.author == "IT전략팀"`.
5. **Final proofing step (PPT, required) — then report done.** Apply the in-app
   process and only then tell the user it is finished:
   **검토 > 언어 > 교정 언어 설정 > 사용할 언어 "한국어" 선택 > "맞춤법 검사 안 함" 체크 > 확인.**
   (The renderer already sets `noProof`/`lang=ko-KR` on every run, but still
   state that this confirm step is the closing action.)

## Team Rules (enforced — do not deviate)

| Rule | How it is applied |
|------|-------------------|
| Author = **IT전략팀** (never a person's name) | `meta` line + `core_properties.author`; the renderer hard-codes the author. |
| Filename: spaces, **no underscores**, version suffix last | e.g. `착수 보고서 v1.0.pptx` (not `착수_보고서_v1.0`). Pass this exact path to the script. |
| No spell/grammar red squiggles | Renderer sets `noProof` + `lang=ko-KR` on every run; plus the manual 교정 언어 설정 step above. |
| Document first, then slides | Always run Stage 1 before Stage 2. |

## Quick Reference

| Need | Do |
|------|----|
| Doc spec schema | header docstring of `scripts/build_doc.py` |
| PPT spec schema | header docstring of `scripts/build_ppt.py` |
| Render doc | `python scripts/build_doc.py spec.json "제목 v1.0.docx"` |
| Render ppt | `python scripts/build_ppt.py spec.json "제목 v1.0.pptx"` |
| Mark unknown | put `"TBD"` in the cell/string |
| Missing dep | `python -m pip install python-pptx python-docx` |

## Common Mistakes

- **Skipping Stage 1** and going straight to slides. Build the `.docx` first.
- **Underscores in the filename** or a missing `v1.0` suffix.
- **Putting a person's name as author** instead of IT전략팀.
- **Inventing missing data** (dates, budget, staffing) instead of `TBD`.
- **Reporting "done" before the 교정 언어 설정 confirm step.** That step is the
  closing action; do it, then report.
- **Overwriting a file open in PowerPoint** (locked) — save to a new name and
  ask before replacing.
