---
name: drafting-project-documents
description: Use when asked to write or draft a Korean project document (착수 보고서/kickoff report, 요구사항 정의서/requirements, 회의록/meeting notes) as a .docx, whether starting from an existing draft file or from scratch through conversation.
---

# Drafting Project Documents

## Overview

Produce a clean, styled Korean `.docx` project document. Content comes from **an existing draft** (read its outline, fill the blanks) **or from conversation** (interview the user). The document is rendered by a single spec-driven engine, so every document shares consistent styling.

**Core principle:** Gather facts → never invent the unknown (mark `TBD`) → match the reader's vocabulary → render from a JSON spec.

## When to Use

- "착수 보고서 / 요구사항 정의서 / 회의록 작성해줘", "이 초안 채워줘", "전환 착수 보고서 초안 만들어줘"
- A `.docx` draft exists and needs filling, OR there is no draft and content must be elicited.
- NOT for free-form prose, slides, or non-document answers.

## Workflow

1. **Identify document type and reader.** Map to a standard outline in `references/document-types.md`. Ask who reads it (임원 vs 실무) — this drives vocabulary. See the tone rules in that file.
2. **Gather content — two paths:**
   - *Draft exists:* read its outline (if it is open in Word and locked, read via Word COM `GetActiveObject("Word.Application")`, not by unzipping the locked file). Keep the section structure; fill each section.
   - *No draft:* interview the user section by section, following the standard outline.
3. **Never invent.** Unknown values become the literal string `"TBD"` in the spec — the engine renders them as grey "(미정 — 추후 확정)". Convert relative dates to absolute. Confirm anything ambiguous instead of guessing.
4. **Match the reader.** For 임원 audiences, strip technical jargon per `references/document-types.md` (effect-first wording; keep concrete tech names in one reference table only).
5. **Build the spec.** Write a UTF-8 JSON spec (schema documented at the top of `scripts/build_doc.py`).
6. **Render.** `python build_doc.py spec.json "<output>.docx"`. Install the dependency first if missing: `python -m pip install python-docx`.
7. **Verify.** Re-open the docx and read it back (`python-docx`) to confirm sections/tables. If the target filename is locked (open in Word), save to a new name rather than force-closing the user's document.

## Quick Reference

| Need | Do |
|------|----|
| Doc outlines + tone rules | `references/document-types.md` |
| Spec JSON schema | header docstring of `scripts/build_doc.py` |
| Mark a value unknown | put `"TBD"` in the cell/string |
| Render | `python build_doc.py spec.json out.docx` |
| Read a locked open draft | Word COM `GetActiveObject`, not zip extraction |

## Common Mistakes

- **Inventing missing data** (dates, staffing, infra) instead of `TBD`. Don't.
- **Force-closing/overwriting the user's open Word file.** Save to a new name; ask before replacing.
- **Leaving developer notes in an executive report.** Move them out for 임원 readers.
- **Jargon for executives.** Translate to plain effect-based language.
- **Skipping the read-back verification** after rendering.
