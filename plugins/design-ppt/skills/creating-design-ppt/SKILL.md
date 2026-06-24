---
name: creating-design-ppt
description: Use when asked to create a Claude-Design-style Korean slide report as a .pptx — assemble a deck.html from section snippets, then render each 1920×1080 <section> to a pixel-faithful slide. Applies the IT전략팀 team rules.
---

# Creating Design PPT (HTML deck → pixel-faithful .pptx)

## Overview

Produce a Korean PowerPoint report that preserves a Claude-Design visual identity
**exactly**, by assembling an HTML deck and rendering each `<section>` (1920×1080)
to a full-bleed slide image. Slide text is not editable in PowerPoint (it is an
image); the original text is preserved in the **speaker notes**.

**Core principle:** The HTML/CSS is the source of truth; the `.pptx` is its render.
Never invent facts — mark unknowns `(미정 — 추후 확정)`.

## When to Use

- "이 디자인으로 PPT 만들어줘", "Claude Design 스타일 보고서를 .pptx로"
- A pixel-faithful designed deck is the deliverable (NOT an editable native deck →
  use `creating-ppt-reports` for that).

## Workflow

1. **Gather content** — report type + reader (임원/실무); interview or read a draft.
   Never invent unknowns → `(미정 — 추후 확정)`.
2. **Assemble `deck.html`** — copy snippets from `assets/sections/` into one file,
   in order. Replace every `[[...]]` placeholder with real content and write each
   `data-speaker-notes`. Follow `assets/design-tokens.md`; the shared `assets/base.css`
   is injected automatically at build time. New layouts must reuse the tokens
   (네이비 `#0b1b3a` / 라이트 `#f5f7fa` / 액센트 `#0b3fd1` / Pretendard).
3. **Build** — `python scripts/build_design_ppt.py deck.html "<제목> v1.0.pptx"`
4. **Verify** — reopen with python-pptx; confirm slide count, notes present, and
   `core_properties.author == "IT전략팀"`.
5. **Report done** — give the file path and slide count.

## Team Rules (enforced — do not deviate)

| Rule | How it is applied |
|------|-------------------|
| Author = **IT전략팀** (never a person's name) | `build_design_ppt.py` hard-codes `core_properties.author`. |
| Filename: spaces, **no underscores**, version suffix last | e.g. `착수 보고서 v1.0.pptx`. Pass this exact path to the script. |
| Don't invent unknowns | Use `(미정 — 추후 확정)` in the slide text. |

## Prerequisites

- **Pretendard font installed system-wide** for full design fidelity. Without it,
  rendering falls back to Malgun Gothic (still works; spacing/weight may differ).
- Chrome or Edge installed (auto-detected; override via `DESIGN_PPT_BROWSER`).
- `python -m pip install python-pptx` if missing.

## Quick Reference

| Need | Do |
|------|----|
| Slide archetypes | `assets/sections/*.html` |
| Design tokens | `assets/design-tokens.md` |
| Build deck | `python scripts/build_design_ppt.py deck.html "제목 v1.0.pptx"` |
| Mark unknown | put `(미정 — 추후 확정)` in the text |
| Browser not found | set `DESIGN_PPT_BROWSER` to chrome.exe/msedge.exe |

## Common Mistakes

- Leaving `[[...]]` placeholders in the final deck.
- Underscores in the filename or a missing `v1.0` suffix.
- A person's name as author instead of IT전략팀.
- Expecting editable text in PowerPoint — slides are images by design; edit the
  HTML and rebuild instead.
