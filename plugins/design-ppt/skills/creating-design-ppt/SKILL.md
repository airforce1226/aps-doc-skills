---
name: creating-design-ppt
description: Use when asked to create an APS-Brand-Presentation-style Korean slide report as a .pptx — assemble a deck.html from the APS section archetypes, then build each 1920×1080 <section> as editable PowerPoint native objects (default), or as a pixel-faithful image with --mode image. Applies the IT전략팀 team rules.
---

# Creating Design PPT (HTML deck → editable native .pptx)

## Overview

Produce a Korean PowerPoint report that preserves the **APS Brand Presentation
System** visual identity, by assembling an HTML deck from the 14 APS section
archetypes (`assets/sections/`) and building each `<section>` (1920×1080) into a
slide. **The default output is editable native PowerPoint objects** (textboxes,
tables, shapes) so the recipient can edit text/colors/tables directly; an
optional **image mode** (`--mode image`) bakes pixel-faithful screenshots instead
(text not editable; original text preserved in the speaker notes).

**Core principle:** The HTML/CSS is the source of truth; the `.pptx` is its render.
Never invent facts — mark unknowns `(미정 — 추후 확정)`.

## Output modes (image vs native)

The same `deck.html` builds two ways — pick per deliverable:

| Mode | Command | Traits | When |
|------|---------|--------|------|
| Native (**default** — used by `/design-ppt`) | `python scripts/build_design_ppt.py deck.html "제목 v1.0.pptx" --mode native` | Text/tables/shapes become **editable PowerPoint native objects**; some decorations (charts/logos) fall back to raster | When the recipient must edit it directly in PowerPoint (most cases) |
| Image (opt-in) | `python scripts/build_design_ppt.py deck.html "제목 v1.0.pptx" --mode image` | Pixel-perfect 100% fidelity; **not editable** (slides are images, original text in speaker notes) | Print/distribution final, pixel-perfect |

Native mode **measures** the same `deck.html` via headless Chrome `--dump-dom`,
reads each element's coordinates/styles, and re-lays them out as python-pptx native
objects (`scripts/native_render.py`). At the end of the build it prints a per-slide
`native=/raster=` report showing what is editable and where it fell back. Colors are
snapped to the `assets/design-tokens.md` palette, the font is Malgun Gothic, and the
team rules (author=IT전략팀, classification badge, noProof notes) apply in both modes.

**What falls back to raster:** only decorations that cannot be reproduced natively —
`conic-gradient` donut charts, gradient logos (SVG), etc. — are rasterized to an image
of that region (work-order §5.7 philosophy). Everything else (backgrounds, boxes,
rules, tables, KPI numbers, bar charts) is generated as editable native objects.
When adding a new non-reproducible element, tag its HTML with `data-ppt="raster"` to
rasterize it explicitly (e.g. the donut in `assets/sections/09-charts.html`).

## When to Use

- "Make a PPT in this design", "an APS-brand slide / APS-template report as .pptx"
- The APS-brand **designed** deck is the deliverable, laid out via the HTML archetypes.
  Default output is editable native objects; add `--mode image` only when pixel-perfect
  screenshots are required. (For a structure-first native deck authored without the HTML
  design layer, `creating-ppt-reports` is the alternative.)

## Workflow

1. **Gather content** — report type + reader (executive/working-level); interview or
   read a draft. Never invent unknowns → `(미정 — 추후 확정)`.
   **Security classification = default "대외비" (do not ask):** all APS internal
   reports are 대외비, so do not ask separately — apply the **"대외비" badge to every
   slide by default**. Change it only when the user **explicitly** requests another
   grade (기밀/대내한/내부용) or "공개" (no badge).
   **Cover/closing slides = always ask (when multi-slide):** if the deck is **2+
   slides**, **before** assembling deck.html, ask the user whether to include the cover
   slide (`01-cover.html`) and the closing slide (`13-closing.html`), and apply per their
   answer — never add or omit them arbitrarily. For a single-slide deck, build only the
   body and do not ask.
2. **Assemble `deck.html`** — copy snippets from `assets/sections/` into one file,
   in order. Replace every `[[...]]` placeholder with real content and write each
   `data-speaker-notes`. Follow `assets/design-tokens.md`; the shared `assets/base.css`
   is injected automatically at build time. New layouts must reuse the tokens
   (Navy `#0b1b3a` / Paper `#f5f7fa` / APS Blue `#0b3fd1` / Brand Gradient
   `#BED600→#2BA6CB` / Malgun Gothic). The 14 archetypes are in `assets/sections/`.
   Place the `assets/sections/_classification.html` badge ("대외비") as the **first
   child** of **every** `<section>` by default (the parent section needs
   `position:relative`). Replace the text only when another grade is explicitly given.
3. **Build (default: editable native)** —
   `python scripts/build_design_ppt.py deck.html "<제목> v1.0.pptx" --mode native`
   Check the closing `native=/raster=` report for editability and fallback locations.
   If a pixel-perfect image deck is needed, build with `--mode image` (or produce both).
4. **Verify** — reopen with python-pptx; confirm slide count and
   `core_properties.author == "IT전략팀"`. In native mode, confirm text/tables/shapes are
   **real objects, not pictures** (e.g. `shape.has_text_frame` / `has_table`). By default,
   confirm **every** slide has the "대외비" badge (except when the user explicitly chose
   공개). (All text runs in slide bodies/tables + notes are auto-set to `noProof` at build.)
5. **Report done** — give the file path and slide count.

## Team Rules (enforced — do not deviate)

| Rule | How it is applied |
|------|-------------------|
| Author = **IT전략팀** (never a person's name) | `build_design_ppt.py` hard-codes `core_properties.author`. |
| Filename: spaces, **no underscores**, version suffix last | e.g. `착수 보고서 v1.0.pptx`. Pass this exact path to the script. |
| Don't invent unknowns | Use `(미정 — 추후 확정)` in the slide text. |
| **개조식 작성 (no conversational tone)** | Slide bodies use **noun-form endings** (개조식). No 합쇼체 endings such as `~있었습니다/~합니다/~됩니다`; end with noun forms like `~잔존/~점유/~정상화`. |
| **Classification = default 대외비 (not asked)** | Place the `_classification.html` ("대외비") badge on **all slides by default**. Change only when the user specifies another grade/공개. |
| Classification badge = red **border only, centered** | The `_classification.html` badge has no fill — red border (2px), red text, text centered in the box. In native builds, `native_render.py` generates a bordered box + centered text (no fill). |
| Disable spellcheck (noProof) | `native_render.py` (**all** text runs in slide bodies/tables) and `build_design_ppt.py` (speaker notes) set `lang="ko-KR"` + `noProof="1"` (검토>언어>교정 언어 설정>맞춤법 검사 안 함). |
| Cover/closing = ask via Q&A when multi-slide | If the deck is **2+ slides**, before assembling deck.html ask the user whether to include the cover (`01-cover.html`) / closing (`13-closing.html`) slides, and apply per their answer. No arbitrary add/omit (single-slide decks are not asked). |

## Prerequisites

- **Pretendard font installed system-wide** for full design fidelity. Without it,
  rendering falls back to Malgun Gothic (still works; spacing/weight may differ).
- Chrome or Edge installed (auto-detected; override via `DESIGN_PPT_BROWSER`).
- `python -m pip install python-pptx` if missing.

## Quick Reference

| Need | Do |
|------|----|
| Slide archetypes | `assets/sections/*.html` |
| Preview all templates (open in browser) | `assets/templates-gallery.html` |
| Regenerate gallery (after editing archetypes) | `python scripts/build_gallery.py` |
| Design tokens | `assets/design-tokens.md` |
| Build deck (native, **default, editable**) | `python scripts/build_design_ppt.py deck.html "제목 v1.0.pptx" --mode native` |
| Build deck (image, pixel-perfect, not editable) | `python scripts/build_design_ppt.py deck.html "제목 v1.0.pptx" --mode image` |
| Rasterize a non-reproducible element | add `data-ppt="raster"` to that HTML element |
| Mark unknown | put `(미정 — 추후 확정)` in the text |
| Classification badge | put `assets/sections/_classification.html` first in each section |
| Browser not found | set `DESIGN_PPT_BROWSER` to chrome.exe/msedge.exe |

## Common Mistakes

- Leaving `[[...]]` placeholders in the final deck.
- Underscores in the filename or a missing `v1.0` suffix.
- A person's name as author instead of IT전략팀.
- Putting the classification badge (대외비/기밀 etc.) on only some slides (place consistently on all).
- In a multi-slide deck, adding or omitting the cover (`01-cover.html`) / closing
  (`13-closing.html`) slides without asking the user (multi-slide always uses Q&A).
- Building with `--mode image` and then complaining text can't be edited in PowerPoint —
  image mode is intentionally non-editable (use the default native mode, or edit the HTML and rebuild).
- Mistaking a chart/logo in native output for a bug because it's an image — that is the
  intended raster fallback (the editable targets are text/tables/shapes/backgrounds). Bar
  charts are native, only the donut is raster.
