# design-ppt Design Tokens (APS Brand Presentation System)

Source: extracted from the **"APS 슬라이드 템플릿" (Turn on the APS ON)** brand
presentation system. These tokens and the `sections/` archetypes are the design
reference; new layouts must conform to them.

## Canvas
- Slide: **1920 × 1080 px** fixed (`<section style="width:1920px;height:1080px;position:relative;...">`)
- Padding: cover/divider/closing (navy) **120px**, all light content slides **110px** (top/bottom), left-right 130px — keep this uniform so the eyebrow/title header and the bottom footer line up across slides
- Content header block (every light slide): eyebrow `font-size:22px; letter-spacing:.26em; color:#0b3fd1`, then `<h2 margin:16px 0 0>`, then optional underline bar `margin-top:22px`; footer row is the last child after a `flex:1` region with `border-top` + `padding-top:22px`
- Every `<section>` is `position:relative` (anchor for the classification badge / watermark)

## Color
| Use | HEX |
|-----|-----|
| Navy (cover/divider/closing background, dark boxes) | `#0b1b3a` |
| APS Blue (accent/emphasis/numbers/links) | `#0b3fd1` |
| Brand Gradient (title underline accent bar, logo) | `#BED600` → `#2BA6CB` |
| Paper (body background) | `#f5f7fa` |
| Body text | `#0b1b3a` |
| Secondary text (slate) | `#5b6b85` |
| Faint secondary / footer text | `#9aa6b8` |
| Border / divider / grid gap | `#e1e7f0` |
| Soft box background (light blue) | `#e7edf8` |
| Light-blue label on navy | `#7fa3ff` |
| Classification (대외비) red | `#c0392b` |
| Alert red | `#e2231a` |

## Typography
- Operating font: **`'Malgun Gothic','맑은 고딕',sans-serif`** — default for all archetype
  bodies/titles (system-guaranteed, matches standard rendering).
- Brand-recommended (cover/large titles): **Noto Serif KR** (titles) · **Pretendard**
  (subtitles/body) — applied when installed system-wide (falls back to Malgun Gothic if not).
- Section labels: uppercase + `letter-spacing:.26em`, APS Blue `#0b3fd1`, 600
  (e.g. `EXECUTIVE SUMMARY`, `03 — PLAN`).
- Title underline accent bar: `width:60px; height:4px; background:linear-gradient(90deg,#BED600,#2BA6CB);`.

## Slide archetypes (snippets)
| File | Use |
|------|-----|
| `sections/00-brand-guide.html` | Brand guide (slogan/core values/colors/typography) |
| `sections/01-cover.html` | Cover (navy + glow + logo + large title) |
| `sections/02-executive-summary.html` | 1-page executive summary (key points + figures + decision request) |
| `sections/03-section-divider.html` | Section divider (big number + section name) |
| `sections/04-body-two-column.html` | Standard body (left narrative / right card list) |
| `sections/05-metrics.html` | Metrics emphasis (large numbers, 4-up) |
| `sections/06-phase-steps.html` | Phases/steps (phase cards, current phase in navy) |
| `sections/07-table.html` | Table (navy header, even-row shading, emphasized total row) |
| `sections/08-components.html` | Component library (chips, numbered circles, process, callout boxes) |
| `sections/09-charts.html` | Chart samples (bar, donut) |
| `sections/10-wbs-gantt.html` | WBS gantt-style schedule |
| `sections/11-org-rnr.html` | Org / R&R |
| `sections/12-asis-tobe.html` | As-Is / To-Be comparison |
| `sections/13-closing.html` | Closing (key message + next steps) |
| `sections/_classification.html` | Classification badge overlay (대외비 etc., first child of each section) |

## Native-mode color snap & role hints
- `--mode native` (editable output) **snaps** extracted colors to the nearest token in the
  palette above, preventing arbitrary colors. `PALETTE` in `scripts/native_render.py` is a
  code copy of this table, so when you change a token, fix both sides.
- The measurer classifies each element heuristically (background/box=box, text leaf=text,
  thin band=rule, table=table, img/svg=raster). If one element has both background+text it
  emits **both** box and text, and a `<section>`'s own background becomes a full-bleed box.
- Elements not reproducible natively (conic-gradient donut, gradient logo, etc.) fall back to
  an image of just that region when marked `data-ppt="raster"` in the HTML. `data-ppt` is
  ignored by image mode, so it is safe. Values: `box|text|rule|table|raster|skip`.

## Meta rules
- Each `<section>` has `data-label` (identifier) and `data-speaker-notes` (presenter notes).
- The build script moves `data-speaker-notes` into PowerPoint speaker notes and sets
  `noProof="1"` (spellcheck off) on the note text.
- When a classification (대외비 etc.) applies, place the `sections/_classification.html` badge
  consistently on **all** slides.
- **Page numbers are injected automatically at build.** The build script stamps a `current /
  total` number at a fixed bottom-right position on every slide, so do not put numbers in the
  archetype footers (they would drift if order changes). **The cover (first slide) is excluded
  by convention** and numbering starts at 1 on the next slide. To exclude a specific slide, set
  `data-page-number="off"` on its `<section>`. On navy-background slides the text color is
  auto-lightened. (Common to both image and native modes.)
