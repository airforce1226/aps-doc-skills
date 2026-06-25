---
name: creating-design-ppt
description: Use when asked to create an APS-Brand-Presentation-style Korean slide report as a .pptx — assemble a deck.html from the APS section archetypes, then render each 1920×1080 <section> to a pixel-faithful slide. Applies the IT전략팀 team rules.
---

# Creating Design PPT (HTML deck → pixel-faithful .pptx)

## Overview

Produce a Korean PowerPoint report that preserves the **APS Brand Presentation
System** visual identity exactly, by assembling an HTML deck from the 14 APS
section archetypes (`assets/sections/`) and rendering each `<section>` (1920×1080)
to a full-bleed slide image. Slide text is not editable in PowerPoint (it is an
image); the original text is preserved in the **speaker notes**.

**Core principle:** The HTML/CSS is the source of truth; the `.pptx` is its render.
Never invent facts — mark unknowns `(미정 — 추후 확정)`.

## When to Use

- "이 디자인으로 PPT 만들어줘", "APS 브랜드 슬라이드 / APS 템플릿 보고서를 .pptx로"
- A pixel-faithful designed deck is the deliverable (NOT an editable native deck →
  use `creating-ppt-reports` for that).

## Workflow

1. **Gather content** — report type + reader (임원/실무); interview or read a draft.
   Never invent unknowns → `(미정 — 추후 확정)`.
   **보안 분류 확인 (필수):** 문서가 대외비/기밀/대내한/내부용에 해당하는지 반드시 확인·결정한다.
   해당하면 모든 슬라이드에 분류 배지를 넣고, 비해당이면 "공개 — 배지 생략"으로 명시한다. 그냥 넘어가지 않는다.
2. **Assemble `deck.html`** — copy snippets from `assets/sections/` into one file,
   in order. Replace every `[[...]]` placeholder with real content and write each
   `data-speaker-notes`. Follow `assets/design-tokens.md`; the shared `assets/base.css`
   is injected automatically at build time. New layouts must reuse the tokens
   (Navy `#0b1b3a` / Paper `#f5f7fa` / APS Blue `#0b3fd1` / Brand Gradient
   `#BED600→#2BA6CB` / Malgun Gothic). 14개 아키타입이 `assets/sections/`에 있다.
   보안 분류가 있으면 `assets/sections/_classification.html` 배지를 **모든** `<section>`의
   맨 앞 자식으로 붙이고 등급 텍스트만 교체한다(부모 section 에 `position:relative` 필요).
3. **Build** — `python scripts/build_design_ppt.py deck.html "<제목> v1.0.pptx"`
4. **Verify** — reopen with python-pptx; confirm slide count, notes present, and
   `core_properties.author == "IT전략팀"`. 보안 분류 대상이면 **모든** 슬라이드에 배지가
   있는지 확인한다(노트 텍스트는 빌드 시 맞춤법 검사 비활성 `noProof` 자동 처리됨).
5. **Report done** — give the file path and slide count.

## Team Rules (enforced — do not deviate)

| Rule | How it is applied |
|------|-------------------|
| Author = **IT전략팀** (never a person's name) | `build_design_ppt.py` hard-codes `core_properties.author`. |
| Filename: spaces, **no underscores**, version suffix last | e.g. `착수 보고서 v1.0.pptx`. Pass this exact path to the script. |
| Don't invent unknowns | Use `(미정 — 추후 확정)` in the slide text. |
| **보안 분류 누락 금지** | 대외비/기밀 등 해당 시 `_classification.html` 배지를 전 슬라이드에 배치. 비해당이면 의도적 생략을 명시. |
| 맞춤법 검사 비활성(noProof) | `build_design_ppt.py` 가 발표자 노트 모든 run 에 `noProof="1"` 설정 (검토>언어>맞춤법 검사 안 함). |

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
| 보안 분류 배지 | `assets/sections/_classification.html` 를 각 section 맨 앞에 |
| Browser not found | set `DESIGN_PPT_BROWSER` to chrome.exe/msedge.exe |

## Common Mistakes

- Leaving `[[...]]` placeholders in the final deck.
- Underscores in the filename or a missing `v1.0` suffix.
- A person's name as author instead of IT전략팀.
- 대외비/기밀 등 보안 분류 배지를 일부 슬라이드에서만 넣고 누락(전 슬라이드 일관 배치).
- Expecting editable text in PowerPoint — slides are images by design; edit the
  HTML and rebuild instead.
