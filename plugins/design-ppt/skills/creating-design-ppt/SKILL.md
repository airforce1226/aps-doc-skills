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

## Output modes (이미지 vs 네이티브)

The same `deck.html` builds two ways — pick per deliverable:

| 모드 | 명령 | 특성 | 언제 |
|------|------|------|------|
| 네이티브 (**기본** — `/design-ppt`가 사용) | `python scripts/build_design_ppt.py deck.html "제목 v1.0.pptx" --mode native` | 텍스트·표·도형이 **편집 가능한 PowerPoint 네이티브 개체** · 차트/로고 등 일부 장식은 래스터 폴백 | 받는 사람이 PowerPoint에서 직접 고쳐야 할 때 (대부분) |
| 이미지 (opt-in) | `python scripts/build_design_ppt.py deck.html "제목 v1.0.pptx" --mode image` | 픽셀 100% 충실 · **편집 불가**(슬라이드는 이미지, 원문은 발표자 노트) | 인쇄·배포용 픽셀 완벽 최종본 |

네이티브 모드는 같은 `deck.html`을 헤드리스 Chrome `--dump-dom`으로 **측정**해 각 요소의
좌표·스타일을 읽고 python-pptx 네이티브 개체로 재배치한다(`scripts/native_render.py`).
빌드 끝에 슬라이드별 `native=/raster=` 리포트를 출력해 편집 가능성과 폴백 위치를 보여준다.
색은 `assets/design-tokens.md` 팔레트로 스냅되고, 폰트는 맑은 고딕, 팀 규칙(author=IT전략팀·
대외비 배지·noProof 노트)은 두 모드 공통이다.

**언제 무엇이 래스터로 떨어지나:** `conic-gradient` 도넛 차트, 그라데이션 로고(SVG) 등
네이티브로 재현 불가한 장식만 그 영역의 이미지로 폴백한다(작업지시서 §5.7 철학). 그 외
배경·박스·룰·표·KPI 숫자·막대 차트는 모두 편집 가능한 네이티브 개체로 생성된다.
재현 불가 요소를 새로 추가할 때는 해당 HTML에 `data-ppt="raster"`를 붙이면 명시적으로
래스터 처리된다(예: `assets/sections/09-charts.html`의 도넛).

## When to Use

- "이 디자인으로 PPT 만들어줘", "APS 브랜드 슬라이드 / APS 템플릿 보고서를 .pptx로"
- The APS-brand **designed** deck is the deliverable, laid out via the HTML archetypes.
  Default output is editable native objects; add `--mode image` only when pixel-perfect
  screenshots are required. (For a structure-first native deck authored without the HTML
  design layer, `creating-ppt-reports` is the alternative.)

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
3. **Build (기본: 편집 가능 네이티브)** —
   `python scripts/build_design_ppt.py deck.html "<제목> v1.0.pptx" --mode native`
   빌드 끝의 `native=/raster=` 리포트로 편집 가능성·폴백 위치를 확인한다.
   픽셀 완벽 이미지본이 필요하면 `--mode image`로 빌드한다(또는 두 개 다 산출).
4. **Verify** — reopen with python-pptx; confirm slide count and
   `core_properties.author == "IT전략팀"`. 네이티브 모드면 텍스트/표/도형이 **picture가 아닌
   실제 개체**인지(예: `shape.has_text_frame` / `has_table`) 확인한다. 보안 분류 대상이면
   **모든** 슬라이드에 배지가 있는지 확인한다(슬라이드 본문·표의 모든 텍스트 run + 노트가 빌드 시 `noProof` 자동 처리됨).
5. **Report done** — give the file path and slide count.

## Team Rules (enforced — do not deviate)

| Rule | How it is applied |
|------|-------------------|
| Author = **IT전략팀** (never a person's name) | `build_design_ppt.py` hard-codes `core_properties.author`. |
| Filename: spaces, **no underscores**, version suffix last | e.g. `착수 보고서 v1.0.pptx`. Pass this exact path to the script. |
| Don't invent unknowns | Use `(미정 — 추후 확정)` in the slide text. |
| **개조식 작성 (구어체 금지)** | 슬라이드 본문은 **명사형 종결**의 개조식으로 쓴다. `~있었습니다/~합니다/~됩니다` 같은 구어체(합쇼체) 종결어미 금지 → `~잔존/~점유/~정상화` 처럼 명사형으로 맺는다. |
| **보안 분류 누락 금지** | 대외비/기밀 등 해당 시 `_classification.html` 배지를 전 슬라이드에 배치. 비해당이면 의도적 생략을 명시. |
| 보안 배지 = 빨강 **테두리만·정중앙** | `_classification.html` 배지는 채움 없이 적색 테두리(2px)·적색 글자, 텍스트는 박스 정중앙. 네이티브 빌드 시 `native_render.py`가 테두리 박스 + 중앙 정렬 텍스트로 생성한다(채움 X). |
| 맞춤법 검사 비활성(noProof) | `native_render.py`(슬라이드 본문·표 **모든** 텍스트 run)와 `build_design_ppt.py`(발표자 노트)가 `lang="ko-KR"` + `noProof="1"` 설정 (검토>언어>교정 언어 설정>맞춤법 검사 안 함). |

## Prerequisites

- **Pretendard font installed system-wide** for full design fidelity. Without it,
  rendering falls back to Malgun Gothic (still works; spacing/weight may differ).
- Chrome or Edge installed (auto-detected; override via `DESIGN_PPT_BROWSER`).
- `python -m pip install python-pptx` if missing.

## Quick Reference

| Need | Do |
|------|----|
| Slide archetypes | `assets/sections/*.html` |
| 전체 템플릿 미리보기 (브라우저로 열기) | `assets/templates-gallery.html` |
| 갤러리 재생성 (아키타입 수정 후) | `python scripts/build_gallery.py` |
| Design tokens | `assets/design-tokens.md` |
| Build deck (native, **기본·편집 가능**) | `python scripts/build_design_ppt.py deck.html "제목 v1.0.pptx" --mode native` |
| Build deck (image, 픽셀 완벽·편집 불가) | `python scripts/build_design_ppt.py deck.html "제목 v1.0.pptx" --mode image` |
| 재현 불가 요소를 래스터로 | 해당 HTML 요소에 `data-ppt="raster"` 부여 |
| Mark unknown | put `(미정 — 추후 확정)` in the text |
| 보안 분류 배지 | `assets/sections/_classification.html` 를 각 section 맨 앞에 |
| Browser not found | set `DESIGN_PPT_BROWSER` to chrome.exe/msedge.exe |

## Common Mistakes

- Leaving `[[...]]` placeholders in the final deck.
- Underscores in the filename or a missing `v1.0` suffix.
- A person's name as author instead of IT전략팀.
- 대외비/기밀 등 보안 분류 배지를 일부 슬라이드에서만 넣고 누락(전 슬라이드 일관 배치).
- `--mode image`로 빌드해 놓고 PowerPoint에서 텍스트가 안 고쳐진다고 하기 — 이미지 모드는
  의도적으로 편집 불가다(기본 네이티브 모드를 쓰거나 HTML을 고쳐 재빌드).
- 네이티브 모드 결과에서 차트/로고가 이미지인 것을 버그로 오인 — 의도된 래스터 폴백이다
  (편집 대상은 텍스트·표·도형·배경). 막대 차트는 네이티브, 도넛만 래스터.
