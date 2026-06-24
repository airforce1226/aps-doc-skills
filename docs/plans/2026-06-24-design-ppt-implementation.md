# design-ppt Plugin Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a new `design-ppt` plugin to the `aps-doc-skills` marketplace that turns a Claude-Design-style `deck.html` (each `<section>` = one 1920×1080 slide) into a pixel-faithful `.pptx` via headless-Chrome rendering + python-pptx full-bleed assembly.

**Architecture:** A deterministic Python build script renders each `<section>` to a 1920×1080 PNG with headless Chrome/Edge and places each PNG full-bleed onto a 16:9 slide, copying `data-speaker-notes` into the notes pane and forcing author = IT전략팀. Claude assembles `deck.html` from a library of section snippets (no rigid JSON). Design tokens (navy/light/accent/Pretendard) are documented and applied via shared CSS.

**Tech Stack:** Python 3.12, python-pptx 1.0.2, pytest 8.2.0 (all installed); headless Chrome (installed, Edge fallback); Pretendard font (system-installed).

**Repo:** `C:\Users\lee.jh3\.claude\plugins\marketplaces\aps-doc-skills` (branch `main`). All paths below are relative to this repo root unless absolute.

---

## File Structure

```
.claude-plugin/marketplace.json                 # MODIFY: register design-ppt
plugins/design-ppt/
  .claude-plugin/plugin.json                     # CREATE
  commands/design-ppt.md                         # CREATE: /design-ppt command
  skills/creating-design-ppt/
    SKILL.md                                      # CREATE: workflow + team rules
    assets/
      design-tokens.md                            # CREATE: extracted design system
      base.css                                    # CREATE: shared section styles
      sections/
        00-cover.html                             # CREATE: cover archetype
        01-section-body.html                      # CREATE: heading + bullets archetype
        02-closing.html                           # CREATE: closing archetype
    scripts/
      build_design_ppt.py                         # CREATE: render + assemble (core)
      tests/test_build_design_ppt.py              # CREATE: unit + integration tests
```

Responsibility split: `build_design_ppt.py` is the only logic-bearing file (split → wrap → capture → assemble), kept small and pure where possible so the unit tests touch no browser. Assets are static design data. SKILL.md/command/plugin.json are configuration.

---

### Task 1: Plugin scaffold + marketplace registration

**Files:**
- Create: `plugins/design-ppt/.claude-plugin/plugin.json`
- Create: `plugins/design-ppt/commands/design-ppt.md`
- Modify: `.claude-plugin/marketplace.json`

- [ ] **Step 1: Create `plugins/design-ppt/.claude-plugin/plugin.json`**

```json
{
  "name": "design-ppt",
  "description": "Render a Claude-Design-style HTML deck into a pixel-faithful Korean .pptx (each <section> = one 1920×1080 slide). Headless-Chrome render + full-bleed assembly. Enforces IT전략팀 team rules.",
  "version": "0.1.0",
  "author": {
    "name": "airforce1226",
    "email": "dhkim3@apsystems.co.kr"
  },
  "homepage": "https://github.com/airforce1226/aps-doc-skills",
  "repository": "https://github.com/airforce1226/aps-doc-skills",
  "license": "MIT",
  "keywords": ["pptx", "powerpoint", "design", "html", "korean", "presentation", "PPT보고서", "claude-design"]
}
```

- [ ] **Step 2: Create `plugins/design-ppt/commands/design-ppt.md`**

```markdown
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
```

- [ ] **Step 3: Register the plugin in `.claude-plugin/marketplace.json`**

Add this object to the end of the `"plugins"` array (after the `ppt-report` entry; add a comma after the `ppt-report` closing brace):

```json
    {
      "name": "design-ppt",
      "description": "Render a Claude-Design-style HTML deck into a pixel-faithful Korean .pptx — each <section> becomes one 1920×1080 slide via headless-Chrome render + full-bleed assembly. Enforces IT전략팀 team rules.",
      "source": "./plugins/design-ppt",
      "category": "productivity"
    }
```

- [ ] **Step 4: Validate JSON**

Run: `python -c "import json; json.load(open('.claude-plugin/marketplace.json', encoding='utf-8')); json.load(open('plugins/design-ppt/.claude-plugin/plugin.json', encoding='utf-8')); print('JSON OK')"`
Expected: `JSON OK`

- [ ] **Step 5: Commit**

```bash
git add .claude-plugin/marketplace.json plugins/design-ppt/.claude-plugin/plugin.json plugins/design-ppt/commands/design-ppt.md
git commit -m "feat: scaffold design-ppt plugin (plugin.json, command, marketplace entry)"
```

---

### Task 2: `split_sections` — parse deck.html into slides (TDD)

**Files:**
- Create: `plugins/design-ppt/skills/creating-design-ppt/scripts/build_design_ppt.py`
- Test: `plugins/design-ppt/skills/creating-design-ppt/scripts/tests/test_build_design_ppt.py`

- [ ] **Step 1: Write the failing test**

Create `scripts/tests/test_build_design_ppt.py`:

```python
# -*- coding: utf-8 -*-
import base64
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import build_design_ppt as b

SAMPLE = (
    '<!DOCTYPE html><html><head></head><body>\n'
    '<section data-label="표지" data-speaker-notes="표지 노트" '
    'style="width:1920px;height:1080px;background:#0b1b3a;">COVER</section>\n'
    '<section data-label="본문" data-speaker-notes="본문 노트" '
    'style="width:1920px;height:1080px;background:#f5f7fa;">BODY</section>\n'
    '</body></html>'
)


def test_split_sections_count_and_fields():
    secs = b.split_sections(SAMPLE)
    assert len(secs) == 2
    assert secs[0]["label"] == "표지"
    assert secs[0]["notes"] == "표지 노트"
    assert "COVER" in secs[0]["html"]
    assert secs[0]["html"].lstrip().startswith("<section")
    assert secs[1]["label"] == "본문"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd plugins/design-ppt/skills/creating-design-ppt/scripts && python -m pytest tests/test_build_design_ppt.py::test_split_sections_count_and_fields -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'build_design_ppt'`

- [ ] **Step 3: Create `build_design_ppt.py` with the module header + `split_sections`**

```python
# -*- coding: utf-8 -*-
"""Build a pixel-faithful .pptx from a Claude-Design-style deck.html.

Each <section> (1920x1080) is rendered to a PNG with headless Chrome/Edge,
then placed full-bleed onto a 16:9 slide. data-speaker-notes -> notes pane.
Author metadata is forced to "IT전략팀".

Usage:
    python build_design_ppt.py deck.html "제목 v1.0.pptx"

Assumptions:
  * <section> elements are siblings (not nested) — true for Claude Design decks.
  * data-label / data-speaker-notes values contain no double-quote character.
"""
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.util import Cm

AUTHOR = "IT전략팀"
W, H = 1920, 1080
SLIDE_W = Cm(33.867)   # 13.333 in  (16:9)
SLIDE_H = Cm(19.05)    # 7.5 in
ASSETS = Path(__file__).resolve().parent.parent / "assets"

SECTION_RE = re.compile(
    r"<section\b(?P<attrs>[^>]*)>(?P<body>.*?)</section>",
    re.DOTALL | re.IGNORECASE,
)


def _attr(attrs, name):
    m = re.search(name + r'\s*=\s*"([^"]*)"', attrs)
    return m.group(1) if m else ""


def split_sections(html):
    out = []
    for m in SECTION_RE.finditer(html):
        attrs = m.group("attrs")
        out.append({
            "label": _attr(attrs, "data-label"),
            "notes": _attr(attrs, "data-speaker-notes"),
            "html": m.group(0),
        })
    return out
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_build_design_ppt.py::test_split_sections_count_and_fields -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add plugins/design-ppt/skills/creating-design-ppt/scripts/
git commit -m "feat: add split_sections for design-ppt deck parsing"
```

---

### Task 3: `wrap_section_page` — wrap one section into a standalone page (TDD)

**Files:**
- Modify: `scripts/build_design_ppt.py`
- Test: `scripts/tests/test_build_design_ppt.py`

- [ ] **Step 1: Add the failing test** (append to test file)

```python
def test_wrap_section_page_is_full_html():
    page = b.wrap_section_page("<section>X</section>", "body{margin:0}")
    assert page.startswith("<!DOCTYPE html>")
    assert "body{margin:0}" in page
    assert "<section>X</section>" in page
    assert 'charset="utf-8"' in page
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_build_design_ppt.py::test_wrap_section_page_is_full_html -v`
Expected: FAIL — `AttributeError: module 'build_design_ppt' has no attribute 'wrap_section_page'`

- [ ] **Step 3: Add `wrap_section_page` to `build_design_ppt.py`** (after `split_sections`)

```python
def wrap_section_page(section_html, css):
    return (
        '<!DOCTYPE html><html><head><meta charset="utf-8">'
        "<style>html,body{margin:0;padding:0;}" + css + "</style></head>"
        "<body>" + section_html + "</body></html>"
    )
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_build_design_ppt.py::test_wrap_section_page_is_full_html -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add plugins/design-ppt/skills/creating-design-ppt/scripts/
git commit -m "feat: add wrap_section_page for per-slide rendering"
```

---

### Task 4: `find_browser` — locate Chrome/Edge with env override (TDD)

**Files:**
- Modify: `scripts/build_design_ppt.py`
- Test: `scripts/tests/test_build_design_ppt.py`

- [ ] **Step 1: Add the failing test**

```python
def test_find_browser_env_override(tmp_path):
    fake = tmp_path / "chrome.exe"
    fake.write_text("x")
    os.environ["DESIGN_PPT_BROWSER"] = str(fake)
    try:
        assert b.find_browser() == str(fake)
    finally:
        del os.environ["DESIGN_PPT_BROWSER"]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_build_design_ppt.py::test_find_browser_env_override -v`
Expected: FAIL — `AttributeError: ... has no attribute 'find_browser'`

- [ ] **Step 3: Add `find_browser` to `build_design_ppt.py`**

```python
BROWSER_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
]


def find_browser():
    override = os.environ.get("DESIGN_PPT_BROWSER")
    if override and os.path.exists(override):
        return override
    for p in BROWSER_CANDIDATES:
        if os.path.exists(p):
            return p
    raise SystemExit(
        "Chrome/Edge를 찾을 수 없습니다. DESIGN_PPT_BROWSER 환경변수로 실행파일 경로를 지정하세요."
    )
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_build_design_ppt.py::test_find_browser_env_override -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add plugins/design-ppt/skills/creating-design-ppt/scripts/
git commit -m "feat: add find_browser with DESIGN_PPT_BROWSER override"
```

---

### Task 5: `assemble_pptx` — full-bleed slides + notes + author (TDD)

**Files:**
- Modify: `scripts/build_design_ppt.py`
- Test: `scripts/tests/test_build_design_ppt.py`

- [ ] **Step 1: Add the failing test** (uses a 1×1 PNG so no PIL dependency)

```python
PNG_1x1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC"
)


def test_assemble_pptx_fullbleed_notes_author(tmp_path):
    png = tmp_path / "s.png"
    png.write_bytes(PNG_1x1)
    out = tmp_path / "sample v1.0.pptx"
    b.assemble_pptx([{"png": str(png), "notes": "노트내용"}], str(out))

    from pptx import Presentation
    prs = Presentation(str(out))
    assert len(prs.slides) == 1
    assert prs.core_properties.author == "IT전략팀"
    assert prs.slides[0].has_notes_slide
    assert "노트내용" in prs.slides[0].notes_slide.notes_text_frame.text
    pic = prs.slides[0].shapes[0]
    assert pic.left == 0 and pic.top == 0
    assert pic.width == prs.slide_width
    assert pic.height == prs.slide_height
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_build_design_ppt.py::test_assemble_pptx_fullbleed_notes_author -v`
Expected: FAIL — `AttributeError: ... has no attribute 'assemble_pptx'`

- [ ] **Step 3: Add `_set_notes` and `assemble_pptx` to `build_design_ppt.py`**

```python
def _set_notes(slide, text):
    tf = slide.notes_slide.notes_text_frame
    tf.text = text
    for p in tf.paragraphs:
        for r in p.runs:
            r._r.get_or_add_rPr().set("lang", "ko-KR")


def assemble_pptx(slides, out_path):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]
    for item in slides:
        s = prs.slides.add_slide(blank)
        s.shapes.add_picture(item["png"], 0, 0, width=SLIDE_W, height=SLIDE_H)
        if item.get("notes"):
            _set_notes(s, item["notes"])
    prs.core_properties.author = AUTHOR
    prs.core_properties.last_modified_by = AUTHOR
    prs.save(out_path)
    return out_path
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_build_design_ppt.py::test_assemble_pptx_fullbleed_notes_author -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add plugins/design-ppt/skills/creating-design-ppt/scripts/
git commit -m "feat: add assemble_pptx (full-bleed pictures, notes, IT전략팀 author)"
```

---

### Task 6: `capture_slide` + `build`/`main` orchestration (integration test)

**Files:**
- Modify: `scripts/build_design_ppt.py`
- Test: `scripts/tests/test_build_design_ppt.py`

- [ ] **Step 1: Add the integration test** (skips cleanly if no browser is present)

```python
import pytest


def _browser_available():
    try:
        b.find_browser()
        return True
    except SystemExit:
        return False


@pytest.mark.skipif(not _browser_available(), reason="no Chrome/Edge installed")
def test_build_end_to_end(tmp_path):
    deck = tmp_path / "deck.html"
    deck.write_text(SAMPLE, encoding="utf-8")
    out = tmp_path / "sample v1.0.pptx"
    b.build(str(deck), str(out), css="body{margin:0}")

    from pptx import Presentation
    prs = Presentation(str(out))
    assert len(prs.slides) == 2
    assert prs.core_properties.author == "IT전략팀"
    assert prs.slides[0].has_notes_slide
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_build_design_ppt.py::test_build_end_to_end -v`
Expected: FAIL — `AttributeError: ... has no attribute 'build'`

- [ ] **Step 3: Add `capture_slide`, `build`, and `main` to `build_design_ppt.py`**

```python
def capture_slide(page_html_path, png_path, browser):
    url = Path(page_html_path).resolve().as_uri()
    cmd = [
        browser, "--headless=new", "--disable-gpu", "--hide-scrollbars",
        "--force-device-scale-factor=1",
        "--run-all-compositor-stages-before-draw",
        "--virtual-time-budget=3000",
        "--window-size=%d,%d" % (W, H),
        "--screenshot=%s" % png_path,
        url,
    ]
    subprocess.run(
        cmd, check=True, timeout=120,
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    if not os.path.exists(png_path):
        raise SystemExit("렌더 실패: %s" % png_path)


def build(deck_path, out_path, css=None):
    html = Path(deck_path).read_text(encoding="utf-8")
    if css is None:
        css_file = ASSETS / "base.css"
        css = css_file.read_text(encoding="utf-8") if css_file.exists() else ""
    sections = split_sections(html)
    if not sections:
        raise SystemExit("deck.html에서 <section>을 찾지 못했습니다.")
    browser = find_browser()
    slides = []
    with tempfile.TemporaryDirectory() as tmp:
        for i, sec in enumerate(sections):
            page_path = os.path.join(tmp, "slide_%02d.html" % i)
            Path(page_path).write_text(wrap_section_page(sec["html"], css), encoding="utf-8")
            png_path = os.path.join(tmp, "slide_%02d.png" % i)
            capture_slide(page_path, png_path, browser)
            slides.append({"png": png_path, "notes": sec["notes"]})
        assemble_pptx(slides, out_path)
    print("Rendered %d slide(s)." % len(slides))
    return out_path


def main():
    if len(sys.argv) < 3:
        print('Usage: python build_design_ppt.py deck.html "제목 v1.0.pptx"', file=sys.stderr)
        sys.exit(2)
    out = build(sys.argv[1], sys.argv[2])
    print("Saved:", out)


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Run the full test file to verify all pass**

Run: `python -m pytest tests/test_build_design_ppt.py -v`
Expected: PASS (5 passed; the integration test runs because Chrome is installed)

- [ ] **Step 5: Commit**

```bash
git add plugins/design-ppt/skills/creating-design-ppt/scripts/
git commit -m "feat: add capture_slide + build/main orchestration for design-ppt"
```

---

### Task 7: Design assets — `design-tokens.md` + `base.css`

**Files:**
- Create: `plugins/design-ppt/skills/creating-design-ppt/assets/design-tokens.md`
- Create: `plugins/design-ppt/skills/creating-design-ppt/assets/base.css`

- [ ] **Step 1: Create `assets/design-tokens.md`**

```markdown
# design-ppt 디자인 토큰

원본: "사내 표준 백엔드 아키텍처와 테스트 LDAP 구축 보고 v1.0" (Claude Design)에서 추출.

## 캔버스
- 슬라이드: **1920 × 1080 px** 고정 (`<section style="width:1920px;height:1080px">`)
- 패딩: 본문 96px, 표지 120px, 좌우 130px

## 컬러
| 용도 | HEX |
|------|-----|
| 네이비(표지/강조 배경) | `#0b1b3a` |
| 표지 텍스트 | `#ffffff` |
| 라이트(본문 배경) | `#f5f7fa` |
| 본문 텍스트 | `#0b1b3a` |
| 액센트 블루 | `#0b3fd1` |
| 표지 글로우(radial) | `rgba(11,63,209,.38)` |
| APS 로고 그라데이션 | `#BED600` → `#2BA6CB` |

## 타이포
- 폰트: **Pretendard** (weight 500/600/700/800), 폴백 `'Malgun Gothic','맑은 고딕',sans-serif`
- Pretendard는 **시스템 설치 전제** (미설치 시 폴백 렌더 → 자간/굵기 차이 가능)

## 슬라이드 아키타입 (스니펫)
| 파일 | 용도 |
|------|------|
| `sections/00-cover.html` | 표지 (네이비 + 글로우 + 로고 + 제목/부제/작성정보) |
| `sections/01-section-body.html` | 섹션 본문 (제목 + 액센트 바 + 불릿/단락) |
| `sections/02-closing.html` | 마무리 (네이비 + 핵심 메시지) |

## 메타 규칙
- 각 `<section>`은 `data-label`(식별)·`data-speaker-notes`(발표 노트)를 가진다.
- 빌드 스크립트가 `data-speaker-notes`를 PowerPoint 발표자 노트로 옮긴다.
```

- [ ] **Step 2: Create `assets/base.css`**

```css
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body {
  font-family: 'Pretendard', 'Malgun Gothic', '맑은 고딕', sans-serif;
  -webkit-font-smoothing: antialiased;
}
section { overflow: hidden; }
```

- [ ] **Step 3: Verify files exist**

Run: `python -c "import os; print(all(os.path.exists(p) for p in ['plugins/design-ppt/skills/creating-design-ppt/assets/design-tokens.md','plugins/design-ppt/skills/creating-design-ppt/assets/base.css']))"`
Expected: `True`

- [ ] **Step 4: Commit**

```bash
git add plugins/design-ppt/skills/creating-design-ppt/assets/
git commit -m "feat: add design-ppt design tokens and base.css"
```

---

### Task 8: Section snippets

**Files:**
- Create: `plugins/design-ppt/skills/creating-design-ppt/assets/sections/00-cover.html`
- Create: `plugins/design-ppt/skills/creating-design-ppt/assets/sections/01-section-body.html`
- Create: `plugins/design-ppt/skills/creating-design-ppt/assets/sections/02-closing.html`

Each snippet is a complete `<section>` using the design tokens, with `[[...]]` placeholders Claude replaces when assembling a deck. Keep `data-label`/`data-speaker-notes` attributes.

- [ ] **Step 1: Create `sections/00-cover.html`**

```html
<section data-label="표지" data-speaker-notes="[[발표 노트: 표지 슬라이드 설명]]"
  style="width:1920px; height:1080px; background:#0b1b3a; color:#fff; padding:120px 130px;
         display:flex; flex-direction:column; justify-content:center; position:relative; overflow:hidden;">
  <div style="position:absolute; right:-180px; top:-180px; width:760px; height:760px; border-radius:50%;
              background:radial-gradient(circle at center, rgba(11,63,209,.38), transparent 68%);"></div>
  <div style="position:relative;">
    <div style="font-size:34px; font-weight:600; color:#9fc1ff; letter-spacing:2px;">[[상단 분류 / 부서: 예) IT전략팀]]</div>
    <h1 style="font-size:96px; font-weight:800; line-height:1.15; margin:28px 0 0;">[[보고서 제목]]</h1>
    <div style="font-size:40px; font-weight:500; color:#c9d6ee; margin-top:24px;">[[부제]]</div>
    <div style="font-size:30px; color:#8da3c7; margin-top:80px;">[[작성일 YYYY-MM-DD]] · 작성 IT전략팀 · [[버전 v1.0]]</div>
  </div>
</section>
```

- [ ] **Step 2: Create `sections/01-section-body.html`**

```html
<section data-label="[[슬라이드 라벨]]" data-speaker-notes="[[발표 노트]]"
  style="width:1920px; height:1080px; background:#f5f7fa; color:#0b1b3a; padding:96px 130px;
         display:flex; flex-direction:column;">
  <h2 style="font-size:64px; font-weight:800; margin:0;">[[섹션 제목]]</h2>
  <div style="width:120px; height:8px; background:#0b3fd1; margin:24px 0 56px; border-radius:4px;"></div>
  <ul style="font-size:38px; line-height:1.6; font-weight:500; margin:0; padding-left:1.2em;">
    <li style="margin-bottom:24px;"><b style="color:#0b3fd1;">[[항목]]</b> — [[설명]]</li>
    <li style="margin-bottom:24px;"><b style="color:#0b3fd1;">[[항목]]</b> — [[설명]]</li>
    <li style="margin-bottom:24px;"><b style="color:#0b3fd1;">[[항목]]</b> — [[설명]]</li>
  </ul>
</section>
```

- [ ] **Step 3: Create `sections/02-closing.html`**

```html
<section data-label="마무리" data-speaker-notes="[[발표 노트: 마무리 메시지]]"
  style="width:1920px; height:1080px; background:#0b1b3a; color:#fff; padding:120px 130px;
         display:flex; flex-direction:column; justify-content:center; align-items:center; text-align:center;">
  <h2 style="font-size:80px; font-weight:800; margin:0;">[[핵심 메시지]]</h2>
  <p style="font-size:40px; color:#c9d6ee; font-weight:500; margin-top:40px; max-width:1200px; line-height:1.5;">[[보조 설명 / 다음 단계]]</p>
  <div style="font-size:28px; color:#8da3c7; margin-top:96px;">작성 IT전략팀</div>
</section>
```

- [ ] **Step 4: Render-sanity check** — build a 3-slide deck from the snippets and confirm it renders without error

```bash
cd plugins/design-ppt/skills/creating-design-ppt
python -c "import pathlib; s=pathlib.Path('assets/sections'); open('snippet_check.html','w',encoding='utf-8').write('<html><body>'+''.join((s/f).read_text(encoding='utf-8') for f in ['00-cover.html','01-section-body.html','02-closing.html'])+'</body></html>')"
python scripts/build_design_ppt.py snippet_check.html "snippet check v1.0.pptx"
```
Expected: `Rendered 3 slide(s).` then `Saved: snippet check v1.0.pptx`

- [ ] **Step 5: Inspect one rendered slide visually** (optional but recommended)

Open `snippet check v1.0.pptx` in PowerPoint, or re-render slide 0 to a PNG and view it:
Run: `python -c "import build_design_ppt as b,sys; sys.path.insert(0,'scripts')"` — if unsure, just open the .pptx.
Confirm: navy cover, blue accent bar on body slide, fonts look like Pretendard (or Malgun Gothic fallback).

- [ ] **Step 6: Clean up scratch files and commit**

```bash
rm -f plugins/design-ppt/skills/creating-design-ppt/snippet_check.html "plugins/design-ppt/skills/creating-design-ppt/snippet check v1.0.pptx"
git add plugins/design-ppt/skills/creating-design-ppt/assets/sections/
git commit -m "feat: add design-ppt section snippets (cover, section-body, closing)"
```

---

### Task 9: `SKILL.md` — the workflow skill

**Files:**
- Create: `plugins/design-ppt/skills/creating-design-ppt/SKILL.md`

- [ ] **Step 1: Create `SKILL.md`**

```markdown
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
```

- [ ] **Step 2: Validate the skill frontmatter parses**

Run: `python -c "import io; t=open('plugins/design-ppt/skills/creating-design-ppt/SKILL.md',encoding='utf-8').read(); assert t.startswith('---') and 'name: creating-design-ppt' in t; print('SKILL.md OK')"`
Expected: `SKILL.md OK`

- [ ] **Step 3: Commit**

```bash
git add plugins/design-ppt/skills/creating-design-ppt/SKILL.md
git commit -m "feat: add creating-design-ppt skill (HTML deck → pptx workflow)"
```

---

### Task 10: End-to-end verification against the real design

**Files:** none (verification only)

- [ ] **Step 1: Run the full test suite**

Run: `cd plugins/design-ppt/skills/creating-design-ppt/scripts && python -m pytest tests/ -v`
Expected: all tests pass.

- [ ] **Step 2: Confirm Pretendard availability** (informational)

Run: `python -c "import os,glob; print([f for f in glob.glob(os.path.expandvars(r'%WINDIR%\\Fonts\\*')) if 'pretendard' in f.lower()][:3])"`
Expected: a non-empty list if Pretendard is installed. If empty, note that rendering will use the Malgun Gothic fallback (acceptable; flag to user).

- [ ] **Step 3: Build a 3-slide sample deck end to end**

Assemble a `sample.html` from the three snippets with placeholders filled with real sample text (cover + one body + closing), then:
Run: `python scripts/build_design_ppt.py sample.html "design ppt 샘플 v1.0.pptx"`
Expected: `Rendered 3 slide(s).` and `Saved: design ppt 샘플 v1.0.pptx`

- [ ] **Step 4: Verify the output programmatically**

Run:
```bash
python -c "from pptx import Presentation; p=Presentation('design ppt 샘플 v1.0.pptx'); print('slides', len(p.slides)); print('author', p.core_properties.author); print('notes0', p.slides[0].notes_slide.notes_text_frame.text[:30])"
```
Expected: `slides 3`, `author IT전략팀`, and a non-empty notes string.

- [ ] **Step 5: Visual confirmation**

Open `design ppt 샘플 v1.0.pptx` in PowerPoint and confirm each slide matches the
intended design (navy cover with glow, blue accent bar, closing slide). If fonts
look wrong, install Pretendard and rebuild.

- [ ] **Step 6: Clean up scratch artifacts and commit any doc tweaks**

```bash
rm -f plugins/design-ppt/skills/creating-design-ppt/sample.html "plugins/design-ppt/skills/creating-design-ppt/design ppt 샘플 v1.0.pptx"
```
If verification surfaced fixes (token tweaks, render flags), commit them:
```bash
git add -A && git commit -m "fix: design-ppt render adjustments from e2e verification"
```

---

## Self-Review

**Spec coverage:**
- §3 design tokens → Task 7 (design-tokens.md, base.css). ✓
- §4 structure (plugin/command/skill/assets/scripts) → Tasks 1, 7, 8, 9, 2–6. ✓
- §5 workflow (Claude assembles, build, verify) → Task 9 SKILL.md + Task 1 command. ✓
- §6 pipeline (split → wrap → headless capture → full-bleed pptx + notes + author) → Tasks 2–6. ✓
- §7 team rules (author/filename/notes lang/미정) → Task 5 (author), Task 9 (filename/미정), Task 5 (`_set_notes` lang). ✓
- §8 YAGNI (no native text, no docx, no reverse converter) → not implemented, as intended. ✓
- §9 risks (font fidelity, render consistency) → Task 10 Step 2 (Pretendard check), capture flags `--force-device-scale-factor=1`/`--virtual-time-budget`. ✓

**Placeholder scan:** Code steps contain full code. `[[...]]` markers exist only inside snippet templates (intended design, replaced at authoring time) and are explicitly called out in SKILL.md "Common Mistakes." No TODO/TBD in logic. ✓

**Type/name consistency:** `split_sections` → dicts with `label`/`notes`/`html`; `build` consumes `notes`/`html`; `assemble_pptx` consumes `png`/`notes`; `find_browser`/`capture_slide`/`wrap_section_page` names match across tasks and tests. Constants `W/H/SLIDE_W/SLIDE_H/AUTHOR/ASSETS` defined once in Task 2. ✓
