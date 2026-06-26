# design-ppt Native Editable Mode Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a `--mode native` build path to design-ppt so the same `deck.html` produces a .pptx whose text/boxes/rules/tables are editable PowerPoint native objects (image mode stays the default, unchanged).

**Architecture:** Reuse the existing headless-Chrome pipeline, but swap `--screenshot` for `--dump-dom` with an injected measurement script that emits per-node geometry+style as JSON. Python maps each node to a python-pptx native shape (TextBox / Rectangle / RoundRect+gradient / Table); unmappable decoration (charts, logo gradient) falls back to a Pillow-cropped raster of that node's bbox taken from a one-shot section screenshot.

**Tech Stack:** Python 3.12, python-pptx (+ its bundled Pillow), lxml (via python-pptx), headless Chrome/Edge. No new pip dependencies.

---

## File Structure

```
plugins/design-ppt/skills/creating-design-ppt/
├─ scripts/
│  ├─ build_design_ppt.py        # MODIFY: add --mode {image|native}; image path untouched
│  ├─ native_render.py           # CREATE: measurement JS, dump-dom parse, node→shape mapping,
│  │                             #         gradient/table/color-snap helpers, build report
│  └─ tests/
│     ├─ test_build_design_ppt.py   # existing (unchanged)
│     └─ test_native_render.py      # CREATE: unit tests (no browser) + 1 browser-gated e2e
├─ assets/
│  ├─ sections/01-cover.html …    # MODIFY (Task 11): add data-ppt role hints
│  └─ design-tokens.md            # MODIFY (Task 12): note color-snap palette is source of truth
└─ SKILL.md                       # MODIFY (Task 12): document output modes
```

**Conventions to follow (from existing code):**
- `# -*- coding: utf-8 -*-` header on every .py file.
- Korean error messages via `raise SystemExit("…")`, matching `build_design_ppt.py`.
- Reuse `build_design_ppt.find_browser`, `split_sections`, `wrap_section_page`, `_set_notes`, `AUTHOR`, `SLIDE_W`, `SLIDE_H` — do **not** duplicate them.
- Tests import the module via `sys.path.insert(0, str(Path(__file__).resolve().parents[1]))`.

**Color/size facts (verified against live Chrome dump-dom):**
- bbox `x/y/w/h` are CSS px on the 1920×1080 canvas.
- `fontFamily` arrives quoted, e.g. `"Malgun Gothic"`; take the first family, strip quotes.
- color arrives as `rgb(r, g, b)`; gradient as `linear-gradient(90deg, rgb(...), rgb(...))`.
- `pt = px × 0.54`; `emu_per_px = SLIDE_W_EMU / 1920`.

---

## Task 1: Conversion + color helpers

**Files:**
- Create: `scripts/native_render.py`
- Test: `scripts/tests/test_native_render.py`

- [ ] **Step 1: Write the failing test**

```python
# -*- coding: utf-8 -*-
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import native_render as nr


def test_px_to_pt():
    assert nr.px_to_pt(104) == 56.16  # 104 * 0.54


def test_px_to_emu_full_width():
    # 1920 px must map to the full 16:9 slide width in EMU.
    assert nr.px_to_emu(1920) == nr.SLIDE_W_EMU


def test_rgb_to_hex():
    assert nr.rgb_to_hex("rgb(255, 255, 255)") == "FFFFFF"
    assert nr.rgb_to_hex("rgb(11, 27, 58)") == "0B1B3A"


def test_snap_color_exact_and_near():
    assert nr.snap_color("0B1B3A") == "0B1B3A"          # exact token
    assert nr.snap_color("0C1C3B") == "0B1B3A"          # within tolerance -> navy
    assert nr.snap_color("00FF00", tol=5) == "00FF00"   # far -> unchanged
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest scripts/tests/test_native_render.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'native_render'`

- [ ] **Step 3: Write minimal implementation**

```python
# -*- coding: utf-8 -*-
"""Native-object renderer for design-ppt (--mode native).

Reads a deck.html section's computed layout (via headless Chrome --dump-dom)
and emits editable python-pptx shapes instead of a baked screenshot.
"""
import re

from pptx.util import Emu

# Slide is 13.333" x 7.5" (16:9). 1 inch = 914400 EMU.
SLIDE_W_EMU = int(13.333 * 914400)
SLIDE_H_EMU = int(7.5 * 914400)
PX_TO_PT = 0.54  # 작업지시서 §3 환산

# Single source of truth for color snapping — mirrors assets/design-tokens.md.
PALETTE = {
    "navy": "0B1B3A", "navy2": "16263F", "blue": "0B3FD1", "blueSoft": "3F6BD6",
    "blueTint": "7FA3FF", "paper": "F5F7FA", "white": "FFFFFF", "slate": "5B6B85",
    "ink": "26354F", "line": "E1E7F0", "line2": "D8DFE9", "gradLime": "BED600",
    "gradCyan": "2BA6CB", "danger": "C0392B", "success": "1F7A47", "softBlue": "E7EDF8",
    "softBlue2": "EEF2FB", "footer": "9AA6B8",
}


def px_to_pt(px):
    return round(float(px) * PX_TO_PT, 2)


def px_to_emu(px):
    return Emu(int(round(float(px) * SLIDE_W_EMU / 1920)))


def rgb_to_hex(css):
    m = re.search(r"rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)", css or "")
    if not m:
        return None
    return "".join("%02X" % int(m.group(i)) for i in (1, 2, 3))


def snap_color(hex6, tol=12):
    """Snap an arbitrary RRGGBB to the nearest palette token within tol; else keep."""
    if not hex6:
        return hex6
    target = hex6.upper()
    tr, tg, tb = (int(target[i:i + 2], 16) for i in (0, 2, 4))
    best, best_d = target, None
    for value in PALETTE.values():
        r, g, bb = (int(value[i:i + 2], 16) for i in (0, 2, 4))
        d = abs(r - tr) + abs(g - tg) + abs(bb - tb)
        if best_d is None or d < best_d:
            best, best_d = value, d
    return best if best_d is not None and best_d <= tol * 3 else target
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest scripts/tests/test_native_render.py -v`
Expected: PASS (4 passed)

- [ ] **Step 5: Commit**

```bash
git add scripts/native_render.py scripts/tests/test_native_render.py
git commit -m "feat: add native_render conversion and color-snap helpers"
```

---

## Task 2: Measurement JS + dump-dom layout extraction

**Files:**
- Modify: `scripts/native_render.py`
- Test: `scripts/tests/test_native_render.py`

The measurement script walks the DOM, assigns each node a `role` (from `data-ppt` if present, else heuristic), and emits geometry+style. `extract_layout` parses the `<pre id="__layout__">` JSON out of a dump-dom string (testable without a browser).

- [ ] **Step 1: Write the failing test**

```python
def test_extract_layout_parses_pre_json():
    dom = (
        '<html><body><section></section>'
        '<pre id="__layout__">[{"role":"text","x":120,"y":300,"w":714,"h":138,'
        '"text":"\\uc81c\\ubaa9","font":"\\"Malgun Gothic\\"","sizePx":104,'
        '"weight":"800","color":"rgb(255, 255, 255)","align":"left"}]</pre>'
        '</body></html>'
    )
    nodes = nr.extract_layout(dom)
    assert len(nodes) == 1
    assert nodes[0]["role"] == "text"
    assert nodes[0]["x"] == 120 and nodes[0]["w"] == 714
    assert nodes[0]["text"] == "제목"


def test_extract_layout_missing_pre_raises():
    import pytest
    with pytest.raises(SystemExit):
        nr.extract_layout("<html><body>no pre here</body></html>")


def test_measure_js_is_nonempty_string():
    assert isinstance(nr.MEASURE_JS, str) and "__layout__" in nr.MEASURE_JS
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest scripts/tests/test_native_render.py -k extract_layout -v`
Expected: FAIL with `AttributeError: module 'native_render' has no attribute 'extract_layout'`

- [ ] **Step 3: Write minimal implementation**

Append to `native_render.py`:

```python
import json

# Injected before </body>. Walks the DOM, classifies each node, emits layout JSON.
MEASURE_JS = r"""
<script>
(function(){
  function firstFont(f){ return (f||'').split(',')[0].replace(/["']/g,'').trim(); }
  function hasElementChildren(el){
    for (var i=0;i<el.children.length;i++){
      var c=el.children[i], s=getComputedStyle(c);
      if (s.display!=='none' && c.getBoundingClientRect().width>0) return true;
    }
    return false;
  }
  function role(el, cs){
    var hint=el.getAttribute('data-ppt'); if(hint) return hint;
    var tag=el.tagName.toLowerCase();
    if(tag==='table') return 'table';
    if(tag==='img'||tag==='svg'||tag==='canvas') return 'raster';
    var r=el.getBoundingClientRect();
    var grad=(cs.backgroundImage||'').indexOf('gradient')>=0;
    if((r.height<=6 || r.width<=6) && (grad || cs.backgroundColor!=='rgba(0, 0, 0, 0)')) return 'rule';
    var txt=(el.textContent||'').trim();
    if(txt && !hasElementChildren(el)) return 'text';
    if(cs.backgroundColor!=='rgba(0, 0, 0, 0)' || cs.borderTopWidth!=='0px'
       || (cs.borderRadius && cs.borderRadius!=='0px')) return 'box';
    return 'skip';
  }
  var out=[], all=document.querySelectorAll('section *');
  for(var i=0;i<all.length;i++){
    var el=all[i], cs=getComputedStyle(el);
    if(cs.display==='none'||cs.visibility==='hidden') continue;
    var ro=role(el,cs); if(ro==='skip') continue;
    var r=el.getBoundingClientRect();
    if(r.width<=0||r.height<=0) continue;
    var node={role:ro, x:r.left, y:r.top, w:r.width, h:r.height,
      color:cs.color, bg:cs.backgroundColor, grad:cs.backgroundImage,
      align:cs.textAlign, font:firstFont(cs.fontFamily), sizePx:parseFloat(cs.fontSize),
      weight:cs.fontWeight, ls:cs.letterSpacing,
      radius:parseFloat(cs.borderTopLeftRadius)||0};
    if(ro==='text'){ node.text=(el.textContent||'').trim(); }
    if(ro==='table'){
      var rows=[];
      el.querySelectorAll('tr').forEach(function(tr){
        var cells=[];
        tr.querySelectorAll('th,td').forEach(function(td){
          var tcs=getComputedStyle(td);
          cells.push({text:(td.textContent||'').trim(), bg:tcs.backgroundColor,
            color:tcs.color, weight:tcs.fontWeight, align:tcs.textAlign,
            header:td.tagName.toLowerCase()==='th'});
        });
        rows.push(cells);
      });
      node.rows=rows;
    }
    out.push(node);
  }
  var pre=document.createElement('pre'); pre.id='__layout__';
  pre.textContent=JSON.stringify(out); document.body.appendChild(pre);
})();
</script>
"""

_PRE_RE = re.compile(r'<pre id="__layout__">(.*?)</pre>', re.DOTALL)


def extract_layout(dom_text):
    m = _PRE_RE.search(dom_text)
    if not m:
        raise SystemExit("native 모드: __layout__ JSON을 dump-dom에서 찾지 못했습니다.")
    raw = m.group(1)
    # dump-dom HTML-escapes &, <, > inside <pre>; undo before JSON parse.
    raw = raw.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    return json.loads(raw)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest scripts/tests/test_native_render.py -k "extract_layout or measure_js" -v`
Expected: PASS (3 passed)

- [ ] **Step 5: Commit**

```bash
git add scripts/native_render.py scripts/tests/test_native_render.py
git commit -m "feat: add measurement JS and dump-dom layout extraction"
```

---

## Task 3: dump-dom browser call

**Files:**
- Modify: `scripts/native_render.py`
- Test: `scripts/tests/test_native_render.py`

- [ ] **Step 1: Write the failing test**

```python
def test_dump_dom_builds_expected_command(monkeypatch):
    captured = {}

    class FakeResult:
        returncode = 0
        stdout = b"<html><pre id=\"__layout__\">[]</pre></html>"
        stderr = b""

    def fake_run(cmd, **kw):
        captured["cmd"] = cmd
        return FakeResult()

    monkeypatch.setattr(nr.subprocess, "run", fake_run)
    dom = nr.dump_dom("C:/tmp/page.html", "chrome.exe")
    assert "--dump-dom" in captured["cmd"]
    assert "--headless=new" in captured["cmd"]
    assert "__layout__" in dom
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest scripts/tests/test_native_render.py -k dump_dom -v`
Expected: FAIL with `AttributeError: module 'native_render' has no attribute 'dump_dom'`

- [ ] **Step 3: Write minimal implementation**

Append to `native_render.py`:

```python
import subprocess
from pathlib import Path

W, H = 1920, 1080


def dump_dom(page_html_path, browser):
    url = Path(page_html_path).resolve().as_uri()
    cmd = [
        browser, "--headless=new", "--disable-gpu", "--hide-scrollbars",
        "--force-device-scale-factor=1",
        "--run-all-compositor-stages-before-draw",
        "--virtual-time-budget=3000",
        "--window-size=%d,%d" % (W, H),
        "--dump-dom", url,
    ]
    result = subprocess.run(cmd, check=False, timeout=120,
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    dom = result.stdout.decode("utf-8", errors="replace") if result.stdout else ""
    if result.returncode != 0 or "__layout__" not in dom:
        err = result.stderr.decode(errors="replace").strip() if result.stderr else ""
        raise SystemExit("native 모드 dump-dom 실패 (exit %d): %s" % (result.returncode, err))
    return dom
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest scripts/tests/test_native_render.py -k dump_dom -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add scripts/native_render.py scripts/tests/test_native_render.py
git commit -m "feat: add dump_dom headless browser call for native mode"
```

---

## Task 4: Text node → TextBox

**Files:**
- Modify: `scripts/native_render.py`
- Test: `scripts/tests/test_native_render.py`

- [ ] **Step 1: Write the failing test**

```python
def test_add_text_creates_editable_textbox(tmp_path):
    from pptx import Presentation
    from pptx.util import Cm
    prs = Presentation()
    prs.slide_width = Cm(33.867); prs.slide_height = Cm(19.05)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    node = {"role": "text", "x": 120, "y": 300, "w": 714, "h": 138,
            "text": "제목", "font": "Malgun Gothic", "sizePx": 104,
            "weight": "800", "color": "rgb(255, 255, 255)", "align": "left",
            "ls": "normal"}
    shp = nr.add_text(slide, node)
    assert shp.has_text_frame
    assert shp.text_frame.paragraphs[0].runs[0].text == "제목"
    assert shp.text_frame.paragraphs[0].runs[0].font.bold is True
    assert round(shp.text_frame.paragraphs[0].runs[0].font.size.pt, 2) == 56.16
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest scripts/tests/test_native_render.py -k add_text -v`
Expected: FAIL with `AttributeError: module 'native_render' has no attribute 'add_text'`

- [ ] **Step 3: Write minimal implementation**

Append to `native_render.py`:

```python
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

_ALIGN = {"left": PP_ALIGN.LEFT, "right": PP_ALIGN.RIGHT,
          "center": PP_ALIGN.CENTER, "justify": PP_ALIGN.JUSTIFY}


def _is_bold(weight):
    try:
        return int(weight) >= 600
    except (TypeError, ValueError):
        return weight == "bold"


def add_text(slide, node):
    box = slide.shapes.add_textbox(px_to_emu(node["x"]), px_to_emu(node["y"]),
                                   px_to_emu(node["w"]), px_to_emu(node["h"]))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = _ALIGN.get(node.get("align"), PP_ALIGN.LEFT)
    run = p.add_run()
    run.text = node.get("text", "")
    run.font.name = node.get("font") or "맑은 고딕"
    run.font.size = __import__("pptx").util.Pt(px_to_pt(node["sizePx"]))
    run.font.bold = _is_bold(node.get("weight"))
    hexc = snap_color(rgb_to_hex(node.get("color")))
    if hexc:
        run.font.color.rgb = RGBColor.from_string(hexc)
    return box
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest scripts/tests/test_native_render.py -k add_text -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add scripts/native_render.py scripts/tests/test_native_render.py
git commit -m "feat: map text nodes to editable native textboxes"
```

---

## Task 5: Box / rule nodes → Rectangle (+ gradient)

**Files:**
- Modify: `scripts/native_render.py`
- Test: `scripts/tests/test_native_render.py`

- [ ] **Step 1: Write the failing test**

```python
def test_add_box_solid_fill(tmp_path):
    from pptx import Presentation
    from pptx.util import Cm
    prs = Presentation()
    prs.slide_width = Cm(33.867); prs.slide_height = Cm(19.05)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    node = {"role": "box", "x": 0, "y": 0, "w": 1920, "h": 1080,
            "bg": "rgb(11, 27, 58)", "grad": "none", "radius": 0}
    shp = nr.add_box(slide, node)
    assert str(shp.fill.fore_color.rgb) == "0B1B3A"
    assert shp.line.fill.type is not None  # no exception accessing line


def test_add_rule_gradient_has_gradfill(tmp_path):
    from pptx import Presentation
    from pptx.util import Cm
    prs = Presentation()
    prs.slide_width = Cm(33.867); prs.slide_height = Cm(19.05)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    node = {"role": "rule", "x": 120, "y": 460, "w": 60, "h": 4,
            "grad": "linear-gradient(90deg, rgb(190, 214, 0), rgb(43, 166, 203))",
            "bg": "rgba(0, 0, 0, 0)", "radius": 0}
    shp = nr.add_rule(slide, node)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    assert shp._element.spPr.find("a:gradFill", ns) is not None
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest scripts/tests/test_native_render.py -k "add_box or add_rule" -v`
Expected: FAIL with `AttributeError` (add_box/add_rule undefined)

- [ ] **Step 3: Write minimal implementation**

Append to `native_render.py`:

```python
from lxml import etree
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

_GRAD_STOP_RE = re.compile(r"rgb\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*\)")


def _add_rect(slide, node, rounded=False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, px_to_emu(node["x"]), px_to_emu(node["y"]),
                                 px_to_emu(node["w"]), px_to_emu(node["h"]))
    shp.line.fill.background()  # no border by default; box rule can re-enable
    shp.shadow.inherit = False
    return shp


def add_box(slide, node):
    rounded = float(node.get("radius") or 0) >= 6
    shp = _add_rect(slide, node, rounded=rounded)
    hexc = rgb_to_hex(node.get("bg"))
    if hexc:
        shp.fill.solid()
        shp.fill.fore_color.rgb = RGBColor.from_string(snap_color(hexc))
    else:
        shp.fill.background()
    return shp


def _set_gradient(shp, stops):
    """Inject a horizontal 2+ stop <a:gradFill> into the shape's spPr via lxml."""
    spPr = shp._element.spPr
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill"):
        existing = spPr.find(qn(tag))
        if existing is not None:
            spPr.remove(existing)
    grad = spPr.makeelement(qn("a:gradFill"), {})
    lst = grad.makeelement(qn("a:gsLst"), {})
    n = len(stops)
    for i, hexc in enumerate(stops):
        gs = lst.makeelement(qn("a:gs"), {"pos": str(int(i * 100000 / max(n - 1, 1)))})
        clr = gs.makeelement(qn("a:srgbClr"), {"val": hexc})
        gs.append(clr); lst.append(gs)
    grad.append(lst)
    lin = grad.makeelement(qn("a:lin"), {"ang": "0", "scaled": "1"})
    grad.append(lin)
    ln = spPr.find(qn("a:ln"))
    spPr.insert(list(spPr).index(ln) if ln is not None else len(spPr), grad)


def add_rule(slide, node):
    shp = _add_rect(slide, node, rounded=False)
    stops = [rgb_to_hex(s) for s in _GRAD_STOP_RE.findall(node.get("grad") or "")]
    if len(stops) >= 2:
        _set_gradient(shp, stops)
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = RGBColor.from_string(
            snap_color(rgb_to_hex(node.get("bg")) or PALETTE["blue"]))
    return shp
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest scripts/tests/test_native_render.py -k "add_box or add_rule" -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add scripts/native_render.py scripts/tests/test_native_render.py
git commit -m "feat: map box/rule nodes to native rectangles with gradient rule"
```

---

## Task 6: Table node → native PowerPoint table

**Files:**
- Modify: `scripts/native_render.py`
- Test: `scripts/tests/test_native_render.py`

Style rules (작업지시서 §5.4): header row navy bg + white bold; even body rows zebra `softBlue2`; a row whose first cell text is "합계"/"계" gets blue bg + white bold; numeric cells right-aligned; negative numbers red text.

- [ ] **Step 1: Write the failing test**

```python
def test_add_table_native_with_header_and_total(tmp_path):
    from pptx import Presentation
    from pptx.util import Cm
    prs = Presentation()
    prs.slide_width = Cm(33.867); prs.slide_height = Cm(19.05)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    node = {"role": "table", "x": 130, "y": 300, "w": 1660, "h": 400,
            "rows": [
                [{"text": "항목", "header": True}, {"text": "금액", "header": True}],
                [{"text": "매출", "header": False}, {"text": "1,200", "header": False}],
                [{"text": "합계", "header": False}, {"text": "-50", "header": False}],
            ]}
    shp = nr.add_table(slide, node)
    assert shp.has_table
    tbl = shp.table
    assert len(tbl.rows) == 3 and len(tbl.columns) == 2
    # header cell white bold on navy
    hdr = tbl.cell(0, 0)
    assert str(hdr.fill.fore_color.rgb) == "0B1B3A"
    # total row blue
    assert str(tbl.cell(2, 0).fill.fore_color.rgb) == "0B3FD1"
    # negative number red
    neg_run = tbl.cell(2, 1).text_frame.paragraphs[0].runs[0]
    assert str(neg_run.font.color.rgb) == "C0392B"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest scripts/tests/test_native_render.py -k add_table -v`
Expected: FAIL with `AttributeError: module 'native_render' has no attribute 'add_table'`

- [ ] **Step 3: Write minimal implementation**

Append to `native_render.py`:

```python
_NUM_RE = re.compile(r"^-?[\d,]+(\.\d+)?$")


def _cell_text(cell, text, *, color, bold, align):
    cell.fill.solid() if color is None else None
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = "맑은 고딕"
    run.font.size = __import__("pptx").util.Pt(SIZE_BODY_PT)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)


SIZE_BODY_PT = 14 * PX_TO_PT  # design body 14px token -> pt


def add_table(slide, node):
    rows = node["rows"]
    nrows, ncols = len(rows), max(len(r) for r in rows)
    gf = slide.shapes.add_table(nrows, ncols, px_to_emu(node["x"]), px_to_emu(node["y"]),
                                px_to_emu(node["w"]), px_to_emu(node["h"]))
    tbl = gf.table
    tbl.first_row = False  # we style explicitly, disable built-in banding
    tbl.horz_banding = False
    for ri, row in enumerate(rows):
        is_header = ri == 0 or any(c.get("header") for c in row)
        first_text = (row[0].get("text") if row else "") or ""
        is_total = first_text.strip() in ("합계", "계", "소계", "Total")
        for ci in range(ncols):
            cell = tbl.cell(ri, ci)
            data = row[ci] if ci < len(row) else {"text": ""}
            text = data.get("text", "")
            numeric = bool(_NUM_RE.match(text.strip()))
            negative = text.strip().startswith("-")
            if is_header:
                bg, fg, bold = PALETTE["navy"], PALETTE["white"], True
            elif is_total:
                bg, fg, bold = PALETTE["blue"], PALETTE["white"], True
            elif ri % 2 == 0:
                bg, fg, bold = PALETTE["softBlue2"], PALETTE["ink"], False
            else:
                bg, fg, bold = PALETTE["white"], PALETTE["ink"], False
            if negative and not is_header and not is_total:
                fg = PALETTE["danger"]
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(bg)
            from pptx.enum.text import PP_ALIGN as _A
            align = _A.RIGHT if (numeric and not is_header) else _A.LEFT
            _cell_text(cell, text, color=fg, bold=bold, align=align)
    return gf
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest scripts/tests/test_native_render.py -k add_table -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add scripts/native_render.py scripts/tests/test_native_render.py
git commit -m "feat: map table nodes to native styled PowerPoint tables"
```

---

## Task 7: Raster fallback (Pillow crop of section screenshot)

**Files:**
- Modify: `scripts/native_render.py`
- Test: `scripts/tests/test_native_render.py`

For `role == "raster"` (charts, logo, anything unmappable) crop that node's bbox out of a one-shot full-section PNG and place it as a picture at the same coordinates.

- [ ] **Step 1: Write the failing test**

```python
def test_crop_node_png(tmp_path):
    from PIL import Image
    src = tmp_path / "section.png"
    Image.new("RGB", (1920, 1080), (255, 0, 0)).save(src)
    node = {"x": 100, "y": 50, "w": 200, "h": 80}
    out = nr.crop_node_png(str(src), node, str(tmp_path / "crop.png"))
    img = Image.open(out)
    assert img.size == (200, 80)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest scripts/tests/test_native_render.py -k crop_node -v`
Expected: FAIL with `AttributeError: module 'native_render' has no attribute 'crop_node_png'`

- [ ] **Step 3: Write minimal implementation**

Append to `native_render.py`:

```python
from PIL import Image


def crop_node_png(section_png, node, out_path):
    img = Image.open(section_png)
    x, y = int(round(node["x"])), int(round(node["y"]))
    w, h = int(round(node["w"])), int(round(node["h"]))
    img.crop((x, y, x + w, y + h)).save(out_path)
    return out_path


def add_raster(slide, node, section_png, tmp_dir, idx):
    out = "%s/raster_%d.png" % (tmp_dir, idx)
    crop_node_png(section_png, node, out)
    return slide.shapes.add_picture(out, px_to_emu(node["x"]), px_to_emu(node["y"]),
                                    px_to_emu(node["w"]), px_to_emu(node["h"]))
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest scripts/tests/test_native_render.py -k crop_node -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add scripts/native_render.py scripts/tests/test_native_render.py
git commit -m "feat: add raster fallback via Pillow crop of section screenshot"
```

---

## Task 8: Slide assembly + build_native orchestration

**Files:**
- Modify: `scripts/native_render.py`
- Test: `scripts/tests/test_native_render.py`

`render_slide` dispatches nodes to the right mapper and returns per-slide counts. `build_native` ties sections → dump-dom → render → notes/author, reusing `build_design_ppt` helpers. The screenshot for raster fallback is captured lazily (only if a slide has any raster node).

- [ ] **Step 1: Write the failing test**

```python
def test_render_slide_dispatches_and_counts(tmp_path):
    from pptx import Presentation
    from pptx.util import Cm
    prs = Presentation()
    prs.slide_width = Cm(33.867); prs.slide_height = Cm(19.05)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    nodes = [
        {"role": "box", "x": 0, "y": 0, "w": 1920, "h": 1080,
         "bg": "rgb(11, 27, 58)", "grad": "none", "radius": 0},
        {"role": "text", "x": 120, "y": 300, "w": 700, "h": 138, "text": "제목",
         "font": "Malgun Gothic", "sizePx": 104, "weight": "800",
         "color": "rgb(255,255,255)", "align": "left", "ls": "normal"},
    ]
    counts = nr.render_slide(slide, nodes, section_png=None, tmp_dir=str(tmp_path), si=0)
    assert counts["native"] == 2 and counts["raster"] == 0
    assert any(s.has_text_frame for s in slide.shapes)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest scripts/tests/test_native_render.py -k render_slide -v`
Expected: FAIL with `AttributeError: module 'native_render' has no attribute 'render_slide'`

- [ ] **Step 3: Write minimal implementation**

Append to `native_render.py`:

```python
import tempfile
import os

# Imported lazily inside functions to avoid a circular import at module load.


def render_slide(slide, nodes, section_png, tmp_dir, si):
    counts = {"native": 0, "raster": 0}
    # Boxes/rules first (background layer), then text/tables on top.
    order = {"box": 0, "rule": 1, "raster": 2, "table": 3, "text": 4}
    for node in sorted(nodes, key=lambda n: order.get(n["role"], 5)):
        role = node["role"]
        if role == "box":
            add_box(slide, node); counts["native"] += 1
        elif role == "rule":
            add_rule(slide, node); counts["native"] += 1
        elif role == "text":
            add_text(slide, node); counts["native"] += 1
        elif role == "table":
            add_table(slide, node); counts["native"] += 1
        elif role == "raster":
            if section_png:
                add_raster(slide, node, section_png, tmp_dir, len(slide.shapes))
                counts["raster"] += 1
    return counts


def build_native(deck_path, out_path, css=None, browser=None):
    import build_design_ppt as b
    from pptx import Presentation
    html = Path(deck_path).read_text(encoding="utf-8")
    if css is None:
        css_file = b.ASSETS / "base.css"
        css = css_file.read_text(encoding="utf-8") if css_file.exists() else ""
    sections = b.split_sections(html)
    if not sections:
        raise SystemExit("deck.html에서 <section>을 찾지 못했습니다.")
    browser = browser or b.find_browser()
    prs = Presentation()
    prs.slide_width = b.SLIDE_W
    prs.slide_height = b.SLIDE_H
    blank = prs.slide_layouts[6]
    report = []
    with tempfile.TemporaryDirectory() as tmp:
        for si, sec in enumerate(sections):
            page = os.path.join(tmp, "slide_%02d.html" % si)
            # Inject MEASURE_JS just before </body>.
            wrapped = b.wrap_section_page(sec["html"], css).replace(
                "</body>", MEASURE_JS + "</body>")
            Path(page).write_text(wrapped, encoding="utf-8")
            nodes = extract_layout(dump_dom(page, browser))
            section_png = None
            if any(n["role"] == "raster" for n in nodes):
                section_png = os.path.join(tmp, "slide_%02d.png" % si)
                b.capture_slide(page, section_png, browser)
            slide = prs.slides.add_slide(blank)
            counts = render_slide(slide, nodes, section_png, tmp, si)
            if sec.get("notes"):
                b._set_notes(slide, sec["notes"])
            report.append((si, sec.get("label", ""), counts))
        prs.core_properties.author = b.AUTHOR
        prs.core_properties.last_modified_by = b.AUTHOR
        prs.save(out_path)
    print_report(report)
    return out_path


def print_report(report):
    tot_n = sum(c["native"] for _, _, c in report)
    tot_r = sum(c["raster"] for _, _, c in report)
    for si, label, c in report:
        print("Slide %02d (%s): native=%d raster=%d"
              % (si + 1, label or "-", c["native"], c["raster"]))
    print("Total: %d slides, %d native objects, %d raster fallbacks."
          % (len(report), tot_n, tot_r))
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest scripts/tests/test_native_render.py -k render_slide -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add scripts/native_render.py scripts/tests/test_native_render.py
git commit -m "feat: add render_slide dispatch and build_native orchestration"
```

---

## Task 9: Wire `--mode` flag into build_design_ppt.py

**Files:**
- Modify: `scripts/build_design_ppt.py:157-163` (the `main()` function)
- Test: `scripts/tests/test_build_design_ppt.py`

- [ ] **Step 1: Write the failing test**

Add to `test_build_design_ppt.py`:

```python
def test_parse_args_defaults_to_image():
    deck, out, mode = b.parse_args(["deck.html", "제목 v1.0.pptx"])
    assert deck == "deck.html" and out == "제목 v1.0.pptx" and mode == "image"


def test_parse_args_native_mode():
    _, _, mode = b.parse_args(["deck.html", "제목 v1.0.pptx", "--mode", "native"])
    assert mode == "native"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest scripts/tests/test_build_design_ppt.py -k parse_args -v`
Expected: FAIL with `AttributeError: module 'build_design_ppt' has no attribute 'parse_args'`

- [ ] **Step 3: Write minimal implementation**

Replace `main()` in `build_design_ppt.py` with:

```python
def parse_args(argv):
    mode = "image"
    rest = []
    i = 0
    while i < len(argv):
        if argv[i] == "--mode":
            mode = argv[i + 1]
            i += 2
            continue
        rest.append(argv[i])
        i += 1
    if len(rest) < 2:
        raise SystemExit('Usage: python build_design_ppt.py deck.html "제목 v1.0.pptx" [--mode image|native]')
    if mode not in ("image", "native"):
        raise SystemExit("--mode 는 image 또는 native 여야 합니다: %s" % mode)
    return rest[0], rest[1], mode


def main():
    if len(sys.argv) < 3:
        print('Usage: python build_design_ppt.py deck.html "제목 v1.0.pptx" [--mode image|native]',
              file=sys.stderr)
        sys.exit(2)
    deck, out, mode = parse_args(sys.argv[1:])
    if mode == "native":
        import native_render
        native_render.build_native(deck, out)
    else:
        build(deck, out)
    print("Saved:", out)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest scripts/tests/test_build_design_ppt.py -v`
Expected: PASS (all existing + 2 new; image path unchanged)

- [ ] **Step 5: Commit**

```bash
git add scripts/build_design_ppt.py scripts/tests/test_build_design_ppt.py
git commit -m "feat: add --mode image|native CLI flag (image stays default)"
```

---

## Task 10: Browser-gated end-to-end native build

**Files:**
- Modify: `scripts/tests/test_native_render.py`

- [ ] **Step 1: Write the failing test**

```python
import pytest


def _browser_available():
    import build_design_ppt as b
    try:
        b.find_browser(); return True
    except SystemExit:
        return False


@pytest.mark.skipif(not _browser_available(), reason="no Chrome/Edge installed")
def test_build_native_end_to_end(tmp_path):
    from pptx import Presentation
    deck = tmp_path / "deck.html"
    deck.write_text(
        '<section data-label="표지" data-speaker-notes="노트" '
        'style="width:1920px;height:1080px;position:relative;background:#0b1b3a;">'
        '<div data-ppt="text" style="position:absolute;left:120px;top:300px;'
        'font-size:104px;font-weight:800;color:#fff;font-family:\'Malgun Gothic\';">제목</div>'
        '<div data-ppt="rule" style="position:absolute;left:120px;top:460px;width:60px;'
        'height:4px;background:linear-gradient(90deg,#BED600,#2BA6CB);"></div>'
        '</section>', encoding="utf-8")
    out = tmp_path / "표지 v1.0.pptx"
    nr.build_native(str(deck), str(out), css="html,body{margin:0}")

    prs = Presentation(str(out))
    assert len(prs.slides) == 1
    assert prs.core_properties.author == "IT전략팀"
    shapes = prs.slides[0].shapes
    assert any(s.has_text_frame and "제목" in s.text_frame.text for s in shapes)
    # No full-bleed picture: native mode must not bake a screenshot.
    assert not any(s.shape_type == 13 for s in shapes)  # 13 = PICTURE
```

- [ ] **Step 2: Run test to verify it fails (or skips without a browser)**

Run: `python -m pytest scripts/tests/test_native_render.py -k end_to_end -v`
Expected: PASS if Chrome/Edge present; SKIP otherwise. If it FAILS, debug with superpowers:systematic-debugging before continuing.

- [ ] **Step 3: No new implementation** — this validates Tasks 1-9 together.

- [ ] **Step 4: Run the full suite**

Run: `python -m pytest scripts/tests/ -v`
Expected: all PASS/SKIP, no failures.

- [ ] **Step 5: Commit**

```bash
git add scripts/tests/test_native_render.py
git commit -m "test: add end-to-end native build verification"
```

---

## Task 11: Annotate archetypes with `data-ppt` role hints

**Files:**
- Modify: `assets/sections/01-cover.html`, `03-section-divider.html`, `13-closing.html` (navy archetypes first)
- Modify: `assets/sections/02-executive-summary.html`, `04-body-two-column.html`, `05-metrics.html`, `07-table.html` (content archetypes)
- Modify remaining: `06`, `08`, `10`, `11`, `12` ; mark `09-charts.html` chart container `data-ppt="raster"`

`data-*` attributes are ignored by the image path, so this is purely additive (zero regression risk). For each archetype, add `data-ppt` to the key elements so classification is deterministic.

- [ ] **Step 1: Read one archetype and identify roles**

Run: open `assets/sections/01-cover.html`. Identify: the navy background container (`data-ppt="box"`), the title/label/footer text nodes (`data-ppt="text"`), the accent rule (`data-ppt="rule"`), the logo lockup (`data-ppt="raster"`).

- [ ] **Step 2: Add hints to 01-cover.html**

Example edits (apply the matching attribute to each element):

```html
<!-- background -->
<div class="cover-bg" data-ppt="box" style="…background:#0b1b3a…">
  <div class="eyebrow" data-ppt="text">…</div>
  <h1 class="cover-title" data-ppt="text">…</h1>
  <div class="accent-rule" data-ppt="rule"></div>
  <div class="logo-lockup" data-ppt="raster">…</div>
  <div class="cover-footer" data-ppt="text">…</div>
</div>
```

- [ ] **Step 3: Build the cover natively and eyeball it**

Run: assemble a 1-section `deck.html` from `01-cover.html`, then
`python scripts/build_design_ppt.py cover_deck.html "표지 v1.0.pptx" --mode native`
Open in PowerPoint: title/label/footer are editable text, accent rule has the gradient, logo is a raster. Adjust hints if any element is mis-rendered.

- [ ] **Step 4: Repeat for the remaining archetypes**

For each `assets/sections/*.html`, add `data-ppt` to backgrounds (`box`), text (`text`), rules (`rule`), tables (`table`), charts/logos (`raster`). Build each natively and confirm. Mark `09-charts.html`'s chart bodies `data-ppt="raster"`.

- [ ] **Step 5: Commit**

```bash
git add assets/sections/
git commit -m "feat: annotate APS archetypes with data-ppt role hints for native mode"
```

---

## Task 12: Documentation (SKILL.md + design-tokens.md)

**Files:**
- Modify: `SKILL.md`
- Modify: `assets/design-tokens.md`

- [ ] **Step 1: Add an "출력 모드" section to SKILL.md**

Insert after the Overview section:

```markdown
## 출력 모드 (Output modes)

| 모드 | 명령 | 특성 | 언제 |
|------|------|------|------|
| 이미지 (기본) | `… "제목 v1.0.pptx"` | 픽셀 100% 충실 · **편집 불가**(슬라이드는 이미지, 원문은 노트) | 인쇄·배포용 최종본 |
| 네이티브 | `… "제목 v1.0.pptx" --mode native` | 텍스트·표·도형이 **편집 가능한 네이티브 개체** · 차트/로고 그라데이션은 래스터 폴백 | 받는 사람이 PowerPoint에서 직접 고쳐야 할 때 |

네이티브 모드는 같은 `deck.html`을 헤드리스 Chrome `--dump-dom`으로 측정해 좌표·스타일을
읽고 python-pptx 네이티브 개체로 재배치한다. 빌드 끝에 슬라이드별 `native=/raster=` 리포트를
출력한다. 차트·복합 로고는 의도된 래스터 폴백이다(작업지시서 §5.7 철학).
```

- [ ] **Step 2: Add Quick Reference rows**

```markdown
| Build native (editable) | `python scripts/build_design_ppt.py deck.html "제목 v1.0.pptx" --mode native` |
| 모드별 트레이드오프 | 이미지=픽셀완벽·편집불가 / 네이티브=편집가능·일부 래스터 |
```

- [ ] **Step 3: Add a Common Mistake**

```markdown
- 네이티브 모드 결과에서 차트/로고가 이미지인 것을 버그로 오인 — 의도된 폴백(편집 대상은 텍스트·표·도형).
```

- [ ] **Step 4: Note the color-snap palette in design-tokens.md**

Add at the end of `design-tokens.md`:

```markdown
## 네이티브 모드 색 스냅
`--mode native`는 추출한 색을 위 팔레트의 최근접 토큰으로 스냅해 임의색 유입을 막는다
(`native_render.PALETTE`가 이 표의 코드 사본 — 토큰 변경 시 양쪽을 함께 고친다).
```

- [ ] **Step 5: Commit**

```bash
git add SKILL.md assets/design-tokens.md
git commit -m "docs: document native output mode and color-snap palette"
```

---

## Self-Review Notes

- **Spec coverage:** §2 추출(Task 2-3), §3 매핑/색스냅(Task 1,4,5,6), §4 data-ppt 힌트(Task 11), §5 코드구조/리포트(Task 8), §6 문서(Task 12), §7 DoD(Task 10 e2e + 단위 테스트), §8 마일스톤 N1-N5 ↔ Tasks. 차트 네이티브화는 spec의 명시적 2차 비목표 → 계획 제외(의도적).
- **PALETTE duplication** between `native_render.py` and `design-tokens.md` is called out in Task 12 Step 4 so future edits stay in sync.
- **Pt import:** code uses `__import__("pptx").util.Pt(...)` to avoid a top-level name clash; an implementer may instead add `from pptx.util import Pt` once at the top — either is fine, keep it consistent.
- **Risk:** the only browser-dependent behavior (dump-dom returning post-script DOM) was verified live before this plan; Task 10 re-checks it end-to-end.
