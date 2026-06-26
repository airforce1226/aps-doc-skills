# -*- coding: utf-8 -*-
"""Native-object renderer for design-ppt (--mode native).

Reads a deck.html section's computed layout (via headless Chrome --dump-dom)
and emits editable python-pptx shapes instead of a baked screenshot.
"""
import json
import re
import subprocess
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

# Standard 16:9 slide = 13.333" x 7.5". Exact EMU per 작업지시서 §3
# (12192000 x 6858000), so px_to_emu(1) == 6350 exactly (12192000 / 1920).
SLIDE_W_EMU = 12192000
SLIDE_H_EMU = 6858000
PX_TO_PT = 0.54  # design canvas px -> slide pt

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
    run.font.size = Pt(px_to_pt(node["sizePx"]))
    run.font.bold = _is_bold(node.get("weight"))
    hexc = snap_color(rgb_to_hex(node.get("color")))
    if hexc:
        run.font.color.rgb = RGBColor.from_string(hexc)
    return box


_GRAD_STOP_RE = re.compile(r"rgb\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*\)")


def _add_rect(slide, node, rounded=False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, px_to_emu(node["x"]), px_to_emu(node["y"]),
                                 px_to_emu(node["w"]), px_to_emu(node["h"]))
    shp.line.fill.background()  # no border by default
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
