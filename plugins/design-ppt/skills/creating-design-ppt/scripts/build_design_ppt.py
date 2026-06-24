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
    m = re.search(re.escape(name) + r'\s*=\s*"([^"]*)"', attrs)
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


def wrap_section_page(section_html, css):
    return (
        '<!DOCTYPE html><html><head><meta charset="utf-8">'
        "<style>html,body{margin:0;padding:0;}" + css + "</style></head>"
        "<body>" + section_html + "</body></html>"
    )


BROWSER_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
]


def find_browser():
    override = os.environ.get("DESIGN_PPT_BROWSER")
    if override:
        if not os.path.exists(override):
            raise SystemExit("DESIGN_PPT_BROWSER 경로를 찾을 수 없습니다: %s" % override)
        return override
    for p in BROWSER_CANDIDATES:
        if os.path.exists(p):
            return p
    raise SystemExit(
        "Chrome/Edge를 찾을 수 없습니다. DESIGN_PPT_BROWSER 환경변수로 실행파일 경로를 지정하세요."
    )


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
    result = subprocess.run(
        cmd, check=False, timeout=120,
        stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
    )
    if result.returncode != 0 or not os.path.exists(png_path):
        err = result.stderr.decode(errors="replace").strip() if result.stderr else ""
        raise SystemExit(
            "Chrome 렌더 실패 (exit %d): %s" % (result.returncode, err or png_path)
        )


def _set_notes(slide, text):
    tf = slide.notes_slide.notes_text_frame
    tf.text = text
    for p in tf.paragraphs:
        for r in p.runs:
            # _r is python-pptx's documented escape hatch to the lxml element;
            # the public API has no language-tagging, so set lang=ko-KR directly.
            r._r.get_or_add_rPr().set("lang", "ko-KR")


def assemble_pptx(slides, out_path):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]  # index 6 = "Blank" in python-pptx's default template
    for item in slides:
        s = prs.slides.add_slide(blank)
        s.shapes.add_picture(item["png"], 0, 0, width=SLIDE_W, height=SLIDE_H)
        if item.get("notes"):
            _set_notes(s, item["notes"])
    prs.core_properties.author = AUTHOR
    prs.core_properties.last_modified_by = AUTHOR
    prs.save(out_path)
    return out_path


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
