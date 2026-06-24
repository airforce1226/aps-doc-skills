# -*- coding: utf-8 -*-
import base64
import os
import sys
from pathlib import Path

import pytest

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

PNG_1x1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC"
)


def test_split_sections_count_and_fields():
    secs = b.split_sections(SAMPLE)
    assert len(secs) == 2
    assert secs[0]["label"] == "표지"
    assert secs[0]["notes"] == "표지 노트"
    assert "COVER" in secs[0]["html"]
    assert secs[0]["html"].lstrip().startswith("<section")
    assert secs[1]["label"] == "본문"


def test_wrap_section_page_is_full_html():
    page = b.wrap_section_page("<section>X</section>", "body{margin:0}")
    assert page.startswith("<!DOCTYPE html>")
    assert "body{margin:0}" in page
    assert "<section>X</section>" in page
    assert 'charset="utf-8"' in page


def test_find_browser_env_override(tmp_path):
    fake = tmp_path / "chrome.exe"
    fake.write_text("x")
    os.environ["DESIGN_PPT_BROWSER"] = str(fake)
    try:
        assert b.find_browser() == str(fake)
    finally:
        del os.environ["DESIGN_PPT_BROWSER"]


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
    assert "표지 노트" in prs.slides[0].notes_slide.notes_text_frame.text
