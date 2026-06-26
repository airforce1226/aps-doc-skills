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
