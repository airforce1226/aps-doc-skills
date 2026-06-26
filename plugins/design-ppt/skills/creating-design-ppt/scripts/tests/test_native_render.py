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
