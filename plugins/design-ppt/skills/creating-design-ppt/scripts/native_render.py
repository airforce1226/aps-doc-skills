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
