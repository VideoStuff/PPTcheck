"""Sizes and aspect ratios.

PowerPoint stores positions and sizes in EMU (English Metric Units):
914,400 EMU = 1 inch. Office treats 1 inch as 96 pixels, so a 13.333 x 7.5 in
widescreen slide is 1280 x 720 px.
"""
from __future__ import annotations

from math import gcd

EMU_PER_INCH = 914_400
PX_PER_INCH = 96

# Common aspect ratios, checked in this order (first close match wins).
NAMED_RATIOS = [(16, 9), (4, 3), (16, 10), (3, 2), (1, 1), (21, 9), (5, 4), (2, 1),
                (9, 16), (3, 4), (10, 16), (2, 3), (4, 5)]
RATIO_TOLERANCE = 0.01      # 1% counts as "the same" ratio

SLIDE_SIZE_NAMES = {(13.333, 7.5): "Widescreen", (10.0, 7.5): "Standard",
                    (10.0, 5.625): "On-screen Show", (10.0, 6.25): "On-screen Show"}


def emu_to_in(emu: int | float | None) -> float | None:
    return None if emu is None else emu / EMU_PER_INCH


def fmt_in(value: float | None) -> str | None:
    """13.3333 -> '13.33', 7.5 -> '7.5'."""
    if value is None:
        return None
    return f"{value:.2f}".rstrip("0").rstrip(".")


def aspect_label(width: float | None, height: float | None) -> str | None:
    """Human-readable aspect ratio, e.g. 1920x1080 -> '16:9', 2.39 -> '2.39:1'."""
    if not width or not height:
        return None
    ratio = width / height
    for a, b in NAMED_RATIOS:
        if abs(ratio - a / b) / (a / b) <= RATIO_TOLERANCE:
            return f"{a}:{b}"
    if float(width).is_integer() and float(height).is_integer():
        g = gcd(int(width), int(height))
        a, b = int(width) // g, int(height) // g
        if a <= 32 and b <= 32:
            return f"{a}:{b}"
    return f"{ratio:.2f}:1" if ratio >= 1 else f"1:{1 / ratio:.2f}"


def slide_size(cx: int | None, cy: int | None, size_type: str | None = None) -> dict | None:
    """Describe the presentation's slide size from <p:sldSz cx cy type>."""
    if not cx or not cy:
        return None
    w_in, h_in = emu_to_in(cx), emu_to_in(cy)
    name = None
    for (w, h), label in SLIDE_SIZE_NAMES.items():
        if abs(w_in - w) < 0.01 and abs(h_in - h) < 0.01:
            name = label
    w_px, h_px = round(w_in * PX_PER_INCH), round(h_in * PX_PER_INCH)
    aspect = aspect_label(w_in, h_in)
    return {
        "width_in": round(w_in, 3), "height_in": round(h_in, 3),
        "width_px": w_px, "height_px": h_px,
        "aspect_ratio": aspect, "name": name or (size_type if size_type and size_type != "custom" else "Custom"),
        "text": f"{fmt_in(w_in)} × {fmt_in(h_in)} in ({w_px} × {h_px} px) · {aspect}",
    }

