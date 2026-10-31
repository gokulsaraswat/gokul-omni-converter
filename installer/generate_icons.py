"""Rasterise the GO Omni Pro mark into the icon files the installers need.

The vector master is ``assets/logo/go_omni_pro_mark.svg``. Because the app ships with
Pillow but no SVG rasteriser, this script redraws that same mark with Pillow so icon
generation works on Windows, Linux, and macOS with zero extra dependencies. The geometry
constants below are the coordinates from the SVG, in its 256x256 user space -- keep the
two in sync if the mark is ever redesigned.

Run from the project root::

    python installer/generate_icons.py

Outputs (all under ``assets/logo/``):
    go_omni_pro.ico        Windows app + installer icon (16/32/48/64/128/256)
    go_omni_pro_256.png    Linux .desktop icon, About page, window icon
    go_omni_pro_48.png     small Linux icon theme size
    go_omni_pro_512.png    hi-dpi About page / macOS
"""

from __future__ import annotations

import math
import sys
from pathlib import Path

from PIL import Image, ImageDraw

# --- geometry, in the SVG's 256x256 user space -------------------------------------
CANVAS = 256
SUPERSAMPLE = 8

BADGE_BOX = (8, 8, 248, 248)
BADGE_RADIUS = 56
BADGE_STOPS = ((0.0, "#4f46e5"), (0.55, "#6d28d9"), (1.0, "#8b5cf6"))

RING_CENTER = (128.0, 124.0)
RING_RADIUS = 76.0
RING_WIDTH = 15.0
RING_START_DEG = -25.0
RING_SWEEP_DEG = 310.0
RING_BRIGHT_DEG = 120.0  # bright tail, ending at the arrow head

ARROW = ((169.0, 56.3), (137.5, 65.4), (146.3, 32.6))

SHEET_BOX = (86, 74, 170, 174)
SHEET_RADIUS = 14
SHEET_FOLD = 28  # size of the folded corner
SHEET_STOPS = ((0.0, "#ffffff"), (1.0, "#ede9fe"))
FOLD_FILL = "#c4b5fd"

TEXT_LINES = (((104, 118), (152, 118)), ((104, 138), (152, 138)), ((104, 158), (132, 158)))
TEXT_WIDTH = 8
TEXT_FILL = (124, 58, 237, 128)  # #7c3aed at 50%

ICO_SIZES = (16, 32, 48, 64, 128, 256)
PNG_SIZES = (48, 256, 512)


def _hex_to_rgb(value: str) -> tuple[int, int, int]:
    value = value.lstrip("#")
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))  # type: ignore[return-value]


def _sample_stops(stops: tuple[tuple[float, str], ...], t: float) -> tuple[int, int, int]:
    """Linear interpolation across an ordered list of (offset, hex) colour stops."""
    t = min(1.0, max(0.0, t))
    for index in range(len(stops) - 1):
        left_pos, left_hex = stops[index]
        right_pos, right_hex = stops[index + 1]
        if left_pos <= t <= right_pos:
            span = right_pos - left_pos
            ratio = 0.0 if span <= 0 else (t - left_pos) / span
            left = _hex_to_rgb(left_hex)
            right = _hex_to_rgb(right_hex)
            return tuple(round(left[c] + (right[c] - left[c]) * ratio) for c in range(3))  # type: ignore[return-value]
    return _hex_to_rgb(stops[-1][1])


def _gradient(size: int, stops: tuple[tuple[float, str], ...], diagonal: bool) -> Image.Image:
    """Build the gradient small and upscale -- far cheaper than per-pixel at 2048px."""
    small = 64
    image = Image.new("RGB", (small, small))
    pixels = image.load()
    for y in range(small):
        for x in range(small):
            t = (x + y) / (2 * (small - 1)) if diagonal else y / (small - 1)
            pixels[x, y] = _sample_stops(stops, t)
    return image.resize((size, size), Image.BICUBIC)


def _masked(size: int, stops: tuple[tuple[float, str], ...], mask: Image.Image, diagonal: bool) -> Image.Image:
    layer = _gradient(size, stops, diagonal).convert("RGBA")
    layer.putalpha(mask)
    return layer


def _arc_box(scale: float) -> list[float]:
    """Pillow strokes arcs *inward* from the bounding box, so inflate by half the stroke
    to put the centreline on RING_RADIUS the way the SVG's stroke does."""
    cx, cy = RING_CENTER
    outer = RING_RADIUS + RING_WIDTH / 2
    return [(cx - outer) * scale, (cy - outer) * scale, (cx + outer) * scale, (cy + outer) * scale]


def render(size: int) -> Image.Image:
    """Draw the mark at ``size`` pixels, supersampled then downscaled for smooth edges."""
    work = size * SUPERSAMPLE
    scale = work / CANVAS
    canvas = Image.new("RGBA", (work, work), (0, 0, 0, 0))

    def s(value: float) -> float:
        return value * scale

    # badge: gradient clipped to a rounded rectangle
    badge_mask = Image.new("L", (work, work), 0)
    ImageDraw.Draw(badge_mask).rounded_rectangle(
        [s(BADGE_BOX[0]), s(BADGE_BOX[1]), s(BADGE_BOX[2]), s(BADGE_BOX[3])],
        radius=s(BADGE_RADIUS),
        fill=255,
    )
    canvas.alpha_composite(_masked(work, BADGE_STOPS, badge_mask, diagonal=True))

    # ImageDraw *replaces* pixels rather than blending, so anything translucent has to be
    # drawn on its own layer and alpha-composited back onto the badge.
    def overlay(paint) -> None:
        layer = Image.new("RGBA", (work, work), (0, 0, 0, 0))
        paint(ImageDraw.Draw(layer))
        canvas.alpha_composite(layer)

    overlay(
        lambda d: d.rounded_rectangle(
            [s(9.5), s(9.5), s(246.5), s(246.5)],
            radius=s(54.5),
            outline=(255, 255, 255, 51),
            width=max(1, round(s(3))),
        )
    )

    ring_box = _arc_box(scale)
    ring_width = max(1, round(s(RING_WIDTH)))
    end_deg = RING_START_DEG + RING_SWEEP_DEG

    def _cap(d, angle: float, alpha: int) -> None:
        """Pillow arcs have flat ends; add discs to match the SVG's round linecap."""
        cx = RING_CENTER[0] + RING_RADIUS * math.cos(math.radians(angle))
        cy = RING_CENTER[1] + RING_RADIUS * math.sin(math.radians(angle))
        d.ellipse(
            [s(cx) - ring_width / 2, s(cy) - ring_width / 2, s(cx) + ring_width / 2, s(cy) + ring_width / 2],
            fill=(255, 255, 255, alpha),
        )

    def _faint_ring(d) -> None:
        d.arc(ring_box, RING_START_DEG, end_deg, fill=(255, 255, 255, 77), width=ring_width)
        _cap(d, RING_START_DEG, 77)
        _cap(d, end_deg, 77)

    def _bright_ring(d) -> None:
        d.arc(ring_box, end_deg - RING_BRIGHT_DEG, end_deg, fill=(255, 255, 255, 255), width=ring_width)
        _cap(d, end_deg - RING_BRIGHT_DEG, 255)
        d.polygon([(s(x), s(y)) for x, y in ARROW], fill=(255, 255, 255, 255))

    overlay(_faint_ring)
    overlay(_bright_ring)

    # document sheet: rounded rect, top-right corner cut away, gradient fill
    sheet_mask = Image.new("L", (work, work), 0)
    sheet_draw = ImageDraw.Draw(sheet_mask)
    left, top, right, bottom = SHEET_BOX
    sheet_draw.rounded_rectangle(
        [s(left), s(top), s(right), s(bottom)], radius=s(SHEET_RADIUS), fill=255
    )
    sheet_draw.polygon(
        [
            (s(right - SHEET_FOLD), s(top) - 1),
            (s(right) + 1, s(top) - 1),
            (s(right) + 1, s(top + SHEET_FOLD)),
        ],
        fill=0,
    )
    canvas.alpha_composite(_masked(work, SHEET_STOPS, sheet_mask, diagonal=True))

    overlay(
        lambda d: d.polygon(
            [
                (s(right - SHEET_FOLD), s(top)),
                (s(right), s(top + SHEET_FOLD)),
                (s(right - SHEET_FOLD), s(top + SHEET_FOLD)),
            ],
            fill=_hex_to_rgb(FOLD_FILL) + (255,),
        )
    )

    line_width = max(1, round(s(TEXT_WIDTH)))

    def _text_lines(d) -> None:
        for (x1, y1), (x2, y2) in TEXT_LINES:
            d.line([s(x1), s(y1), s(x2), s(y2)], fill=TEXT_FILL, width=line_width)
            for x in (x1, x2):  # round caps
                d.ellipse(
                    [s(x) - line_width / 2, s(y1) - line_width / 2, s(x) + line_width / 2, s(y1) + line_width / 2],
                    fill=TEXT_FILL,
                )

    overlay(_text_lines)
    return canvas.resize((size, size), Image.LANCZOS)


def main() -> int:
    out_dir = Path(__file__).resolve().parent.parent / "assets" / "logo"
    out_dir.mkdir(parents=True, exist_ok=True)

    master = render(1024)
    ico_path = out_dir / "go_omni_pro.ico"
    master.save(ico_path, format="ICO", sizes=[(n, n) for n in ICO_SIZES])
    print(f"wrote {ico_path.relative_to(out_dir.parent.parent)}")

    for size in PNG_SIZES:
        png_path = out_dir / f"go_omni_pro_{size}.png"
        render(size).save(png_path, format="PNG")
        print(f"wrote {png_path.relative_to(out_dir.parent.parent)}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
