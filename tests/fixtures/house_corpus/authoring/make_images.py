"""Generate the small synthetic PNG assets embedded in the house decks.

Three deterministic 480x360 illustrations drawn with Pillow in house
palette colors only, quantized to 8-bit palette PNGs so each asset stays
in the low single-digit KB range (corpus size budget < 8 MB total).

Assets land in ``tests/fixtures/house_corpus/authoring/assets/`` and are
committed so the decks can be rebuilt byte-for-byte without Pillow.
"""

from __future__ import annotations

import os

from PIL import Image, ImageDraw

import _bootstrap
from house_style import SCHEME_HEX

SIZE = (480, 360)  # 4:3, placed at 240x180pt in the sidebar image zone

ASSETS = ("exhibit_bars.png", "exhibit_blocks.png", "exhibit_wave.png")


def _rgb(token: str) -> tuple[int, int, int]:
    value = SCHEME_HEX[token]
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


def _canvas() -> tuple[Image.Image, ImageDraw.ImageDraw]:
    image = Image.new("RGB", SIZE, _rgb("lt1"))
    draw = ImageDraw.Draw(image)
    draw.rectangle([0, 0, SIZE[0] - 1, SIZE[1] - 1], outline=_rgb("lt2"),
                   width=4)
    return image, draw


def _save(image: Image.Image, path: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    image.convert("P", palette=Image.ADAPTIVE, colors=16).save(
        path, format="PNG", optimize=True)


def make_bars(path: str) -> None:
    """Bar-chart style exhibit in the four accent colors."""
    image, draw = _canvas()
    heights = (140, 200, 110, 250)
    tokens = ("accent1", "accent2", "accent3", "accent5")
    for i, (height, token) in enumerate(zip(heights, tokens)):
        x0 = 60 + i * 100
        draw.rectangle([x0, 300 - height, x0 + 60, 300], fill=_rgb(token))
    draw.line([40, 300, 440, 300], fill=_rgb("dk2"), width=4)
    _save(image, path)


def make_blocks(path: str) -> None:
    """Process-diagram style exhibit: three blocks with connectors."""
    image, draw = _canvas()
    for i, token in enumerate(("accent1", "accent3", "accent4")):
        x0 = 40 + i * 150
        draw.rounded_rectangle([x0, 130, x0 + 110, 230], radius=14,
                               fill=_rgb("lt2"), outline=_rgb(token),
                               width=6)
        if i < 2:
            draw.line([x0 + 110, 180, x0 + 150, 180], fill=_rgb("dk2"),
                      width=6)
    draw.rectangle([40, 60, 440, 90], fill=_rgb("dk2"))
    _save(image, path)


def make_wave(path: str) -> None:
    """Trend-line style exhibit: polyline with marker dots."""
    image, draw = _canvas()
    points = [(40, 260), (120, 200), (200, 230), (280, 140), (360, 170),
              (440, 90)]
    draw.line(points, fill=_rgb("accent1"), width=6)
    for x, y in points:
        draw.ellipse([x - 9, y - 9, x + 9, y + 9], fill=_rgb("accent2"))
    draw.line([40, 300, 440, 300], fill=_rgb("dk2"), width=4)
    _save(image, path)


def main() -> list[str]:
    builders = {"exhibit_bars.png": make_bars,
                "exhibit_blocks.png": make_blocks,
                "exhibit_wave.png": make_wave}
    written = []
    for filename in ASSETS:
        path = _bootstrap.asset_path(filename)
        builders[filename](path)
        written.append(path)
        print(f"Wrote {path} ({os.path.getsize(path)} bytes)")
    return written


if __name__ == "__main__":
    main()
