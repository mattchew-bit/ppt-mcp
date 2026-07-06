"""Visual comparison of rendered slide PNGs (pixelmatch recipe).

Ports the Playwright/jest visual-gate recipe via the pure-Python
``pixelmatch`` package (ISC-licensed port of mapbox/pixelmatch):

* per-pixel color-distance threshold 0.1 (YIQ space)
* anti-aliasing detection ON (detected AA pixels are not counted as diffs)
* gate at diff-ratio 0.5% (strict) / 1% (lenient)
* emits a diff PNG artifact (mismatches painted red)

Renderer provenance travels IN the PNG itself (no sidecar files): every PNG
this package renders carries a ``tEXt`` metadata key ``ppt-mcp-renderer``
(``powerpoint`` or ``libreoffice``). When both images carry tags and they
differ, the comparison is refused -- cross-renderer pixel diffs measure the
renderers, not the decks.

Simple metrics only: mean per-channel delta via Pillow ``ImageStat`` --
deliberately no SSIM (would drag in numpy/scipy/skimage, none of which are
allowed as dependencies).
"""

from __future__ import annotations

import os
from typing import Any, Dict, Optional

RENDERER_TAG_KEY = "ppt-mcp-renderer"

DEFAULT_PIXEL_THRESHOLD = 0.1
STRICT_DIFF_RATIO = 0.005    # 0.5% -- strict gate
LENIENT_DIFF_RATIO = 0.01    # 1%   -- lenient gate


# -------------------------------------------------------------- PNG tagging


def tag_png_renderer(png_path: str, renderer: str) -> None:
    """Embed (or replace) the ``ppt-mcp-renderer`` tEXt key in a PNG.

    Pixels are untouched; existing text chunks other than the renderer tag
    are preserved.
    """
    from PIL import Image
    from PIL.PngImagePlugin import PngInfo

    if not renderer or not isinstance(renderer, str):
        raise ValueError(f"renderer must be a non-empty string, got {renderer!r}")
    abs_path = os.path.abspath(png_path)
    if not os.path.isfile(abs_path):
        raise FileNotFoundError(f"PNG not found: {abs_path}")

    with Image.open(abs_path) as img:
        img.load()
        info = PngInfo()
        for key, value in getattr(img, "text", {}).items():
            if key != RENDERER_TAG_KEY and isinstance(value, str):
                info.add_text(key, value)
        info.add_text(RENDERER_TAG_KEY, renderer)
        img.save(abs_path, "PNG", pnginfo=info)


def read_renderer_tag(png_path: str) -> Optional[str]:
    """The ``ppt-mcp-renderer`` tag of a PNG, or ``None`` when untagged."""
    from PIL import Image

    abs_path = os.path.abspath(png_path)
    if not os.path.isfile(abs_path):
        raise FileNotFoundError(f"PNG not found: {abs_path}")
    with Image.open(abs_path) as img:
        value = getattr(img, "text", {}).get(RENDERER_TAG_KEY)
    return value if isinstance(value, str) else None


# ---------------------------------------------------------------- comparison


def _load_image(path: str):
    from PIL import Image

    abs_path = os.path.abspath(path)
    if not os.path.isfile(abs_path):
        raise FileNotFoundError(f"Image not found: {abs_path}")
    with Image.open(abs_path) as img:
        return img.convert("RGBA")


def _check_same_renderer(image_a: str, image_b: str) -> Dict[str, Optional[str]]:
    tag_a = read_renderer_tag(image_a)
    tag_b = read_renderer_tag(image_b)
    if tag_a and tag_b and tag_a != tag_b:
        raise ValueError(
            f"Refusing cross-renderer comparison: '{image_a}' was rendered "
            f"by '{tag_a}' but '{image_b}' by '{tag_b}'. Pixel diffs across "
            "renderers measure font substitution and rasterizer differences, "
            "not deck differences -- re-render both images with the same "
            "renderer."
        )
    return {"renderer_a": tag_a, "renderer_b": tag_b}


def _mean_channel_delta(img_a, img_b) -> Dict[str, float]:
    """Mean absolute per-channel delta (RGB), pure Pillow -- no numpy."""
    from PIL import ImageChops, ImageStat

    diff = ImageChops.difference(img_a.convert("RGB"), img_b.convert("RGB"))
    means = ImageStat.Stat(diff).mean  # [r, g, b]
    r, g, b = (round(value, 4) for value in means)
    return {"r": r, "g": g, "b": b,
            "overall": round((r + g + b) / 3.0, 4)}


def _default_diff_path(image_a: str, image_b: str) -> str:
    from pathlib import Path

    a = Path(os.path.abspath(image_a))
    b_stem = Path(image_b).stem
    return str(a.parent / f"{a.stem}__vs__{b_stem}__diff.png")


def compare_renders(image_a: str,
                    image_b: str,
                    threshold: float = DEFAULT_PIXEL_THRESHOLD,
                    diff_path: Optional[str] = None) -> Dict[str, Any]:
    """Pixel-compare two same-renderer PNGs; returns metrics + a diff PNG.

    Args:
        image_a: First PNG (typically the reference render).
        image_b: Second PNG (typically the candidate render).
        threshold: Per-pixel color-distance threshold in [0, 1]
            (pixelmatch YIQ metric; 0.1 = Playwright default).
        diff_path: Where to write the diff artifact (default: next to
            ``image_a``).

    Returns:
        dict with ``diff_ratio``, ``diff_pixel_count``, ``dimensions``,
        ``verdict`` (``pass`` <=0.5%, ``borderline`` <=1%, ``fail``),
        ``passes_strict``/``passes_lenient``, ``diff_png_path``,
        ``mean_channel_delta`` and the renderer tags.

    Raises:
        ValueError: dimension mismatch, cross-renderer tags, bad threshold.
        FileNotFoundError: either image missing.
    """
    from PIL import Image

    from pixelmatch.contrib.PIL import pixelmatch

    if not isinstance(threshold, (int, float)) or not 0.0 <= threshold <= 1.0:
        raise ValueError(
            f"threshold must be a number in [0, 1], got {threshold!r}")

    img_a = _load_image(image_a)
    img_b = _load_image(image_b)
    if img_a.size != img_b.size:
        raise ValueError(
            f"Image dimensions differ: '{image_a}' is "
            f"{img_a.size[0]}x{img_a.size[1]} but '{image_b}' is "
            f"{img_b.size[0]}x{img_b.size[1]}. compare_renders requires "
            "equal dimensions -- re-render both decks at the same width."
        )
    renderers = _check_same_renderer(image_a, image_b)

    width, height = img_a.size
    diff_img = Image.new("RGBA", img_a.size)
    # includeAA=False == anti-aliasing detection ON (AA pixels not counted).
    diff_pixel_count = pixelmatch(
        img_a, img_b, output=diff_img,
        threshold=float(threshold), includeAA=False,
    )
    diff_ratio = diff_pixel_count / float(width * height)

    resolved_diff_path = os.path.abspath(
        diff_path if diff_path else _default_diff_path(image_a, image_b))
    os.makedirs(os.path.dirname(resolved_diff_path), exist_ok=True)
    diff_img.save(resolved_diff_path, "PNG")

    return _build_result(
        diff_ratio, diff_pixel_count, (width, height), float(threshold),
        resolved_diff_path, _mean_channel_delta(img_a, img_b), renderers)


def _build_result(diff_ratio, diff_pixel_count, size, threshold,
                  diff_png_path, mean_channel_delta, renderers) -> Dict[str, Any]:
    """Assemble the compare_renders result envelope (verdict included)."""
    passes_strict = diff_ratio <= STRICT_DIFF_RATIO
    passes_lenient = diff_ratio <= LENIENT_DIFF_RATIO
    verdict = ("pass" if passes_strict
               else "borderline" if passes_lenient
               else "fail")
    return {
        "diff_ratio": round(diff_ratio, 6),
        "diff_pixel_count": diff_pixel_count,
        "dimensions": {"width": size[0], "height": size[1]},
        "verdict": verdict,
        "passes_strict": passes_strict,
        "passes_lenient": passes_lenient,
        "thresholds": {
            "per_pixel": threshold,
            "strict_ratio": STRICT_DIFF_RATIO,
            "lenient_ratio": LENIENT_DIFF_RATIO,
        },
        "diff_png_path": diff_png_path,
        "mean_channel_delta": mean_channel_delta,
        "renderer_a": renderers["renderer_a"],
        "renderer_b": renderers["renderer_b"],
    }
