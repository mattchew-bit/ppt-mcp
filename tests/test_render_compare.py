"""Tests for ``utils/render_compare.py`` -- pixelmatch-based visual compare.

Non-COM: every test here runs on any platform (Pillow + pixelmatch are core
dependencies). Images are generated with Pillow on the fly; no rendered
fixtures are required.
"""

import pytest
from PIL import Image, ImageDraw

from utils.render_compare import (
    LENIENT_DIFF_RATIO,
    RENDERER_TAG_KEY,
    STRICT_DIFF_RATIO,
    compare_renders,
    read_renderer_tag,
    tag_png_renderer,
)


# ---------------------------------------------------------------- helpers


def _save(img, path):
    img.save(str(path), "PNG")
    return str(path)


def _flat_image(size=(200, 150), color=(240, 240, 240)):
    return Image.new("RGB", size, color)


def _image_with_rect(size=(200, 150), color=(240, 240, 240),
                     rect=(40, 40, 120, 100), rect_color=(192, 80, 77)):
    img = Image.new("RGB", size, color)
    draw = ImageDraw.Draw(img)
    draw.rectangle(rect, fill=rect_color)
    return img


def _aliased_circle(size=(400, 400), radius=80):
    """A hard-edged (aliased) circle."""
    img = Image.new("RGB", size, (255, 255, 255))
    draw = ImageDraw.Draw(img)
    cx, cy = size[0] // 2, size[1] // 2
    draw.ellipse((cx - radius, cy - radius, cx + radius, cy + radius),
                 fill=(31, 78, 121))
    return img


def _antialiased_circle(size=(400, 400), radius=80):
    """The same circle drawn at 4x and downscaled -> antialiased edges."""
    big = _aliased_circle((size[0] * 4, size[1] * 4), radius * 4)
    return big.resize(size, Image.LANCZOS)


# ------------------------------------------------------------- identical


def test_identical_images_zero_diff(tmp_path):
    img = _image_with_rect()
    a = _save(img, tmp_path / "a.png")
    b = _save(img.copy(), tmp_path / "b.png")

    result = compare_renders(a, b)

    assert result["diff_pixel_count"] == 0
    assert result["diff_ratio"] == 0.0
    assert result["verdict"] == "pass"
    assert result["passes_strict"] is True
    assert result["passes_lenient"] is True
    assert result["dimensions"] == {"width": 200, "height": 150}


def test_identical_images_diff_png_written(tmp_path):
    import os

    img = _flat_image()
    a = _save(img, tmp_path / "a.png")
    b = _save(img.copy(), tmp_path / "b.png")

    result = compare_renders(a, b)

    assert os.path.isfile(result["diff_png_path"])


def test_explicit_diff_path_respected(tmp_path):
    import os

    img = _flat_image()
    a = _save(img, tmp_path / "a.png")
    b = _save(img.copy(), tmp_path / "b.png")
    diff = str(tmp_path / "out" / "custom_diff.png")

    result = compare_renders(a, b, diff_path=diff)

    assert result["diff_png_path"] == os.path.abspath(diff)
    assert os.path.isfile(diff)


# ------------------------------------------------------- seeded difference


def test_one_region_change_detected(tmp_path):
    a = _save(_flat_image(), tmp_path / "a.png")
    b = _save(_image_with_rect(), tmp_path / "b.png")

    result = compare_renders(a, b)

    assert result["diff_pixel_count"] > 0
    assert result["diff_ratio"] > 0.0
    # An 80x60 block on 200x150 is 16% of pixels -> well past lenient.
    assert result["verdict"] == "fail"
    assert result["passes_strict"] is False
    assert result["passes_lenient"] is False


def test_diff_png_marks_changed_region(tmp_path):
    a = _save(_flat_image(), tmp_path / "a.png")
    b = _save(_image_with_rect(), tmp_path / "b.png")

    result = compare_renders(a, b)

    diff_img = Image.open(result["diff_png_path"]).convert("RGB")
    # pixelmatch paints mismatching pixels red by default.
    assert diff_img.getpixel((80, 70)) == (255, 0, 0)


def test_mean_channel_delta_reported(tmp_path):
    a = _save(_flat_image(color=(100, 100, 100)), tmp_path / "a.png")
    b = _save(_flat_image(color=(110, 100, 100)), tmp_path / "b.png")

    result = compare_renders(a, b)

    delta = result["mean_channel_delta"]
    assert delta["r"] == pytest.approx(10.0, abs=0.5)
    assert delta["g"] == pytest.approx(0.0, abs=0.5)
    assert delta["b"] == pytest.approx(0.0, abs=0.5)
    assert delta["overall"] == pytest.approx(10.0 / 3.0, abs=0.5)


# ------------------------------------------------------------ AA robustness


def test_antialiased_edges_stay_under_strict(tmp_path):
    a = _save(_aliased_circle(), tmp_path / "a.png")
    b = _save(_antialiased_circle(), tmp_path / "b.png")

    result = compare_renders(a, b)

    assert result["diff_ratio"] <= STRICT_DIFF_RATIO
    assert result["verdict"] == "pass"


# ------------------------------------------------------- validation errors


def test_dimension_mismatch_raises(tmp_path):
    a = _save(_flat_image(size=(200, 150)), tmp_path / "a.png")
    b = _save(_flat_image(size=(100, 150)), tmp_path / "b.png")

    with pytest.raises(ValueError) as excinfo:
        compare_renders(a, b)

    message = str(excinfo.value)
    assert "200x150" in message
    assert "100x150" in message


def test_missing_file_raises(tmp_path):
    a = _save(_flat_image(), tmp_path / "a.png")

    with pytest.raises(FileNotFoundError):
        compare_renders(a, str(tmp_path / "nope.png"))


def test_bad_threshold_rejected(tmp_path):
    img = _flat_image()
    a = _save(img, tmp_path / "a.png")
    b = _save(img.copy(), tmp_path / "b.png")

    with pytest.raises(ValueError):
        compare_renders(a, b, threshold=1.5)


# ----------------------------------------------------------- renderer tags


def test_tag_roundtrip(tmp_path):
    path = _save(_flat_image(), tmp_path / "a.png")

    assert read_renderer_tag(path) is None
    tag_png_renderer(path, "powerpoint")
    assert read_renderer_tag(path) == "powerpoint"

    # Re-tagging replaces, not duplicates.
    tag_png_renderer(path, "libreoffice")
    assert read_renderer_tag(path) == "libreoffice"


def test_tag_preserves_pixels(tmp_path):
    img = _image_with_rect()
    path = _save(img, tmp_path / "a.png")

    tag_png_renderer(path, "powerpoint")

    reread = Image.open(path).convert("RGB")
    assert reread.tobytes() == img.tobytes()


def test_cross_renderer_comparison_refused(tmp_path):
    img = _flat_image()
    a = _save(img, tmp_path / "a.png")
    b = _save(img.copy(), tmp_path / "b.png")
    tag_png_renderer(a, "powerpoint")
    tag_png_renderer(b, "libreoffice")

    with pytest.raises(ValueError) as excinfo:
        compare_renders(a, b)

    message = str(excinfo.value)
    assert "powerpoint" in message
    assert "libreoffice" in message


def test_same_renderer_tags_allowed(tmp_path):
    img = _flat_image()
    a = _save(img, tmp_path / "a.png")
    b = _save(img.copy(), tmp_path / "b.png")
    tag_png_renderer(a, "powerpoint")
    tag_png_renderer(b, "powerpoint")

    result = compare_renders(a, b)

    assert result["diff_pixel_count"] == 0
    assert result["renderer_a"] == "powerpoint"
    assert result["renderer_b"] == "powerpoint"


def test_single_tagged_image_allowed(tmp_path):
    img = _flat_image()
    a = _save(img, tmp_path / "a.png")
    b = _save(img.copy(), tmp_path / "b.png")
    tag_png_renderer(a, "powerpoint")

    result = compare_renders(a, b)

    assert result["renderer_a"] == "powerpoint"
    assert result["renderer_b"] is None


# ------------------------------------------------------------- constants


def test_threshold_constants():
    assert STRICT_DIFF_RATIO == 0.005
    assert LENIENT_DIFF_RATIO == 0.01
    assert RENDERER_TAG_KEY == "ppt-mcp-renderer"
