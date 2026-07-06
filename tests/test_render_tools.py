"""Tests for ``tools/render_tools.py`` -- the MCP-facing render/compare tools.

Non-COM: the actual renderers are monkeypatched; only boundary validation,
response envelopes, and inline image content are exercised here. COM-backed
end-to-end rendering lives in ``test_render_com.py`` (marked ``com``).
"""

import asyncio

import pytest
from PIL import Image as PILImage

from tools.render_tools import MAX_INLINE_IMAGES, register_render_tools


class _RecorderApp:
    """Minimal stand-in for FastMCP capturing the registered tool functions."""

    def __init__(self):
        self.tools = {}

    def tool(self, *args, **kwargs):
        def decorator(fn):
            self.tools[fn.__name__] = fn
            return fn

        return decorator


@pytest.fixture
def tools():
    app = _RecorderApp()
    register_render_tools(app)
    return app.tools


def _fake_render_result(paths, width=1280, height=720):
    return {
        "paths": list(paths),
        "width": width,
        "height": height,
        "slide_count": max(4, len(paths)),
        "renderer": "powerpoint",
    }


def _make_png(path, size=(64, 36)):
    PILImage.new("RGB", size, (200, 200, 200)).save(str(path), "PNG")
    return str(path)


def _make_deck_file(path):
    with open(str(path), "wb") as handle:
        handle.write(b"PK\x03\x04" + b"\x00" * 32)
    return str(path)


# ------------------------------------------------------------ registration


def test_tools_register_on_real_fastmcp_app():
    from mcp.server.fastmcp import FastMCP

    app = FastMCP(name="render-tools-test")
    register_render_tools(app)
    tool_names = {tool.name for tool in asyncio.run(app.list_tools())}
    assert {"render_slide", "render_deck", "compare_renders"} <= tool_names


# ------------------------------------------------------ boundary validation


def test_render_slide_missing_file(tools):
    result = tools["render_slide"](file_path="Z:/nope/absent.pptx", slide_index=0)
    assert "error" in result
    assert "absent.pptx" in result["error"]


def test_render_slide_negative_index(tools, tmp_path):
    deck = _make_deck_file(tmp_path / "deck.pptx")
    result = tools["render_slide"](file_path=deck, slide_index=-1)
    assert "error" in result
    assert "slide_index" in result["error"]


def test_render_slide_bad_width(tools, tmp_path):
    deck = _make_deck_file(tmp_path / "deck.pptx")
    result = tools["render_slide"](file_path=deck, slide_index=0, width=10)
    assert "error" in result
    assert "width" in result["error"]


def test_render_slide_encrypted_deck(tools, tmp_path):
    locked = tmp_path / "locked.pptx"
    with open(str(locked), "wb") as handle:
        handle.write(b"\xd0\xcf\x11\xe0" + b"\x00" * 32)

    result = tools["render_slide"](file_path=str(locked), slide_index=0)

    assert "error" in result
    assert "password" in result["error"].lower()


def test_render_deck_missing_file(tools):
    result = tools["render_deck"](file_path="Z:/nope/absent.pptx")
    assert "error" in result


def test_render_slide_zero_byte_deck(tools, tmp_path):
    """Regression: a zero-byte .pptx used to reach COM, where render_deck
    'succeeded' with 0 slides and render_slide errored nonsensically."""
    deck = tmp_path / "empty.pptx"
    deck.write_bytes(b"")

    result = tools["render_slide"](file_path=str(deck), slide_index=0)

    assert "error" in result
    assert "empty" in result["error"].lower()


def test_render_deck_zero_byte_deck_is_not_a_silent_success(tools, tmp_path):
    deck = tmp_path / "empty.pptx"
    deck.write_bytes(b"")

    result = tools["render_deck"](file_path=str(deck))

    assert "error" in result
    assert "empty" in result["error"].lower()


def test_render_deck_garbage_bytes_actionable_error(tools, tmp_path):
    """Regression: garbage input surfaced as a raw com_error hex tuple."""
    deck = tmp_path / "garbage.pptx"
    deck.write_bytes(b"this is not a presentation" * 8)

    result = tools["render_deck"](file_path=str(deck))

    assert "error" in result
    assert "corrupt" in result["error"].lower()
    assert "-2147352567" not in result["error"]


def test_compare_renders_missing_file(tools, tmp_path):
    a = _make_png(tmp_path / "a.png")
    result = tools["compare_renders"](image_a=a, image_b=str(tmp_path / "b.png"))
    assert "error" in result


def test_compare_renders_dimension_mismatch(tools, tmp_path):
    a = _make_png(tmp_path / "a.png", size=(64, 36))
    b = _make_png(tmp_path / "b.png", size=(32, 36))
    result = tools["compare_renders"](image_a=a, image_b=b)
    assert "error" in result


# --------------------------------------------------- envelopes + inline images


def test_render_slide_success_envelope(tools, tmp_path, monkeypatch):
    from mcp.server.fastmcp import Image as FastMCPImage

    from utils import render_com

    deck = _make_deck_file(tmp_path / "deck.pptx")
    png = _make_png(tmp_path / "slide_001.png")

    monkeypatch.setattr(
        "tools.render_tools._select_renderer", lambda: "powerpoint")
    monkeypatch.setattr(
        render_com, "render_slides",
        lambda *args, **kwargs: _fake_render_result([png]))

    result = tools["render_slide"](file_path=deck, slide_index=0)

    assert isinstance(result, list)
    envelope, image = result[0], result[1]
    assert envelope["image_path"] == png
    assert envelope["renderer"] == "powerpoint"
    assert envelope["width"] == 1280
    assert envelope["height"] == 720
    assert isinstance(image, FastMCPImage)


def test_render_deck_inline_image_cap(tools, tmp_path, monkeypatch):
    from mcp.server.fastmcp import Image as FastMCPImage

    from utils import render_com

    deck = _make_deck_file(tmp_path / "deck.pptx")
    pngs = [_make_png(tmp_path / f"slide_{i:03d}.png")
            for i in range(MAX_INLINE_IMAGES + 3)]

    monkeypatch.setattr(
        "tools.render_tools._select_renderer", lambda: "powerpoint")
    monkeypatch.setattr(
        render_com, "render_slides",
        lambda *args, **kwargs: _fake_render_result(pngs))

    result = tools["render_deck"](file_path=deck)

    envelope = result[0]
    inline_images = [item for item in result[1:]
                     if isinstance(item, FastMCPImage)]
    assert envelope["image_paths"] == pngs
    assert len(inline_images) == MAX_INLINE_IMAGES
    assert envelope["inline_images"] == MAX_INLINE_IMAGES


def test_compare_renders_success_envelope(tools, tmp_path):
    from mcp.server.fastmcp import Image as FastMCPImage

    img = PILImage.new("RGB", (100, 80), (220, 220, 220))
    a = tmp_path / "a.png"
    b = tmp_path / "b.png"
    img.save(str(a), "PNG")
    img.save(str(b), "PNG")

    result = tools["compare_renders"](image_a=str(a), image_b=str(b))

    assert isinstance(result, list)
    envelope, diff_image = result[0], result[1]
    assert envelope["diff_pixel_count"] == 0
    assert envelope["verdict"] == "pass"
    assert isinstance(diff_image, FastMCPImage)


def test_capability_error_surfaces_in_envelope(tools, tmp_path, monkeypatch):
    import sys

    from utils import render_lo

    deck = _make_deck_file(tmp_path / "deck.pptx")
    for name in ("win32com", "win32com.client", "pythoncom", "win32process"):
        monkeypatch.setitem(sys.modules, name, None)
    monkeypatch.setattr(render_lo, "find_soffice", lambda: None)

    result = tools["render_slide"](file_path=deck, slide_index=0)

    assert "error" in result
    assert "ppt-mcp[render]" in result["error"]
