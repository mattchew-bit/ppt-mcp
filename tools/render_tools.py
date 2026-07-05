"""Slide rendering and visual-compare tools for PowerPoint MCP Server.

Renders slides to PNG via desktop PowerPoint COM (full fidelity; Windows +
pip install 'ppt-mcp[render]') with an automatic LibreOffice fallback
(approximate fidelity) when COM is unavailable. Comparison uses the
pixelmatch recipe (see utils/render_compare.py) and works on any platform.

Render tools return the PNG path(s) in a JSON envelope AND inline MCP image
content (FastMCP ``Image``) so the model can *see* the rendered slides.
"""

from typing import Dict, List, Optional

from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp import Image as FastMCPImage
from mcp.types import ToolAnnotations

from utils.render_com import (
    DEFAULT_RENDER_WIDTH,
    RenderCapabilityError,
    validate_render_source,
    validate_render_width,
)

#: Cap on inline image blocks per render_deck response -- every rendered
#: slide is still returned as a path; only the inline previews are capped
#: (base64 image content is expensive in model context).
MAX_INLINE_IMAGES = 8


def _select_renderer() -> str:
    """``"powerpoint"`` when COM is available, else ``"libreoffice"``.

    Raises a combined :class:`RenderCapabilityError` when neither renderer
    can run on this machine.
    """
    from utils import render_com, render_lo

    try:
        render_com.ensure_com_capability()
        return "powerpoint"
    except RenderCapabilityError as com_error:
        try:
            render_lo.ensure_lo_capability()
            return "libreoffice"
        except RenderCapabilityError as lo_error:
            raise RenderCapabilityError(
                f"No renderer is available on this machine. "
                f"PowerPoint COM: {com_error} "
                f"LibreOffice fallback: {lo_error}"
            ) from lo_error


def _run_renderer(file_path: str, width: int,
                  slide_range: Optional[str]) -> Dict:
    """Dispatch to the selected renderer; shared by both render tools."""
    renderer = _select_renderer()
    if renderer == "powerpoint":
        from utils.render_com import render_slides

        return render_slides(file_path, width=width, slide_range=slide_range)
    from utils.render_lo import render_slides_lo

    return render_slides_lo(file_path, width=width, slide_range=slide_range)


def _validated_render_inputs(file_path: str, width: int) -> Optional[Dict]:
    """Boundary validation shared by the render tools; error envelope or None."""
    try:
        validate_render_source(file_path)
        validate_render_width(width)
    except (FileNotFoundError, ValueError) as exc:
        return {"error": str(exc)}
    return None


def register_render_tools(app: FastMCP):
    """Register rendering and visual-compare tools with the FastMCP app."""

    @app.tool(
        annotations=ToolAnnotations(
            title="Render Slide to PNG",
            readOnlyHint=True,
        ),
    )
    def render_slide(
        file_path: str,
        slide_index: int,
        width: int = DEFAULT_RENDER_WIDTH,
    ):
        """Render one slide of a .pptx file to PNG and return it inline.

        Uses desktop PowerPoint via COM when available (full fidelity;
        requires Windows + pip install 'ppt-mcp[render]'), otherwise falls
        back to LibreOffice (approximate fidelity). The deck is rendered
        from a temp copy, so it may be open in PowerPoint at the same time.
        Password-protected decks are refused. Height follows the slide
        aspect ratio.

        Args:
            file_path: Path to the .pptx file on disk
            slide_index: Index of the slide to render (0-based)
            width: Output width in pixels, 32-4096 (default: 1280)
        """
        if slide_index < 0:
            return {"error": f"slide_index must be >= 0, got {slide_index}"}
        error = _validated_render_inputs(file_path, width)
        if error is not None:
            return error
        try:
            result = _run_renderer(file_path, width,
                                   slide_range=str(slide_index + 1))
        except RenderCapabilityError as exc:
            return {"error": str(exc)}
        except ValueError as exc:
            return {"error": str(exc)}
        except Exception as exc:
            return {"error": f"Failed to render slide: {exc}"}

        image_path = result["paths"][0]
        envelope = {
            "message": (
                f"Rendered slide {slide_index} of '{file_path}' at "
                f"{result['width']}x{result['height']} via {result['renderer']}"
            ),
            "file_path": file_path,
            "slide_index": slide_index,
            "image_path": image_path,
            "width": result["width"],
            "height": result["height"],
            "renderer": result["renderer"],
        }
        return [envelope, FastMCPImage(path=image_path)]

    @app.tool(
        annotations=ToolAnnotations(
            title="Render Deck to PNGs",
            readOnlyHint=True,
        ),
    )
    def render_deck(
        file_path: str,
        width: int = DEFAULT_RENDER_WIDTH,
        slide_range: Optional[str] = None,
    ):
        """Render slides of a .pptx file to PNGs (one file per slide).

        Same renderer selection and temp-copy behavior as render_slide.
        All rendered PNG paths are returned; the first few slides are also
        returned inline as image content (capped to keep responses small).

        Args:
            file_path: Path to the .pptx file on disk
            slide_range: 1-based selection like "1-3,5" (default: all slides)
            width: Output width in pixels, 32-4096 (default: 1280)
        """
        error = _validated_render_inputs(file_path, width)
        if error is not None:
            return error
        try:
            result = _run_renderer(file_path, width, slide_range=slide_range)
        except RenderCapabilityError as exc:
            return {"error": str(exc)}
        except ValueError as exc:
            return {"error": str(exc)}
        except Exception as exc:
            return {"error": f"Failed to render deck: {exc}"}

        paths: List[str] = result["paths"]
        inline_count = min(len(paths), MAX_INLINE_IMAGES)
        envelope = {
            "message": (
                f"Rendered {len(paths)} slide(s) of '{file_path}' at "
                f"{result['width']}x{result['height']} via {result['renderer']}"
                + (f"; inlining first {inline_count} of {len(paths)} images"
                   if inline_count < len(paths) else "")
            ),
            "file_path": file_path,
            "image_paths": paths,
            "rendered_slides": len(paths),
            "deck_slide_count": result["slide_count"],
            "width": result["width"],
            "height": result["height"],
            "renderer": result["renderer"],
            "inline_images": inline_count,
        }
        return [envelope] + [FastMCPImage(path=p) for p in paths[:inline_count]]

    @app.tool(
        annotations=ToolAnnotations(
            title="Compare Rendered Slides",
            readOnlyHint=True,
        ),
    )
    def compare_renders(
        image_a: str,
        image_b: str,
        threshold: float = 0.1,
    ):
        """Pixel-compare two rendered slide PNGs (pixelmatch, AA-aware).

        Returns diff metrics (diff_ratio, diff_pixel_count, verdict at the
        strict 0.5% / lenient 1% gates, mean channel delta) plus a diff PNG
        (mismatched pixels in red) inline. Both images must have identical
        dimensions and come from the same renderer (PNG tag enforced) --
        cross-renderer diffs measure the renderer, not the deck.

        Args:
            image_a: Path to the first PNG (e.g. the reference render)
            image_b: Path to the second PNG (e.g. the candidate render)
            threshold: Per-pixel color-distance threshold 0-1 (default 0.1)
        """
        from utils.render_compare import compare_renders as _compare

        try:
            result = _compare(image_a, image_b, threshold=threshold)
        except (FileNotFoundError, ValueError) as exc:
            return {"error": str(exc)}
        except Exception as exc:
            return {"error": f"Failed to compare renders: {exc}"}

        envelope = dict(result)
        envelope["message"] = (
            f"Compared '{image_a}' vs '{image_b}': "
            f"{result['diff_pixel_count']} differing pixels "
            f"({result['diff_ratio'] * 100:.3f}%) -> verdict "
            f"'{result['verdict']}'"
        )
        return [envelope, FastMCPImage(path=result["diff_png_path"])]
