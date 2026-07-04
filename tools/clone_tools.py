"""
Slide cloning tools for PowerPoint MCP Server.
Handles same-deck slide duplication and cross-deck slide copying with
full relationship rewriting (see utils/clone_utils.py).
"""
from typing import Dict, Optional, Tuple
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations


def _slide_index_error(pres, slide_index: int) -> Optional[Dict]:
    """Error envelope for an out-of-range slide index, or None when valid."""
    if slide_index < 0 or slide_index >= len(pres.slides):
        return {
            "error": f"Invalid slide index: {slide_index}. "
                     f"Available slides: 0-{len(pres.slides) - 1}"
        }
    return None


def _resolve_copy_decks(
    presentations: Dict,
    source_presentation_id: str,
    destination_id: Optional[str],
) -> Tuple[Optional[object], Optional[object], Optional[Dict]]:
    """Resolve source and destination decks for copy_slide.

    Returns ``(src_pres, dst_pres, None)`` on success, or
    ``(None, None, error_envelope)`` when either deck cannot be resolved.
    """
    if source_presentation_id not in presentations:
        return None, None, {
            "error": f"Source presentation '{source_presentation_id}' not found. "
                     f"Available presentations: {list(presentations.keys())}"
        }
    if destination_id is None or destination_id not in presentations:
        return None, None, {
            "error": "No destination presentation is currently loaded or the specified ID is invalid"
        }
    return presentations[source_presentation_id], presentations[destination_id], None


def _copy_slide_response(
    src_pres, dst_pres, source_presentation_id: str, dst_id: str, slide_index: int
) -> Dict:
    """Run the cross-deck copy and build the tool response envelope."""
    from utils.clone_utils import copy_slide as _copy_slide

    try:
        new_slide = _copy_slide(src_pres, slide_index, dst_pres)
        new_slide_index = len(dst_pres.slides) - 1
        return {
            "message": (
                f"Copied slide {slide_index} from '{source_presentation_id}' "
                f"to '{dst_id}' at index {new_slide_index}"
            ),
            "source_presentation_id": source_presentation_id,
            "source_slide_index": slide_index,
            "destination_presentation_id": dst_id,
            "new_slide_index": new_slide_index,
            "layout_name": new_slide.slide_layout.name,
        }
    except NotImplementedError as e:
        return {"error": str(e)}
    except ValueError as e:
        return {"error": str(e)}
    except Exception as e:
        return {"error": f"Failed to copy slide: {str(e)}"}


def register_clone_tools(app: FastMCP, presentations: Dict, get_current_presentation_id):
    """Register slide cloning tools with the FastMCP app"""

    @app.tool(
        annotations=ToolAnnotations(
            title="Duplicate Slide",
        ),
    )
    def duplicate_slide(
        slide_index: int,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """Duplicate a slide within the same presentation, preserving formatting.

        The duplicate is appended after the last slide. The slide XML is
        deep-copied while images, media, and the slide layout are shared with
        the source slide. Speaker notes are not copied. Slides containing
        charts, SmartArt, OLE objects, or ActiveX controls are not supported
        in v1.

        Args:
            slide_index: Index of the slide to duplicate (0-based)
            presentation_id: ID of the presentation (default: current)
        """
        from utils.clone_utils import duplicate_slide as _duplicate_slide

        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        index_error = _slide_index_error(pres, slide_index)
        if index_error is not None:
            return index_error

        try:
            new_slide = _duplicate_slide(pres, slide_index)
            new_slide_index = len(pres.slides) - 1
            return {
                "message": f"Duplicated slide {slide_index}; duplicate appended at index {new_slide_index}",
                "source_slide_index": slide_index,
                "new_slide_index": new_slide_index,
                "layout_name": new_slide.slide_layout.name,
                "presentation_id": pres_id,
            }
        except NotImplementedError as e:
            return {"error": str(e)}
        except ValueError as e:
            return {"error": str(e)}
        except Exception as e:
            return {"error": f"Failed to duplicate slide: {str(e)}"}

    @app.tool(
        annotations=ToolAnnotations(
            title="Copy Slide Between Presentations",
        ),
    )
    def copy_slide(
        source_presentation_id: str,
        slide_index: int,
        destination_presentation_id: Optional[str] = None
    ) -> Dict:
        """Copy a slide from one loaded presentation into another, preserving formatting.

        The copied slide is appended to the destination and re-bound to the
        destination layout matched by layout name -- both decks must share
        template lineage (v1 constraint; start the destination from the same
        template as the source). Images and media are re-created in the
        destination (deduplicated by content hash), external hyperlinks are
        preserved, internal slide-jump hyperlinks are dropped (text kept),
        speaker notes are not copied. Slides containing charts, SmartArt,
        OLE objects, or ActiveX controls are not supported in v1.

        Args:
            source_presentation_id: ID of the presentation to copy from
            slide_index: Index of the slide to copy (0-based)
            destination_presentation_id: ID of the presentation to copy into (default: current)
        """
        dst_id = (
            destination_presentation_id
            if destination_presentation_id is not None
            else get_current_presentation_id()
        )
        src_pres, dst_pres, error = _resolve_copy_decks(
            presentations, source_presentation_id, dst_id
        )
        if error is not None:
            return error

        index_error = _slide_index_error(src_pres, slide_index)
        if index_error is not None:
            return index_error

        return _copy_slide_response(
            src_pres, dst_pres, source_presentation_id, dst_id, slide_index
        )
