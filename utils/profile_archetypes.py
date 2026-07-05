"""Geometric slide-archetype classification for house-style profiles.

Classifies each slide into a CLOSED set of layout archetypes using only
title-band geometry, shape-count/type signatures and text density -- no
ML, per the style-fidelity plan (Step 3). The closed set:

    title, agenda, section_divider, content, two_column, full_bleed,
    closing

Rules (fixture-arbitrated against the Meridian house corpus):

* a picture covering most of the slide       -> ``full_bleed``
* title band in the TOP quarter of the slide -> a working slide:
    - two or more tall, narrow text panels side by side -> ``two_column``
    - body region spanning most of the slide width      -> ``content``
    - otherwise (list + sidebar)                        -> ``agenda``
* title band LOWER on the slide -> a display slide:
    - a text shape ABOVE the title (a kicker)  -> ``section_divider``
    - body text below starting near the left   -> ``title`` (cover)
    - body text below starting toward center   -> ``closing``
    - no body text at all                      -> ``title``

Input is the "slide facts" structure built by ``utils.profile_extract``;
output geometry is in inches (unrounded -- ``utils.profile_schema``
rounds and DTCG-shapes it). Pure functions, no mutation.
"""

from typing import Dict, List, Optional, Sequence, Tuple

from .profile_grid import POINTS_PER_INCH
from .style_roles import ph_type_label_name, placeholder_role

#: The closed archetype set (hard-capped by the plan; no semantic types).
ARCHETYPE_NAMES = (
    "title", "agenda", "section_divider", "content", "two_column",
    "full_bleed", "closing",
)

#: Title tops below this fraction of the slide height mean "working
#: slide" (agenda/content/two_column); at or below, a display slide.
TOP_BAND_FRACTION = 0.25

#: Shapes whose top edge sits below this fraction are footer furniture.
FOOTER_FRACTION = 0.9

#: A picture covering at least this share of the slide is a full bleed.
FULL_BLEED_COVERAGE = 0.7

#: Column panels: at least this tall (fraction of slide height) ...
COLUMN_MIN_HEIGHT_FRACTION = 0.35

#: ... and at most this wide (fraction of slide width).
COLUMN_MAX_WIDTH_FRACTION = 0.45

#: Body regions at least this wide (fraction) mean ``content``.
CONTENT_MIN_WIDTH_FRACTION = 0.7

#: Cover subtitles start within this fraction of the slide width.
COVER_TEXT_LEFT_FRACTION = 0.25


def _rect(shape) -> Optional[Tuple[float, float, float, float]]:
    geometry = shape["geometry"]
    left, top = geometry.get("left_pt"), geometry.get("top_pt")
    width, height = geometry.get("width_pt"), geometry.get("height_pt")
    if None in (left, top, width, height):
        return None
    return (left, top, width, height)


def _max_run_size(shape) -> Optional[float]:
    sizes = [
        run["font"]["size_pt"]
        for paragraph in shape.get("paragraphs", [])
        for run in paragraph.get("runs", [])
        if run["font"].get("size_pt") is not None
    ]
    return max(sizes) if sizes else None


def _has_text(shape) -> bool:
    return any(
        run.get("text", "").strip()
        for paragraph in shape.get("paragraphs", [])
        for run in paragraph.get("runs", [])
    )


def find_title_shape(slide: Dict) -> Optional[Dict]:
    """The slide's title shape: title placeholder, else biggest-type text.

    The fallback (largest max run size, topmost on ties) covers decks
    whose authors detached titles from placeholders. Placeholder types
    map through the shared exact-match table (``utils.style_roles``),
    so a SUBTITLE placeholder is never mistaken for the title.
    """
    placeholders = [
        shape for shape in slide["shapes"]
        if placeholder_role(ph_type_label_name(shape.get("ph_type")))
        == "title" and _rect(shape) is not None
    ]
    if placeholders:
        return placeholders[0]
    candidates = [
        shape for shape in slide["shapes"]
        if _has_text(shape) and _rect(shape) is not None
        and _max_run_size(shape) is not None
    ]
    if not candidates:
        return None
    return min(
        candidates,
        key=lambda shape: (-_max_run_size(shape), _rect(shape)[1]),
    )


def _below_title(slide: Dict, title: Dict) -> List[Dict]:
    """Non-footer shapes whose top edge sits below the title band."""
    title_rect = _rect(title)
    title_bottom = title_rect[1] + title_rect[3] * 0.5
    footer_top = slide["height_pt"] * FOOTER_FRACTION
    below = []
    for shape in slide["shapes"]:
        if shape is title:
            continue
        rect = _rect(shape)
        if rect is None:
            continue
        if rect[1] > title_bottom and rect[1] < footer_top:
            below.append(shape)
    return below


def _text_above_title(slide: Dict, title: Dict) -> bool:
    title_top = _rect(title)[1]
    for shape in slide["shapes"]:
        if shape is title or shape.get("is_picture") or not _has_text(shape):
            continue
        rect = _rect(shape)
        if rect is not None and rect[1] + rect[3] <= title_top:
            return True
    return False


def _full_bleed_picture(slide: Dict) -> bool:
    slide_area = slide["width_pt"] * slide["height_pt"]
    for shape in slide["shapes"]:
        if not shape.get("is_picture"):
            continue
        rect = _rect(shape)
        if rect is not None and (rect[2] * rect[3]) >= (
                FULL_BLEED_COVERAGE * slide_area):
            return True
    return False


def body_region_rect(slide: Dict, title: Dict,
                     body_size_pt: float) -> Optional[Tuple]:
    """The slide's body region in points, or ``None``.

    Preference order (each step only runs when the previous found
    nothing): bounding box of text-bearing PLACEHOLDERS below the title;
    bounding box of body-sized text shapes below the title (captions and
    footers fall out on size); the largest remaining non-picture shape
    (covers divider rules).
    """
    below = [s for s in _below_title(slide, title)
             if not s.get("is_picture")]
    placeholders = [s for s in below if s.get("is_placeholder")
                    and _has_text(s)]
    if placeholders:
        return _bounding_box(placeholders)
    body_sized = [
        shape for shape in below
        if _has_text(shape)
        and (_max_run_size(shape) or 0) >= body_size_pt
    ]
    if body_sized:
        return _bounding_box(body_sized)
    others = [s for s in below if _rect(s) is not None]
    if not others:
        return None
    largest = max(others, key=lambda s: _rect(s)[2] * _rect(s)[3])
    return _rect(largest)


def _bounding_box(shapes: Sequence[Dict]) -> Tuple[float, float,
                                                   float, float]:
    rects = [_rect(shape) for shape in shapes]
    left = min(rect[0] for rect in rects)
    top = min(rect[1] for rect in rects)
    right = max(rect[0] + rect[2] for rect in rects)
    bottom = max(rect[1] + rect[3] for rect in rects)
    return (left, top, right - left, bottom - top)


def _column_panel_count(slide: Dict, title: Dict) -> int:
    count = 0
    for shape in _below_title(slide, title):
        if shape.get("is_picture") or shape.get("is_placeholder"):
            continue
        rect = _rect(shape)
        if rect is None or not _has_text(shape):
            continue
        if (rect[3] >= COLUMN_MIN_HEIGHT_FRACTION * slide["height_pt"]
                and rect[2] <= COLUMN_MAX_WIDTH_FRACTION
                * slide["width_pt"]):
            count += 1
    return count


def _classify_display_slide(slide: Dict, title: Dict,
                            body_size_pt: float) -> str:
    """Title sits low: cover, section divider, or closing slide."""
    if _text_above_title(slide, title):
        return "section_divider"
    below_text = [
        shape for shape in _below_title(slide, title)
        if not shape.get("is_picture") and _has_text(shape)
        and (_max_run_size(shape) or 0) >= body_size_pt
    ]
    if not below_text:
        return "title"
    leftmost = min(_rect(shape)[0] for shape in below_text)
    if leftmost < COVER_TEXT_LEFT_FRACTION * slide["width_pt"]:
        return "title"
    return "closing"


def _classify_working_slide(slide: Dict, title: Dict,
                            body_size_pt: float) -> str:
    """Title sits in the top band: agenda, content, or two_column."""
    if _column_panel_count(slide, title) >= 2:
        return "two_column"
    body = body_region_rect(slide, title, body_size_pt)
    if body is not None and body[2] >= (
            CONTENT_MIN_WIDTH_FRACTION * slide["width_pt"]):
        return "content"
    return "agenda"


def classify_slide(slide: Dict, body_size_pt: float) -> str:
    """Classify one slide-facts record into the closed archetype set.

    ``body_size_pt`` is the corpus's learned body size (used to tell
    body text from captions/labels). Slides with no title and no
    dominant picture fall back to ``content``.
    """
    if body_size_pt <= 0:
        raise ValueError(f"body_size_pt must be positive: {body_size_pt!r}")
    if _full_bleed_picture(slide):
        return "full_bleed"
    title = find_title_shape(slide)
    if title is None:
        return "content"
    title_fraction = _rect(title)[1] / slide["height_pt"]
    if title_fraction < TOP_BAND_FRACTION:
        return _classify_working_slide(slide, title, body_size_pt)
    return _classify_display_slide(slide, title, body_size_pt)


def _median_rect(rects: List[Tuple]) -> Dict[str, float]:
    """Per-coordinate median of point rects, converted to inches."""
    def median(values: List[float]) -> float:
        ordered = sorted(values)
        mid = len(ordered) // 2
        if len(ordered) % 2:
            return ordered[mid]
        return (ordered[mid - 1] + ordered[mid]) / 2.0

    keys = ("x", "y", "w", "h")
    return {
        key: median([rect[index] for rect in rects]) / POINTS_PER_INCH
        for index, key in enumerate(keys)
    }


def learn_archetypes(slides: Sequence[Dict],
                     body_size_pt: float) -> Dict[str, Dict]:
    """Classify every slide and learn per-archetype geometry boxes.

    Returns ``{archetype: {"title_band": rect_in, "body_region":
    rect_in|None, "count": n}}`` for archetypes actually observed.
    Rects are ``{"x", "y", "w", "h"}`` in inches (median across the
    archetype's slides -- robust to one-off nudges).
    """
    if not slides:
        raise ValueError("learn_archetypes requires at least one slide")
    grouped: Dict[str, Dict[str, List]] = {}
    for slide in slides:
        name = classify_slide(slide, body_size_pt)
        bucket = grouped.setdefault(
            name, {"titles": [], "bodies": [], "count": 0})
        bucket["count"] += 1
        title = find_title_shape(slide)
        if title is not None:
            bucket["titles"].append(_rect(title))
            body = body_region_rect(slide, title, body_size_pt)
            if body is not None:
                bucket["bodies"].append(body)
    result: Dict[str, Dict] = {}
    for name in ARCHETYPE_NAMES:  # deterministic, closed-set order
        bucket = grouped.get(name)
        if bucket is None:
            continue
        result[name] = {
            "title_band": (_median_rect(bucket["titles"])
                           if bucket["titles"] else None),
            "body_region": (_median_rect(bucket["bodies"])
                            if bucket["bodies"] else None),
            "count": bucket["count"],
        }
    return result
