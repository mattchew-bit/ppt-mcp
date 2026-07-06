"""Geometry / media lint rules (Step 5 rule catalog v1, second half).

Implementations for the registry in ``utils.lint_rules``: grid
alignment, footer-zone stragglers, off-slide objects, per-archetype
geometry, image aspect/DPI, footer presence, and the pptx-lint-seeded
overlap rule. Pure functions over ``LintContext``; strictly read-only.
"""

from itertools import combinations
from typing import Dict, Iterator, List, Optional, Tuple

from .lint_engine import (
    BORDER_TOLERANCE_PT,
    LintContext,
    format_scale,
    iter_shapes,
    make_finding,
    pt_to_in,
    rect_of,
    shape_has_text,
)
from .profile_archetypes import body_region_rect, find_title_shape
from .profile_grid import POINTS_PER_INCH
from .style_apply import leaf_points, leaf_value

Finding = Dict

#: Off-slide slack, points (sub-point float noise is not a bleed).
OFF_SLIDE_TOLERANCE_PT = 1.0

#: Per-coordinate tolerance for archetype box comparisons, inches
#: (median-vs-instance jitter from auto-sized boxes stays inside this).
ARCHETYPE_TOLERANCE_IN = 0.25

#: Minimum overlap extent (points) in BOTH axes before a pair counts.
OVERLAP_MIN_PT = 2.0

#: Frame-vs-native aspect ratio drift beyond this flags distortion.
ASPECT_TOLERANCE = 0.02

#: Effective-resolution floor for placed images (plan rule catalog).
MIN_EFFECTIVE_DPI = 96.0

#: Archetypes whose house convention includes footer furniture; the
#: display slides (title / section_divider / full_bleed) are exempt.
FOOTER_REQUIRED_ARCHETYPES = ("agenda", "content", "two_column",
                              "closing")


# ---------------------------------------------------------------------------
# border-style
# ---------------------------------------------------------------------------

def rule_border_style(ctx: LintContext) -> Iterator[Finding]:
    """Visible borders off the house weight / color / dash."""
    rule = ctx.border_rule
    if not rule:
        return
    weight_rule = (leaf_points(rule["weight"]) if "weight" in rule
                   else None)
    color_rule = (str(leaf_value(rule["color"])).lstrip("#").upper()
                  if "color" in rule else None)
    dash_rule = leaf_value(rule["dash"]) if "dash" in rule else None

    for slide, shape in iter_shapes(ctx):
        if shape.get("is_placeholder") or shape.get("is_picture"):
            continue
        line = shape.get("line") or {}
        if not line.get("visible"):
            continue
        yield from _border_deltas(ctx, slide, shape, line, weight_rule,
                                  color_rule, dash_rule)


def _border_deltas(ctx, slide, shape, line, weight_rule, color_rule,
                   dash_rule) -> Iterator[Finding]:
    weight = line.get("weight_pt")
    if (weight_rule is not None and weight is not None
            and abs(weight - weight_rule) > BORDER_TOLERANCE_PT):
        yield make_finding(
            "border-style", "error", slide["slide_number"],
            shape["name"], "line.weight_pt", weight_rule, weight,
            f"shape border is {weight:g}pt; house border is "
            f"{weight_rule:g}pt",
            shape_id=shape["shape_id"],
        )
    color = (line.get("color_hex") or "").upper()
    if color_rule is not None and color and color != color_rule:
        yield make_finding(
            "border-style", "error", slide["slide_number"],
            shape["name"], "line.color_hex", f"#{color_rule}",
            f"#{color}",
            f"shape border color is #{color}; house border color is "
            f"#{color_rule}",
            shape_id=shape["shape_id"],
        )
    dash = line.get("dash")
    if dash_rule is not None and dash is not None and dash != dash_rule:
        yield make_finding(
            "border-style", "error", slide["slide_number"],
            shape["name"], "line.dash", dash_rule, dash,
            f"shape border dash is {dash!r}; house border dash is "
            f"{dash_rule!r}",
            shape_id=shape["shape_id"],
        )


# ---------------------------------------------------------------------------
# off-grid / straggler-textbox
# ---------------------------------------------------------------------------

def _nearest(value_in: float,
             lines: List[float]) -> Tuple[Optional[float], float]:
    """(nearest line, distance) in inches; (None, inf) when no lines."""
    if not lines:
        return None, float("inf")
    line = min(lines, key=lambda edge: abs(value_in - edge))
    return line, abs(value_in - line)


def _grid_distances(ctx: LintContext,
                    rect: Tuple[float, float, float, float]):
    """Per-family (edge_label, shape_value_in, nearest, distance)."""
    left_in = pt_to_in(rect[0])
    right_in = pt_to_in(rect[0] + rect[2])
    center_in = pt_to_in(rect[0] + rect[2] / 2.0)
    for family, value in (("left", left_in), ("right", right_in),
                          ("center", center_in)):
        line, distance = _nearest(value, ctx.grid_edges.get(family, []))
        if line is not None:
            yield family, value, line, distance


def rule_off_grid(ctx: LintContext) -> Iterator[Finding]:
    """Shapes anchored to none of the learned grid families.

    A shape conforms when its left edge sits on a left gridline, OR its
    right edge on a right gridline, OR its center on a center gridline
    (within the learned tolerance). Footer-zone shapes are exempt --
    the grid was learned footerless (see ``utils.profile_extract``) and
    footer furniture is policed by ``straggler-textbox``. Rotated
    shapes are skipped: their xfrm box does not describe rendered
    edges.
    """
    tolerance = ctx.grid_tolerance_in
    if tolerance is None or not any(ctx.grid_edges.values()):
        return
    for slide, shape in iter_shapes(ctx):
        rect = rect_of(shape)
        if rect is None or ctx.in_footer_zone(slide, shape):
            continue
        if shape.get("rotation_deg"):
            continue
        distances = list(_grid_distances(ctx, rect))
        if not distances or any(d <= tolerance for _, _, _, d in distances):
            continue
        family, value, line, distance = min(
            distances, key=lambda entry: entry[3])
        yield make_finding(
            "off-grid", "error", slide["slide_number"], shape["name"],
            f"geometry.{family}_pt", round(line * POINTS_PER_INCH, 2),
            round(value * POINTS_PER_INCH, 2),
            f"{family} edge {value * POINTS_PER_INCH:.1f}pt is "
            f"{distance * POINTS_PER_INCH:.1f}pt from the nearest house "
            f"{family} gridline ({line * POINTS_PER_INCH:.1f}pt); house "
            f"{family} gridlines are "
            f"{format_scale(ctx.grid_edges[family], 'in')} "
            f"(tolerance {tolerance:g}in)",
            shape_id=shape["shape_id"],
        )


def rule_straggler_textbox(ctx: LintContext) -> Iterator[Finding]:
    """Non-placeholder text in the footer zone off the grid anchors.

    House footer furniture (source notes, page numbers) anchors to the
    learned vertical grid like everything else; a stray "draft" box
    pasted into the footer zone does not. Placeholder footers are
    sanctioned by the layout and never flagged.
    """
    tolerance = ctx.grid_tolerance_in
    grid_lines = ctx.all_grid_lines_in()
    if tolerance is None or not grid_lines:
        return
    for slide, shape in iter_shapes(ctx):
        if shape.get("is_placeholder") or not shape_has_text(shape):
            continue
        rect = rect_of(shape)
        if rect is None or not ctx.in_footer_zone(slide, shape):
            continue
        _, distance = _nearest(pt_to_in(rect[0]), grid_lines)
        if distance <= tolerance:
            continue
        left, top, width, height = rect
        yield make_finding(
            "straggler-textbox", "error", slide["slide_number"],
            shape["name"], "geometry",
            "footer zone reserved for grid-anchored footer furniture",
            f"text box at ({left:g}, {top:g}, {width:g}, {height:g})pt",
            f"non-placeholder text box parked in the footer zone "
            f"{distance * POINTS_PER_INCH:.1f}pt off the nearest grid "
            "anchor -- house footers carry a source note and page "
            "number only",
            shape_id=shape["shape_id"],
        )


# ---------------------------------------------------------------------------
# off-slide
# ---------------------------------------------------------------------------

def rule_off_slide(ctx: LintContext) -> Iterator[Finding]:
    """Shapes extending beyond the slide canvas."""
    for slide, shape in iter_shapes(ctx):
        rect = rect_of(shape)
        if rect is None:
            continue
        left, top, width, height = rect
        overhang = max(
            -left, -top,
            (left + width) - slide["width_pt"],
            (top + height) - slide["height_pt"],
        )
        if overhang <= OFF_SLIDE_TOLERANCE_PT:
            continue
        yield make_finding(
            "off-slide", "warn", slide["slide_number"], shape["name"],
            "geometry",
            f"within the {slide['width_pt']:g}x{slide['height_pt']:g}pt "
            "slide",
            f"({left:g}, {top:g}, {width:g}, {height:g})pt",
            f"shape extends {overhang:.1f}pt beyond the slide bounds",
            shape_id=shape["shape_id"],
        )


# ---------------------------------------------------------------------------
# archetype-geometry / footer-presence
# ---------------------------------------------------------------------------

def _box_delta_in(spec_box: Dict, actual_pt: Tuple) -> Tuple[str, float]:
    """(worst coordinate, worst |delta| in inches) vs an archetype box."""
    worst_key, worst = "", -1.0
    for index, key in enumerate(("x", "y", "w", "h")):
        expected = float(leaf_value(spec_box[key]))
        delta = abs(actual_pt[index] / POINTS_PER_INCH - expected)
        if delta > worst:
            worst_key, worst = key, delta
    return worst_key, worst


def _box_text(box_pt: Tuple) -> str:
    return ("(" + ", ".join(f"{v / POINTS_PER_INCH:.2f}" for v in box_pt)
            + ")in")


def _spec_text(spec_box: Dict) -> str:
    return ("(" + ", ".join(f"{float(leaf_value(spec_box[k])):g}"
                            for k in ("x", "y", "w", "h")) + ")in")


def _archetype_box_findings(ctx, slide, name, spec) -> Iterator[Finding]:
    title = find_title_shape(slide)
    boxes = []
    if title is not None:
        title_rect = rect_of(title)
        if spec.get("title_band") and title_rect is not None:
            boxes.append(("title_band", spec["title_band"], title_rect,
                          title["name"], title["shape_id"]))
        body = body_region_rect(slide, title, ctx.body_size_pt)
        if spec.get("body_region") and body is not None:
            boxes.append(("body_region", spec["body_region"], body,
                          None, None))
    for label, spec_box, actual, shape_name, shape_id in boxes:
        key, delta = _box_delta_in(spec_box, actual)
        if delta <= ARCHETYPE_TOLERANCE_IN:
            continue
        yield make_finding(
            "archetype-geometry", "warn", slide["slide_number"],
            shape_name, f"archetype.{label}", _spec_text(spec_box),
            _box_text(actual),
            f"{label} of this {name!r} slide is off the house box by "
            f"{delta:.2f}in on {key} (house {name} {label} is "
            f"{_spec_text(spec_box)})",
            shape_id=shape_id,
        )


def rule_archetype_geometry(ctx: LintContext) -> Iterator[Finding]:
    """Title band / body region vs the per-archetype house boxes."""
    if not ctx.archetypes:
        return
    for slide in ctx.slides:
        name = ctx.archetype_by_slide[slide["slide_number"]]
        spec = ctx.archetypes.get(name)
        if spec is None:
            yield make_finding(
                "archetype-geometry", "info", slide["slide_number"],
                None, "archetype", sorted(ctx.archetypes), name,
                f"slide classifies as {name!r}, an archetype the house "
                "profile never learned",
            )
            continue
        yield from _archetype_box_findings(ctx, slide, name, spec)


def rule_footer_presence(ctx: LintContext) -> Iterator[Finding]:
    """Working slides missing footer furniture.

    Applies to archetypes the house profile actually learned from the
    corpus; display slides (title / section_divider / full_bleed) are
    conventionally footer-free and exempt.
    """
    required = [name for name in FOOTER_REQUIRED_ARCHETYPES
                if name in ctx.archetypes]
    if not required:
        return
    for slide in ctx.slides:
        name = ctx.archetype_by_slide[slide["slide_number"]]
        if name not in required:
            continue
        has_footer = any(
            shape_has_text(shape) and ctx.in_footer_zone(slide, shape)
            for shape in slide["shapes"]
        )
        if has_footer:
            continue
        yield make_finding(
            "footer-presence", "warn", slide["slide_number"], None,
            "footer.presence", "footer text in the footer zone", "none",
            f"{name!r} slides carry house footer furniture (source note "
            "/ page number); this slide's footer zone is empty",
        )


# ---------------------------------------------------------------------------
# image-distortion / image-dpi
# ---------------------------------------------------------------------------

def _picture_metrics(shape_record) -> Optional[Dict]:
    """Native/effective pixel dims + frame inches, or None (linked)."""
    shape = shape_record["_shape"]
    rect = rect_of(shape_record)
    if rect is None:
        return None
    try:
        native_w, native_h = shape.image.size
        crops = (shape.crop_left, shape.crop_right,
                 shape.crop_top, shape.crop_bottom)
    except Exception:
        return None  # linked / unreadable image: nothing to measure
    eff_w = native_w * max(0.0, 1.0 - crops[0] - crops[1])
    eff_h = native_h * max(0.0, 1.0 - crops[2] - crops[3])
    if eff_w <= 0 or eff_h <= 0:
        return None
    return {
        "eff_px": (eff_w, eff_h),
        "frame_in": (rect[2] / POINTS_PER_INCH,
                     rect[3] / POINTS_PER_INCH),
        "cropped": any(c > 0 for c in crops),
    }


def rule_image_distortion(ctx: LintContext) -> Iterator[Finding]:
    """Picture frames stretched off the native aspect ratio."""
    for slide, shape in iter_shapes(ctx):
        if not shape.get("is_picture"):
            continue
        metrics = _picture_metrics(shape)
        if metrics is None:
            continue
        native_aspect = metrics["eff_px"][0] / metrics["eff_px"][1]
        frame_w, frame_h = metrics["frame_in"]
        if frame_h <= 0:
            continue
        frame_aspect = frame_w / frame_h
        drift = abs(frame_aspect / native_aspect - 1.0)
        if drift <= ASPECT_TOLERANCE:
            continue
        yield make_finding(
            "image-distortion", "warn", slide["slide_number"],
            shape["name"], "image.aspect_ratio",
            round(native_aspect, 3), round(frame_aspect, 3),
            f"picture frame aspect {frame_aspect:.3f} distorts the "
            f"native (crop-adjusted) aspect {native_aspect:.3f} by "
            f"{drift * 100:.0f}%",
            shape_id=shape["shape_id"],
        )


def rule_image_dpi(ctx: LintContext) -> Iterator[Finding]:
    """Pictures placed below the effective-resolution floor."""
    for slide, shape in iter_shapes(ctx):
        if not shape.get("is_picture"):
            continue
        metrics = _picture_metrics(shape)
        if metrics is None:
            continue
        frame_w, frame_h = metrics["frame_in"]
        if frame_w <= 0 or frame_h <= 0:
            continue
        dpi = min(metrics["eff_px"][0] / frame_w,
                  metrics["eff_px"][1] / frame_h)
        if dpi >= MIN_EFFECTIVE_DPI:
            continue
        yield make_finding(
            "image-dpi", "warn", slide["slide_number"], shape["name"],
            "image.effective_dpi", f">= {MIN_EFFECTIVE_DPI:g}",
            round(dpi, 1),
            f"picture renders at ~{dpi:.0f} DPI at its placed size -- "
            f"below the {MIN_EFFECTIVE_DPI:g} DPI floor",
            shape_id=shape["shape_id"],
        )


# ---------------------------------------------------------------------------
# overlap (pptx-lint seed)
# ---------------------------------------------------------------------------

def rule_overlap(ctx: LintContext) -> Iterator[Finding]:
    """Overlapping floating shapes (non-placeholder pairs only).

    Placeholders are layout containers -- panels and exhibits
    legitimately sit inside a body placeholder's box -- so only
    non-placeholder x non-placeholder pairs count (documented
    adaptation of the pptx-lint overlap rule).
    """
    for slide in ctx.slides:
        floats = [
            (shape, rect_of(shape)) for shape in slide["shapes"]
            if not shape.get("is_placeholder") and rect_of(shape)
        ]
        for (a, ra), (b, rb) in combinations(floats, 2):
            inter_w = (min(ra[0] + ra[2], rb[0] + rb[2])
                       - max(ra[0], rb[0]))
            inter_h = (min(ra[1] + ra[3], rb[1] + rb[3])
                       - max(ra[1], rb[1]))
            if inter_w <= OVERLAP_MIN_PT or inter_h <= OVERLAP_MIN_PT:
                continue
            yield make_finding(
                "overlap", "warn", slide["slide_number"], a["name"],
                "geometry.overlap", "no overlap between floating shapes",
                f"{a['name']!r} overlaps {b['name']!r} by "
                f"{inter_w:.0f}x{inter_h:.0f}pt",
                f"floating shapes {a['name']!r} and {b['name']!r} "
                f"overlap by {inter_w:.0f}x{inter_h:.0f}pt",
                shape_id=shape_id_of(a),
            )


def shape_id_of(shape: Dict) -> Optional[int]:
    return shape.get("shape_id")
