"""Style-lint rule catalog v1: registry + text/style rules (Step 5).

Every rule is a small generator taking a ``LintContext`` and yielding
finding dicts (``utils.lint_engine.make_finding``). The ``RULES`` table
at the bottom is the single registry the engine executes -- id, default
severity and a one-line doc per rule; geometry/media rules live in
``utils.lint_rules_geometry`` and register through the same table.

Catalog notes (research 2026-07-03 §3, plan Step 5):

* Messages are distribution-style wherever the profile learned a
  distribution (type scale, spacing quanta, palette shares).
* ``overlap`` / ``empty-slide`` / ``tiny-font`` port the rule semantics
  of leppardwang/pptx-lint (MIT -- see NOTICE), adapted to resolved
  values: overlap considers non-placeholder shapes only (placeholders
  are layout containers that floating panels legitimately sit inside),
  and the tiny-font floor sits below normal footer/caption sizes.
* ``hardcoded-color`` is scoped to TEXT colors: desktop PowerPoint
  itself writes literal ``srgbClr`` for shape fills/lines assigned via
  RGB, so flagging those would indict conformant house decks;
  re-linking shape fills is ``apply_style_profile`` territory.
* The ``font-family`` buFont facet is a deck-consistency check (the
  pinned house-profile/1 contract carries no bullet font key), and the
  ea/cs/sym facet flags explicit literal typefaces outside the house
  set -- the Slidewise complete-coverage trick.

Out of scope v1 (documented, per plan): spellcheck and cross-slide
numeric consistency -- LLM territory at generation time, not lint.
"""

from collections import Counter
from dataclasses import dataclass
from typing import Callable, Dict, Iterator, Optional, Tuple

from .lint_engine import (
    LINE_SPACING_TOLERANCE,
    SIZE_TOLERANCE_PT,
    SPACING_TOLERANCE_PT,
    LintContext,
    format_scale,
    iter_runs,
    iter_shapes,
    make_finding,
)
from .lint_xml import (
    autofit_of,
    iter_run_props,
    run_lang_of,
    run_typefaces_of,
    solid_srgb_of,
)
from .style_apply import leaf_points, leaf_value

#: pptx-lint-seeded readability floor -- deliberately below normal
#: footer/caption sizes (house corpora run 9-11pt furniture).
TINY_FONT_FLOOR_PT = 9.0

Finding = Dict
RuleFn = Callable[[LintContext], Iterator[Finding]]


@dataclass(frozen=True)
class Rule:
    """One registry row: id, default severity, one-line doc, impl."""

    rule_id: str
    severity: str
    doc: str
    fn: RuleFn


# ---------------------------------------------------------------------------
# font-scale
# ---------------------------------------------------------------------------

def rule_font_scale(ctx: LintContext) -> Iterator[Finding]:
    """Run sizes off the learned house type scale."""
    scale = ctx.font_size_scale
    if not scale:
        return
    for slide, shape, p_idx, r_idx, run in iter_runs(ctx):
        size = run["font"].get("size_pt")
        if size is None or not run.get("text", "").strip():
            continue
        if min(abs(size - member) for member in scale) <= SIZE_TOLERANCE_PT:
            continue
        role = ctx.shape_role(slide, shape) or "text"
        yield make_finding(
            "font-scale", "error", slide["slide_number"], shape["name"],
            "font.size_pt", scale, size,
            f"{role} run is {size:g}pt; house type scale is "
            f"{format_scale(scale, 'pt')}",
            shape_id=shape["shape_id"], paragraph=p_idx, run=r_idx,
        )


# ---------------------------------------------------------------------------
# font-family (complete run coverage: latin + ea/cs/sym + buFont)
# ---------------------------------------------------------------------------

def _latin_findings(ctx: LintContext) -> Iterator[Finding]:
    fonts = set(ctx.house_fonts)
    for slide, shape, p_idx, r_idx, run in iter_runs(ctx):
        name = run["font"].get("name")
        if not run.get("text", "").strip() or name is None:
            continue
        if name in fonts:
            continue
        yield make_finding(
            "font-family", "error", slide["slide_number"], shape["name"],
            "font.name", ctx.house_fonts, name,
            f"run font is {name!r}; house fonts are {ctx.house_fonts}",
            shape_id=shape["shape_id"], paragraph=p_idx, run=r_idx,
        )


def _extra_script_findings(ctx: LintContext) -> Iterator[Finding]:
    """Explicit ea/cs/sym typefaces outside the house set (raw XML)."""
    fonts = set(ctx.house_fonts)
    for slide, shape in iter_shapes(ctx):
        for p_idx, r_idx, rpr in iter_run_props(shape["_shape"]._element):
            for tag, typeface in run_typefaces_of(rpr).items():
                if tag == "latin" or typeface in fonts:
                    continue
                yield make_finding(
                    "font-family", "warn", slide["slide_number"],
                    shape["name"], f"font.{tag}", ctx.house_fonts,
                    typeface,
                    f"run declares {tag} typeface {typeface!r} outside "
                    f"the house fonts {ctx.house_fonts} -- fonts must "
                    "match across every script slot (latin/ea/cs/sym)",
                    shape_id=shape["shape_id"], paragraph=p_idx,
                    run=r_idx,
                )


def _bullet_font_findings(ctx: LintContext) -> Iterator[Finding]:
    """Explicit bullet fonts inconsistent with the deck's convention."""
    counts: Counter = Counter()
    for slide, shape in iter_shapes(ctx):
        for paragraph in shape.get("paragraphs", []):
            bullet = paragraph.get("bullet") or {}
            if not bullet.get("font_follows_text") and bullet.get("font"):
                counts[bullet["font"]] += 1
    if len(counts) < 2:
        return
    modal = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
    for slide, shape in iter_shapes(ctx):
        for p_idx, paragraph in enumerate(shape.get("paragraphs", []),
                                          start=1):
            bullet = paragraph.get("bullet") or {}
            font = bullet.get("font")
            if (bullet.get("font_follows_text") or font is None
                    or font == modal):
                continue
            yield make_finding(
                "font-family", "warn", slide["slide_number"],
                shape["name"], "bullet.font", modal, font,
                f"bullet font is {font!r}; the deck's bullet convention "
                f"is {modal!r}",
                shape_id=shape["shape_id"], paragraph=p_idx,
            )


def rule_font_family(ctx: LintContext) -> Iterator[Finding]:
    """Fonts off the house set, across every script slot + buFont."""
    if not ctx.house_fonts:
        return
    yield from _latin_findings(ctx)
    yield from _extra_script_findings(ctx)
    yield from _bullet_font_findings(ctx)


# ---------------------------------------------------------------------------
# bullet-style / spacing (body placeholder cascade, level-scoped)
# ---------------------------------------------------------------------------

def _iter_body_paragraphs(ctx: LintContext):
    """(slide, shape, p_idx_1based, paragraph) for body-role shapes."""
    for slide, shape in iter_shapes(ctx):
        if ctx.shape_role(slide, shape) != "body":
            continue
        for p_idx, paragraph in enumerate(shape.get("paragraphs", []),
                                          start=1):
            if any(run.get("text", "").strip()
                   for run in paragraph.get("runs", [])):
                yield slide, shape, p_idx, paragraph


def rule_bullet_style(ctx: LintContext) -> Iterator[Finding]:
    """Body bullet characters off the per-level house convention."""
    rules = ctx.paragraph_rules.get("bullets") or {}
    if not rules:
        return
    for slide, shape, p_idx, paragraph in _iter_body_paragraphs(ctx):
        rule = rules.get(f"l{paragraph['indent_level']}")
        if rule is None:
            continue
        expected = leaf_value(rule["char"])
        bullet = paragraph.get("bullet") or {}
        actual = bullet.get("char") if bullet.get("type") == "char" else None
        if actual is None or actual == expected:
            continue
        level = paragraph["indent_level"]
        yield make_finding(
            "bullet-style", "error", slide["slide_number"], shape["name"],
            "bullet.char", expected, actual,
            f"level {level} bullet is {actual!r}; house level {level} "
            f"bullet is {expected!r}",
            shape_id=shape["shape_id"], paragraph=p_idx,
        )


def _spacing_checks(ctx: LintContext) -> Tuple:
    rules = ctx.paragraph_rules
    checks = []
    for key in ("space_before", "space_after"):
        if key in rules:
            checks.append((key, leaf_points(rules[key])))
    return tuple(checks)


def _spacing_scale_text(ctx: LintContext, key: str,
                        rule_pt: float) -> str:
    if key == "space_after" and ctx.space_after_scale:
        return (f"house space_after scale is "
                f"{format_scale(ctx.space_after_scale, 'pt')}")
    return f"house {key} is {rule_pt:g}pt"


def rule_spacing(ctx: LintContext) -> Iterator[Finding]:
    """Level-1 body paragraph spacing off the house rules."""
    checks = _spacing_checks(ctx)
    line_rule = ctx.paragraph_rules.get("line_spacing")
    line_expected = leaf_value(line_rule) if line_rule else None
    if not checks and line_expected is None:
        return
    for slide, shape, p_idx, paragraph in _iter_body_paragraphs(ctx):
        if paragraph["indent_level"] != 1:
            continue  # the profile learns level-1 spacing only
        for key, rule_pt in checks:
            actual = paragraph[key].get("points")
            if actual is None or abs(actual - rule_pt) <= SPACING_TOLERANCE_PT:
                continue
            yield make_finding(
                "spacing", "error", slide["slide_number"], shape["name"],
                f"paragraph.{key}_pt", rule_pt, actual,
                f"paragraph {key} is {actual:g}pt; "
                f"{_spacing_scale_text(ctx, key, rule_pt)}",
                shape_id=shape["shape_id"], paragraph=p_idx,
            )
        actual_line = paragraph["line_spacing"].get("multiple")
        if (line_expected is not None and actual_line is not None
                and abs(actual_line - line_expected)
                > LINE_SPACING_TOLERANCE):
            yield make_finding(
                "spacing", "error", slide["slide_number"], shape["name"],
                "paragraph.line_spacing", line_expected, actual_line,
                f"line spacing is {actual_line:g}x; house line spacing "
                f"is {line_expected:g}x",
                shape_id=shape["shape_id"], paragraph=p_idx,
            )


# ---------------------------------------------------------------------------
# color-palette / hardcoded-color (raw srgbClr probes)
# ---------------------------------------------------------------------------

def _off_palette(ctx: LintContext, hex_value: Optional[str]) -> bool:
    return (hex_value is not None
            and hex_value not in ctx.scheme_hexes
            and hex_value not in ctx.usage_hexes)


def rule_color_palette(ctx: LintContext) -> Iterator[Finding]:
    """Literal srgbClr colors outside the house palette."""
    if not ctx.scheme_hexes:
        return
    allowed = sorted(ctx.scheme_hexes | ctx.usage_hexes)
    from .lint_xml import shape_fill_srgb, shape_line_srgb

    def finding(slide, shape, prop, hex_value, label, **refs):
        return make_finding(
            "color-palette", "error", slide["slide_number"],
            shape["name"], prop, allowed, f"#{hex_value}",
            f"{label} #{hex_value} is not one of the "
            f"{len(ctx.scheme_hexes)} scheme colors or the house "
            "top-usage palette",
            shape_id=shape["shape_id"], **refs,
        )

    for slide, shape in iter_shapes(ctx):
        elem = shape["_shape"]._element
        fill_hex = shape_fill_srgb(elem)
        if _off_palette(ctx, fill_hex):
            yield finding(slide, shape, "fill.color_hex", fill_hex,
                          "shape fill")
        line_hex = shape_line_srgb(elem)
        if _off_palette(ctx, line_hex):
            yield finding(slide, shape, "line.color_hex", line_hex,
                          "shape line")
        for p_idx, r_idx, rpr in iter_run_props(elem):
            run_hex = solid_srgb_of(rpr)
            if run_hex in ctx.scheme_hexes:
                continue  # theme-equal hardcodes belong to hardcoded-color
            if _off_palette(ctx, run_hex):
                yield finding(slide, shape, "font.color_hex", run_hex,
                              "run color", paragraph=p_idx, run=r_idx)


def rule_hardcoded_color(ctx: LintContext) -> Iterator[Finding]:
    """Text colored via literal srgbClr equal to a theme color."""
    if not ctx.scheme_hexes:
        return
    for slide, shape in iter_shapes(ctx):
        for p_idx, r_idx, rpr in iter_run_props(shape["_shape"]._element):
            hex_value = solid_srgb_of(rpr)
            if hex_value is None or hex_value not in ctx.scheme_hexes:
                continue
            token = ctx.token_for_hex(hex_value)
            yield make_finding(
                "hardcoded-color", "error", slide["slide_number"],
                shape["name"], "font.color_source",
                f"schemeClr {token}", f"srgbClr {hex_value}",
                f"run color #{hex_value} equals theme color {token} but "
                f"is hardcoded as srgbClr; use schemeClr {token} so the "
                "text follows the theme",
                shape_id=shape["shape_id"], paragraph=p_idx, run=r_idx,
            )


# ---------------------------------------------------------------------------
# proofing-language / tiny-font / autofit-shrink / empty-slide
# ---------------------------------------------------------------------------

def rule_proofing_language(ctx: LintContext) -> Iterator[Finding]:
    """Run proofing languages inconsistent across the deck."""
    langs: Counter = Counter()
    for slide, shape in iter_shapes(ctx):
        for _, _, rpr in iter_run_props(shape["_shape"]._element):
            lang = run_lang_of(rpr)
            if lang:
                langs[lang] += 1
    if len(langs) < 2:
        return
    modal = sorted(langs.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
    for slide, shape in iter_shapes(ctx):
        for p_idx, r_idx, rpr in iter_run_props(shape["_shape"]._element):
            lang = run_lang_of(rpr)
            if lang is None or lang == modal:
                continue
            yield make_finding(
                "proofing-language", "warn", slide["slide_number"],
                shape["name"], "font.lang", modal, lang,
                f"run proofing language is {lang}; the deck's dominant "
                f"language is {modal}",
                shape_id=shape["shape_id"], paragraph=p_idx, run=r_idx,
            )


def rule_tiny_font(ctx: LintContext) -> Iterator[Finding]:
    """Run sizes below the readability floor (pptx-lint seed)."""
    for slide, shape, p_idx, r_idx, run in iter_runs(ctx):
        size = run["font"].get("size_pt")
        if (size is None or size >= TINY_FONT_FLOOR_PT
                or not run.get("text", "").strip()):
            continue
        yield make_finding(
            "tiny-font", "warn", slide["slide_number"], shape["name"],
            "font.size_pt", f">= {TINY_FONT_FLOOR_PT:g}pt", size,
            f"run is {size:g}pt -- below the {TINY_FONT_FLOOR_PT:g}pt "
            "readability floor",
            shape_id=shape["shape_id"], paragraph=p_idx, run=r_idx,
        )


def rule_autofit_shrink(ctx: LintContext) -> Iterator[Finding]:
    """normAutofit fontScale < 100% -- the overflow proxy."""
    for slide, shape in iter_shapes(ctx):
        kind, scale_pct = autofit_of(shape["_shape"]._element)
        if kind != "normAutofit" or scale_pct is None or scale_pct >= 100:
            continue
        yield make_finding(
            "autofit-shrink", "warn", slide["slide_number"],
            shape["name"], "autofit.font_scale", 100.0, scale_pct,
            f"normAutofit shrinks this frame's text to {scale_pct:g}% -- "
            "the declared sizes overflow the frame",
            shape_id=shape["shape_id"],
        )


def rule_empty_slide(ctx: LintContext) -> Iterator[Finding]:
    """Slides with no text and no pictures (pptx-lint seed)."""
    from .lint_engine import shape_has_text

    for slide in ctx.slides:
        has_content = any(
            shape_has_text(shape) or shape.get("is_picture")
            for shape in slide["shapes"]
        )
        if has_content:
            continue
        yield make_finding(
            "empty-slide", "warn", slide["slide_number"], None,
            "slide.content", "text or image content", "none",
            "slide carries no text and no images",
        )


# ---------------------------------------------------------------------------
# text-overflow-predicted (utils.text_fit)
# ---------------------------------------------------------------------------

def rule_text_overflow(ctx: LintContext) -> Iterator[Finding]:
    """Frames whose declared text is predicted not to fit."""
    from .lint_engine import shape_has_text
    from .text_fit import assess_frame_record

    for slide, shape in iter_shapes(ctx):
        if not shape_has_text(shape):
            continue
        result = assess_frame_record(shape)
        verdict = result.get("verdict")
        if verdict not in ("overflow", "borderline"):
            continue
        severity = "warn" if verdict == "overflow" else "info"
        required = result.get("required_pt")
        available = result.get("available_pt")
        suffix = (" -- within the ±5% band, confirm by render"
                  if verdict == "borderline" else "")
        yield make_finding(
            "text-overflow-predicted", severity, slide["slide_number"],
            shape["name"], "text.fit", "text fits the frame",
            f"requires ~{required:g}pt of {available:g}pt available",
            f"predicted text height ~{required:g}pt vs "
            f"{available:g}pt frame height ({verdict}){suffix}",
            shape_id=shape["shape_id"],
        )


# ---------------------------------------------------------------------------
# Registry (engine executes in this order; output is re-sorted anyway)
# ---------------------------------------------------------------------------

def _geometry_rules():
    from . import lint_rules_geometry as geo

    return (
        Rule("border-style", "error",
             "Visible shape borders off the house weight/color/dash",
             geo.rule_border_style),
        Rule("off-grid", "error",
             "Shape edges off the learned alignment grid "
             "(distance-to-nearest-gridline reported)",
             geo.rule_off_grid),
        Rule("straggler-textbox", "error",
             "Non-placeholder text parked in the footer zone off the "
             "grid anchors", geo.rule_straggler_textbox),
        Rule("off-slide", "warn",
             "Shapes extending beyond the slide bounds",
             geo.rule_off_slide),
        Rule("archetype-geometry", "warn",
             "Title band / body region off the per-archetype house "
             "boxes", geo.rule_archetype_geometry),
        Rule("image-distortion", "warn",
             "Picture frame aspect off the native (srcRect-adjusted) "
             "aspect", geo.rule_image_distortion),
        Rule("image-dpi", "warn",
             "Picture effective resolution under the 96 DPI floor",
             geo.rule_image_dpi),
        Rule("footer-presence", "warn",
             "Working slides missing the house footer furniture",
             geo.rule_footer_presence),
        Rule("overlap", "warn",
             "Overlapping floating shapes (pptx-lint seed)",
             geo.rule_overlap),
    )


def _build_rules():
    return (
        Rule("font-scale", "error",
             "Run sizes off the learned house type scale",
             rule_font_scale),
        Rule("font-family", "error",
             "Fonts off the house set across latin/ea/cs/sym + buFont "
             "(complete run coverage)", rule_font_family),
        Rule("bullet-style", "error",
             "Body bullet characters off the per-level house bullets",
             rule_bullet_style),
        Rule("spacing", "error",
             "Body paragraph space before/after/line off the house "
             "rules", rule_spacing),
        Rule("color-palette", "error",
             "Literal srgbClr colors outside the house palette",
             rule_color_palette),
        Rule("hardcoded-color", "error",
             "Text hardcoded srgbClr where a schemeClr theme link is "
             "expected", rule_hardcoded_color),
        *_geometry_rules(),
        Rule("autofit-shrink", "warn",
             "normAutofit fontScale < 100% (overflow proxy)",
             rule_autofit_shrink),
        Rule("text-overflow-predicted", "warn",
             "Declared text predicted to overflow its frame "
             "(borderline -> info, confirm by render)",
             rule_text_overflow),
        Rule("proofing-language", "warn",
             "Run proofing languages inconsistent across the deck",
             rule_proofing_language),
        Rule("empty-slide", "warn",
             "Slides with no text and no images (pptx-lint seed)",
             rule_empty_slide),
        Rule("tiny-font", "warn",
             "Run sizes below the readability floor (pptx-lint seed)",
             rule_tiny_font),
    )


#: The rule catalog v1 -- the engine runs these in order.
RULES: Tuple[Rule, ...] = _build_rules()

#: id -> Rule, for tools and docs.
RULES_BY_ID: Dict[str, Rule] = {rule.rule_id: rule for rule in RULES}
