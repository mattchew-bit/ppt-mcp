"""Public effective-style resolver API (Step 2 of the style-fidelity plan).

Returns what PowerPoint actually displays for runs, paragraphs and shapes
when python-pptx returns ``None`` because the value is inherited. Built
on the Stage 1 modules: ``resolve_core`` (the one generic cascade
walker), ``resolve_theme`` (per-slide theme engine) and ``resolve_colors``
(color math), plus the Stage 2 extractors in ``resolve_text_props`` and
the shape layer in ``resolve_shape_props``.

DOCUMENTED PRECEDENCE ORDER (fixture-arbitrated -- the Step 0
COM-extracted expected values reproduce exactly under this order):

Text properties, placeholder shapes:
    1. run ``a:rPr``
    2. paragraph ``a:pPr`` (its ``a:defRPr`` for character properties)
    3. shape's own ``p:txBody/a:lstStyle``
    4. layout placeholder ``a:lstStyle``   (SCPShapeTree type+idx match,
       type-only fallback)
    5. master placeholder ``a:lstStyle``   (matched via the layout ph)
    6. master ``p:txStyles`` (titleStyle / bodyStyle / otherStyle by ph
       type; subTitle/obj/absent -> bodyStyle)
    7. shape ``p:style/a:fontRef``         (color + typeface class only;
       Apache POI position -- AFTER the placeholder chain, BEFORE deck
       defaults. ShapeCrawler consults it before step 4; the fixtures
       cannot arbitrate because no fixture shape carries both, so the
       POI order was chosen and is pinned by unit test.)
    8. presentation ``p:defaultTextStyle``
    9. theme ``a:objectDefaults/a:txDef``  (the SLIDE'S MASTER'S theme)

Non-placeholder shapes (floating text boxes): 1, 2, 3, 7, 8, 9 -- the
master's ``p:otherStyle`` is deliberately NOT consulted for slide shapes
(see resolve_core). Master ``p:txStyles`` sits BEFORE the presentation
``p:defaultTextStyle`` (the ShapeCrawler-vs-POI disagreement, resolved
by fixture).

``a:normAutofit/@fontScale`` scales a run's size ONLY when the size came
from the run's own ``a:rPr`` (ShapeCrawler ``PortionFontSize`` rule);
inherited sizes are template values and are reported unscaled.

Hard defaults when the whole cascade is silent (ECMA-376 / observed
PowerPoint behavior, matched against COM-reported effective values):
size 18pt, bold/italic False, alignment left, spacing before/after 0pt,
line spacing 100%, bullet none, line invisible, fill none, geometry
preset rect.

Shape properties resolve through ``p:style`` references into the theme
format scheme with phClr substitution (see ``resolve_shape_props``).

Everything is strictly read-only.
"""

from dataclasses import dataclass
from typing import Any, Dict, Optional, Tuple

from .resolve_core import (
    build_text_cascade_sources,
    indent_level_of,
    resolve_run_property,
    resolve_paragraph_property,
)
from .resolve_shape_props import (
    font_ref_color_hex,
    font_ref_typeface,
    resolve_shape_fill,
    resolve_shape_geometry,
    resolve_shape_line,
)
from .resolve_text_props import (
    BULLET_AUTONUM,
    BULLET_CHAR,
    BULLET_NONE,
    autofit_font_scale,
    extract_alignment,
    extract_bold,
    extract_bullet_autonum,
    extract_bullet_char,
    extract_bullet_color,
    extract_bullet_font,
    extract_bullet_size,
    extract_bullet_type,
    extract_italic,
    extract_latin_typeface,
    extract_line_spacing,
    extract_size_hundredths,
    extract_solid_fill,
    extract_space_after,
    extract_space_before,
    scaled_size_pt,
    split_deck_defaults,
)
from .resolve_theme import ThemeContext

#: ECMA / PowerPoint hard defaults (documented in the module docstring).
DEFAULT_FONT_SIZE_PT = 18.0
DEFAULT_ALIGNMENT = "left"
DEFAULT_SPACE = {"points": 0.0}
DEFAULT_LINE_SPACING = {"multiple": 1.0}


@dataclass(frozen=True)
class TextContext:
    """One run's / paragraph's place in a deck (immutable).

    ``theme`` is optional: pass a prebuilt ``ThemeContext`` to avoid
    re-parsing the theme part per run when walking a whole deck -- it
    is threaded through the cascade builder (the theme-txDef tail) and
    every color / typeface resolution, so with it supplied the theme
    XML is parsed zero additional times per run.
    """

    slide: Any
    shape: Any
    paragraph: Any = None
    run: Any = None
    theme: Optional[ThemeContext] = None

    def __post_init__(self):
        if self.slide is None or self.shape is None:
            raise ValueError("TextContext requires slide and shape")

    def theme_context(self) -> ThemeContext:
        if self.theme is not None:
            return self.theme
        return ThemeContext.for_slide(self.slide)


def _cascade(context: TextContext,
             theme: Optional[ThemeContext] = None) -> Tuple:
    return build_text_cascade_sources(
        context.slide, context.shape, context.paragraph, context.run,
        theme=theme,
    )


def _level(context: TextContext) -> int:
    if context.paragraph is None:
        return 1
    return indent_level_of(context.paragraph)


# ---------------------------------------------------------------------------
# Run font
# ---------------------------------------------------------------------------

def _resolve_size_pt(context, sources, level) -> float:
    """Size cascade with the run-only normAutofit scaling rule."""
    if sources and sources[0].name == "run":
        run_size = extract_size_hundredths(sources[0].element)
        if run_size is not None:
            return scaled_size_pt(run_size, autofit_font_scale(context.shape))
        sources = sources[1:]
    hundredths = resolve_run_property(sources, level,
                                      extract_size_hundredths)
    if hundredths is not None:
        return hundredths / 100.0
    return DEFAULT_FONT_SIZE_PT


def _resolve_with_font_ref(context, sources, level, extract, from_font_ref):
    """Walk shape chain, then ``fontRef``, then deck defaults (POI order)."""
    chain, deck_defaults = split_deck_defaults(sources)
    value = resolve_run_property(chain, level, extract)
    if value is not None:
        return value
    ref_value = from_font_ref(context.shape, context.theme_context())
    if ref_value is not None:
        return ref_value
    return resolve_run_property(deck_defaults, level, extract)


def resolve_run_font(context: TextContext) -> Dict:
    """Effective font of one run: name, size, bold, italic, color.

    ``{"name": str|None, "size_pt": float, "bold": bool, "italic": bool,
    "color_hex": "RRGGBB"|None}``. See the module docstring for the
    cascade order and defaults.
    """
    if context.run is None or context.paragraph is None:
        raise ValueError("resolve_run_font requires paragraph and run")
    theme = context.theme_context()
    sources = _cascade(context, theme)
    level = _level(context)

    typeface = _resolve_with_font_ref(
        context, sources, level, extract_latin_typeface, font_ref_typeface,
    )
    chain, deck_defaults = split_deck_defaults(sources)
    color_hex = None
    chain_fill = resolve_run_property(chain, level, extract_solid_fill)
    if chain_fill is not None:
        color_hex = theme.resolve_solid_fill(chain_fill)
    else:
        color_hex = font_ref_color_hex(context.shape, theme)
        if color_hex is None:
            tail_fill = resolve_run_property(deck_defaults, level,
                                             extract_solid_fill)
            if tail_fill is not None:
                color_hex = theme.resolve_solid_fill(tail_fill)

    return {
        "name": None if typeface is None else theme.resolve_typeface(typeface),
        "size_pt": _resolve_size_pt(context, sources, level),
        "bold": bool(resolve_run_property(sources, level, extract_bold)),
        "italic": bool(resolve_run_property(sources, level, extract_italic)),
        "color_hex": color_hex,
    }


# ---------------------------------------------------------------------------
# Paragraph
# ---------------------------------------------------------------------------

def _spacing_out(tagged, pct_key: str) -> Dict:
    """Tagged extractor dict -> public form (pct fraction re-labeled)."""
    if "points" in tagged:
        return {"points": tagged["points"]}
    return {pct_key: tagged["pct"]}


def _resolve_bullet(sources, level, theme: ThemeContext) -> Dict:
    """Bullet kind + independently-cascading char/font/size/color."""
    kind = resolve_paragraph_property(sources, level, extract_bullet_type)
    if kind is None or kind == BULLET_NONE:
        return {"type": BULLET_NONE}

    bullet: Dict = {"type": kind}
    if kind == BULLET_CHAR:
        char = resolve_paragraph_property(sources, level, extract_bullet_char)
        # Malformed decks can carry ``a:buChar char=""`` -- never trust
        # file content; report no char/code instead of raising IndexError.
        bullet["char"] = char or None
        bullet["char_code"] = ord(char[0]) if char else None
    elif kind == BULLET_AUTONUM:
        autonum = resolve_paragraph_property(
            sources, level, extract_bullet_autonum
        ) or {}
        bullet["scheme"] = autonum.get("scheme")
        bullet["start_at"] = autonum.get("start_at")

    font = resolve_paragraph_property(sources, level, extract_bullet_font)
    bullet["font_follows_text"] = font is None or "follow_text" in font
    bullet["font"] = (
        None if bullet["font_follows_text"]
        else theme.resolve_typeface(font["typeface"])
    )

    size = resolve_paragraph_property(sources, level, extract_bullet_size)
    bullet["size_follows_text"] = size is None or "follow_text" in size
    bullet["size_pct"] = None if bullet["size_follows_text"] else size.get("pct")
    bullet["size_pt"] = None if bullet["size_follows_text"] else size.get("points")

    color = resolve_paragraph_property(sources, level, extract_bullet_color)
    bullet["color_follows_text"] = color is None or "follow_text" in color
    bullet["color_hex"] = (
        None if bullet["color_follows_text"]
        else theme.resolve_color(color["color_element"])
    )
    return bullet


def resolve_paragraph(context: TextContext) -> Dict:
    """Effective paragraph format: alignment, spacing, level, bullet.

    ``space_before``/``space_after`` are tagged ``{"points": pt}`` or
    ``{"lines": multiple}`` (``a:spcPct``); ``line_spacing`` is
    ``{"points": pt}`` or ``{"multiple": f}``. Bullet ``follow_text``
    flags mirror the OOXML ``bu*Tx`` defaults -- the consumer supplies
    the paragraph's text font/size/color where a flag is ``True``.
    """
    if context.paragraph is None:
        raise ValueError("resolve_paragraph requires a paragraph")
    theme = context.theme_context()
    sources = _cascade(context, theme)
    level = _level(context)

    alignment = resolve_paragraph_property(sources, level, extract_alignment)
    before = resolve_paragraph_property(sources, level, extract_space_before)
    after = resolve_paragraph_property(sources, level, extract_space_after)
    line = resolve_paragraph_property(sources, level, extract_line_spacing)

    return {
        "indent_level": level,
        "alignment": alignment or DEFAULT_ALIGNMENT,
        "space_before": (dict(DEFAULT_SPACE) if before is None
                         else _spacing_out(before, "lines")),
        "space_after": (dict(DEFAULT_SPACE) if after is None
                        else _spacing_out(after, "lines")),
        "line_spacing": (dict(DEFAULT_LINE_SPACING) if line is None
                         else _spacing_out(line, "multiple")),
        "bullet": _resolve_bullet(sources, level, theme),
    }


# ---------------------------------------------------------------------------
# Shape
# ---------------------------------------------------------------------------

def resolve_shape(shape, slide=None,
                  theme: Optional[ThemeContext] = None) -> Dict:
    """Effective line / fill / geometry of a shape.

    ``slide`` defaults to the shape's own slide part and enables
    placeholder geometry inheritance plus per-slide theme resolution;
    pass ``theme`` to reuse a prebuilt ``ThemeContext`` across shapes.
    """
    if shape is None:
        raise ValueError("shape must not be None")
    if slide is None:
        slide = getattr(shape.part, "slide", None)
        if slide is None:
            raise ValueError(
                "resolve_shape cannot derive a slide from the shape's "
                f"part ({type(shape.part).__name__}); shapes on layout/"
                "master parts need an explicit slide= argument"
            )
    if theme is None:
        theme = ThemeContext.for_slide(slide)
    geometry = resolve_shape_geometry(shape, slide)
    return {
        "geometry": {
            "preset": geometry["preset"],
            "left_pt": geometry["left_pt"],
            "top_pt": geometry["top_pt"],
            "width_pt": geometry["width_pt"],
            "height_pt": geometry["height_pt"],
        },
        "rotation_deg": geometry["rotation_deg"],
        "adjustments": geometry["adjustments"],
        "line": resolve_shape_line(shape, theme),
        "fill": resolve_shape_fill(shape, theme),
    }
