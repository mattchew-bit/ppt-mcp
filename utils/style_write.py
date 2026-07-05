"""Write-side XML helpers for slide-level style application (Step 3).

The Step 2 resolver (``utils/resolve_*``) is strictly read-only; these
helpers are its write-side counterpart: set run font name/size/bold/
italic/color, paragraph space-before/space-after/line-spacing, bullet
char/color/size per paragraph, and shape border weight/color/dash, fill
and corner radius -- at run / paragraph / shape level ONLY.

HARD RULE (style-fidelity plan, Step 3 write strategy): never mutate
slide masters, slide layouts or theme parts of an existing deck. Every
public helper checks that its target lives on a slide part
(``/ppt/slides/...``) and raises ``ValueError`` otherwise. Master-level
styling is reached via the clone-first workflow (Step 1), never by
editing those parts in place.

python-pptx's public API is used where it exists (run font, paragraph
spacing, line format, solid fill) because its oxml layer maintains
ECMA-376 child ordering. Bullets have no python-pptx API, so they are
written with lxml against an explicit ``CT_TextParagraphProperties``
child-order table (ECMA-376 21.1.2.2.7).
"""

import re
from typing import Optional

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE, MSO_THEME_COLOR
from pptx.oxml.ns import qn
from pptx.util import Pt

#: Only parts under this prefix may ever be written to.
SLIDE_PARTNAME_PREFIX = "/ppt/slides/"

#: ECMA-376 CT_TextParagraphProperties child sequence (21.1.2.2.7).
#: Bullet elements must land between the spacing block and defRPr.
_PPR_CHILD_ORDER = (
    "a:lnSpc", "a:spcBef", "a:spcAft",
    "a:buClrTx", "a:buClr",
    "a:buSzTx", "a:buSzPct", "a:buSzPts",
    "a:buFontTx", "a:buFont",
    "a:buNone", "a:buAutoNum", "a:buChar",
    "a:tabLst", "a:defRPr", "a:extLst",
)

#: The 12 ``a:schemeClr`` slots writable as theme-linked colors.
SCHEME_TOKENS = (
    "dk1", "lt1", "dk2", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
)

#: ``a:buSzPct/@val`` is in thousandths of a percent (95% -> 95000).
_BULLET_SIZE_PCT_DENOMINATOR = 1000

#: ``a:gd fmla="val N"`` adjustment values are fractions * 100000.
_ADJUSTMENT_DENOMINATOR = 100_000

_HEX_COLOR_RE = re.compile(r"\A[0-9A-Fa-f]{6}\Z")


# ---------------------------------------------------------------------------
# Guards / small shared validators
# ---------------------------------------------------------------------------

def ensure_slide_part(target) -> None:
    """Raise unless ``target`` (shape / paragraph / run proxy) lives on a
    slide part. This is the never-touch-masters hard rule."""
    if target is None:
        raise ValueError("style_write target must not be None")
    part = getattr(target, "part", None)
    if part is None:
        raise ValueError(
            f"style_write target {type(target).__name__} exposes no part"
        )
    partname = str(part.partname)
    if not partname.startswith(SLIDE_PARTNAME_PREFIX):
        raise ValueError(
            f"style_write refuses to modify {partname}: only slide parts "
            f"({SLIDE_PARTNAME_PREFIX}...) may be written -- masters, "
            "layouts and themes are never mutated (clone-first workflow "
            "is the sanctioned route to master-level styling)"
        )


def _clean_hex(value: str, what: str) -> str:
    """``'#RRGGBB'``/``'RRGGBB'`` -> ``'RRGGBB'`` upper, or ValueError."""
    if not isinstance(value, str):
        raise ValueError(f"{what} must be a hex color string: {value!r}")
    candidate = value.lstrip("#")
    if not _HEX_COLOR_RE.match(candidate):
        raise ValueError(f"{what} is not a RRGGBB hex color: {value!r}")
    return candidate.upper()


def _check_scheme_token(token: str) -> str:
    if token not in SCHEME_TOKENS:
        raise ValueError(
            f"unknown scheme color token {token!r}; valid: {SCHEME_TOKENS}"
        )
    return token


def _one_color_arg(hex_value: Optional[str], scheme: Optional[str],
                   what: str) -> None:
    if (hex_value is None) == (scheme is None):
        raise ValueError(
            f"{what} requires exactly one of hex_value= or scheme="
        )


# ---------------------------------------------------------------------------
# Run-level writes
# ---------------------------------------------------------------------------

def set_run_font(run, name: Optional[str] = None,
                 size_pt: Optional[float] = None,
                 bold: Optional[bool] = None,
                 italic: Optional[bool] = None) -> None:
    """Write explicit ``a:rPr`` font properties on one run.

    Only the arguments provided are written; ``None`` leaves the
    existing (possibly inherited) value alone.
    """
    ensure_slide_part(run)
    if name is None and size_pt is None and bold is None and italic is None:
        raise ValueError("set_run_font called with nothing to set")
    if name is not None:
        if not isinstance(name, str) or not name.strip():
            raise ValueError(f"font name must be a non-empty string: {name!r}")
        run.font.name = name
    if size_pt is not None:
        if not isinstance(size_pt, (int, float)) or size_pt <= 0:
            raise ValueError(f"size_pt must be a positive number: {size_pt!r}")
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bool(bold)
    if italic is not None:
        run.font.italic = bool(italic)


def set_run_color(run, hex_value: Optional[str] = None,
                  scheme: Optional[str] = None) -> None:
    """Set a run's solid text color: explicit hex OR a theme token.

    ``scheme`` writes ``<a:schemeClr val="...">`` so the run stays
    theme-linked; ``hex_value`` writes ``<a:srgbClr>``.
    """
    ensure_slide_part(run)
    _one_color_arg(hex_value, scheme, "set_run_color")
    if scheme is not None:
        run.font.color.theme_color = MSO_THEME_COLOR.from_xml(
            _check_scheme_token(scheme))
    else:
        run.font.color.rgb = RGBColor.from_string(
            _clean_hex(hex_value, "run color"))


# ---------------------------------------------------------------------------
# Paragraph-level writes
# ---------------------------------------------------------------------------

def set_paragraph_spacing(paragraph,
                          space_before_pt: Optional[float] = None,
                          space_after_pt: Optional[float] = None,
                          line_spacing: Optional[float] = None) -> None:
    """Write ``a:pPr`` spacing on one paragraph (points / multiple).

    ``line_spacing`` is a multiple (1.2 -> ``spcPct val="120000"``).
    Only the arguments provided are written.
    """
    ensure_slide_part(paragraph)
    if (space_before_pt is None and space_after_pt is None
            and line_spacing is None):
        raise ValueError("set_paragraph_spacing called with nothing to set")
    for label, value in (("space_before_pt", space_before_pt),
                         ("space_after_pt", space_after_pt),
                         ("line_spacing", line_spacing)):
        if value is not None and (not isinstance(value, (int, float))
                                  or value < 0):
            raise ValueError(f"{label} must be a non-negative number: "
                             f"{value!r}")
    if space_before_pt is not None:
        paragraph.space_before = Pt(space_before_pt)
    if space_after_pt is not None:
        paragraph.space_after = Pt(space_after_pt)
    if line_spacing is not None:
        paragraph.line_spacing = float(line_spacing)


def _insert_ppr_child(p_pr: etree._Element, tag: str) -> etree._Element:
    """Get-or-create ``tag`` in ``a:pPr`` at its ECMA-376 position."""
    existing = p_pr.find(qn(tag))
    if existing is not None:
        return existing
    element = p_pr.makeelement(qn(tag), {})
    successors = _PPR_CHILD_ORDER[_PPR_CHILD_ORDER.index(tag) + 1:]
    for successor_tag in successors:
        successor = p_pr.find(qn(successor_tag))
        if successor is not None:
            successor.addprevious(element)
            return element
    p_pr.append(element)
    return element


def _remove_ppr_children(p_pr: etree._Element, *tags: str) -> None:
    for tag in tags:
        element = p_pr.find(qn(tag))
        if element is not None:
            p_pr.remove(element)


def _set_bullet_char(p_pr: etree._Element, char: str) -> None:
    if not isinstance(char, str) or len(char) != 1:
        raise ValueError(f"bullet char must be a single character: {char!r}")
    _remove_ppr_children(p_pr, "a:buNone", "a:buAutoNum")
    _insert_ppr_child(p_pr, "a:buChar").set("char", char)


def _set_bullet_color(p_pr: etree._Element, hex_value: Optional[str],
                      scheme: Optional[str]) -> None:
    _remove_ppr_children(p_pr, "a:buClrTx")
    bu_clr = _insert_ppr_child(p_pr, "a:buClr")
    for child in list(bu_clr):
        bu_clr.remove(child)
    if scheme is not None:
        color = bu_clr.makeelement(qn("a:schemeClr"),
                                   {"val": _check_scheme_token(scheme)})
    else:
        color = bu_clr.makeelement(
            qn("a:srgbClr"), {"val": _clean_hex(hex_value, "bullet color")})
    bu_clr.append(color)


def _set_bullet_size_pct(p_pr: etree._Element, size_pct: float) -> None:
    if not isinstance(size_pct, (int, float)) or not 25 <= size_pct <= 400:
        raise ValueError(
            f"bullet size_pct must be 25..400 percent: {size_pct!r}")
    _remove_ppr_children(p_pr, "a:buSzTx", "a:buSzPts")
    _insert_ppr_child(p_pr, "a:buSzPct").set(
        "val", str(int(round(size_pct * _BULLET_SIZE_PCT_DENOMINATOR))))


def set_paragraph_bullet(paragraph, char: Optional[str] = None,
                         color_hex: Optional[str] = None,
                         color_scheme: Optional[str] = None,
                         size_pct: Optional[float] = None,
                         font: Optional[str] = None) -> None:
    """Write explicit bullet formatting on one paragraph's ``a:pPr``.

    No python-pptx API exists for bullets; children are placed at their
    ECMA-376 positions. Setting ``char`` removes a conflicting
    ``buNone``/``buAutoNum``. ``size_pct`` is percent of the text size
    (95 -> ``buSzPct val="95000"``). Only provided arguments are written.
    """
    ensure_slide_part(paragraph)
    if (char is None and color_hex is None and color_scheme is None
            and size_pct is None and font is None):
        raise ValueError("set_paragraph_bullet called with nothing to set")
    if color_hex is not None and color_scheme is not None:
        raise ValueError(
            "set_paragraph_bullet takes color_hex= or color_scheme=, "
            "not both")
    p_pr = paragraph._p.get_or_add_pPr()
    if char is not None:
        _set_bullet_char(p_pr, char)
    if color_hex is not None or color_scheme is not None:
        _set_bullet_color(p_pr, color_hex, color_scheme)
    if size_pct is not None:
        _set_bullet_size_pct(p_pr, size_pct)
    if font is not None:
        if not isinstance(font, str) or not font.strip():
            raise ValueError(f"bullet font must be a non-empty string: "
                             f"{font!r}")
        _remove_ppr_children(p_pr, "a:buFontTx")
        _insert_ppr_child(p_pr, "a:buFont").set("typeface", font)


# ---------------------------------------------------------------------------
# Shape-level writes
# ---------------------------------------------------------------------------

def set_shape_border(shape, weight_pt: Optional[float] = None,
                     color_hex: Optional[str] = None,
                     color_scheme: Optional[str] = None,
                     dash: Optional[str] = None) -> None:
    """Write explicit ``a:ln`` outline properties on one shape.

    ``dash`` is an ECMA ``ST_PresetLineDashVal`` token (``"solid"``,
    ``"dash"``, ``"sysDash"``, ...). Only provided arguments are written.
    """
    ensure_slide_part(shape)
    if (weight_pt is None and color_hex is None and color_scheme is None
            and dash is None):
        raise ValueError("set_shape_border called with nothing to set")
    if color_hex is not None and color_scheme is not None:
        raise ValueError(
            "set_shape_border takes color_hex= or color_scheme=, not both")
    if weight_pt is not None:
        if not isinstance(weight_pt, (int, float)) or weight_pt <= 0:
            raise ValueError(
                f"border weight_pt must be a positive number: {weight_pt!r}")
        shape.line.width = Pt(weight_pt)
    if color_scheme is not None:
        shape.line.color.theme_color = MSO_THEME_COLOR.from_xml(
            _check_scheme_token(color_scheme))
    elif color_hex is not None:
        shape.line.color.rgb = RGBColor.from_string(
            _clean_hex(color_hex, "border color"))
    if dash is not None:
        try:
            style = MSO_LINE_DASH_STYLE.from_xml(dash)
        except (KeyError, ValueError) as exc:
            raise ValueError(f"unknown border dash token {dash!r}") from exc
        shape.line.dash_style = style


def set_shape_fill(shape, hex_value: Optional[str] = None,
                   scheme: Optional[str] = None) -> None:
    """Set a shape's solid fill: explicit hex OR a theme token."""
    ensure_slide_part(shape)
    _one_color_arg(hex_value, scheme, "set_shape_fill")
    shape.fill.solid()
    if scheme is not None:
        shape.fill.fore_color.theme_color = MSO_THEME_COLOR.from_xml(
            _check_scheme_token(scheme))
    else:
        shape.fill.fore_color.rgb = RGBColor.from_string(
            _clean_hex(hex_value, "fill color"))


def set_shape_corner_radius(shape, fraction: float) -> None:
    """Set a ``roundRect`` shape's corner-radius adjustment fraction.

    ``fraction`` matches the resolver's / COM's reporting convention
    (``a:gd fmla="val 12000"`` == 0.12). Raises on non-roundRect shapes
    -- corner radius is meaningless there and silence would hide bugs.
    """
    ensure_slide_part(shape)
    if not isinstance(fraction, (int, float)) or not 0 <= fraction <= 0.5:
        raise ValueError(
            f"corner-radius fraction must be within 0..0.5: {fraction!r}")
    sp_pr = shape._element.find(qn("p:spPr"))
    geom = None if sp_pr is None else sp_pr.find(qn("a:prstGeom"))
    if geom is None or geom.get("prst") != "roundRect":
        preset = None if geom is None else geom.get("prst")
        raise ValueError(
            f"set_shape_corner_radius requires a roundRect shape; "
            f"{shape.name!r} has preset {preset!r}")
    av_lst = geom.find(qn("a:avLst"))
    if av_lst is None:
        av_lst = geom.makeelement(qn("a:avLst"), {})
        geom.append(av_lst)
    for gd in list(av_lst):
        av_lst.remove(gd)
    gd = av_lst.makeelement(qn("a:gd"), {
        "name": "adj",
        "fmla": f"val {int(round(fraction * _ADJUSTMENT_DENOMINATOR))}",
    })
    av_lst.append(gd)
