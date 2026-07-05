"""Per-property extractors for the effective-style text cascade (read-only).

Small pure functions handed to ``utils.resolve_core.resolve_run_property``
/ ``resolve_paragraph_property``: each receives one rPr/defRPr-like or
pPr-like lxml element from a cascade source and returns the property's
explicit value there, or ``None`` so the walker keeps descending. This is
the per-property-lambda half of the Apache POI fetcher architecture
(Apache-2.0, re-implemented); the extraction details port ShapeCrawler's
verified lookups (MIT -- ``PortionFontSize``, ``IndentFonts``).

Value conventions (documented so the fixture adapter stays thin):
    * font size        -- OOXML hundredths of a point (``sz``), callers
      divide by 100
    * spacing          -- tagged dicts: ``{"points": pt}`` for
      ``a:spcPts`` (val is hundredths of a point) or ``{"pct": f}`` for
      ``a:spcPct`` (val is thousandths of a percent, so 115000 -> 1.15)
    * bullets          -- tagged dicts, with ``{"follow_text": True}``
      for the ``bu*Tx`` "inherit from the paragraph text" markers
    * booleans         -- OOXML ``"1"/"true"`` / ``"0"/"false"``

The ``a:normAutofit`` ``fontScale`` helper implements ShapeCrawler's
scaling rule: it applies ONLY to sizes that came from the run's own
``a:rPr`` (deeper cascade levels store unscaled template sizes).

Everything here is pure and never mutates any element.
"""

from typing import Dict, List, Optional, Union

from lxml import etree
from pptx.oxml.ns import qn

#: ST_Percentage denominator (100000 == 100%).
PERCENT_DENOMINATOR = 100000.0

#: ``a:spcPts``/``sz`` style values are hundredths of a point.
HUNDREDTHS_PER_POINT = 100.0

#: ``a:pPr/@algn`` -> resolver alignment token.
ALIGNMENT_TOKENS: Dict[str, str] = {
    "l": "left",
    "ctr": "center",
    "r": "right",
    "just": "justify",
    "dist": "distribute",
    "thaiDist": "thai_distribute",
    "justLow": "justify_low",
}

_XML_TRUE = frozenset({"1", "true"})
_XML_FALSE = frozenset({"0", "false"})

TaggedValue = Dict[str, Union[float, bool, str, etree._Element]]


def _require_element(element, name: str) -> None:
    if element is None:
        raise ValueError(f"{name} must not be None")


def _int_attr(element: etree._Element, attr: str) -> int:
    """Required integer attribute; explicit ValueError when absent.

    Never trust file content: a malformed deck may drop a mandatory
    attribute (e.g. ``a:spcPts`` without ``val``), which must surface as
    a clear ``ValueError``, not a bare ``TypeError`` from ``int(None)``.
    """
    raw = element.get(attr)
    if raw is None:
        tag = element.tag.rsplit("}", 1)[-1]
        raise ValueError(
            f"<{tag}> element is missing required attribute {attr!r}"
        )
    return int(raw)


def xml_bool(raw: Optional[str]) -> Optional[bool]:
    """OOXML boolean attribute -> ``True``/``False``, ``None`` if absent."""
    if raw is None:
        return None
    if raw in _XML_TRUE:
        return True
    if raw in _XML_FALSE:
        return False
    raise ValueError(f"invalid OOXML boolean: {raw!r}")


# ---------------------------------------------------------------------------
# Run (character) property extractors -- receive rPr/defRPr-like elements
# ---------------------------------------------------------------------------

def extract_size_hundredths(rpr: etree._Element) -> Optional[int]:
    """``@sz`` in hundredths of a point, or ``None``."""
    raw = rpr.get("sz")
    return None if raw is None else int(raw)


def extract_latin_typeface(rpr: etree._Element) -> Optional[str]:
    """``a:latin/@typeface`` (may be a ``+mj-lt`` theme token)."""
    latin = rpr.find(qn("a:latin"))
    return None if latin is None else latin.get("typeface")


def extract_bold(rpr: etree._Element) -> Optional[bool]:
    """``@b`` as a bool; explicit ``0`` stops the cascade with False."""
    return xml_bool(rpr.get("b"))


def extract_italic(rpr: etree._Element) -> Optional[bool]:
    """``@i`` as a bool; explicit ``0`` stops the cascade with False."""
    return xml_bool(rpr.get("i"))


def extract_solid_fill(rpr: etree._Element) -> Optional[etree._Element]:
    """The ``a:solidFill`` child carrying a color (text color), or ``None``.

    A childless ``a:solidFill`` is schema-valid (the color choice is
    optional in CT_SolidColorFillProperties) and python-pptx emits one
    as a side effect of merely READING ``run.font.color``; it specifies
    no color, so the cascade keeps descending.
    """
    solid_fill = rpr.find(qn("a:solidFill"))
    if solid_fill is None or len(solid_fill) == 0:
        return None
    return solid_fill


# ---------------------------------------------------------------------------
# Paragraph property extractors -- receive pPr-like elements
# ---------------------------------------------------------------------------

def extract_alignment(ppr: etree._Element) -> Optional[str]:
    """``@algn`` mapped to a resolver token (``left``/``center``/...)."""
    raw = ppr.get("algn")
    if raw is None:
        return None
    token = ALIGNMENT_TOKENS.get(raw)
    if token is None:
        raise ValueError(f"unknown a:pPr/@algn value: {raw!r}")
    return token


def _extract_spacing(ppr: etree._Element, tag: str) -> Optional[TaggedValue]:
    """Shared ``a:spcBef``/``a:spcAft``/``a:lnSpc`` reader (tagged dict)."""
    spacing = ppr.find(qn(tag))
    if spacing is None:
        return None
    points = spacing.find(qn("a:spcPts"))
    if points is not None:
        return {"points": _int_attr(points, "val") / HUNDREDTHS_PER_POINT}
    percent = spacing.find(qn("a:spcPct"))
    if percent is not None:
        return {"pct": _int_attr(percent, "val") / PERCENT_DENOMINATOR}
    return None


def extract_space_before(ppr: etree._Element) -> Optional[TaggedValue]:
    """``a:spcBef`` -> ``{"points": pt}`` or ``{"pct": fraction}``."""
    return _extract_spacing(ppr, "a:spcBef")


def extract_space_after(ppr: etree._Element) -> Optional[TaggedValue]:
    """``a:spcAft`` -> ``{"points": pt}`` or ``{"pct": fraction}``."""
    return _extract_spacing(ppr, "a:spcAft")


def extract_line_spacing(ppr: etree._Element) -> Optional[TaggedValue]:
    """``a:lnSpc`` -> ``{"points": pt}`` or ``{"pct": multiple}``."""
    return _extract_spacing(ppr, "a:lnSpc")


# ---------------------------------------------------------------------------
# Bullet extractors -- receive pPr-like elements
# ---------------------------------------------------------------------------

BULLET_NONE = "none"
BULLET_CHAR = "char"
BULLET_AUTONUM = "autonum"


def extract_bullet_type(ppr: etree._Element) -> Optional[str]:
    """First explicit bullet *kind* on this pPr, or ``None``.

    Only ``a:buNone`` / ``a:buChar`` / ``a:buAutoNum`` decide the kind;
    ``buFont``/``buSzPct``/``buClr`` alone do not (they cascade
    independently, matching PowerPoint behavior).
    """
    if ppr.find(qn("a:buNone")) is not None:
        return BULLET_NONE
    if ppr.find(qn("a:buChar")) is not None:
        return BULLET_CHAR
    if ppr.find(qn("a:buAutoNum")) is not None:
        return BULLET_AUTONUM
    return None


def extract_bullet_char(ppr: etree._Element) -> Optional[str]:
    """``a:buChar/@char``, or ``None``."""
    bu_char = ppr.find(qn("a:buChar"))
    return None if bu_char is None else bu_char.get("char")


def extract_bullet_autonum(ppr: etree._Element) -> Optional[TaggedValue]:
    """``a:buAutoNum`` -> ``{"scheme": ..., "start_at": int|None}``."""
    autonum = ppr.find(qn("a:buAutoNum"))
    if autonum is None:
        return None
    start = autonum.get("startAt")
    return {
        "scheme": autonum.get("type"),
        "start_at": None if start is None else int(start),
    }


def extract_bullet_font(ppr: etree._Element) -> Optional[TaggedValue]:
    """``a:buFont``/``a:buFontTx`` -> tagged typeface or follow-text."""
    if ppr.find(qn("a:buFontTx")) is not None:
        return {"follow_text": True}
    bu_font = ppr.find(qn("a:buFont"))
    if bu_font is None:
        return None
    typeface = bu_font.get("typeface")
    return None if typeface is None else {"typeface": typeface}


def extract_bullet_size(ppr: etree._Element) -> Optional[TaggedValue]:
    """``a:buSzPct``/``a:buSzPts``/``a:buSzTx`` -> tagged size."""
    if ppr.find(qn("a:buSzTx")) is not None:
        return {"follow_text": True}
    percent = ppr.find(qn("a:buSzPct"))
    if percent is not None:
        return {"pct": _int_attr(percent, "val") / PERCENT_DENOMINATOR}
    points = ppr.find(qn("a:buSzPts"))
    if points is not None:
        return {"points": _int_attr(points, "val") / HUNDREDTHS_PER_POINT}
    return None


def extract_bullet_color(ppr: etree._Element) -> Optional[TaggedValue]:
    """``a:buClr``/``a:buClrTx`` -> tagged color element or follow-text."""
    if ppr.find(qn("a:buClrTx")) is not None:
        return {"follow_text": True}
    bu_clr = ppr.find(qn("a:buClr"))
    if bu_clr is None:
        return None
    for child in bu_clr:
        return {"color_element": child}  # the single color child
    return None


# ---------------------------------------------------------------------------
# normAutofit fontScale (ShapeCrawler PortionFontSize semantics)
# ---------------------------------------------------------------------------

def autofit_font_scale(shape) -> Optional[float]:
    """The shape's ``a:normAutofit/@fontScale`` as a fraction, or ``None``.

    Reads the shape's own ``p:txBody/a:bodyPr``. Accepts both the modern
    thousandths-of-a-percent form (``"62500"``) and the legacy percent
    string (``"62.5%"``). Returns ``None`` when there is no normAutofit
    or it carries no fontScale (scale 100%).
    """
    _require_element(shape, "shape")
    body_pr = shape._element.find(
        qn("p:txBody") + "/" + qn("a:bodyPr")
    )
    if body_pr is None:
        return None
    autofit = body_pr.find(qn("a:normAutofit"))
    if autofit is None:
        return None
    raw = autofit.get("fontScale")
    if raw is None:
        return None
    if raw.endswith("%"):
        return float(raw[:-1]) / 100.0
    return int(raw) / PERCENT_DENOMINATOR


def scaled_size_pt(size_hundredths: int, font_scale: Optional[float]) -> float:
    """Hundredths-of-a-point size -> points, with optional autofit scale."""
    if not isinstance(size_hundredths, int) or size_hundredths <= 0:
        raise ValueError(f"invalid size in hundredths: {size_hundredths!r}")
    size_pt = size_hundredths / HUNDREDTHS_PER_POINT
    if font_scale is None:
        return size_pt
    if not 0.0 < font_scale <= 1.0:
        raise ValueError(f"invalid normAutofit fontScale: {font_scale!r}")
    return size_pt * font_scale


# ---------------------------------------------------------------------------
# Source-list partition helper (fontRef insertion point)
# ---------------------------------------------------------------------------

#: Cascade-source names that constitute the deck-default tail.
DECK_DEFAULT_SOURCE_NAMES = frozenset({
    "presentation-defaultTextStyle",
    "theme-txDef",
})


def split_deck_defaults(sources) -> tuple:
    """Split cascade sources into (shape-chain, deck-default tail).

    The shape ``p:style/a:fontRef`` layer is consulted between the two
    (Apache POI order -- after the full placeholder chain, before
    presentation/theme defaults; see ``utils.resolve_utils``).
    """
    chain: List = []
    tail: List = []
    for source in sources:
        if source.name in DECK_DEFAULT_SOURCE_NAMES:
            tail.append(source)
        else:
            chain.append(source)
    return tuple(chain), tuple(tail)
