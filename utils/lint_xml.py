"""Raw-XML probes for the style lint (Step 5).

The Step 2 resolver reports EFFECTIVE values -- what PowerPoint
displays. Several lint rules additionally need to know *how* a value is
stored in the slide XML (hardcoded ``a:srgbClr`` vs theme-linked
``a:schemeClr``, explicit east-asian/complex-script/symbol typefaces,
per-run proofing language, ``txBox`` markers, autofit settings). These
small read-only probes provide exactly that; every color/typeface probe
here reads the SLIDE part only -- inherited values are template
territory and conformant by construction.
"""

from typing import Dict, Iterator, Optional, Tuple

from pptx.oxml.ns import qn

#: Explicit-typeface tags a run property can carry (Slidewise-style
#: complete coverage: latin + east-asian + complex-script + symbol).
RUN_TYPEFACE_TAGS = ("latin", "ea", "cs", "sym")


def _txbody(shape_elem):
    return shape_elem.find(qn("p:txBody"))


def iter_run_props(shape_elem) -> Iterator[Tuple[int, int, object]]:
    """Yield ``(paragraph_1based, run_1based, a:rPr)`` for explicit rPr.

    Paragraph indices count every ``a:p`` in document order (matching
    the resolver's paragraph records); runs without an ``a:rPr`` are
    skipped -- they carry no explicit properties to probe.
    """
    txbody = _txbody(shape_elem)
    if txbody is None:
        return
    for p_idx, para in enumerate(txbody.findall(qn("a:p")), start=1):
        for r_idx, run in enumerate(para.findall(qn("a:r")), start=1):
            rpr = run.find(qn("a:rPr"))
            if rpr is not None:
                yield p_idx, r_idx, rpr


def solid_srgb_of(prop_elem) -> Optional[str]:
    """Uppercase hex of a direct ``a:solidFill/a:srgbClr`` child."""
    fill = prop_elem.find(qn("a:solidFill"))
    if fill is None:
        return None
    srgb = fill.find(qn("a:srgbClr"))
    if srgb is None or not srgb.get("val"):
        return None
    return srgb.get("val").upper()


def run_lang_of(rpr) -> Optional[str]:
    """The run's proofing language (``a:rPr/@lang``), if declared."""
    return rpr.get("lang")


def run_typefaces_of(rpr) -> Dict[str, str]:
    """Explicit run typefaces by tag (latin/ea/cs/sym), literal values.

    Theme references (``+mj-lt`` / ``+mn-ea`` ...) are excluded: they
    resolve to the deck's theme fonts and are theme-linked by
    definition.
    """
    faces: Dict[str, str] = {}
    for tag in RUN_TYPEFACE_TAGS:
        node = rpr.find(qn(f"a:{tag}"))
        if node is not None:
            typeface = node.get("typeface")
            if typeface and not typeface.startswith("+"):
                faces[tag] = typeface
    return faces


def shape_fill_srgb(shape_elem) -> Optional[str]:
    """Uppercase hex of an explicit ``p:spPr`` solid srgbClr fill."""
    sppr = shape_elem.find(qn("p:spPr"))
    if sppr is None:
        return None
    return solid_srgb_of(sppr)


def shape_line_srgb(shape_elem) -> Optional[str]:
    """Uppercase hex of an explicit ``p:spPr/a:ln`` solid srgbClr."""
    sppr = shape_elem.find(qn("p:spPr"))
    if sppr is None:
        return None
    ln = sppr.find(qn("a:ln"))
    if ln is None:
        return None
    return solid_srgb_of(ln)


def is_textbox(shape_elem) -> bool:
    """True for shapes marked ``<p:cNvSpPr txBox="1">``."""
    cnvsppr = shape_elem.find(
        f"{qn('p:nvSpPr')}/{qn('p:cNvSpPr')}")
    return cnvsppr is not None and cnvsppr.get("txBox") == "1"


def autofit_of(shape_elem) -> Tuple[Optional[str], Optional[float]]:
    """``(kind, font_scale_pct)`` from the shape's own ``a:bodyPr``.

    ``kind`` is ``"normAutofit"`` / ``"spAutoFit"`` / ``None``;
    ``font_scale_pct`` is the normAutofit shrink percentage (100.0 when
    the attribute is absent, ``None`` for non-normAutofit frames).
    """
    txbody = _txbody(shape_elem)
    if txbody is None:
        return None, None
    bodypr = txbody.find(qn("a:bodyPr"))
    if bodypr is None:
        return None, None
    norm = bodypr.find(qn("a:normAutofit"))
    if norm is not None:
        raw = norm.get("fontScale")
        return "normAutofit", (100.0 if raw is None
                               else int(raw) / 1000.0)
    if bodypr.find(qn("a:spAutoFit")) is not None:
        return "spAutoFit", None
    return None, None


def wrap_of(shape_elem) -> Optional[str]:
    """The shape's ``a:bodyPr/@wrap`` (``"square"``/``"none"``/None)."""
    txbody = _txbody(shape_elem)
    if txbody is None:
        return None
    bodypr = txbody.find(qn("a:bodyPr"))
    if bodypr is None:
        return None
    return bodypr.get("wrap")
