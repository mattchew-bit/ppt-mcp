"""Per-slide theme engine for the effective-style resolver (read-only).

Resolves the theme that actually governs a slide by walking the OPC
relationship chain -- slide -> its ``sldLayout`` -> its ``sldMaster`` ->
*that master's* theme part. Real decks carry several masters, each with
its own theme part (``multi_master.pptx`` in the test fixtures has two),
so nothing here ever assumes a global ``theme1.xml``.

What a ``ThemeContext`` exposes per theme part:
    * ``color_scheme``    -- the 12 ``a:clrScheme`` slots resolved to hex
    * ``font_scheme``     -- major/minor latin/ea/cs typefaces
    * ``format_scheme``   -- the raw ``a:fmtScheme`` element plus its
      fill / line / effect / background-fill style lists (deep resolution
      of those lists is a later stage; Stage 1 exposes the elements)
    * ``clr_map``         -- the master's ``p:clrMap`` remapping (e.g.
      ``tx1 -> dk1``); ``scheme_color_hex`` applies the indirection
    * ``resolve_color``   -- any DrawingML color element (``a:srgbClr``,
      ``a:schemeClr``, ``a:sysClr``, ``a:prstClr``, ...) with lumMod /
      lumOff / tint / shade / satMod / hueMod transforms applied ->
      final ``RRGGBB`` uppercase hex (math in ``utils.resolve_colors``)
    * ``theme_font``      -- ``+mj-lt`` / ``+mn-lt`` / ``+mj-ea`` / etc.
      token -> concrete typeface (ShapeCrawler ``TextPortionFont``
      semantics, re-implemented)
    * ``text_default_list_style`` -- the theme part's
      ``a:objectDefaults/a:txDef/a:lstStyle`` for the tail of the
      non-placeholder text cascade

Precedence note (documents the ShapeCrawler-vs-POI disagreement): this
engine resolves the ``txDef`` object defaults from the *slide's master's*
theme part, not from the presentation part's theme rel that ShapeCrawler
uses -- in a multi-master deck the master chain is the one PowerPoint
follows. The cascade *order* in which txDef is consulted (after
``p:defaultTextStyle``) is documented in ``utils.resolve_core``; the
Step 0 fixtures are the arbiter for both decisions.

Ported semantics (re-implemented, not copied):
    * ShapeCrawler (MIT): ``PresentationColor`` / ``HexParser`` schemeClr
      -> clrMap -> clrScheme resolution; ``TextPortionFont`` +mj-lt /
      +mn-lt theme-font mapping.

Strictly read-only: no presentation part or element is ever mutated.
"""

from typing import Dict, List, Optional

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

from .resolve_colors import (
    apply_color_transforms,
    find_color_child,
    normalize_hex,
    resolve_raw_color,
    system_color_hex,
)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

#: The 12 clrScheme slot names, in schema order.
COLOR_SCHEME_SLOTS = (
    "dk1", "lt1", "dk2", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
)

#: schemeClr values that go through the master's clrMap before hitting
#: the clrScheme (ECMA-376 19.3.1.6).
_CLR_MAP_SLOTS = ("bg1", "tx1", "bg2", "tx2")

#: Identity clrMap used when a master carries no ``p:clrMap``.
_DEFAULT_CLR_MAP: Dict[str, str] = {
    "bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2",
    "accent1": "accent1", "accent2": "accent2", "accent3": "accent3",
    "accent4": "accent4", "accent5": "accent5", "accent6": "accent6",
    "hlink": "hlink", "folHlink": "folHlink",
}

#: Theme font tokens -> (major|minor, latin|ea|cs).
_FONT_TOKENS: Dict[str, tuple] = {
    "+mj-lt": ("major", "latin"), "+mn-lt": ("minor", "latin"),
    "+mj-ea": ("major", "ea"), "+mn-ea": ("minor", "ea"),
    "+mj-cs": ("major", "cs"), "+mn-cs": ("minor", "cs"),
}


# ---------------------------------------------------------------------------
# Part navigation
# ---------------------------------------------------------------------------

def master_for_slide(slide):
    """python-pptx ``SlideMaster`` governing a slide (via its layout)."""
    if slide is None:
        raise ValueError("slide must not be None")
    return slide.slide_layout.slide_master


def theme_element_for_master(slide_master) -> etree._Element:
    """Parse and return the ``a:theme`` root of a master's theme part.

    Follows the master part's THEME relationship -- never a hardcoded
    ``theme1.xml`` -- so each master resolves to its own theme.
    """
    if slide_master is None:
        raise ValueError("slide_master must not be None")
    try:
        theme_part = slide_master.part.part_related_by(RT.THEME)
    except KeyError as exc:
        raise ValueError(
            f"slide master {slide_master.part.partname} has no theme "
            "relationship"
        ) from exc
    return etree.fromstring(theme_part.blob)


def presentation_default_text_style(slide) -> Optional[etree._Element]:
    """The deck-level ``p:defaultTextStyle`` element, or ``None``.

    Lives on ``ppt/presentation.xml``; used by the non-placeholder text
    cascade (floating text boxes).
    """
    if slide is None:
        raise ValueError("slide must not be None")
    presentation_part = slide.part.package.main_document_part
    return presentation_part._element.find(qn("p:defaultTextStyle"))


# ---------------------------------------------------------------------------
# ThemeContext
# ---------------------------------------------------------------------------

class ThemeContext:
    """Resolved theme + clrMap for one slide master (read-only).

    Build with ``ThemeContext.for_slide(slide)`` (resolves the slide's
    governing master through the relationship chain) or
    ``ThemeContext.for_master(slide_master)``.
    """

    def __init__(self, theme_element: etree._Element,
                 master_element: etree._Element):
        if theme_element is None or master_element is None:
            raise ValueError("theme_element and master_element are required")
        self._theme = theme_element
        self._master = master_element
        theme_elements = theme_element.find(qn("a:themeElements"))
        if theme_elements is None:
            raise ValueError("theme part has no a:themeElements")
        self._theme_elements = theme_elements
        self._clr_map = self._parse_clr_map()
        self._color_scheme = self._parse_color_scheme()
        self._font_scheme = self._parse_font_scheme()

    # -- construction ------------------------------------------------------

    @classmethod
    def for_slide(cls, slide) -> "ThemeContext":
        """ThemeContext for the master that governs ``slide``."""
        master = master_for_slide(slide)
        return cls.for_master(master)

    @classmethod
    def for_master(cls, slide_master) -> "ThemeContext":
        """ThemeContext for a python-pptx ``SlideMaster``."""
        theme_element = theme_element_for_master(slide_master)
        return cls(theme_element, slide_master.element)

    # -- parsing -----------------------------------------------------------

    def _parse_clr_map(self) -> Dict[str, str]:
        clr_map_element = self._master.find(qn("p:clrMap"))
        if clr_map_element is None:
            return dict(_DEFAULT_CLR_MAP)
        mapping = dict(_DEFAULT_CLR_MAP)
        for key in mapping:
            value = clr_map_element.get(key)
            if value is not None:
                mapping[key] = value
        return mapping

    def _parse_color_scheme(self) -> Dict[str, str]:
        scheme_element = self._theme_elements.find(qn("a:clrScheme"))
        if scheme_element is None:
            raise ValueError("theme has no a:clrScheme")
        scheme: Dict[str, str] = {}
        for slot in COLOR_SCHEME_SLOTS:
            slot_element = scheme_element.find(qn(f"a:{slot}"))
            if slot_element is None:
                raise ValueError(f"a:clrScheme is missing slot {slot!r}")
            scheme[slot] = self._scheme_slot_hex(slot_element, slot)
        return scheme

    @staticmethod
    def _scheme_slot_hex(slot_element: etree._Element, slot: str) -> str:
        srgb = slot_element.find(qn("a:srgbClr"))
        if srgb is not None:
            return normalize_hex(srgb.get("val"))
        sys_clr = slot_element.find(qn("a:sysClr"))
        if sys_clr is not None:
            return system_color_hex(sys_clr)
        raise ValueError(f"clrScheme slot {slot!r} has no srgbClr/sysClr")

    def _parse_font_scheme(self) -> Dict[str, Dict[str, str]]:
        font_scheme_element = self._theme_elements.find(qn("a:fontScheme"))
        if font_scheme_element is None:
            raise ValueError("theme has no a:fontScheme")
        parsed: Dict[str, Dict[str, str]] = {}
        for group, tag in (("major", "a:majorFont"), ("minor", "a:minorFont")):
            group_element = font_scheme_element.find(qn(tag))
            if group_element is None:
                raise ValueError(f"a:fontScheme is missing <{tag}>")
            fonts: Dict[str, str] = {}
            for script in ("latin", "ea", "cs"):
                script_element = group_element.find(qn(f"a:{script}"))
                fonts[script] = (
                    script_element.get("typeface", "")
                    if script_element is not None else ""
                )
            parsed[group] = fonts
        return parsed

    # -- read-only views ---------------------------------------------------

    @property
    def theme_element(self) -> etree._Element:
        """The parsed ``a:theme`` root element."""
        return self._theme

    @property
    def master_element(self) -> etree._Element:
        """The governing ``p:sldMaster`` element."""
        return self._master

    @property
    def theme_name(self) -> str:
        """The theme's ``name`` attribute (may be empty)."""
        return self._theme.get("name", "")

    @property
    def clr_map(self) -> Dict[str, str]:
        """Copy of the effective clrMap (identity where unmapped)."""
        return dict(self._clr_map)

    @property
    def color_scheme(self) -> Dict[str, str]:
        """Copy of the 12 clrScheme slots as ``RRGGBB`` hex."""
        return dict(self._color_scheme)

    @property
    def font_scheme(self) -> Dict[str, Dict[str, str]]:
        """Copy of ``{'major'|'minor': {'latin'|'ea'|'cs': typeface}}``."""
        return {group: dict(fonts) for group, fonts in self._font_scheme.items()}

    @property
    def format_scheme(self) -> Optional[etree._Element]:
        """The raw ``a:fmtScheme`` element (fills/lines/effects), if any."""
        return self._theme_elements.find(qn("a:fmtScheme"))

    def fill_styles(self) -> List[etree._Element]:
        """Children of ``a:fillStyleLst`` (theme default fills)."""
        return self._format_scheme_list("a:fillStyleLst")

    def line_styles(self) -> List[etree._Element]:
        """Children of ``a:lnStyleLst`` (theme default lines)."""
        return self._format_scheme_list("a:lnStyleLst")

    def effect_styles(self) -> List[etree._Element]:
        """Children of ``a:effectStyleLst`` (theme default effects)."""
        return self._format_scheme_list("a:effectStyleLst")

    def background_fill_styles(self) -> List[etree._Element]:
        """Children of ``a:bgFillStyleLst`` (theme background fills)."""
        return self._format_scheme_list("a:bgFillStyleLst")

    def _format_scheme_list(self, tag: str) -> List[etree._Element]:
        scheme = self.format_scheme
        if scheme is None:
            return []
        container = scheme.find(qn(tag))
        return list(container) if container is not None else []

    # -- color resolution --------------------------------------------------

    def scheme_color_hex(self, name: str) -> str:
        """Scheme color name -> ``RRGGBB`` hex, with clrMap indirection.

        Accepts direct slots (``dk1``, ``accent3``, ...), mapped slots
        (``bg1``/``tx1``/``bg2``/``tx2`` go through the master's clrMap
        first) and the ``folHlink`` name. ``phClr`` cannot be resolved
        without a caller-supplied placeholder color and raises.
        """
        if not isinstance(name, str) or not name:
            raise ValueError(f"invalid scheme color name: {name!r}")
        if name == "phClr":
            raise ValueError(
                "a:schemeClr val='phClr' needs the style-matrix reference "
                "color; resolve it at the shape-style layer"
            )
        if name in _CLR_MAP_SLOTS:
            name = self._clr_map.get(name, name)
        if name not in self._color_scheme:
            raise ValueError(f"unknown scheme color name: {name!r}")
        return self._color_scheme[name]

    def resolve_color(self, color_element: etree._Element) -> str:
        """Any DrawingML color element -> final ``RRGGBB`` uppercase hex.

        Handles ``a:schemeClr`` (through clrMap + clrScheme) and all raw
        color forms, applying child transforms (tint / shade / lumMod /
        lumOff / satMod / hueMod) per ``utils.resolve_colors``.
        """
        if color_element is None:
            raise ValueError("color_element must not be None")
        raw = resolve_raw_color(color_element)  # None => schemeClr
        if raw is not None:
            return raw
        scheme_name = color_element.get("val")
        if scheme_name is None:
            raise ValueError("a:schemeClr has no val attribute")
        base = self.scheme_color_hex(scheme_name)
        return apply_color_transforms(base, list(color_element))

    def resolve_solid_fill(self, solid_fill_element: etree._Element) -> str:
        """``a:solidFill``-like element -> hex of its color child."""
        color_child = find_color_child(solid_fill_element)
        if color_child is None:
            raise ValueError(
                f"<{solid_fill_element.tag}> has no color child element"
            )
        return self.resolve_color(color_child)

    # -- fonts -------------------------------------------------------------

    def theme_font(self, token: str) -> Optional[str]:
        """``+mj-lt``-style token -> concrete typeface, else ``None``.

        Non-token typeface strings (already concrete, e.g. ``Georgia``)
        return ``None`` so callers know no substitution happened; use
        ``resolve_typeface`` to always get a usable name.
        """
        mapping = _FONT_TOKENS.get(token)
        if mapping is None:
            return None
        group, script = mapping
        typeface = self._font_scheme[group][script]
        return typeface or None

    def resolve_typeface(self, typeface: str) -> str:
        """Typeface or theme token -> concrete typeface name.

        ``+mj-*`` / ``+mn-*`` tokens are substituted from the font
        scheme; anything else passes through unchanged.
        """
        if not isinstance(typeface, str) or not typeface:
            raise ValueError(f"invalid typeface: {typeface!r}")
        resolved = self.theme_font(typeface)
        if resolved is not None:
            return resolved
        if typeface.startswith("+"):
            raise ValueError(
                f"theme font token {typeface!r} has no typeface in the "
                "theme's font scheme"
            )
        return typeface

    @property
    def major_latin_font(self) -> str:
        """The major (heading) latin typeface."""
        return self._font_scheme["major"]["latin"]

    @property
    def minor_latin_font(self) -> str:
        """The minor (body) latin typeface."""
        return self._font_scheme["minor"]["latin"]

    # -- text defaults -----------------------------------------------------

    def text_default_list_style(self) -> Optional[etree._Element]:
        """``a:objectDefaults/a:txDef/a:lstStyle`` of this theme, if any.

        Tail of the non-placeholder text cascade (consulted after the
        deck's ``p:defaultTextStyle`` -- order documented in
        ``utils.resolve_core``).
        """
        object_defaults = self._theme.find(qn("a:objectDefaults"))
        if object_defaults is None:
            return None
        tx_def = object_defaults.find(qn("a:txDef"))
        if tx_def is None:
            return None
        return tx_def.find(qn("a:lstStyle"))
