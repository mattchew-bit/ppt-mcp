"""DrawingML color math for the effective-style resolver (read-only).

Resolves raw (non-scheme) OOXML color elements to final ``RRGGBB``
uppercase hex, and applies the DrawingML color transforms that PowerPoint
layers on top of a base color. Scheme-color indirection (``a:schemeClr``
through ``p:clrMap`` and the theme ``a:clrScheme``) lives one level up in
``utils.resolve_theme``; this module is the shared math beneath it.

Supported base color elements (children of ``a:solidFill`` etc.):
    * ``a:srgbClr``  -- literal hex
    * ``a:sysClr``   -- system color; ``lastClr`` preferred, small
      fallback table otherwise
    * ``a:prstClr``  -- ECMA-376 preset color name (full ST_PresetColorVal
      table, including the ``dk``/``lt``/``med`` aliases)
    * ``a:scrgbClr`` -- percent RGB (interpreted as direct sRGB percents,
      the same simplification Apache POI makes)
    * ``a:hslClr``   -- hue in 60000ths of a degree, sat/lum percents

Transform math (child elements of the color element, applied in document
order, values are ST_Percentage thousandths-of-a-percent so 100000 == 1.0):
    * ``a:tint``   -- lighten toward white: ``c' = c*v + (1 - v)``
      per sRGB channel (ECMA-376 20.1.2.3.34)
    * ``a:shade``  -- darken toward black: ``c' = c*v`` per sRGB channel
      (ECMA-376 20.1.2.3.31)
    * ``a:lumMod`` / ``a:lumOff`` -- HSL luminance modulate / offset
      (the pair behind PowerPoint's "Lighter 40%" theme variants)
    * ``a:satMod`` -- HSL saturation modulate
    * ``a:hueMod`` -- HSL hue modulate (wraps at 360)
    * alpha family (``a:alpha``/``a:alphaMod``/``a:alphaOff``) is ignored:
      the resolver reports opaque RGB, matching what the COM object model
      reports for font colors.
    Other transforms (``a:gray``, ``a:comp``, ``a:inv``, gamma pair) are
    not applied in v1; they are listed in ``_IGNORED_TRANSFORMS`` so their
    presence is deliberate, not an oversight.

The tint/shade-in-sRGB + lum/sat-in-HSL split follows Apache POI's
``XSLFColor`` semantics (Apache-2.0, re-implemented -- not copied). The
Step 0 fixtures are the final arbiter: if a COM-recorded effective color
ever disagrees, the fixture value wins and this math gets adjusted.

Everything in this module is pure: no element is ever mutated.
"""

from typing import Dict, Iterable, Optional, Tuple

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

#: ST_Percentage denominator: 100000 == 100%.
PERCENT_DENOMINATOR = 100000.0

#: ST_Angle-style hue attribute denominator: 60000 units == 1 degree.
HUE_DENOMINATOR = 60000.0

#: Fallbacks for ``a:sysClr`` when the authoring app recorded no
#: ``lastClr``. Values match Windows default (light) system colors.
SYSTEM_COLOR_FALLBACKS: Dict[str, str] = {
    "windowText": "000000",
    "window": "FFFFFF",
    "menuText": "000000",
    "captionText": "000000",
    "btnFace": "F0F0F0",
    "btnText": "000000",
    "grayText": "6D6D6D",
    "highlight": "0078D7",
    "highlightText": "FFFFFF",
    "hotLight": "0066CC",
    "infoBk": "FFFFE1",
    "infoText": "000000",
}

#: ECMA-376 ST_PresetColorVal -> RRGGBB (X11/CSS values, camelCase keys).
PRESET_COLORS: Dict[str, str] = {
    "aliceBlue": "F0F8FF", "antiqueWhite": "FAEBD7", "aqua": "00FFFF",
    "aquamarine": "7FFFD4", "azure": "F0FFFF", "beige": "F5F5DC",
    "bisque": "FFE4C4", "black": "000000", "blanchedAlmond": "FFEBCD",
    "blue": "0000FF", "blueViolet": "8A2BE2", "brown": "A52A2A",
    "burlyWood": "DEB887", "cadetBlue": "5F9EA0", "chartreuse": "7FFF00",
    "chocolate": "D2691E", "coral": "FF7F50", "cornflowerBlue": "6495ED",
    "cornsilk": "FFF8DC", "crimson": "DC143C", "cyan": "00FFFF",
    "darkBlue": "00008B", "darkCyan": "008B8B", "darkGoldenrod": "B8860B",
    "darkGray": "A9A9A9", "darkGreen": "006400", "darkKhaki": "BDB76B",
    "darkMagenta": "8B008B", "darkOliveGreen": "556B2F",
    "darkOrange": "FF8C00", "darkOrchid": "9932CC", "darkRed": "8B0000",
    "darkSalmon": "E9967A", "darkSeaGreen": "8FBC8F",
    "darkSlateBlue": "483D8B", "darkSlateGray": "2F4F4F",
    "darkTurquoise": "00CED1", "darkViolet": "9400D3",
    "deepPink": "FF1493", "deepSkyBlue": "00BFFF", "dimGray": "696969",
    "dodgerBlue": "1E90FF", "firebrick": "B22222",
    "floralWhite": "FFFAF0", "forestGreen": "228B22", "fuchsia": "FF00FF",
    "gainsboro": "DCDCDC", "ghostWhite": "F8F8FF", "gold": "FFD700",
    "goldenrod": "DAA520", "gray": "808080", "green": "008000",
    "greenYellow": "ADFF2F", "honeydew": "F0FFF0", "hotPink": "FF69B4",
    "indianRed": "CD5C5C", "indigo": "4B0082", "ivory": "FFFFF0",
    "khaki": "F0E68C", "lavender": "E6E6FA", "lavenderBlush": "FFF0F5",
    "lawnGreen": "7CFC00", "lemonChiffon": "FFFACD",
    "lightBlue": "ADD8E6", "lightCoral": "F08080", "lightCyan": "E0FFFF",
    "lightGoldenrodYellow": "FAFAD2", "lightGray": "D3D3D3",
    "lightGreen": "90EE90", "lightPink": "FFB6C1",
    "lightSalmon": "FFA07A", "lightSeaGreen": "20B2AA",
    "lightSkyBlue": "87CEFA", "lightSlateGray": "778899",
    "lightSteelBlue": "B0C4DE", "lightYellow": "FFFFE0",
    "lime": "00FF00", "limeGreen": "32CD32", "linen": "FAF0E6",
    "magenta": "FF00FF", "maroon": "800000", "mediumAquamarine": "66CDAA",
    "mediumBlue": "0000CD", "mediumOrchid": "BA55D3",
    "mediumPurple": "9370DB", "mediumSeaGreen": "3CB371",
    "mediumSlateBlue": "7B68EE", "mediumSpringGreen": "00FA9A",
    "mediumTurquoise": "48D1CC", "mediumVioletRed": "C71585",
    "midnightBlue": "191970", "mintCream": "F5FFFA",
    "mistyRose": "FFE4E1", "moccasin": "FFE4B5", "navajoWhite": "FFDEAD",
    "navy": "000080", "oldLace": "FDF5E6", "olive": "808000",
    "oliveDrab": "6B8E23", "orange": "FFA500", "orangeRed": "FF4500",
    "orchid": "DA70D6", "paleGoldenrod": "EEE8AA", "paleGreen": "98FB98",
    "paleTurquoise": "AFEEEE", "paleVioletRed": "DB7093",
    "papayaWhip": "FFEFD5", "peachPuff": "FFDAB9", "peru": "CD853F",
    "pink": "FFC0CB", "plum": "DDA0DD", "powderBlue": "B0E0E6",
    "purple": "800080", "red": "FF0000", "rosyBrown": "BC8F8F",
    "royalBlue": "4169E1", "saddleBrown": "8B4513", "salmon": "FA8072",
    "sandyBrown": "F4A460", "seaGreen": "2E8B57", "seaShell": "FFF5EE",
    "sienna": "A0522D", "silver": "C0C0C0", "skyBlue": "87CEEB",
    "slateBlue": "6A5ACD", "slateGray": "708090", "snow": "FFFAFA",
    "springGreen": "00FF7F", "steelBlue": "4682B4", "tan": "D2B48C",
    "teal": "008080", "thistle": "D8BFD8", "tomato": "FF6347",
    "turquoise": "40E0D0", "violet": "EE82EE", "wheat": "F5DEB3",
    "white": "FFFFFF", "whiteSmoke": "F5F5F5", "yellow": "FFFF00",
    "yellowGreen": "9ACD32",
}

#: ECMA-376 abbreviated preset aliases -> canonical names above.
_PRESET_ALIASES: Dict[str, str] = {
    "dkBlue": "darkBlue", "dkCyan": "darkCyan",
    "dkGoldenrod": "darkGoldenrod", "dkGray": "darkGray",
    "dkGreen": "darkGreen", "dkKhaki": "darkKhaki",
    "dkMagenta": "darkMagenta", "dkOliveGreen": "darkOliveGreen",
    "dkOrange": "darkOrange", "dkOrchid": "darkOrchid",
    "dkRed": "darkRed", "dkSalmon": "darkSalmon",
    "dkSeaGreen": "darkSeaGreen", "dkSlateBlue": "darkSlateBlue",
    "dkSlateGray": "darkSlateGray", "dkTurquoise": "darkTurquoise",
    "dkViolet": "darkViolet",
    "ltBlue": "lightBlue", "ltCoral": "lightCoral",
    "ltCyan": "lightCyan", "ltGoldenrodYellow": "lightGoldenrodYellow",
    "ltGray": "lightGray", "ltGreen": "lightGreen",
    "ltPink": "lightPink", "ltSalmon": "lightSalmon",
    "ltSeaGreen": "lightSeaGreen", "ltSkyBlue": "lightSkyBlue",
    "ltSlateGray": "lightSlateGray", "ltSteelBlue": "lightSteelBlue",
    "ltYellow": "lightYellow",
    "medAquamarine": "mediumAquamarine", "medBlue": "mediumBlue",
    "medOrchid": "mediumOrchid", "medPurple": "mediumPurple",
    "medSeaGreen": "mediumSeaGreen", "medSlateBlue": "mediumSlateBlue",
    "medSpringGreen": "mediumSpringGreen",
    "medTurquoise": "mediumTurquoise", "medVioletRed": "mediumVioletRed",
}

#: Transforms deliberately skipped in v1 (alpha family: output is opaque
#: RGB; the rest are rare and unimplemented, not forgotten).
_IGNORED_TRANSFORMS = frozenset({
    "alpha", "alphaMod", "alphaOff",
    "gray", "comp", "inv", "gamma", "invGamma",
    "red", "redMod", "redOff",
    "green", "greenMod", "greenOff",
    "blue", "blueMod", "blueOff",
    "hue", "hueOff", "sat", "satOff", "lum",
})

#: Local names of the raw (non-scheme) base color element tags.
RAW_COLOR_TAGS = frozenset({"srgbClr", "sysClr", "prstClr", "scrgbClr", "hslClr"})


# ---------------------------------------------------------------------------
# Hex / channel helpers
# ---------------------------------------------------------------------------

def normalize_hex(value: str) -> str:
    """Normalize a hex color string to ``RRGGBB`` uppercase.

    Accepts an optional leading ``#`` and 8-digit ARGB (leading alpha
    byte is dropped, mirroring ShapeCrawler's ``FontColor`` hex cleanup).
    """
    if not isinstance(value, str):
        raise TypeError(f"hex color must be str, got {type(value).__name__}")
    cleaned = value[1:] if value.startswith("#") else value
    if len(cleaned) == 8:
        cleaned = cleaned[2:]
    if len(cleaned) != 6:
        raise ValueError(f"invalid hex color: {value!r}")
    int(cleaned, 16)  # raises ValueError on non-hex characters
    return cleaned.upper()


def hex_to_rgb(hex_color: str) -> Tuple[float, float, float]:
    """``"RRGGBB"`` -> normalized ``(r, g, b)`` floats in 0..1."""
    cleaned = normalize_hex(hex_color)
    return (
        int(cleaned[0:2], 16) / 255.0,
        int(cleaned[2:4], 16) / 255.0,
        int(cleaned[4:6], 16) / 255.0,
    )


def rgb_to_hex(rgb: Tuple[float, float, float]) -> str:
    """Normalized ``(r, g, b)`` floats -> ``RRGGBB`` uppercase.

    Channels are clamped to 0..1 and rounded half-away-from-zero, the
    same rounding ShapeCrawler applies to normAutofit-scaled sizes.
    """
    channels = []
    for channel in rgb:
        clamped = min(1.0, max(0.0, channel))
        channels.append(int(clamped * 255.0 + 0.5))
    return "{:02X}{:02X}{:02X}".format(*channels)


# ---------------------------------------------------------------------------
# HSL conversions (h in degrees 0..360, s/l in 0..1)
# ---------------------------------------------------------------------------

def rgb_to_hsl(rgb: Tuple[float, float, float]) -> Tuple[float, float, float]:
    """Standard sRGB -> HSL conversion on normalized channels."""
    r, g, b = rgb
    high, low = max(r, g, b), min(r, g, b)
    lum = (high + low) / 2.0
    if high == low:
        return (0.0, 0.0, lum)
    delta = high - low
    if lum < 0.5:
        sat = delta / (high + low)
    else:
        sat = delta / (2.0 - high - low)
    if high == r:
        hue = ((g - b) / delta) % 6.0
    elif high == g:
        hue = (b - r) / delta + 2.0
    else:
        hue = (r - g) / delta + 4.0
    return (hue * 60.0, sat, lum)


def _hue_component(p: float, q: float, t: float) -> float:
    """One channel of the HSL -> RGB reconstruction."""
    if t < 0.0:
        t += 1.0
    if t > 1.0:
        t -= 1.0
    if t < 1.0 / 6.0:
        return p + (q - p) * 6.0 * t
    if t < 1.0 / 2.0:
        return q
    if t < 2.0 / 3.0:
        return p + (q - p) * (2.0 / 3.0 - t) * 6.0
    return p


def hsl_to_rgb(hsl: Tuple[float, float, float]) -> Tuple[float, float, float]:
    """Standard HSL -> sRGB conversion; returns normalized channels."""
    hue, sat, lum = hsl
    if sat == 0.0:
        return (lum, lum, lum)
    q = lum * (1.0 + sat) if lum < 0.5 else lum + sat - lum * sat
    p = 2.0 * lum - q
    h = (hue % 360.0) / 360.0
    return (
        _hue_component(p, q, h + 1.0 / 3.0),
        _hue_component(p, q, h),
        _hue_component(p, q, h - 1.0 / 3.0),
    )


# ---------------------------------------------------------------------------
# Transforms
# ---------------------------------------------------------------------------

def _percent_value(element) -> float:
    """Read a ST_Percentage ``val`` attribute as a 0..n float fraction."""
    raw = element.get("val")
    if raw is None:
        raise ValueError(
            f"color transform <{_local_name(element)}> has no val attribute"
        )
    return int(raw) / PERCENT_DENOMINATOR


def _local_name(element) -> str:
    """Tag name without its namespace prefix."""
    tag = element.tag
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def apply_color_transforms(hex_color: str, transform_elements: Iterable) -> str:
    """Apply DrawingML transform child elements to a base hex color.

    ``transform_elements`` is an iterable of lxml elements (typically the
    children of an ``a:srgbClr``/``a:schemeClr`` element) applied in
    order. Returns the transformed ``RRGGBB`` uppercase hex. Unknown or
    deliberately-ignored transforms are skipped (see module docstring).
    """
    rgb = hex_to_rgb(hex_color)
    for transform in transform_elements:
        name = _local_name(transform)
        if name in _IGNORED_TRANSFORMS:
            continue
        if name == "tint":
            value = _percent_value(transform)
            rgb = tuple(c * value + (1.0 - value) for c in rgb)
        elif name == "shade":
            value = _percent_value(transform)
            rgb = tuple(c * value for c in rgb)
        elif name in ("lumMod", "lumOff", "satMod", "hueMod"):
            rgb = _apply_hsl_transform(rgb, name, _percent_value(transform))
        # anything else: not a transform we act on -- skip silently, the
        # supported/ignored split is documented at module level.
    return rgb_to_hex(rgb)


def _apply_hsl_transform(
    rgb: Tuple[float, float, float], name: str, value: float
) -> Tuple[float, float, float]:
    """Apply one HSL-space transform and convert back to RGB."""
    hue, sat, lum = rgb_to_hsl(rgb)
    if name == "lumMod":
        lum *= value
    elif name == "lumOff":
        lum += value
    elif name == "satMod":
        sat *= value
    elif name == "hueMod":
        hue = (hue * value) % 360.0
    sat = min(1.0, max(0.0, sat))
    lum = min(1.0, max(0.0, lum))
    return hsl_to_rgb((hue, sat, lum))


# ---------------------------------------------------------------------------
# Raw (non-scheme) base color resolution
# ---------------------------------------------------------------------------

def preset_color_hex(name: str) -> str:
    """ECMA-376 preset color name -> ``RRGGBB`` hex (aliases included)."""
    canonical = _PRESET_ALIASES.get(name, name)
    try:
        return PRESET_COLORS[canonical]
    except KeyError:
        raise ValueError(f"unknown preset color name: {name!r}") from None


def system_color_hex(element) -> str:
    """Resolve an ``a:sysClr`` element; prefers its ``lastClr`` snapshot."""
    last_color = element.get("lastClr")
    if last_color is not None:
        return normalize_hex(last_color)
    name = element.get("val")
    if name in SYSTEM_COLOR_FALLBACKS:
        return SYSTEM_COLOR_FALLBACKS[name]
    raise ValueError(f"cannot resolve a:sysClr val={name!r} without lastClr")


def _scrgb_hex(element) -> str:
    """``a:scrgbClr`` percent channels -> hex (direct-percent reading)."""
    channels = []
    for attr in ("r", "g", "b"):
        raw = element.get(attr)
        if raw is None:
            raise ValueError(f"a:scrgbClr missing {attr!r} attribute")
        channels.append(int(raw) / PERCENT_DENOMINATOR)
    return rgb_to_hex(tuple(channels))


def _hsl_hex(element) -> str:
    """``a:hslClr`` attributes -> hex."""
    raw_hue = element.get("hue")
    raw_sat = element.get("sat")
    raw_lum = element.get("lum")
    if raw_hue is None or raw_sat is None or raw_lum is None:
        raise ValueError("a:hslClr requires hue, sat and lum attributes")
    hsl = (
        (int(raw_hue) / HUE_DENOMINATOR) % 360.0,
        int(raw_sat) / PERCENT_DENOMINATOR,
        int(raw_lum) / PERCENT_DENOMINATOR,
    )
    return rgb_to_hex(hsl_to_rgb(hsl))


def resolve_raw_color(color_element) -> Optional[str]:
    """Resolve a non-scheme color element to transformed ``RRGGBB`` hex.

    Handles ``a:srgbClr`` / ``a:sysClr`` / ``a:prstClr`` / ``a:scrgbClr``
    / ``a:hslClr`` including their child transforms. Returns ``None`` for
    ``a:schemeClr`` (the caller must resolve it with theme context; see
    ``utils.resolve_theme.ThemeContext.resolve_color``). Raises
    ``ValueError`` for elements that are not color elements at all.
    """
    if color_element is None:
        raise ValueError("color_element must not be None")
    name = _local_name(color_element)
    if name == "schemeClr":
        return None
    if name == "srgbClr":
        raw = color_element.get("val")
        if raw is None:
            raise ValueError("a:srgbClr has no val attribute")
        base = normalize_hex(raw)
    elif name == "sysClr":
        base = system_color_hex(color_element)
    elif name == "prstClr":
        preset_name = color_element.get("val")
        if preset_name is None:
            raise ValueError("a:prstClr has no val attribute")
        base = preset_color_hex(preset_name)
    elif name == "scrgbClr":
        base = _scrgb_hex(color_element)
    elif name == "hslClr":
        base = _hsl_hex(color_element)
    else:
        raise ValueError(f"<{name}> is not a DrawingML color element")
    return apply_color_transforms(base, list(color_element))


def find_color_child(parent_element):
    """First color child (raw or schemeClr) of e.g. an ``a:solidFill``.

    Returns the lxml element or ``None`` when the parent has no color
    child. Never mutates ``parent_element``.
    """
    if parent_element is None:
        raise ValueError("parent_element must not be None")
    for child in parent_element:
        if _local_name(child) in RAW_COLOR_TAGS or _local_name(child) == "schemeClr":
            return child
    return None
