"""House-profile schema helpers: DTCG leaves, size budget, DTCG export.

PROFILE SCHEMA CONTRACT (``house-profile/1`` -- pinned; the profile
builder and every consumer code to this):

    schema_version, name, source_decks, slide_size,
    typography: {title|body|footer: {font, size, bold, color}},
    paragraph:  {space_before, space_after, line_spacing,
                 bullets: {l1..l3: {char, color?, size_pct?}}},
    palette:    {scheme: <12 theme colors>,
                 usage: [{color, share, role}]},
    shape_defaults: {border: {weight, color, dash}, corner_radius, fill},
    grid:       {edges: {left, right, center}, unit, tolerance},
    archetypes: {<name>: {title_band: rect, body_region: rect, count}},
    images:     {count_per_slide, size_distribution, zones},
    distributions: {font_sizes_pt, space_after_pt, palette_shares}

Leaf values are DTCG-shaped: ``{"value": X, "unit": "pt"|"in"|None}``
for dimensions/unitless numbers, ``{"value": "#RRGGBB"}`` for colors,
``{"value": <str|bool>}`` for tokens. Grid edge arrays and distribution
arrays stay plain (their containers name the unit) per the contract.

CONSUMPTION SPLIT: ``apply_style_profile`` consumes ONLY typography /
paragraph / palette / shape_defaults (the deterministically applicable
rules); grid / archetypes / distributions / images inform generation
and the Step 5 lint.

The serializer enforces the plan's hard budget: a profile must fit in
``MAX_PROFILE_BYTES`` when compactly serialized, or the writer raises a
``ValueError`` listing the heaviest sections to trim.
"""

import json
from typing import Any, Dict, List, Optional

#: Hard cap from the plan: house profiles are prompts, not dumps.
MAX_PROFILE_BYTES = 8192

SCHEMA_VERSION = "house-profile/1"

#: Points per pixel at CSS reference density (96 px / 72 pt).
_PX_PER_PT = 96.0 / 72.0


def round2(value: float) -> float:
    """Uniform 2-decimal rounding used for every serialized number."""
    return round(float(value), 2)


def dim(value: float, unit: Optional[str]) -> Dict[str, Any]:
    """DTCG-shaped dimension leaf: ``{"value": 14.0, "unit": "pt"}``.

    ``unit=None`` marks unitless quantities (multiples, fractions).
    """
    if unit not in ("pt", "in", None):
        raise ValueError(f"unsupported unit: {unit!r}")
    return {"value": round2(value), "unit": unit}


def color(hex_value: str) -> Dict[str, str]:
    """DTCG-shaped color leaf: ``{"value": "#RRGGBB"}`` (normalized)."""
    raw = hex_value.lstrip("#")
    if len(raw) != 6 or any(c not in "0123456789abcdefABCDEF" for c in raw):
        raise ValueError(f"not an RRGGBB color: {hex_value!r}")
    return {"value": f"#{raw.upper()}"}


def token(value: Any) -> Dict[str, Any]:
    """DTCG-shaped plain leaf for strings / booleans."""
    return {"value": value}


def rect_in(x: float, y: float, w: float, h: float) -> Dict[str, Dict]:
    """Archetype-box rect: x/y/w/h as DTCG inch leaves."""
    return {"x": dim(x, "in"), "y": dim(y, "in"),
            "w": dim(w, "in"), "h": dim(h, "in")}


def serialize_profile(profile: Dict) -> bytes:
    """Compact canonical serialization (the persisted / measured form).

    UTF-8, no whitespace, keys in construction order (the builder is
    deterministic, so equal inputs give byte-identical output).
    """
    if profile.get("schema_version") != SCHEMA_VERSION:
        raise ValueError(
            "profile is missing schema_version "
            f"{SCHEMA_VERSION!r}: got {profile.get('schema_version')!r}"
        )
    return json.dumps(
        profile, separators=(",", ":"), ensure_ascii=False,
    ).encode("utf-8")


def _section_sizes(profile: Dict) -> List[str]:
    sizes = sorted(
        ((len(json.dumps(value, separators=(",", ":"),
                         ensure_ascii=False).encode("utf-8")), key)
         for key, value in profile.items()),
        reverse=True,
    )
    return [f"{key}={size}B" for size, key in sizes[:5]]


def enforce_size_budget(profile: Dict,
                        max_bytes: int = MAX_PROFILE_BYTES) -> bytes:
    """Serialize and enforce the hard byte budget.

    Raises ``ValueError`` naming the heaviest sections when the profile
    does not fit -- a house profile that blows the budget degrades
    generation instead of helping it (plan risk table).
    """
    payload = serialize_profile(profile)
    if len(payload) > max_bytes:
        raise ValueError(
            f"house profile is {len(payload)} bytes, over the "
            f"{max_bytes}-byte budget; heaviest sections: "
            f"{', '.join(_section_sizes(profile))} -- trim top-N caps "
            "(palette usage, grid edges, distributions) or archetypes"
        )
    return payload


# ---------------------------------------------------------------------------
# DTCG (W3C Design Tokens) exporter -- colors/typography subset
# ---------------------------------------------------------------------------

def _dtcg_color_group(scheme: Dict[str, Dict]) -> Dict[str, Dict]:
    return {
        name: {"$type": "color", "$value": leaf["value"]}
        for name, leaf in scheme.items()
    }


def _dtcg_font_size(size_leaf: Dict) -> Dict[str, Any]:
    """Profile pt size -> DTCG dimension (px -- the only spec'd unit)."""
    return {
        "value": round2(size_leaf["value"] * _PX_PER_PT),
        "unit": "px",
    }


#: Profile keys a role must carry to form a DTCG typography composite.
_DTCG_TYPOGRAPHY_KEYS = ("font", "size", "bold")


def _has_typography_composite(spec: Dict) -> bool:
    """True when a role spec carries every DTCG-composite key.

    The profile builder emits only modal-FOUND keys, so a role (e.g. a
    footer with no resolvable font names/sizes in the corpus) may be a
    partial spec -- valid in the house schema but inexpressible as a
    DTCG typography composite. Such roles are skipped by the exporter
    instead of crashing it.
    """
    return all(key in spec for key in _DTCG_TYPOGRAPHY_KEYS)


def _dtcg_typography(role: Dict[str, Dict]) -> Dict[str, Any]:
    value: Dict[str, Any] = {
        "fontFamily": role["font"]["value"],
        "fontSize": _dtcg_font_size(role["size"]),
        "fontWeight": 700 if role["bold"]["value"] else 400,
    }
    return {"$type": "typography", "$value": value}


def to_dtcg(profile: Dict) -> Dict:
    """Export the colors/typography subset as W3C DTCG token groups.

    DTCG proper has no geometry or pt-dimension types (plan note), so
    this exporter covers exactly what DTCG can express: the theme color
    scheme, per-role text colors, and per-role typography composites
    (font sizes converted pt -> px at CSS density). Everything else in
    the house profile stays in the bespoke schema. Roles whose spec is
    partial (missing font/size/bold -- the builder emits only
    modal-found keys) are omitted from the typography group; their
    text color, when present, still exports.
    """
    if profile.get("schema_version") != SCHEMA_VERSION:
        raise ValueError(
            f"to_dtcg expects a {SCHEMA_VERSION!r} profile, got "
            f"{profile.get('schema_version')!r}"
        )
    typography = profile.get("typography", {})
    tokens: Dict[str, Any] = {
        "$description": (
            f"Exported from house profile {profile.get('name', '?')!r}"
        ),
        "color": {
            "scheme": _dtcg_color_group(
                profile.get("palette", {}).get("scheme", {})),
            "text": {
                role: {"$type": "color",
                       "$value": spec["color"]["value"]}
                for role, spec in typography.items()
                if "color" in spec
            },
        },
        "typography": {
            role: _dtcg_typography(spec)
            for role, spec in typography.items()
            if _has_typography_composite(spec)
        },
    }
    return tokens
