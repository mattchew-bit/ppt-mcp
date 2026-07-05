"""Unit tests for utils.resolve_colors -- DrawingML color math.

Transform expectations are hand-computed from the documented math
(tint/shade in sRGB, lumMod/lumOff/satMod in HSL). Where PowerPoint's
own rounding differs by one channel step (see the lumMod/lumOff case),
the fixture decks -- COM-extracted effective values -- remain the final
arbiter in the Step 2 integration tests.
"""

import pytest
from lxml import etree

from utils.resolve_colors import (
    PRESET_COLORS,
    apply_color_transforms,
    find_color_child,
    hex_to_rgb,
    hsl_to_rgb,
    normalize_hex,
    preset_color_hex,
    resolve_raw_color,
    rgb_to_hex,
    rgb_to_hsl,
    system_color_hex,
)

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def a_element(fragment: str) -> etree._Element:
    """Parse a DrawingML fragment with the ``a:`` namespace bound."""
    return etree.fromstring(
        f'<root xmlns:a="{_A_NS}">{fragment}</root>'
    )[0]


# ---------------------------------------------------------------------------
# Hex / channel plumbing
# ---------------------------------------------------------------------------

class TestHexPlumbing:
    def test_normalize_plain(self):
        assert normalize_hex("c0504d") == "C0504D"

    def test_normalize_hash_prefix(self):
        assert normalize_hex("#1f4e79") == "1F4E79"

    def test_normalize_argb_drops_alpha(self):
        assert normalize_hex("FF1A1A2E") == "1A1A2E"

    @pytest.mark.parametrize("bad", ["", "FFF", "GGGGGG", "12345"])
    def test_normalize_rejects_invalid(self, bad):
        with pytest.raises(ValueError):
            normalize_hex(bad)

    def test_normalize_rejects_non_string(self):
        with pytest.raises(TypeError):
            normalize_hex(0xC0504D)

    def test_hex_rgb_round_trip(self):
        assert rgb_to_hex(hex_to_rgb("C0504D")) == "C0504D"

    def test_rgb_to_hex_clamps(self):
        assert rgb_to_hex((1.5, -0.2, 0.5)) == "FF0080"


# ---------------------------------------------------------------------------
# HSL conversions
# ---------------------------------------------------------------------------

class TestHslConversions:
    def test_red_round_trip(self):
        h, s, l = rgb_to_hsl((1.0, 0.0, 0.0))
        assert (h, s, l) == (0.0, 1.0, 0.5)
        assert rgb_to_hex(hsl_to_rgb((h, s, l))) == "FF0000"

    def test_gray_has_zero_saturation(self):
        h, s, l = rgb_to_hsl(hex_to_rgb("808080"))
        assert s == 0.0
        assert abs(l - 128 / 255) < 1e-9

    def test_green_hue_is_120(self):
        h, _, _ = rgb_to_hsl((0.0, 1.0, 0.0))
        assert h == 120.0


# ---------------------------------------------------------------------------
# Transforms (hand-computed cases)
# ---------------------------------------------------------------------------

class TestTransforms:
    def test_tint_50_on_red(self):
        # FF*0.5+127.5=255 -> FF; 00*0.5+127.5=127.5 -> rounds to 128=80
        el = a_element('<a:srgbClr val="FF0000"><a:tint val="50000"/></a:srgbClr>')
        assert resolve_raw_color(el) == "FF8080"

    def test_shade_50_on_ff8080(self):
        # 255*0.5=127.5 -> 128=80; 128*0.5=64=40
        el = a_element('<a:srgbClr val="FF8080"><a:shade val="50000"/></a:srgbClr>')
        assert resolve_raw_color(el) == "804040"

    def test_lummod_50_on_white(self):
        # white: s=0, l=1.0; l*0.5=0.5 -> 127.5 -> 128 = 80
        assert apply_color_transforms(
            "FFFFFF", [a_element('<a:lumMod val="50000"/>')]
        ) == "808080"

    def test_lumoff_20_on_black(self):
        # black: l=0 + 0.2 -> 51 = 33
        assert apply_color_transforms(
            "000000", [a_element('<a:lumOff val="20000"/>')]
        ) == "333333"

    def test_satmod_zero_desaturates(self):
        # red at l=0.5 with s -> 0 collapses to mid gray
        assert apply_color_transforms(
            "FF0000", [a_element('<a:satMod val="0"/>')]
        ) == "808080"

    def test_lummod_lumoff_lighter_40(self):
        # PowerPoint "Accent, Lighter 40%" = lumMod 60% + lumOff 40%.
        # Hand-computed via documented HSL math for 4472C4:
        # H=218.4375, S=0.520325, L=0.517647 -> L'=0.710588 ->
        # R=142.80->143(8F), G=170.40->170(AA), B=219.60->220(DC).
        # (PowerPoint's own picker reports 8EAADC -- one rounding step
        # off on R; fixture-driven integration tests arbitrate.)
        assert apply_color_transforms(
            "4472C4",
            [a_element('<a:lumMod val="60000"/>'),
             a_element('<a:lumOff val="40000"/>')],
        ) == "8FAADC"

    def test_transforms_apply_in_document_order(self):
        # tint-then-shade != shade-then-tint
        tint_first = apply_color_transforms(
            "FF0000",
            [a_element('<a:tint val="50000"/>'),
             a_element('<a:shade val="50000"/>')],
        )
        shade_first = apply_color_transforms(
            "FF0000",
            [a_element('<a:shade val="50000"/>'),
             a_element('<a:tint val="50000"/>')],
        )
        # tint50: (1.0, 0.5, 0.5) -> shade50: (0.5, 0.25, 0.25) = 804040
        # shade50: (0.5, 0, 0) -> tint50: (0.75, 0.5, 0.5) = BF8080
        assert tint_first == "804040"
        assert shade_first == "BF8080"

    def test_alpha_is_ignored(self):
        el = a_element('<a:srgbClr val="C0504D"><a:alpha val="50000"/></a:srgbClr>')
        assert resolve_raw_color(el) == "C0504D"

    def test_transform_without_val_raises(self):
        with pytest.raises(ValueError):
            apply_color_transforms("FF0000", [a_element("<a:tint/>")])


# ---------------------------------------------------------------------------
# Base color forms
# ---------------------------------------------------------------------------

class TestBaseColors:
    def test_srgb(self):
        assert resolve_raw_color(a_element('<a:srgbClr val="1f4e79"/>')) == "1F4E79"

    def test_srgb_missing_val_raises(self):
        with pytest.raises(ValueError):
            resolve_raw_color(a_element("<a:srgbClr/>"))

    def test_sysclr_prefers_lastclr(self):
        el = a_element('<a:sysClr val="windowText" lastClr="1A1A2E"/>')
        assert resolve_raw_color(el) == "1A1A2E"

    def test_sysclr_fallback_table(self):
        assert system_color_hex(a_element('<a:sysClr val="window"/>')) == "FFFFFF"

    def test_sysclr_unknown_without_lastclr_raises(self):
        with pytest.raises(ValueError):
            system_color_hex(a_element('<a:sysClr val="noSuchColor"/>'))

    def test_prstclr(self):
        assert resolve_raw_color(a_element('<a:prstClr val="red"/>')) == "FF0000"

    def test_preset_alias_dk_blue(self):
        assert preset_color_hex("dkBlue") == PRESET_COLORS["darkBlue"] == "00008B"

    def test_preset_unknown_raises(self):
        with pytest.raises(ValueError):
            preset_color_hex("clownNose")

    def test_scrgbclr_percent_channels(self):
        el = a_element('<a:scrgbClr r="100000" g="0" b="0"/>')
        assert resolve_raw_color(el) == "FF0000"

    def test_hslclr_green(self):
        el = a_element('<a:hslClr hue="7200000" sat="100000" lum="50000"/>')
        assert resolve_raw_color(el) == "00FF00"

    def test_schemeclr_returns_none_for_caller(self):
        assert resolve_raw_color(a_element('<a:schemeClr val="accent1"/>')) is None

    def test_non_color_element_raises(self):
        with pytest.raises(ValueError):
            resolve_raw_color(a_element("<a:latin/>"))

    def test_none_raises(self):
        with pytest.raises(ValueError):
            resolve_raw_color(None)


# ---------------------------------------------------------------------------
# find_color_child
# ---------------------------------------------------------------------------

class TestFindColorChild:
    def test_finds_scheme_child(self):
        fill = a_element(
            '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>'
        )
        assert find_color_child(fill).get("val") == "accent1"

    def test_empty_fill_returns_none(self):
        assert find_color_child(a_element("<a:solidFill/>")) is None

    def test_none_parent_raises(self):
        with pytest.raises(ValueError):
            find_color_child(None)
