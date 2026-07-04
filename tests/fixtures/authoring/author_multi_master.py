"""Author tests/fixtures/multi_master.pptx via PowerPoint COM.

Fixture (d): TWO slide masters with deliberately different themes, slides
alternating between them -- mimics "Frankenstein" decks assembled from
multiple sources.  Per-slide theme resolution must follow each slide's own
master chain (slide -> layout -> that master -> that master's theme part),
never a global theme1.xml.

Master 1 "FixtureBase": shared base theme (Georgia/Arial, accent1 C0504D,
title 40pt bold centered, body 19/16/13pt, bullets en-dash/square/guillemet).

Master 2 "FixtureAlt" (design "FixtureAlt"): Times New Roman/Georgia fonts,
its own scheme (accent1 0B6E4F, dark2 4A1942, ...), title 34pt italic
right-aligned, body 17/14/11pt with spacing 6/11, 4/8, 2/4 and bullets
white-circle / en-dash / white-bullet.  Every value differs from BOTH the
base master and Office defaults.
"""

from __future__ import annotations

import os

from com_helpers import (
    ALIGN_RIGHT,
    BULLET_EN_DASH,
    BULLET_WHITE_BULLET,
    BULLET_WHITE_CIRCLE,
    LINE_DASH_DOT,
    LINE_SOLID,
    MSO_THEME_COLOR_ACCENT1,
    MSO_THEME_COLOR_ACCENT2,
    MSO_THEME_COLOR_DARK2,
    PH_BODY,
    PH_OBJECT,
    PH_SUBTITLE,
    SHAPE_OVAL,
    SHAPE_ROUNDED_RECTANGLE,
    THEME_ACCENT1,
    THEME_ACCENT2,
    THEME_ACCENT3,
    THEME_ACCENT4,
    THEME_ACCENT5,
    THEME_ACCENT6,
    THEME_DARK1,
    THEME_DARK2,
    THEME_FOLLOWED_HYPERLINK,
    THEME_HYPERLINK,
    THEME_LIGHT1,
    THEME_LIGHT2,
    add_autoshape,
    add_slide,
    add_textbox,
    apply_base_master,
    apply_color_scheme,
    apply_font_scheme,
    find_layout,
    find_placeholder,
    new_presentation,
    no_line,
    run_standalone,
    save_pptx,
    set_adjustment,
    set_bullet_text,
    set_title,
    solid_fill,
    style_line,
    style_master_body_level,
    style_master_title,
)

FIXTURE_NAME = "multi_master"

ALT_COLOR_SCHEME: dict[int, str] = {
    THEME_DARK1: "20232A",
    THEME_LIGHT1: "FBF7F0",
    THEME_DARK2: "4A1942",
    THEME_LIGHT2: "DCD6CC",
    THEME_ACCENT1: "0B6E4F",
    THEME_ACCENT2: "9A2B2B",
    THEME_ACCENT3: "2F5D8A",
    THEME_ACCENT4: "C98A12",
    THEME_ACCENT5: "5E4B8B",
    THEME_ACCENT6: "3B7A57",
    THEME_HYPERLINK: "0B6E4F",
    THEME_FOLLOWED_HYPERLINK: "5E4B8B",
}

ALT_MAJOR_FONT = "Times New Roman"
ALT_MINOR_FONT = "Georgia"


def apply_alt_master(master) -> None:
    """Second theme -- must differ from the base master on every seed."""
    apply_color_scheme(master, ALT_COLOR_SCHEME)
    apply_font_scheme(master, ALT_MAJOR_FONT, ALT_MINOR_FONT)

    style_master_title(
        master,
        size=34,
        italic=True,
        alignment=ALIGN_RIGHT,
        theme_color=MSO_THEME_COLOR_ACCENT1,  # resolves to 0B6E4F
    )
    style_master_body_level(
        master, 1,
        size=17,
        theme_color=MSO_THEME_COLOR_DARK2,     # resolves to 4A1942
        space_before_pt=6, space_after_pt=11, space_within_multiple=1.20,
        bullet_char=BULLET_WHITE_CIRCLE, bullet_rel_size=0.85,
    )
    style_master_body_level(
        master, 2,
        size=14,
        color_hex="4E3B31",
        space_before_pt=4, space_after_pt=8, space_within_multiple=1.12,
        bullet_char=BULLET_EN_DASH, bullet_rel_size=1.0,
    )
    style_master_body_level(
        master, 3,
        size=11,
        space_before_pt=2, space_after_pt=4, space_within_multiple=1.06,
        bullet_char=BULLET_WHITE_BULLET, bullet_rel_size=0.9,
    )


def build(app, output_path: str) -> str:
    with new_presentation(app) as pres:
        design_base = pres.Designs(1)
        design_base.Name = "FixtureBase"
        master_base = design_base.SlideMaster
        master_base.Name = "FixtureBase"
        apply_base_master(master_base)

        design_alt = pres.Designs.Add("FixtureAlt")
        master_alt = design_alt.SlideMaster
        master_alt.Name = "FixtureAlt"
        apply_alt_master(master_alt)

        _slide_opening(pres, master_base)          # master 1
        _slide_legacy_results(pres, master_alt)    # master 2
        _slide_growth_results(pres, master_base)   # master 1
        _slide_floating_alt(pres, master_alt)      # master 2, floating-heavy
        _slide_outlook(pres, master_base)          # master 1

        return save_pptx(pres, output_path)


def _slide_opening(pres, master_base):
    slide = add_slide(pres, find_layout(master_base, "Title Slide"))
    set_title(slide, "Acme Corp portfolio day")
    subtitle = find_placeholder(slide, PH_SUBTITLE, PH_BODY)
    subtitle.TextFrame2.TextRange.Text = (
        "Combined review across both operating groups"
    )


def _slide_legacy_results(pres, master_alt):
    slide = add_slide(pres, find_layout(master_alt, "Title and Content"))
    set_title(slide, "Legacy division results")
    body = find_placeholder(slide, PH_BODY, PH_OBJECT)
    set_bullet_text(body, [
        (1, "Installed base renewals held at ninety percent"),
        (2, "Churn concentrated in the smallest tier"),
        (1, "Service margin improved on lower field costs"),
        (2, "Remote resolution rate reached two thirds"),
        (3, "Parts logistics outsourced in two regions"),
        (1, "Sunset roadmap communicated to all accounts"),
    ])


def _slide_growth_results(pres, master_base):
    slide = add_slide(pres, find_layout(master_base, "Title and Content"))
    set_title(slide, "Growth division results")
    body = find_placeholder(slide, PH_BODY, PH_OBJECT)
    set_bullet_text(body, [
        (1, "New platform bookings doubled year over year"),
        (2, "Land and expand motion now standard"),
        (3, "Average expansion closes within two quarters"),
        (1, "Partner sourced pipeline reached one third"),
        (2, "Two alliances promoted to strategic tier"),
        (1, "Attrition held below the sector benchmark"),
    ])


def _slide_floating_alt(pres, master_alt):
    """Floating-heavy slide bound to the ALT master."""
    slide = add_slide(pres, find_layout(master_alt, "Title Only"))
    set_title(slide, "Integration checkpoints")

    add_textbox(
        slide, 60, 140, 330, 80,
        "Shared back office cutover in the autumn\r"
        "Systems freeze holds for six weeks",
        name="FloatNoteCutover",
    )
    add_textbox(
        slide, 60, 250, 330, 40,
        "Synergy tracking moves to monthly cadence",
        name="FloatNoteSynergy",
    )
    add_textbox(
        slide, 60, 470, 420, 32,
        "Source: integration office, synthetic fixture data",
        name="FloatCaptionSource",
    )

    milestone = add_autoshape(
        slide, SHAPE_ROUNDED_RECTANGLE, 470, 150, 150, 90,
        text="Day 100", name="FloatShapeDay100",
    )
    set_adjustment(milestone, 1, 0.35)
    solid_fill(milestone, "E4EFE9")
    style_line(milestone, weight_pt=2.25, dash_style=LINE_DASH_DOT,
               color_hex="4A1942")

    target = add_autoshape(
        slide, SHAPE_OVAL, 650, 150, 150, 90,
        text="Target", name="FloatShapeTarget",
    )
    # Theme-colored fill: resolves through THIS slide's master (accent2
    # -> 9A2B2B on FixtureAlt, not 6B9F59 as on FixtureBase).
    solid_fill(target, theme_color=MSO_THEME_COLOR_ACCENT2)
    style_line(target, weight_pt=1.25, dash_style=LINE_SOLID,
               color_hex="C98A12")

    banner = add_autoshape(
        slide, SHAPE_ROUNDED_RECTANGLE, 470, 270, 330, 60,
        text="One operating model by year end",
        name="FloatShapeBanner",
    )
    set_adjustment(banner, 1, 0.18)
    solid_fill(banner, "5E4B8B")
    no_line(banner)


def _slide_outlook(pres, master_base):
    slide = add_slide(pres, find_layout(master_base, "Title and Content"))
    set_title(slide, "Combined outlook")
    body = find_placeholder(slide, PH_BODY, PH_OBJECT)
    set_bullet_text(body, [
        (1, "Guidance raised on stronger second half"),
        (2, "Currency assumptions unchanged"),
        (1, "Capital allocation favors the growth division"),
        (2, "Dividend policy reviewed at year end"),
        (1, "Integration costs tracked within envelope"),
        (2, "Synergy target raised after the first wave"),
    ])


def main() -> str:
    fixtures_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    output = os.path.join(fixtures_dir, f"{FIXTURE_NAME}.pptx")
    path = run_standalone(build, output)
    print(f"Wrote {path}")
    return path


if __name__ == "__main__":
    main()
