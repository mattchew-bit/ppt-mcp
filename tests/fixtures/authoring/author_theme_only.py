"""Author tests/fixtures/theme_only.pptx via PowerPoint COM.

Fixture (a): pure theme/master inheritance.  The slide master carries a
custom color scheme, a Georgia/Arial font scheme, and distinctive master
text styles (see com_helpers.apply_base_master).  Slides contain NO local
text-style overrides -- every effective value must resolve through the
placeholder -> layout -> master -> theme chain.  Slide 4 is heavy with
floating (non-placeholder) text boxes and styled autoshapes whose text is
also left un-overridden (it inherits via the defaultTextStyle chain).
"""

from __future__ import annotations

import os

from com_helpers import (
    BULLET_EN_DASH,  # noqa: F401 (documented seed, applied via master)
    LINE_DASH,
    LINE_SOLID,
    MSO_THEME_COLOR_ACCENT4,
    PH_BODY,
    PH_OBJECT,
    PH_SUBTITLE,
    SHAPE_OVAL,
    SHAPE_RECTANGLE,
    SHAPE_ROUNDED_RECTANGLE,
    add_autoshape,
    add_slide,
    add_textbox,
    apply_base_master,
    find_layout,
    find_placeholder,
    gradient_fill,
    new_presentation,
    no_line,
    run_standalone,
    save_pptx,
    set_adjustment,
    set_bullet_text,
    set_title,
    solid_fill,
    style_line,
)

FIXTURE_NAME = "theme_only"


def build(app, output_path: str) -> str:
    with new_presentation(app) as pres:
        master = pres.Designs(1).SlideMaster
        master.Name = "FixtureBase"
        apply_base_master(master)

        _slide_title(pres, master)
        _slide_market_overview(pres, master)
        _slide_priorities(pres, master)
        _slide_floating(pres, master)

        return save_pptx(pres, output_path)


def _slide_title(pres, master):
    slide = add_slide(pres, find_layout(master, "Title Slide"))
    set_title(slide, "Acme Corp Market Overview")
    subtitle = find_placeholder(slide, PH_SUBTITLE, PH_BODY)
    subtitle.TextFrame2.TextRange.Text = (
        "Prepared by the Strategy Team\rQ3 planning cycle"
    )


def _slide_market_overview(pres, master):
    slide = add_slide(pres, find_layout(master, "Title and Content"))
    set_title(slide, "Market overview")
    body = find_placeholder(slide, PH_BODY, PH_OBJECT)
    set_bullet_text(body, [
        (1, "Global demand recovered across all regions"),
        (1, "Pricing pressure eased through the second quarter"),
        (2, "Raw material costs declined four percent"),
        (2, "Logistics lead times returned to baseline"),
        (3, "Ocean freight rates reached a two year low"),
        (1, "Competitive intensity remains elevated"),
        (2, "Two new entrants appeared in the value segment"),
        (1, "Regulatory timeline unchanged for the year"),
    ])


def _slide_priorities(pres, master):
    slide = add_slide(pres, find_layout(master, "Title and Content"))
    set_title(slide, "Strategic priorities")
    body = find_placeholder(slide, PH_BODY, PH_OBJECT)
    set_bullet_text(body, [
        (1, "Expand the premium product line"),
        (2, "Launch two flagship variants next quarter"),
        (1, "Consolidate regional distribution partners"),
        (2, "Reduce partner count from twelve to five"),
        (3, "Retain coverage in all core markets"),
        (1, "Invest in demand forecasting capability"),
        (2, "Stand up a pricing analytics cell"),
        (3, "Quarterly reviews with regional leads"),
    ])


def _slide_floating(pres, master):
    """Floating-heavy slide: non-placeholder text boxes + styled shapes.

    Text inside boxes and shapes stays unformatted (inherits the
    defaultTextStyle chain); shape geometry/line/fill styling is local by
    nature and deliberately non-default (2.25pt dashed borders, custom
    fills, rounded-corner adjustment 0.32).
    """
    slide = add_slide(pres, find_layout(master, "Title Only"))
    set_title(slide, "Operating model at a glance")

    add_textbox(
        slide, 60, 130, 320, 80,
        "Regional hubs consolidate procurement\r"
        "Shared services move to a single center",
        name="FloatNoteHubs",
    )
    add_textbox(
        slide, 60, 240, 320, 60,
        "Four delivery pods aligned to customer segments\r"
        "Each pod owns its own quality gate",
        name="FloatNotePods",
    )
    add_textbox(
        slide, 60, 340, 320, 60,
        "Central functions provide shared tooling\r"
        "Cadence reviews run on a six week cycle",
        name="FloatNoteCadence",
    )
    add_textbox(
        slide, 60, 480, 400, 30,
        "Source: internal analysis, synthetic fixture data",
        name="FloatCaptionSource",
    )

    plan = add_autoshape(
        slide, SHAPE_ROUNDED_RECTANGLE, 470, 140, 130, 90,
        text="Plan", name="FloatPanelPlan",
    )
    set_adjustment(plan, 1, 0.32)  # corner radius (default is 0.16667)
    solid_fill(plan, "E8F0E4")
    style_line(plan, weight_pt=2.25, dash_style=LINE_DASH, color_hex="1F4E79")

    build_shape = add_autoshape(
        slide, SHAPE_RECTANGLE, 630, 140, 130, 90,
        text="Build", name="FloatPanelBuild",
    )
    gradient_fill(build_shape, "C0504D", degree=0.6)
    style_line(build_shape, weight_pt=1.5, dash_style=LINE_SOLID,
               color_hex="6B9F59")

    run_shape = add_autoshape(
        slide, SHAPE_OVAL, 790, 140, 130, 90,
        text="Run", name="FloatPanelRun",
    )
    solid_fill(run_shape, theme_color=MSO_THEME_COLOR_ACCENT4)  # -> E2A33D
    no_line(run_shape)


def main() -> str:
    fixtures_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    output = os.path.join(fixtures_dir, f"{FIXTURE_NAME}.pptx")
    path = run_standalone(build, output)
    print(f"Wrote {path}")
    return path


if __name__ == "__main__":
    main()
