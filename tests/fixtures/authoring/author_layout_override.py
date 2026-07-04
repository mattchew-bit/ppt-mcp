"""Author tests/fixtures/layout_override.pptx via PowerPoint COM.

Fixture (b): same base master/theme as theme_only, plus TWO custom layouts
("Fixture Content A" / "Fixture Content B") whose placeholders OVERRIDE the
master text styles -- different sizes, colors, spacing, and bullet
characters at the layout level.  Slides carry no local overrides: their
effective values must resolve through the layout layer, not the master.

Layout-level seeds (all differ from BOTH the master seeds and Office
defaults -- null-test guard):

  Layout A title: 30pt italic, color 7C5CA6, right-aligned
  Layout A body L1: 21pt, color 2E5E4E, bullet black-diamond, spc 2/12, 1.30
  Layout A body L2: 17pt, color C0504D, bullet small-square,  spc 3/8
  Layout A body L3: 14pt,               bullet multiply-sign, spc 2/6
  Layout B title: 26pt, color E2A33D
  Layout B body L1: 15pt Times New Roman, justified, bullet guillemet
                    (color 1F4E79, rel size 1.1), spc 6/13
  Layout B body L2: 12pt, bullet white-bullet, spc 2/6
"""

from __future__ import annotations

import os

from com_helpers import (
    ALIGN_JUSTIFY,
    ALIGN_RIGHT,
    BULLET_BLACK_DIAMOND,
    BULLET_MULTIPLY,
    BULLET_RIGHT_GUILLEMET,
    BULLET_SMALL_SQUARE,
    BULLET_WHITE_BULLET,
    LINE_DASH,
    LINE_LONG_DASH_DOT,
    PH_BODY,
    PH_CENTER_TITLE,
    PH_OBJECT,
    PH_TITLE,
    SHAPE_CHEVRON,
    SHAPE_ROUNDED_RECTANGLE,
    add_autoshape,
    add_slide,
    add_textbox,
    apply_base_master,
    find_layout,
    find_placeholder,
    new_presentation,
    paragraph,
    paragraphs_count,
    run_standalone,
    save_pptx,
    set_adjustment,
    set_bullet_text,
    set_font,
    set_paragraph,
    set_title,
    solid_fill,
    style_line,
)

FIXTURE_NAME = "layout_override"


def build(app, output_path: str) -> str:
    with new_presentation(app) as pres:
        master = pres.Designs(1).SlideMaster
        master.Name = "FixtureBase"
        apply_base_master(master)

        layout_a = _make_layout_a(master)
        layout_b = _make_layout_b(master)

        _slide_portfolio(pres, layout_a)
        _slide_financials(pres, layout_b)
        _slide_roadmap(pres, layout_a)
        _slide_floating(pres, layout_b)

        return save_pptx(pres, output_path)


def _style_layout_title(layout, **font_kwargs):
    alignment = font_kwargs.pop("alignment", None)
    title = find_placeholder(layout, PH_TITLE, PH_CENTER_TITLE)
    text_range = title.TextFrame2.TextRange
    set_font(text_range, **font_kwargs)
    if alignment is not None:
        for i in range(1, paragraphs_count(text_range) + 1):
            paragraph(text_range, i).ParagraphFormat.Alignment = alignment


def _style_layout_body(layout, level_specs):
    """Rewrite the body placeholder prompt paragraphs, one per level, and
    format each -- this is how layout-level overrides are authored in the
    PowerPoint UI (Slide Master view)."""
    body = find_placeholder(layout, PH_BODY, PH_OBJECT)
    text_range = body.TextFrame2.TextRange
    text_range.Text = "\r".join(spec["prompt"] for spec in level_specs)
    for i, spec in enumerate(level_specs, start=1):
        para = paragraph(text_range, i)
        para.ParagraphFormat.IndentLevel = spec["level"]
        set_font(para, **spec.get("font", {}))
        set_paragraph(para, **spec.get("para", {}))


def _make_layout_a(master):
    layout = find_layout(master, "Title and Content").Duplicate()
    layout.Name = "Fixture Content A"
    _style_layout_title(
        layout, size=30, italic=True, color_hex="7C5CA6",
        alignment=ALIGN_RIGHT,
    )
    _style_layout_body(layout, [
        {
            "prompt": "Layout A first level", "level": 1,
            "font": {"size": 21, "color_hex": "2E5E4E"},
            "para": {"bullet_char": BULLET_BLACK_DIAMOND,
                     "bullet_rel_size": 0.8,
                     "space_before_pt": 2, "space_after_pt": 12,
                     "space_within_multiple": 1.30},
        },
        {
            "prompt": "Layout A second level", "level": 2,
            "font": {"size": 17, "color_hex": "C0504D"},
            "para": {"bullet_char": BULLET_SMALL_SQUARE,
                     "space_before_pt": 3, "space_after_pt": 8},
        },
        {
            "prompt": "Layout A third level", "level": 3,
            "font": {"size": 14},
            "para": {"bullet_char": BULLET_MULTIPLY,
                     "space_before_pt": 2, "space_after_pt": 6},
        },
    ])
    return layout


def _make_layout_b(master):
    layout = find_layout(master, "Title and Content").Duplicate()
    layout.Name = "Fixture Content B"
    _style_layout_title(layout, size=26, color_hex="E2A33D")
    _style_layout_body(layout, [
        {
            "prompt": "Layout B first level", "level": 1,
            "font": {"size": 15, "name": "Times New Roman"},
            "para": {"alignment": ALIGN_JUSTIFY,
                     "bullet_char": BULLET_RIGHT_GUILLEMET,
                     "bullet_rel_size": 1.1,
                     "bullet_color_hex": "1F4E79",
                     "space_before_pt": 6, "space_after_pt": 13},
        },
        {
            "prompt": "Layout B second level", "level": 2,
            "font": {"size": 12},
            "para": {"bullet_char": BULLET_WHITE_BULLET,
                     "space_before_pt": 2, "space_after_pt": 6},
        },
    ])
    return layout


def _slide_portfolio(pres, layout_a):
    slide = add_slide(pres, layout_a)
    set_title(slide, "Portfolio review")
    body = find_placeholder(slide, PH_BODY, PH_OBJECT)
    set_bullet_text(body, [
        (1, "Core segment holds share in a flat market"),
        (2, "Volume steady while list prices firmed"),
        (3, "Discounting concentrated in two accounts"),
        (1, "Growth segment doubled its pipeline"),
        (2, "Six pilot programs converted to contracts"),
        (1, "Legacy products enter managed decline"),
        (2, "Support commitments honored through next year"),
    ])


def _slide_financials(pres, layout_b):
    slide = add_slide(pres, layout_b)
    set_title(slide, "Financial summary")
    body = find_placeholder(slide, PH_BODY, PH_OBJECT)
    set_bullet_text(body, [
        (1, "Revenue expanded ahead of the annual plan"),
        (2, "Services mix improved gross margin"),
        (1, "Operating costs held below budget"),
        (2, "Hiring deferred in two support functions"),
        (1, "Cash conversion strengthened quarter over quarter"),
        (2, "Receivables aging returned to target"),
        (1, "Full year guidance affirmed at the midpoint"),
        (2, "Sensitivity range narrowed on stable demand"),
    ])


def _slide_roadmap(pres, layout_a):
    slide = add_slide(pres, layout_a)
    set_title(slide, "Execution roadmap")
    body = find_placeholder(slide, PH_BODY, PH_OBJECT)
    set_bullet_text(body, [
        (1, "Phase one stabilizes the operating baseline"),
        (2, "Critical process gaps closed by design sprints"),
        (1, "Phase two scales the delivery model"),
        (2, "Playbooks rolled out to every region"),
        (3, "Certification required for pod leads"),
        (1, "Phase three institutionalizes improvement"),
    ])


def _slide_floating(pres, layout_b):
    """Floating-heavy slide on Layout B: non-placeholder boxes + styled shapes."""
    slide = add_slide(pres, layout_b)
    set_title(slide, "Delivery checkpoints")
    body = find_placeholder(slide, PH_BODY, PH_OBJECT)
    set_bullet_text(body, [
        (1, "Milestones tracked against the master schedule"),
        (2, "Escalations reviewed at the weekly forum"),
    ])

    add_textbox(
        slide, 500, 150, 300, 70,
        "Checkpoint owners rotate monthly\rMinutes filed within two days",
        name="FloatNoteGovernance",
    )
    add_textbox(
        slide, 500, 250, 300, 40,
        "All dates are working assumptions",
        name="FloatCaptionDates",
    )
    add_textbox(
        slide, 60, 420, 380, 60,
        "Readiness reviews precede every gate\r"
        "Exceptions require sponsor approval",
        name="FloatNoteReadiness",
    )

    gate = add_autoshape(
        slide, SHAPE_ROUNDED_RECTANGLE, 500, 320, 140, 80,
        text="Gate 1", name="FloatShapeGate1",
    )
    set_adjustment(gate, 1, 0.28)
    solid_fill(gate, "F2E8D8")
    style_line(gate, weight_pt=2.25, dash_style=LINE_DASH, color_hex="A63D57")

    arrow = add_autoshape(
        slide, SHAPE_CHEVRON, 660, 320, 140, 80,
        text="Gate 2", name="FloatShapeGate2",
    )
    solid_fill(arrow, "3E8FB0")
    style_line(arrow, weight_pt=1.75, dash_style=LINE_LONG_DASH_DOT,
               color_hex="1F4E79")


def main() -> str:
    fixtures_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    output = os.path.join(fixtures_dir, f"{FIXTURE_NAME}.pptx")
    path = run_standalone(build, output)
    print(f"Wrote {path}")
    return path


if __name__ == "__main__":
    main()
