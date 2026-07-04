"""Author tests/fixtures/explicit_override.pptx via PowerPoint COM.

Fixture (c): a corporate-style deck whose slides carry EXPLICIT run-,
paragraph-, and shape-level overrides on top of the shared base template
(same master/theme as theme_only).  This is what real decks edited by
hand look like: local formatting layered over inheritance, including
mixed formatting inside a single paragraph (multiple runs).

Explicit seeds (all differ from base-master seeds AND Office defaults):
  Title runs: 36pt 1F4E79 / italic E2A33D (over 40pt bold C0504D master)
  Body para overrides: 23pt bold run; color FF6B35 run; centered para
    with spc 12/15; check-mark bullet 6B9F59; bullet suppressed (buNone)
  Floating text: Times New Roman 21pt bold italic 4A1942; Arial 12 6B9F59;
    14.5pt right-aligned caption
  Shapes: rounded-rect adj 0.40 fill DDEBF7 line 3pt solid C0504D;
    gradient rect line 2.25pt dashed 1F4E79; oval fill 2E86AB no line
"""

from __future__ import annotations

import os

from com_helpers import (
    ALIGN_CENTER,
    ALIGN_RIGHT,
    BULLET_CHECK_MARK,
    LINE_DASH,
    LINE_SOLID,
    PH_BODY,
    PH_OBJECT,
    SHAPE_OVAL,
    SHAPE_RECTANGLE,
    SHAPE_ROUNDED_RECTANGLE,
    add_autoshape,
    add_slide,
    add_textbox,
    apply_base_master,
    characters,
    find_layout,
    find_placeholder,
    gradient_fill,
    new_presentation,
    no_line,
    paragraph,
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

FIXTURE_NAME = "explicit_override"


def _run_in(paragraph_range, text: str, whole_text: str):
    """Return the TextRange2 covering `text` inside a paragraph."""
    start = whole_text.index(text) + 1  # COM Characters() is 1-based
    return characters(paragraph_range, start, len(text))


def build(app, output_path: str) -> str:
    with new_presentation(app) as pres:
        master = pres.Designs(1).SlideMaster
        master.Name = "FixtureBase"
        apply_base_master(master)

        _slide_program_update(pres, master)
        _slide_vendor_detail(pres, master)
        _slide_section_break(pres, master)
        _slide_floating(pres, master)

        return save_pptx(pres, output_path)


def _slide_program_update(pres, master):
    slide = add_slide(pres, find_layout(master, "Title and Content"))

    # Title with two explicitly-formatted runs (mixed formatting).
    title_text = "Transformation program update"
    set_title(slide, title_text)
    title_range = slide.Shapes.Title.TextFrame2.TextRange
    set_font(_run_in(title_range, "Transformation", title_text),
             size=36, color_hex="1F4E79")
    set_font(_run_in(title_range, "program", title_text),
             italic=True, color_hex="E2A33D")

    body = find_placeholder(slide, PH_BODY, PH_OBJECT)
    set_bullet_text(body, [
        (1, "Baseline scope confirmed with all sponsors"),
        (1, "Wave one delivery is trending ahead of plan"),
        (2, "Benefits case revalidated by finance"),
        (1, "Decision needed on the second wave start"),
        (1, "Completed workstreams archived this month"),
        (1, "Risks and issues log reset for wave two"),
        (1, "Change network extended to every site"),
        (2, "Local champions nominated by month end"),
    ])
    text_range = body.TextFrame2.TextRange

    # Paragraph 2: explicit run override (size + weight).
    set_font(paragraph(text_range, 2), size=23, bold=True)
    # Paragraph 3: explicit color override.
    set_font(paragraph(text_range, 3), color_hex="FF6B35")
    # Paragraph 4: explicit paragraph metrics.
    set_paragraph(paragraph(text_range, 4), alignment=ALIGN_CENTER,
                  space_before_pt=12, space_after_pt=15)
    # Paragraph 5: explicit bullet character override.
    set_paragraph(paragraph(text_range, 5), bullet_char=BULLET_CHECK_MARK,
                  bullet_rel_size=1.0, bullet_color_hex="6B9F59")
    # Paragraph 6: bullet suppressed (buNone).
    set_paragraph(paragraph(text_range, 6), bullet_visible=False)


def _slide_vendor_detail(pres, master):
    slide = add_slide(pres, find_layout(master, "Title and Content"))
    set_title(slide, "Vendor consolidation detail")

    body = find_placeholder(slide, PH_BODY, PH_OBJECT)
    line1 = "Vendor consolidation reduces overhead across regions"
    set_bullet_text(body, [
        (1, line1),
        (2, "Preferred supplier list trimmed to eight names"),
        (2, "Standard terms applied to every renewal"),
        (1, "Savings redirected into the delivery pods"),
    ])
    text_range = body.TextFrame2.TextRange

    # Three runs inside a single paragraph -- mixed formatting the
    # extractor must split per-run.
    para1 = paragraph(text_range, 1)
    set_font(_run_in(para1, "Vendor consolidation", line1),
             name="Georgia", size=21, bold=True)
    set_font(_run_in(para1, "overhead", line1),
             name="Times New Roman", size=15, italic=True,
             color_hex="7C5CA6")


def _slide_section_break(pres, master):
    slide = add_slide(pres, find_layout(master, "Section Header"))
    section_title = "Wave two priorities"
    set_title(slide, section_title)
    set_font(_run_in(slide.Shapes.Title.TextFrame2.TextRange,
                     "two", section_title),
             color_hex="A63D57", italic=True)

    body = find_placeholder(slide, PH_BODY, PH_OBJECT)
    if body is not None:
        body.TextFrame2.TextRange.Text = (
            "Scale what worked, retire what did not"
        )
        set_font(body.TextFrame2.TextRange, size=17.5, italic=True,
                 color_hex="3F3F66")


def _slide_floating(pres, master):
    """Floating-heavy slide with explicit run/paragraph/shape overrides."""
    slide = add_slide(pres, find_layout(master, "Title Only"))
    set_title(slide, "Decision summary")

    note = add_textbox(
        slide, 60, 140, 340, 90,
        "Steering committee approved the revised case\r"
        "Funding released for the next two quarters",
        name="FloatNoteDecision",
    )
    note_range = note.TextFrame2.TextRange
    set_font(paragraph(note_range, 1), name="Times New Roman", size=21,
             bold=True, italic=True, color_hex="4A1942")
    set_font(paragraph(note_range, 2), name="Arial", size=12,
             color_hex="6B9F59")
    set_paragraph(paragraph(note_range, 2), space_before_pt=6)

    conditions = add_textbox(
        slide, 60, 260, 340, 70,
        "Conditions attached to the approval\r"
        "Benefits tracking reports monthly",
        name="FloatNoteConditions",
    )
    set_font(paragraph(conditions.TextFrame2.TextRange, 1),
             name="Georgia", size=13, italic=True, color_hex="1F4E79")

    caption = add_textbox(
        slide, 60, 470, 400, 32,
        "Recorded at the June steering committee",
        name="FloatCaptionRecord",
    )
    set_font(caption.TextFrame2.TextRange, size=14.5, color_hex="8C5E93")
    set_paragraph(paragraph(caption.TextFrame2.TextRange, 1),
                  alignment=ALIGN_RIGHT)

    approve = add_autoshape(
        slide, SHAPE_ROUNDED_RECTANGLE, 470, 150, 140, 90,
        text="Approve", name="FloatShapeApprove",
    )
    set_adjustment(approve, 1, 0.40)
    solid_fill(approve, "DDEBF7")
    style_line(approve, weight_pt=3.0, dash_style=LINE_SOLID,
               color_hex="C0504D")
    set_font(approve.TextFrame2.TextRange, bold=True, color_hex="1F4E79")

    review = add_autoshape(
        slide, SHAPE_RECTANGLE, 630, 150, 140, 90,
        text="Review", name="FloatShapeReview",
    )
    gradient_fill(review, "3E8FB0", degree=0.55)
    style_line(review, weight_pt=2.25, dash_style=LINE_DASH,
               color_hex="1F4E79")

    deploy = add_autoshape(
        slide, SHAPE_OVAL, 790, 150, 140, 90,
        text="Deploy", name="FloatShapeDeploy",
    )
    solid_fill(deploy, "2E86AB")
    no_line(deploy)
    set_font(deploy.TextFrame2.TextRange, bold=True, color_hex="FDFBF7")


def main() -> str:
    fixtures_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    output = os.path.join(fixtures_dir, f"{FIXTURE_NAME}.pptx")
    path = run_standalone(build, output)
    print(f"Wrote {path}")
    return path


if __name__ == "__main__":
    main()
