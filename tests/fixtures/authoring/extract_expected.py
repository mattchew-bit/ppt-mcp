"""Record PowerPoint-reported EFFECTIVE style values for each fixture deck.

Walks every fixture via COM and writes
tests/fixtures/expected_values/<fixture>.json.  The values recorded here
are what desktop PowerPoint itself reports through TextRange2 /
ParagraphFormat2 / Shape -- i.e. fully inheritance-resolved "ground truth"
that the Step 2 effective-style resolver must reproduce.

Recording rules:
- Iterate PER RUN so mixed formatting never yields msoUndefined (-2)
  sentinels.  Any tri-state or numeric property that still reads as
  mixed/undefined is recorded as null and logged in the "anomalies" list
  instead of writing junk.
- Floats rounded to 2 decimal places.
- Colors recorded as RRGGBB hex (converted from OLE BGR longs).
- Bullet colors: COM's Bullet.Font.Fill.ForeColor.RGB reads 0 (black)
  when NO bullet color is set anywhere in the inheritance chain, so it
  is never read blindly.  Bullet.UseTextColor is checked first: msoFalse
  means a buClr is in force (recorded with color_source "explicit");
  msoTrue means the OOXML default buClrTx applies -- the bullet paints
  with the color of the paragraph's first text run (render-verified) and
  is recorded with that run's color and color_source "follow_text".
- Space before/after recorded as *_pt when the paragraph line-rule flag is
  off (points), as *_lines otherwise.  Line spacing (SpaceWithin) is
  recorded with its rule ("multiple" or "points").
- Theme font references ("+mj-lt"/"+mn-lt") are recorded raw plus a
  "name_resolved" field looked up from that slide's own master theme
  (multi-master decks resolve per slide).
"""

from __future__ import annotations

import contextlib
import datetime as _dt
import json
import os

from com_helpers import (
    MSO_FALSE,
    MSO_TRUE,
    SHAPE_TYPE_PLACEHOLDER,
    hex_from_ole,
    open_presentation,
    paragraph as get_paragraph,
    paragraphs_count,
    powerpoint_app,
    run_item,
    runs_count,
)

FIXTURES = ["theme_only", "layout_override", "explicit_override", "multi_master"]

MSO_UNDEFINED = -2
MSO_FILL_SOLID = 1

ALIGNMENT_NAMES = {
    1: "left", 2: "center", 3: "right", 4: "justify",
    5: "distribute", 6: "thai_distribute", 7: "justify_low",
}
BULLET_TYPE_NAMES = {0: "none", 1: "unnumbered", 2: "numbered", 3: "picture"}
DASH_STYLE_NAMES = {
    1: "solid", 2: "square_dot", 3: "round_dot", 4: "dash", 5: "dash_dot",
    6: "dash_dot_dot", 7: "long_dash", 8: "long_dash_dot",
}
FILL_TYPE_NAMES = {
    1: "solid", 2: "patterned", 3: "gradient", 4: "textured",
    5: "background", 6: "picture",
}


def _rnd(value) -> float:
    return round(float(value), 2)


class Anomalies:
    """Collects extraction anomalies instead of writing junk values."""

    def __init__(self):
        self.items: list[str] = []

    def add(self, context: str, message: str) -> None:
        self.items.append(f"{context}: {message}")


def _tri_state(value, context: str, prop: str, anomalies: Anomalies):
    """msoTrue/msoFalse -> bool; mixed/undefined -> None + anomaly."""
    if value == MSO_TRUE:
        return True
    if value == MSO_FALSE:
        return False
    anomalies.add(context, f"{prop} returned tri-state {value} (mixed/undefined)")
    return None


def _safe_color(color_format, context: str, prop: str, anomalies: Anomalies):
    try:
        return hex_from_ole(int(color_format.RGB))
    except Exception as exc:  # noqa: BLE001 -- COM errors are opaque
        anomalies.add(context, f"{prop} unreadable: {exc}")
        return None


def _theme_latin_fonts(slide) -> tuple[str | None, str | None]:
    with contextlib.suppress(Exception):
        scheme = slide.Master.Theme.ThemeFontScheme
        return (scheme.MajorFont.Item(1).Name, scheme.MinorFont.Item(1).Name)
    return (None, None)


def _extract_font(font, slide, context: str, anomalies: Anomalies) -> dict:
    name = str(font.Name)
    record = {
        "name": name,
        "size_pt": _rnd(font.Size),
        "bold": _tri_state(font.Bold, context, "bold", anomalies),
        "italic": _tri_state(font.Italic, context, "italic", anomalies),
        "color_rgb": _safe_color(font.Fill.ForeColor, context, "font color",
                                 anomalies),
    }
    if name.startswith("+"):
        major, minor = _theme_latin_fonts(slide)
        record["name_resolved"] = major if name.startswith("+mj") else minor
    if record["size_pt"] is not None and record["size_pt"] <= 0:
        anomalies.add(context, f"font size read as {record['size_pt']}")
        record["size_pt"] = None
    return record


def _bullet_color(bullet, context: str, anomalies: Anomalies,
                  first_run_color: str | None) -> dict:
    """Bullet color with follow-text semantics.

    ``Bullet.Font.Fill.ForeColor.RGB`` returns 0 (black) when no bullet
    color is set anywhere in the chain, so reading it blindly would leak
    a junk ``000000``.  ``Bullet.UseTextColor`` is msoTrue exactly in
    that unset state (OOXML default buClrTx): the bullet paints with the
    color of the paragraph's first text run (render-verified against
    PowerPoint output).  msoFalse means a buClr is in force somewhere in
    the inheritance chain and the reported RGB is real.
    """
    try:
        use_text_color = int(bullet.UseTextColor)
    except Exception as exc:  # noqa: BLE001
        anomalies.add(context, f"bullet UseTextColor unreadable: {exc}")
        return {"color_rgb": None}
    if use_text_color == MSO_FALSE:
        return {
            "color_rgb": _safe_color(bullet.Font.Fill.ForeColor, context,
                                     "bullet color", anomalies),
            "color_source": "explicit",
        }
    if use_text_color == MSO_TRUE:
        if first_run_color is None:
            anomalies.add(context, "bullet follows text color but the "
                                   "paragraph has no readable first-run color")
        return {"color_rgb": first_run_color, "color_source": "follow_text"}
    anomalies.add(context, f"bullet UseTextColor tri-state {use_text_color} "
                           "(mixed/undefined)")
    return {"color_rgb": None}


def _extract_bullet(paragraph_format, context: str, anomalies: Anomalies,
                    first_run_color: str | None = None) -> dict:
    bullet = paragraph_format.Bullet
    record: dict = {}
    try:
        btype = int(bullet.Type)
    except Exception as exc:  # noqa: BLE001
        anomalies.add(context, f"bullet type unreadable: {exc}")
        return {"type": None}
    if btype == MSO_UNDEFINED:
        anomalies.add(context, "bullet type is mixed (-2)")
        return {"type": None}
    record["type"] = btype
    record["type_name"] = BULLET_TYPE_NAMES.get(btype, str(btype))
    with contextlib.suppress(Exception):
        record["visible"] = _tri_state(bullet.Visible, context,
                                       "bullet visible", anomalies)
    if btype == 1:  # unnumbered (character) bullet
        try:
            code = int(bullet.Character)
            record["char_code"] = code
            record["char"] = chr(code) if 0 < code < 0x110000 else None
        except Exception as exc:  # noqa: BLE001
            anomalies.add(context, f"bullet character unreadable: {exc}")
        with contextlib.suppress(Exception):
            record["font_name"] = str(bullet.Font.Name)
        record.update(_bullet_color(bullet, context, anomalies,
                                    first_run_color))
        with contextlib.suppress(Exception):
            record["relative_size"] = _rnd(bullet.RelativeSize)
    return record


def _extract_paragraph(paragraph, index: int, slide, context: str,
                       anomalies: Anomalies) -> dict:
    pf = paragraph.ParagraphFormat
    record: dict = {"idx": index, "level": int(pf.IndentLevel)}

    alignment = int(pf.Alignment)
    if alignment == MSO_UNDEFINED:
        anomalies.add(context, "alignment is mixed (-2)")
        record["alignment"] = None
    else:
        record["alignment"] = alignment
        record["alignment_name"] = ALIGNMENT_NAMES.get(alignment,
                                                       str(alignment))

    rule_before = pf.LineRuleBefore
    if rule_before == MSO_FALSE:
        record["space_before_pt"] = _rnd(pf.SpaceBefore)
    elif rule_before == MSO_TRUE:
        record["space_before_lines"] = _rnd(pf.SpaceBefore)
    else:
        anomalies.add(context, f"LineRuleBefore tri-state {rule_before}")

    rule_after = pf.LineRuleAfter
    if rule_after == MSO_FALSE:
        record["space_after_pt"] = _rnd(pf.SpaceAfter)
    elif rule_after == MSO_TRUE:
        record["space_after_lines"] = _rnd(pf.SpaceAfter)
    else:
        anomalies.add(context, f"LineRuleAfter tri-state {rule_after}")

    rule_within = pf.LineRuleWithin
    if rule_within == MSO_TRUE:
        record["space_within"] = _rnd(pf.SpaceWithin)
        record["space_within_rule"] = "multiple"
    elif rule_within == MSO_FALSE:
        record["space_within"] = _rnd(pf.SpaceWithin)
        record["space_within_rule"] = "points"
    else:
        anomalies.add(context, f"LineRuleWithin tri-state {rule_within}")

    # Runs are extracted BEFORE the bullet: a bullet with no explicit
    # color follows the first run's text color (see _bullet_color).
    runs = []
    try:
        run_count = runs_count(paragraph)
    except Exception:  # noqa: BLE001 -- empty paragraphs have no runs
        run_count = 0
    for j in range(1, run_count + 1):
        run = run_item(paragraph, j)
        run_context = f"{context}/run{j}"
        text = str(run.Text).replace("\r", "").replace("\x0b", " ")
        runs.append({
            "text": text,
            "font": _extract_font(run.Font, slide, run_context, anomalies),
        })

    first_run_color = runs[0]["font"]["color_rgb"] if runs else None
    record["bullet"] = _extract_bullet(pf, context, anomalies,
                                       first_run_color)
    record["runs"] = runs
    return record


def _extract_line(shape, context: str, anomalies: Anomalies) -> dict:
    line = shape.Line
    visible = _tri_state(line.Visible, context, "line visible", anomalies)
    record: dict = {"visible": visible}
    if visible:
        with contextlib.suppress(Exception):
            record["weight_pt"] = _rnd(line.Weight)
        with contextlib.suppress(Exception):
            dash = int(line.DashStyle)
            record["dash_style"] = dash
            record["dash_style_name"] = DASH_STYLE_NAMES.get(dash, str(dash))
        record["color_rgb"] = _safe_color(line.ForeColor, context,
                                          "line color", anomalies)
    return record


def _extract_fill(shape, context: str, anomalies: Anomalies) -> dict:
    fill = shape.Fill
    visible = _tri_state(fill.Visible, context, "fill visible", anomalies)
    record: dict = {"visible": visible}
    if visible:
        with contextlib.suppress(Exception):
            fill_type = int(fill.Type)
            record["type"] = fill_type
            record["type_name"] = FILL_TYPE_NAMES.get(fill_type,
                                                      str(fill_type))
            if fill_type == MSO_FILL_SOLID:
                record["color_rgb"] = _safe_color(fill.ForeColor, context,
                                                  "fill color", anomalies)
    return record


def _extract_shape(shape, slide, context: str, anomalies: Anomalies) -> dict:
    is_placeholder = int(shape.Type) == SHAPE_TYPE_PLACEHOLDER
    record: dict = {
        "name": str(shape.Name),
        "shape_type": int(shape.Type),
        "is_placeholder": is_placeholder,
        "ph_type": None,
        "geometry": {
            "left_pt": _rnd(shape.Left),
            "top_pt": _rnd(shape.Top),
            "width_pt": _rnd(shape.Width),
            "height_pt": _rnd(shape.Height),
            "rotation": _rnd(shape.Rotation),
        },
    }
    if is_placeholder:
        record["ph_type"] = int(shape.PlaceholderFormat.Type)
    with contextlib.suppress(Exception):
        record["geometry"]["auto_shape_type"] = int(shape.AutoShapeType)
    with contextlib.suppress(Exception):
        adjustments = shape.Adjustments
        record["geometry"]["adjustments"] = [
            _rnd(adjustments.Item(i))
            for i in range(1, int(adjustments.Count) + 1)
        ]

    record["line"] = _extract_line(shape, context, anomalies)
    record["fill"] = _extract_fill(shape, context, anomalies)

    paragraphs = []
    has_text = False
    with contextlib.suppress(Exception):
        has_text = (shape.HasTextFrame == MSO_TRUE
                    and shape.TextFrame2.HasText == MSO_TRUE)
    if has_text:
        text_range = shape.TextFrame2.TextRange
        para_count = paragraphs_count(text_range)
        for i in range(1, para_count + 1):
            paragraph = get_paragraph(text_range, i)
            paragraphs.append(
                _extract_paragraph(paragraph, i, slide,
                                   f"{context}/para{i}", anomalies))
    record["paragraphs"] = paragraphs
    return record


def extract_deck(app, deck_path: str, fixture_name: str) -> dict:
    anomalies = Anomalies()
    with open_presentation(app, deck_path, read_only=True) as pres:
        document: dict = {
            "fixture": fixture_name,
            "source": os.path.basename(deck_path),
            "generated_by": "tests/fixtures/authoring/extract_expected.py "
                            "(PowerPoint COM, effective values)",
            "generated_on": _dt.date.today().isoformat(),
            "powerpoint_version": str(app.Version),
            "slide_size_pt": {
                "width": _rnd(pres.PageSetup.SlideWidth),
                "height": _rnd(pres.PageSetup.SlideHeight),
            },
            "slides": [],
        }
        for s in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(s)
            slide_record: dict = {
                "index": s,
                "layout_name": str(slide.CustomLayout.Name),
                "master_name": str(slide.Master.Name),
                "shapes": [],
            }
            for i in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(i)
                context = f"slide{s}/shape{i}({shape.Name})"
                slide_record["shapes"].append(
                    _extract_shape(shape, slide, context, anomalies))
            document["slides"].append(slide_record)
        document["anomalies"] = anomalies.items
    return document


def extract_all(app, fixtures_dir: str, expected_dir: str,
                fixtures: list[str] | None = None) -> dict[str, str]:
    os.makedirs(expected_dir, exist_ok=True)
    written: dict[str, str] = {}
    for fixture_name in fixtures or FIXTURES:
        deck_path = os.path.join(fixtures_dir, f"{fixture_name}.pptx")
        document = extract_deck(app, deck_path, fixture_name)
        out_path = os.path.join(expected_dir, f"{fixture_name}.json")
        with open(out_path, "w", encoding="utf-8") as handle:
            json.dump(document, handle, indent=2, ensure_ascii=False)
            handle.write("\n")
        written[fixture_name] = out_path
        print(f"Wrote {out_path} "
              f"({len(document['slides'])} slides, "
              f"{len(document['anomalies'])} anomalies)")
    return written


def main() -> dict[str, str]:
    fixtures_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    expected_dir = os.path.join(fixtures_dir, "expected_values")
    with powerpoint_app() as app:
        return extract_all(app, fixtures_dir, expected_dir)


if __name__ == "__main__":
    main()
