"""Self-check for the generated fixture decks and expected-value JSONs.

Pure Python (no COM):
1. Every fixture .pptx opens cleanly via python-pptx.
2. Every expected_values/<fixture>.json is valid, has >= 3 slides and
   >= 30 recorded runs, and contains zero extraction anomalies.
3. Seeded values actually appear in the recorded EFFECTIVE values --
   the null-test guard.  A resolver (or a broken authoring run) that
   yields Office defaults can never satisfy these checks:
   - theme_only: 19pt body runs, en-dash bullets, spc 5/9, 40pt bold title
   - layout_override: 21pt L1 runs + black-diamond bullets (layout A),
     Times New Roman 15pt justified runs (layout B) -- proves the
     layout layer overrode the master (master seeds are 19/16pt)
   - explicit_override: 23pt bold run, suppressed bullet, centered
     paragraph with spc 12/15, check-mark bullet
   - multi_master: two distinct master names; 17pt body on alt-master
     slides vs 19pt on base-master slides; white-circle vs en-dash bullets
4. Bullet-color integrity: every visible character bullet has a
   color_source of "explicit" or "follow_text" and a non-null color;
   follow_text colors equal the paragraph's first-run text color (guards
   against the COM RGB-0 sentinel leaking as '000000').

Exits non-zero on any failure.
"""

from __future__ import annotations

import json
import os
import sys

FIXTURES = ["theme_only", "layout_override", "explicit_override", "multi_master"]

FIXTURES_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXPECTED_DIR = os.path.join(FIXTURES_DIR, "expected_values")

EN_DASH = "–"
BLACK_DIAMOND = "◆"
WHITE_CIRCLE = "○"
CHECK_MARK = "✓"


def _iter_paragraphs(doc):
    for slide in doc["slides"]:
        for shape in slide["shapes"]:
            for para in shape["paragraphs"]:
                yield slide, shape, para


def _iter_runs(doc):
    for slide, shape, para in _iter_paragraphs(doc):
        for run in para["runs"]:
            yield slide, shape, para, run


def _check(errors: list[str], condition: bool, message: str) -> None:
    if not condition:
        errors.append(message)


def _load(fixture: str, errors: list[str]):
    path = os.path.join(EXPECTED_DIR, f"{fixture}.json")
    if not os.path.exists(path):
        errors.append(f"{fixture}: missing expected values JSON at {path}")
        return None
    with open(path, encoding="utf-8") as handle:
        return json.load(handle)


def verify_pptx_opens(fixture: str, errors: list[str]) -> None:
    from pptx import Presentation

    path = os.path.join(FIXTURES_DIR, f"{fixture}.pptx")
    if not os.path.exists(path):
        errors.append(f"{fixture}: missing deck at {path}")
        return
    try:
        pres = Presentation(path)
        slide_count = len(pres.slides)
    except Exception as exc:  # noqa: BLE001
        errors.append(f"{fixture}: python-pptx failed to open deck: {exc}")
        return
    _check(errors, slide_count >= 3,
           f"{fixture}: expected >= 3 slides, python-pptx sees {slide_count}")


def verify_json_shape(fixture: str, doc, errors: list[str]) -> None:
    _check(errors, len(doc["slides"]) >= 3,
           f"{fixture}: expected >= 3 slides in JSON, got {len(doc['slides'])}")
    runs = list(_iter_runs(doc))
    _check(errors, len(runs) >= 30,
           f"{fixture}: expected >= 30 recorded runs, got {len(runs)}")
    _check(errors, not doc.get("anomalies"),
           f"{fixture}: extraction anomalies present: {doc.get('anomalies')}")

    floating_heavy = any(
        sum(1 for shape in slide["shapes"] if not shape["is_placeholder"]) >= 4
        for slide in doc["slides"]
    )
    _check(errors, floating_heavy,
           f"{fixture}: no slide with >= 4 floating (non-placeholder) shapes")

    styled_border = any(
        shape["line"].get("visible")
        and shape["line"].get("weight_pt", 0) >= 1.25
        and shape["line"].get("dash_style") is not None
        for slide in doc["slides"] for shape in slide["shapes"]
    )
    _check(errors, styled_border,
           f"{fixture}: no shape with a styled (>=1.25pt) border recorded")


def verify_bullet_colors(fixture: str, doc, errors: list[str]) -> None:
    """Every visible character bullet carries an honest color record.

    Guards against the COM sentinel leak: Bullet.Font.Fill.ForeColor.RGB
    reads 0 (black) when no bullet color is set, so the extractor must
    resolve unset bullets via UseTextColor to the paragraph's first-run
    color (color_source "follow_text") instead of recording '000000'.
    No fixture seeds a black bullet, so any leak is detectable.
    """
    for slide, shape, para in _iter_paragraphs(doc):
        bullet = para["bullet"]
        if bullet.get("type") != 1 or bullet.get("visible") is not True:
            continue
        where = (f"{fixture}: slide{slide['index']}/{shape['name']}"
                 f"/para{para['idx']}")
        source = bullet.get("color_source")
        if source not in ("explicit", "follow_text"):
            errors.append(f"{where}: bullet color_source missing or "
                          f"invalid ({source!r})")
            continue
        color = bullet.get("color_rgb")
        if color is None:
            errors.append(f"{where}: bullet color_rgb is null")
            continue
        if source == "follow_text":
            runs = para["runs"]
            first = runs[0]["font"].get("color_rgb") if runs else None
            if color != first:
                errors.append(f"{where}: follow_text bullet color {color} "
                              f"!= first-run text color {first}")


def _has_run(doc, **font_expect) -> bool:
    for _, _, _, run in _iter_runs(doc):
        font = run["font"]
        if all(font.get(key) == value for key, value in font_expect.items()):
            return True
    return False


def _has_bullet_char(doc, char: str) -> bool:
    return any(para["bullet"].get("char") == char
               for _, _, para in _iter_paragraphs(doc))


def verify_theme_only(doc, errors: list[str]) -> None:
    f = "theme_only"
    _check(errors, _has_run(doc, size_pt=19.0),
           f"{f}: no 19pt body run (master body L1 seed did not cascade)")
    _check(errors, _has_run(doc, size_pt=40.0, bold=True),
           f"{f}: no 40pt bold title run (master title seed did not cascade)")
    _check(errors, _has_bullet_char(doc, EN_DASH),
           f"{f}: no en-dash bullet recorded (master bullet seed missing)")
    spacing = any(
        para.get("space_before_pt") == 5.0 and para.get("space_after_pt") == 9.0
        for _, _, para in _iter_paragraphs(doc)
    )
    _check(errors, spacing, f"{f}: no paragraph with seeded spacing 5/9 pt")
    _check(errors, _has_run(doc, size_pt=28.0) is False,
           f"{f}: found a 28pt run -- Office default body leaked through")


def verify_layout_override(doc, errors: list[str]) -> None:
    f = "layout_override"
    _check(errors, _has_run(doc, size_pt=21.0),
           f"{f}: no 21pt run (layout A body L1 override did not cascade)")
    _check(errors, _has_bullet_char(doc, BLACK_DIAMOND),
           f"{f}: no black-diamond bullet (layout A bullet override missing)")
    _check(errors, _has_run(doc, name="Times New Roman", size_pt=15.0),
           f"{f}: no Times New Roman 15pt run (layout B font override missing)")
    justified = any(
        para.get("alignment_name") == "justify" and para["runs"]
        for _, _, para in _iter_paragraphs(doc)
    )
    _check(errors, justified,
           f"{f}: no justified paragraph (layout B alignment override missing)")
    explicit_bullets = sum(
        1 for _, _, para in _iter_paragraphs(doc)
        if para["bullet"].get("color_source") == "explicit"
        and para["bullet"].get("color_rgb") == "1F4E79"
    )
    _check(errors, explicit_bullets >= 5,
           f"{f}: expected >= 5 explicit 1F4E79 bullet colors seeded at "
           f"layout level, found {explicit_bullets}")


def verify_explicit_override(doc, errors: list[str]) -> None:
    f = "explicit_override"
    _check(errors, _has_run(doc, size_pt=23.0, bold=True),
           f"{f}: no explicit 23pt bold run")
    _check(errors, _has_run(doc, color_rgb="FF6B35"),
           f"{f}: no explicit FF6B35 colored run")
    centered = any(
        para.get("alignment_name") == "center"
        and para.get("space_before_pt") == 12.0
        and para.get("space_after_pt") == 15.0
        for _, _, para in _iter_paragraphs(doc)
    )
    _check(errors, centered,
           f"{f}: no centered paragraph with explicit spacing 12/15 pt")
    _check(errors, _has_bullet_char(doc, CHECK_MARK),
           f"{f}: no check-mark bullet override")

    bullets = {
        (slide["index"], shape["name"], para["idx"]): para["bullet"]
        for slide, shape, para in _iter_paragraphs(doc)
    }
    # Render-verified follow-text bullet colors on slide 1 (bullets with
    # no buClr paint with the first run's text color): para 1 inherits
    # dk2 1F4E79, para 3 follows its explicit FF6B35 run, para 8 follows
    # the master body L2 color 3F3F66.
    for idx, expected in ((1, "1F4E79"), (3, "FF6B35"), (8, "3F3F66")):
        bullet = bullets.get((1, "Content Placeholder 2", idx), {})
        _check(errors,
               bullet.get("color_rgb") == expected
               and bullet.get("color_source") == "follow_text",
               f"{f}: slide1 para{idx} bullet expected follow_text "
               f"{expected}, got {bullet.get('color_source')!r} "
               f"{bullet.get('color_rgb')!r}")
    seeded = bullets.get((1, "Content Placeholder 2", 5), {})
    _check(errors,
           seeded.get("color_rgb") == "6B9F59"
           and seeded.get("color_source") == "explicit",
           f"{f}: check-mark bullet color not recorded as explicit 6B9F59 "
           f"(got {seeded.get('color_source')!r} {seeded.get('color_rgb')!r})")
    suppressed = any(
        para["bullet"].get("visible") is False and para["runs"]
        for _, _, para in _iter_paragraphs(doc)
    )
    _check(errors, suppressed, f"{f}: no paragraph with suppressed bullet")
    mixed = any(
        len(para["runs"]) >= 3
        for _, _, para in _iter_paragraphs(doc)
    )
    _check(errors, mixed,
           f"{f}: no paragraph with >= 3 runs (mixed formatting missing)")


def verify_multi_master(doc, errors: list[str]) -> None:
    f = "multi_master"
    masters = {slide["master_name"] for slide in doc["slides"]}
    _check(errors, len(masters) >= 2,
           f"{f}: expected slides on >= 2 masters, saw {sorted(masters)}")
    _check(errors, _has_run(doc, size_pt=17.0),
           f"{f}: no 17pt run (alt-master body seed did not cascade)")
    _check(errors, _has_run(doc, size_pt=19.0),
           f"{f}: no 19pt run (base-master body seed did not cascade)")
    _check(errors, _has_bullet_char(doc, WHITE_CIRCLE),
           f"{f}: no white-circle bullet (alt-master bullet seed missing)")
    _check(errors, _has_bullet_char(doc, EN_DASH),
           f"{f}: no en-dash bullet (base-master bullet seed missing)")
    alt_fill = any(
        shape["fill"].get("color_rgb") == "9A2B2B"
        for slide in doc["slides"] for shape in slide["shapes"]
    )
    _check(errors, alt_fill,
           f"{f}: accent2 theme fill did not resolve to 9A2B2B via alt master")


SEED_CHECKS = {
    "theme_only": verify_theme_only,
    "layout_override": verify_layout_override,
    "explicit_override": verify_explicit_override,
    "multi_master": verify_multi_master,
}


def run() -> list[str]:
    errors: list[str] = []
    for fixture in FIXTURES:
        verify_pptx_opens(fixture, errors)
        doc = _load(fixture, errors)
        if doc is None:
            continue
        verify_json_shape(fixture, doc, errors)
        verify_bullet_colors(fixture, doc, errors)
        SEED_CHECKS[fixture](doc, errors)
    return errors


def main() -> int:
    errors = run()
    if errors:
        print(f"FAIL: {len(errors)} problem(s)")
        for error in errors:
            print(f"  - {error}")
        return 1
    print("OK: all fixture decks and expected-value JSONs pass self-checks")
    return 0


if __name__ == "__main__":
    sys.exit(main())
