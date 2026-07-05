# Fixture authoring scripts (Step 0 — style-fidelity upgrade)

The four decks in `tests/fixtures/` are the ground-truth fixtures for the
effective-style inheritance resolver (plan: style-fidelity-upgrade, Step 0).
They are **authored by desktop PowerPoint via COM** (pywin32) — never by
python-pptx — so the files contain real inheritance structures
(theme → master `txStyles` → layout → slide) written by PowerPoint itself.
Testing python-pptx-generated files would test the library against itself
and skip every hard inheritance path.

## Fixtures

| Deck | Inheritance layer exercised |
|---|---|
| `theme_only.pptx` | Everything inherits from master/theme. Custom color scheme, Georgia/Arial font scheme, distinctive master text styles. Slides carry **no** local text-style overrides. |
| `layout_override.pptx` | Same base master, plus two custom layouts ("Fixture Content A"/"B") whose placeholders override the master (size/color/spacing/bullets/alignment/font at layout level). Slides inherit from the layouts. |
| `explicit_override.pptx` | Corporate-style deck: explicit run/paragraph/shape overrides layered on the base template, including mixed formatting (multiple runs per paragraph), bullet overrides, and a suppressed bullet. |
| `multi_master.pptx` | TWO masters ("FixtureBase"/"FixtureAlt") with different themes (colors + fonts + text styles); slides alternate. Per-slide theme resolution: e.g. an accent2 fill resolves to `9A2B2B` on alt-master slides, `6B9F59` on base-master slides. |

Every deck has ≥3 slides and at least one slide heavy with floating
(non-placeholder) text boxes and styled autoshapes (dashed/solid borders at
non-default weights, custom fills, gradient fill, rounded-corner
adjustments).

## Null-test guard

**Every seeded value differs from the Office defaults** so a resolver that
returns defaults can never accidentally pass:

| Property | Office default | Base master seed | Alt master seed |
|---|---|---|---|
| Fonts (major/minor) | Calibri Light / Calibri | Georgia / Arial | Times New Roman / Georgia |
| Title | 44pt regular, dk1 | 40pt **bold**, centered, accent1 (`C0504D`) | 34pt *italic*, right, accent1 (`0B6E4F`) |
| Body L1/L2/L3 size | 28/24/20pt | 19/16/13pt | 17/14/11pt |
| Space before/after L1 | 10/0pt | 5/9pt | 6/11pt |
| Line spacing | 0.9 | 1.15/1.10/1.05 | 1.20/1.12/1.06 |
| Bullets L1/L2/L3 | • | – (en dash) / ■ / » | ○ / – / ◦ |
| Accent1 | `4472C4` | `C0504D` | `0B6E4F` |

Layout/explicit seeds (21pt ◆, Times New Roman 15pt justified », 23pt bold
run, `FF6B35` run, 12/15pt centered paragraph, ✓ bullet, buNone, 2.25pt
dashed borders, adj 0.28–0.40) likewise avoid all default values. See each
`author_*.py` docstring for the full seed list.

## Expected values

`extract_expected.py` walks each deck via COM and records the **effective**
values PowerPoint reports (TextRange2/Font2, ParagraphFormat2, Shape
line/fill/geometry) into `tests/fixtures/expected_values/<fixture>.json`:

```
{fixture, powerpoint_version, slide_size_pt,
 slides: [{index, layout_name, master_name,
   shapes: [{name, shape_type, is_placeholder, ph_type,
     geometry: {left_pt, top_pt, width_pt, height_pt, rotation,
                auto_shape_type, adjustments},
     line: {visible, weight_pt, dash_style, dash_style_name, color_rgb},
     fill: {visible, type, type_name, color_rgb},
     paragraphs: [{idx, level, alignment, alignment_name,
       space_before_pt|space_before_lines, space_after_pt|space_after_lines,
       space_within, space_within_rule,
       bullet: {type, type_name, visible, char, char_code, font_name,
                color_rgb, color_source, relative_size},
       runs: [{text, font: {name, size_pt, bold, italic, color_rgb,
                            name_resolved?}}]}]}]}],
 anomalies: []}
```

Notes:

- Floats are rounded to 2dp; colors are `RRGGBB` hex.
- Extraction iterates **per run**, so mixed formatting never produces
  msoUndefined (-2) sentinels; anything that still reads as mixed is
  recorded as `null` and logged in `anomalies` (all four files currently
  have zero anomalies).
- `space_*_pt` vs `space_*_lines` reflects the paragraph's LineRule flags;
  `space_within_rule` is `"multiple"` or `"points"`.
- Bullet `color_rgb` is never read blindly: COM reports RGB 0 (black)
  when no bullet color is set, so `Bullet.UseTextColor` is checked first.
  `color_source: "explicit"` means a `buClr` is in force somewhere in the
  chain and the recorded color is real; `"follow_text"` means the OOXML
  default `buClrTx` applies — the bullet paints with the paragraph's
  first-run text color (render-verified), and that color is recorded.
- Tests must assert **exact values** from these files — never just
  non-None (the null-test trap).

## Regenerating

Requires Windows + desktop PowerPoint (Office16) + `pywin32`
(`py -3 -m pip install pywin32`). Then:

```
py -3 tests/fixtures/authoring/build_all.py
```

This reauthors all four decks, re-extracts the expected-value JSONs,
exports per-slide PNG previews (to `%TEMP%/ppt_mcp_fixture_previews` by
default; use `--previews-dir` to redirect — previews are never committed),
and runs `verify_fixtures.py` (python-pptx round-trip + JSON sanity +
seeded-value guards). Everything is idempotent — outputs are overwritten
in place.

Individual steps: `py -3 author_<fixture>.py`, `py -3 extract_expected.py`,
`py -3 export_previews.py [--out DIR]`, `py -3 verify_fixtures.py`.

## COM hygiene

`com_helpers.powerpoint_app()` enforces the rules from the plan:
`CoInitialize`/`CoUninitialize` pairing, never sets `Visible = False`
(PowerPoint throws), `AutomationSecurity = ForceDisable` while working
(restored after), `DisplayAlerts` off (restored after), operates only on
presentations it creates/opens, absolute paths everywhere, and calls
`app.Quit()` only when `Presentations.Count == 0` afterwards, so any deck
the user has open is never disturbed.

pywin32 note: PowerPoint text access needs makepy static wrappers for both
the PowerPoint **and Office (MSO)** type libraries — `TextRange2`'s
`Paragraphs`/`Runs`/`Characters` are parameterized properties that dynamic
dispatch cannot invoke. `powerpoint_app()` generates both automatically;
`com_helpers.com_get()` performs the low-level parameterized property gets.

## Content policy

Synthetic content only: generic business text ("Acme Corp", "Market
overview"), universally available fonts (Calibri, Arial, Georgia,
Times New Roman), no client names, no firm names, no personal data.
