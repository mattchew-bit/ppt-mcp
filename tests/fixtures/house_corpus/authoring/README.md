# House corpus authoring (Step 3 of the style-fidelity plan)

Regenerable source for `tests/fixtures/house_corpus/`: five conformant
"Meridian" house decks, one deviant deck with seeded violations, and the
machine-readable ground truth (`corpus_truth.json` / `deviations.json`)
that Step 3's `create_house_profile` / `apply_style_profile` / lint work
is built and verified against.

Like the Step 0 fixtures, **every deck is written by desktop PowerPoint
via COM** (never python-pptx), reusing `tests/fixtures/authoring/
com_helpers.py` and its hygiene rules: `CoInitialize`, never
`Visible = False`, own presentations only, `Quit()` only when
`Presentations.Count == 0`, absolute paths, `SaveAs` format 24,
`AutomationSecurity = ForceDisable`.

## Rebuild

```
cd tests/fixtures/house_corpus/authoring
python build_all.py                 # decks + metadata + checks
python build_all.py --previews DIR  # also export contact-sheet PNGs
```

Requires desktop PowerPoint, `pywin32`, `python-pptx`, `Pillow` (assets
are committed, so Pillow is only needed to regenerate them).

## Layout

| File | Role |
|---|---|
| `_bootstrap.py` | sys.path bridge to repo root + Step 0 authoring kit |
| `house_style.py` | **single source of truth** for every seeded convention |
| `make_images.py` | Pillow-drawn PNG exhibits (`assets/`, ~3 KB total) |
| `slide_archetypes.py` | one builder per archetype, all grid-snapped |
| `author_house_decks.py` | deck specs + builds `house_01..05.pptx` |
| `author_deviant.py` | `deviant_01.pptx` + the `DEVIATIONS` registry |
| `write_metadata.py` | serializes constants -> `corpus_truth.json`, `deviations.json` |
| `transform_check.py` | COM-arbitrates the resolver's lumMod/lumOff math |
| `verify_corpus.py` | python-pptx self-check (opens, cross-refs, grid audit) |
| `export_corpus_previews.py` | contact-sheet PNGs (outside the repo) |

## The Meridian conventions (all deliberately non-default)

Recorded in full in `../corpus_truth.json`; headline values:

- **Theme**: accent1 `#1B7F79` teal, accent2 `#D97C2B` amber, dk2
  `#14324F` navy, lt1 `#FAF9F6` warm white; Georgia major / Calibri
  minor. Differs from both the Office default and the Step 0 fixture
  theme (`C0504D` / Georgia+Arial).
- **Type scale**: {11, 14, 20, 30}pt — no member equals the 18pt
  default. Title 30pt Georgia regular, left, dk2 (master titleStyle).
- **Body (master bodyStyle)**: l1 14pt / l2+l3 11pt dk1;
  space before/after 2/8, 2/5, 1/4 pt; line spacing 1.20; bullets
  l1 `—` (em dash), l2 `·` (middle dot), l3 `>` — typable, non-default.
- **Shapes**: 1.25pt dashed `#14324F` borders, corner radius adj 0.12,
  panel fill lt2 `#DCE3E8`.
- **Grid**: 3 columns, left edges 60/360/660pt, right edges
  300/600/900pt (0.8333/5/9.1667 in and 4.1667/8.3333/12.5 in), gutter
  60pt, tolerance 4pt. Every shape on every house slide snaps.
- **Archetypes** (27 labeled slides): title x5, agenda x4,
  section_divider x4, content x6, two_column x5, closing x3. Per-slide
  labels + per-archetype geometry live in the truth file.
- **Images**: sidebar exhibit zone (660, 120, 240, 180)pt on all 9
  agenda/two_column slides; assets drawn in palette colors.
- **Footer**: zone y 500–528pt; source note left (11pt accent3), page
  number right-aligned; on agenda/content/two_column/closing.

## Theme-color tints and the transform check

Three slides carry PowerPoint "Lighter/Darker N%" theme variants
(`a:lumMod`/`a:lumOff` in the XML): house_01 slide 3 (rule fill
accent1 +40%, kicker font accent2 −25%), house_01 slide 5 (both column
panels accent1 +80%), house_02 slide 5 (rule accent1 +60%).

`transform_check.py` closes the Step 2 known gap (transform math
unit-tested but never fixture-arbitrated): it records COM effective
values for house_01 with the **existing** Step 0 extractor into
`../expected_values/house_01.json` and asserts COM truth == resolver
output == `apply_color_transforms` prediction, exactly, for every
tinted element — plus an XML guard that the tints are live `schemeClr`
references, not baked literals.

Empirical note (112-case COM probe, 2026-07-05): the repo's HSL float
math matches PowerPoint's effective RGB except when a channel lands
exactly on a `x.5` boundary, where PowerPoint's own rounding is
direction-inconsistent (its UI picker label can differ from its
effective color, e.g. the documented `8EAADC` vs `8FAADC` case). The
corpus therefore seeds brightness values verified stable (+0.8, +0.6,
+0.4, −0.25) — the checked values are boundary-free.

## Deviant deck

`deviant_01.pptx` — same Meridian template, 9 seeded violations
(off-scale 13pt run, off-grid panel at +22pt, `•` bullet, hardcoded
`srgbClr`, footer straggler box, 12pt space_after, 2.5pt border,
off-palette `#8E44AD` fill, Times New Roman run), each recorded in
`../deviations.json` with slide/shape/paragraph refs, expected and
actual values. Everything else in the deck is house-conformant, so a
linter must flag exactly these and nothing else.
