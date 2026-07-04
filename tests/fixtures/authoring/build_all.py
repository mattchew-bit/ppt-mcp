"""Regenerate all four style-fidelity fixture decks end to end.

Pipeline (single PowerPoint COM session for speed):
  1. Author the four decks   -> tests/fixtures/<fixture>.pptx
  2. Extract effective values -> tests/fixtures/expected_values/<fixture>.json
  3. Export slide previews    -> --previews-dir (outside the repo)
  4. Self-check (python-pptx open + JSON sanity + seeded-value guards)

Usage:
  py -3 build_all.py [--previews-dir DIR] [--skip-previews] [--skip-verify]

Idempotent: every output is overwritten in place.
"""

from __future__ import annotations

import argparse
import os
import sys

import author_explicit_override
import author_layout_override
import author_multi_master
import author_theme_only
import export_previews
import extract_expected
import verify_fixtures
from com_helpers import powerpoint_app

AUTHORS = [
    author_theme_only,
    author_layout_override,
    author_explicit_override,
    author_multi_master,
]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--previews-dir",
                        default=export_previews.default_output_dir(),
                        help="where slide PNGs are written (not the repo)")
    parser.add_argument("--skip-previews", action="store_true")
    parser.add_argument("--skip-verify", action="store_true")
    args = parser.parse_args()

    fixtures_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    expected_dir = os.path.join(fixtures_dir, "expected_values")

    with powerpoint_app() as app:
        for author in AUTHORS:
            output = os.path.join(fixtures_dir, f"{author.FIXTURE_NAME}.pptx")
            path = author.build(app, output)
            print(f"Authored {path}")

        extract_expected.extract_all(app, fixtures_dir, expected_dir)

        if not args.skip_previews:
            exported = export_previews.export_all(app, fixtures_dir,
                                                  args.previews_dir)
            print(f"Wrote {len(exported)} preview PNGs to {args.previews_dir}")

    if not args.skip_verify:
        return verify_fixtures.main()
    return 0


if __name__ == "__main__":
    sys.exit(main())
