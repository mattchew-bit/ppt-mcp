"""Rebuild the entire house corpus from source (one command).

Pipeline (each stage fails loudly; nothing is committed by this script):
    1. make_images         -- Pillow PNG assets (committed, deterministic)
    2. author_house_decks  -- house_01..house_05 via PowerPoint COM
    3. author_deviant      -- deviant_01 with 9 seeded violations
    4. write_metadata      -- corpus_truth.json + deviations.json
    5. transform_check     -- COM-extract house_01, arbitrate the
                              resolver's lumMod/lumOff math (exact match)
    6. verify_corpus       -- python-pptx self-check + cross-references
    7. export_corpus_previews -- optional, --previews DIR

Requires desktop PowerPoint + pywin32 (same as tests/fixtures/authoring).
"""

from __future__ import annotations

import argparse
import sys

import _bootstrap  # noqa: F401

import author_deviant
import author_house_decks
import make_images
import transform_check
import verify_corpus
import write_metadata


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--previews", metavar="DIR", default=None,
                        help="also export contact-sheet PNGs to DIR")
    args = parser.parse_args()

    print("== 1/6 image assets ==")
    make_images.main()
    print("== 2/6 house decks ==")
    author_house_decks.main()
    print("== 3/6 deviant deck ==")
    author_deviant.main()
    print("== 4/6 metadata ==")
    write_metadata.main()
    print("== 5/6 transform check ==")
    transform_ok = transform_check.main()
    print("== 6/6 corpus verify ==")
    verify_ok = verify_corpus.main()

    if args.previews:
        import export_corpus_previews
        export_corpus_previews.export_all(args.previews)

    return 0 if (transform_ok and verify_ok) else 1


if __name__ == "__main__":
    sys.exit(main())
