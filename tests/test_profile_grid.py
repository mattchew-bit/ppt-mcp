"""Tests for ``utils/profile_grid.py`` -- 1-D clustering + grid inference.

Unit tests pin the pure-Python sort-and-merge clustering; the corpus
tests assert the seeded Meridian 3-column grid (corpus_truth.json) is
recovered within 0.05in, per the Step 3 acceptance criterion.
"""

import pytest

from tests.conftest import load_corpus_truth, skip_if_house_corpus_missing
from utils.profile_grid import (
    POINTS_PER_INCH,
    cluster_1d,
    distance_to_nearest,
    infer_grid,
)

# ------------------------------------------------------------- helpers


def _slide(deck, number, rects):
    """Minimal slide-facts record from (left, top, w, h) point rects."""
    return {
        "deck": deck,
        "slide_number": number,
        "width_pt": 960.0,
        "height_pt": 540.0,
        "shapes": [
            {"geometry": {"left_pt": left, "top_pt": top,
                          "width_pt": width, "height_pt": height}}
            for left, top, width, height in rects
        ],
    }


# ------------------------------------------------------------- cluster_1d


def test_cluster_merges_within_tolerance():
    points = [(1.00, ("d", 1)), (1.03, ("d", 2)), (5.0, ("d", 3))]
    clusters = cluster_1d(points, tolerance=0.06)
    assert len(clusters) == 2
    assert clusters[0]["center"] == pytest.approx(1.015)
    assert clusters[0]["count"] == 2
    assert clusters[0]["slides"] == frozenset({("d", 1), ("d", 2)})
    assert clusters[1]["center"] == pytest.approx(5.0)


def test_cluster_splits_beyond_tolerance():
    points = [(1.0, ("d", 1)), (1.1, ("d", 2))]
    clusters = cluster_1d(points, tolerance=0.06)
    assert [cluster["count"] for cluster in clusters] == [1, 1]


def test_cluster_chains_adjacent_gaps():
    # Sort-and-merge chains: consecutive gaps each within tolerance
    # join one cluster even when the extremes are further apart.
    points = [(0.0, ("d", 1)), (0.05, ("d", 2)), (0.10, ("d", 3))]
    clusters = cluster_1d(points, tolerance=0.06)
    assert len(clusters) == 1
    assert clusters[0]["spread"] == pytest.approx(0.10)


def test_cluster_does_not_mutate_input():
    points = [(2.0, ("d", 2)), (1.0, ("d", 1))]
    snapshot = list(points)
    cluster_1d(points, tolerance=0.5)
    assert points == snapshot


def test_cluster_rejects_bad_tolerance():
    with pytest.raises(ValueError, match="tolerance"):
        cluster_1d([(1.0, ("d", 1))], tolerance=0.0)


def test_cluster_empty_points():
    assert cluster_1d([], tolerance=0.1) == []


# ------------------------------------------------------------- infer_grid


def test_infer_grid_keeps_only_supported_clusters():
    # Left edge at 60pt on three slides; 500pt on one slide only.
    slides = [
        _slide("a", 1, [(60, 100, 240, 50)]),
        _slide("a", 2, [(60, 100, 240, 50)]),
        _slide("b", 1, [(60, 100, 240, 50), (500, 100, 100, 50)]),
    ]
    grid = infer_grid(slides, min_slides=3)
    assert grid["edges"]["left"] == [pytest.approx(60 / POINTS_PER_INCH)]


def test_infer_grid_counts_distinct_slides_not_shapes():
    # Five identical shapes on ONE slide must not fake support.
    slides = [_slide("a", 1, [(60, 100, 240, 50)] * 5)]
    grid = infer_grid(slides, min_slides=2)
    assert grid["edges"]["left"] == []


def test_infer_grid_caps_edges_by_support():
    slides = [
        _slide("a", number, [
            (60, 100, 100, 50),          # every slide
            (60 + 100 * number, 200, 50, 50),  # unique per slide pair
        ])
        for number in range(1, 7)
    ] + [
        _slide("b", number, [(60, 100, 100, 50)])
        for number in range(1, 7)
    ]
    grid = infer_grid(slides, min_slides=1, max_edges=1)
    # Only the strongest cluster survives the cap.
    assert grid["edges"]["left"] == [pytest.approx(60 / POINTS_PER_INCH)]


def test_infer_grid_skips_shapes_without_geometry():
    slides = [_slide("a", 1, [(60, 100, 240, 50)])]
    slides[0]["shapes"].append({"geometry": {"left_pt": None,
                                             "width_pt": None}})
    grid = infer_grid(slides, min_slides=1)
    assert len(grid["edges"]["left"]) == 1


def test_infer_grid_validates_inputs():
    with pytest.raises(ValueError, match="at least one slide"):
        infer_grid([])
    with pytest.raises(ValueError, match="min_slides"):
        infer_grid([_slide("a", 1, [])], min_slides=0)
    with pytest.raises(ValueError, match="max_edges"):
        infer_grid([_slide("a", 1, [])], max_edges=0)


def test_distance_to_nearest():
    assert distance_to_nearest(1.1, [1.0, 5.0]) == pytest.approx(0.1)
    with pytest.raises(ValueError):
        distance_to_nearest(1.0, [])


# ------------------------------------------- span-center adjacency rule


def test_spanning_shape_contributes_no_center_edge():
    # SubtitleBox regression: a shape spanning columns 1-2 lands its
    # left/right on NON-ADJACENT column boundaries; its midpoint
    # (330pt = 4.583in) is a layout artifact, not a column center.
    column_1 = (60, 100, 240, 50)     # center 180pt = 2.5in
    column_2 = (360, 100, 240, 50)    # center 480pt = 6.667in
    spanner = (60, 200, 540, 60)      # left col1, right col2 -> spans
    slides = [_slide("a", number, [column_1, column_2, spanner])
              for number in (1, 2, 3)]
    grid = infer_grid(slides, min_slides=3)
    assert grid["edges"]["center"] == [
        pytest.approx(180 / POINTS_PER_INCH),
        pytest.approx(480 / POINTS_PER_INCH),
    ]


def test_off_grid_shape_still_contributes_center_edge():
    # A shape whose edges snap to NO inferred boundary cannot be proven
    # to span columns, so its midpoint keeps its say.
    column = (60, 100, 240, 50)       # boundaries 0.833in / 4.167in
    floater = (400, 300, 100, 40)     # left 5.556in, right 6.944in
    slides = [_slide("a", number, [column, floater])
              for number in (1, 2, 3)]
    grid = infer_grid(slides, min_slides=3)
    assert grid["edges"]["center"] == [
        pytest.approx(180 / POINTS_PER_INCH),
        pytest.approx(450 / POINTS_PER_INCH),
    ]


# ------------------------------------------- footer-furniture exclusion


def test_grid_section_excludes_footer_zone_shapes():
    """Footer boxes must not mint column edges (left/right/center)."""
    from utils.profile_extract import _grid_section

    body = (60.0, 100.0, 240.0, 50.0)
    footer_page = (780.0, 502.0, 120.0, 22.0)  # top at 93% of height
    slides = [_slide("a", number, [body, footer_page])
              for number in (1, 2, 3)]
    section = _grid_section(slides)
    assert section["edges"]["left"] == [pytest.approx(0.83)]
    assert section["edges"]["right"] == [pytest.approx(4.17)]
    assert section["edges"]["center"] == [pytest.approx(2.5)]


def test_footerless_filter_does_not_mutate_slide_facts():
    from utils.profile_extract import _footerless_slides

    slides = [_slide("a", 1, [(60, 100, 240, 50), (780, 502, 120, 22)])]
    shape_count = len(slides[0]["shapes"])
    filtered = _footerless_slides(slides)
    assert len(filtered[0]["shapes"]) == 1
    assert len(slides[0]["shapes"]) == shape_count  # input untouched


# ------------------------------------------------- corpus grid recovery


@skip_if_house_corpus_missing()
def test_seeded_grid_recovered_within_tolerance(house_profile):
    """Every seeded Meridian grid edge is learned within 0.05in."""
    truth = load_corpus_truth()["grid"]
    edges = house_profile["grid"]["edges"]
    for family, truth_key in (("left", "left_edges"),
                              ("right", "right_edges"),
                              ("center", "center_edges")):
        for seeded in truth[truth_key]["in"]:
            nearest = min(abs(seeded - learned)
                          for learned in edges[family])
            assert nearest <= 0.05, (
                f"{family} edge {seeded}in not recovered: {edges[family]}"
            )


@skip_if_house_corpus_missing()
def test_inferred_grid_has_zero_extra_edges(house_profile):
    """Converse of the recovery test: the learned COLUMN grid is the
    seeded 9 edges EXACTLY -- footer furniture and multi-column
    spanners must not mint extras (the Step 3 gate failure)."""
    truth = load_corpus_truth()["grid"]
    edges = house_profile["grid"]["edges"]
    for family, truth_key in (("left", "left_edges"),
                              ("right", "right_edges"),
                              ("center", "center_edges")):
        seeded = truth[truth_key]["in"]
        learned = edges[family]
        assert len(learned) == len(seeded), (
            f"{family}: learned {len(learned)} edges, seeded "
            f"{len(seeded)}: {learned}"
        )
        for edge in learned:
            nearest = min(abs(edge - value) for value in seeded)
            assert nearest <= 0.05, (
                f"extra {family} edge {edge}in not in seeded {seeded}"
            )


@skip_if_house_corpus_missing()
def test_grid_section_shape_and_tolerance(house_profile):
    grid = house_profile["grid"]
    assert grid["unit"] == "in"
    assert set(grid["edges"]) == {"left", "right", "center"}
    assert grid["tolerance"]["unit"] == "in"
    assert 0.02 <= grid["tolerance"]["value"] <= 0.1
    for family in grid["edges"].values():
        assert family == sorted(family)
        assert len(family) <= 8


@skip_if_house_corpus_missing()
def test_off_grid_position_is_detectable(house_profile):
    """The deviant deck's 682pt panel left is off the learned grid."""
    lefts = house_profile["grid"]["edges"]["left"]
    off_grid_in = 682.0 / POINTS_PER_INCH
    tolerance = house_profile["grid"]["tolerance"]["value"]
    assert distance_to_nearest(off_grid_in, lefts) > tolerance
