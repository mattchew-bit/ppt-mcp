"""Alignment-grid inference for house-style profiles (Step 3).

Real corporate templates carry an implicit column grid nobody writes
down: shape left / right / center edges repeat at the same x positions
across every well-built slide. This module recovers that grid from a
corpus of slides with a pure-Python 1-D sort-and-merge clustering pass
(the repo deliberately carries no numpy/scipy):

    1. collect every shape's left / right / horizontal-center edge (in
       inches) tagged with the (deck, slide) it came from
    2. sort each edge family and merge neighbours closer than the
       tolerance into clusters
    3. keep clusters supported by at least ``min_slides`` DISTINCT
       slides (a one-slide artifact is not a convention)
    4. report each surviving cluster's mean as a grid line, capped at
       ``max_edges`` per family (highest support first)

Center edges get one extra rule: a shape whose left AND right edges
snap to NON-ADJACENT inferred column boundaries spans more than one
column, so its midpoint is a layout artifact -- not a column center --
and is excluded from center clustering (a subtitle spanning columns
1-2 must not mint a phantom center line between the real ones).

Input is the "slide facts" structure built by ``utils.profile_extract``
(resolved shape geometry in points); output is plain floats in inches --
DTCG leaf shaping happens in ``utils.profile_schema``.

Everything here is pure data-in / data-out: no I/O, no mutation of the
inputs.
"""

from typing import Dict, Iterable, List, Optional, Sequence, Tuple

#: Points per inch.
POINTS_PER_INCH = 72.0

#: Default merge tolerance, in inches (~4.3pt -- comfortably tighter
#: than a gutter, looser than PowerPoint nudge noise).
DEFAULT_TOLERANCE_IN = 0.06

#: A cluster must appear on at least this many distinct slides.
DEFAULT_MIN_SLIDES = 3

#: Cap per edge family (left / right / center) to bound profile size.
DEFAULT_MAX_EDGES = 8

#: An edge observation: (position_in_inches, (deck, slide_number)).
EdgePoint = Tuple[float, Tuple[str, int]]

#: One shape's horizontal extent: (left_in, right_in, (deck, slide)).
ShapeSpan = Tuple[float, float, Tuple[str, int]]


def cluster_1d(points: Sequence[EdgePoint],
               tolerance: float) -> List[Dict]:
    """Sort-and-merge 1-D clustering.

    ``points`` are ``(value, slide_key)`` pairs; consecutive sorted
    values whose gap is <= ``tolerance`` join the same cluster. Returns
    new cluster dicts ``{"center": mean, "count": n, "slides": frozenset,
    "spread": max-min}`` ordered by ascending center. Inputs are never
    mutated.
    """
    if tolerance <= 0:
        raise ValueError(f"tolerance must be positive: {tolerance!r}")
    ordered = sorted(points, key=lambda point: point[0])
    clusters: List[Dict] = []
    current: List[EdgePoint] = []
    for point in ordered:
        if current and point[0] - current[-1][0] > tolerance:
            clusters.append(_close_cluster(current))
            current = []
        current = current + [point]
    if current:
        clusters.append(_close_cluster(current))
    return clusters


def _close_cluster(members: List[EdgePoint]) -> Dict:
    values = [value for value, _ in members]
    return {
        "center": sum(values) / len(values),
        "count": len(values),
        "slides": frozenset(key for _, key in members),
        "spread": max(values) - min(values),
    }


def _shape_spans(slides: Iterable[Dict]) -> List[ShapeSpan]:
    """Collect every shape's horizontal extent (inches) from slide facts."""
    spans: List[ShapeSpan] = []
    for slide in slides:
        key = (slide["deck"], slide["slide_number"])
        for shape in slide["shapes"]:
            geometry = shape["geometry"]
            left = geometry.get("left_pt")
            width = geometry.get("width_pt")
            if left is None or width is None:
                continue  # shapes with no resolvable xfrm say nothing
            left_in = left / POINTS_PER_INCH
            spans.append((left_in, left_in + width / POINTS_PER_INCH, key))
    return spans


def _column_boundaries(left_lines: Sequence[float],
                       right_lines: Sequence[float],
                       tolerance: float) -> List[float]:
    """Sorted, de-duplicated column boundary positions.

    Gutterless grids put a column's right edge and the next column's
    left edge on the same line; merging within the tolerance makes them
    ONE boundary so the adjacency rule below stays correct.
    """
    lines = [(value, ("", 0)) for value in [*left_lines, *right_lines]]
    return [cluster["center"] for cluster in cluster_1d(lines, tolerance)]


def _snap_index(value: float, boundaries: Sequence[float],
                tolerance: float) -> Optional[int]:
    """Index of the nearest boundary within tolerance, else ``None``."""
    best_index, best_distance = None, None
    for index, line in enumerate(boundaries):
        distance = abs(value - line)
        if distance <= tolerance and (best_distance is None
                                      or distance < best_distance):
            best_index, best_distance = index, distance
    return best_index


def _center_points(spans: Sequence[ShapeSpan],
                   boundaries: Sequence[float],
                   tolerance: float) -> List[EdgePoint]:
    """Center-edge observations, minus multi-column spanners.

    A shape whose left and right edges both snap to inferred column
    boundaries contributes its midpoint only when those boundaries are
    ADJACENT (the shape occupies exactly one column). Non-adjacent
    boundaries mean the shape spans columns and its midpoint is a
    layout artifact. Shapes that do not snap on both sides keep their
    say -- absence of evidence is not spanning.
    """
    points: List[EdgePoint] = []
    for left, right, key in spans:
        left_index = _snap_index(left, boundaries, tolerance)
        right_index = _snap_index(right, boundaries, tolerance)
        if (left_index is not None and right_index is not None
                and right_index - left_index != 1):
            continue
        points.append(((left + right) / 2.0, key))
    return points


def _grid_lines(points: List[EdgePoint], tolerance: float,
                min_slides: int, max_edges: int) -> List[float]:
    """Cluster one edge family and keep the supported grid lines."""
    supported = [
        cluster for cluster in cluster_1d(points, tolerance)
        if len(cluster["slides"]) >= min_slides
    ]
    # Highest support wins the cap; ties break on position for
    # deterministic output.
    strongest = sorted(
        supported,
        key=lambda cluster: (-len(cluster["slides"]), cluster["center"]),
    )[:max_edges]
    return sorted(cluster["center"] for cluster in strongest)


def infer_grid(slides: Sequence[Dict],
               tolerance_in: float = DEFAULT_TOLERANCE_IN,
               min_slides: int = DEFAULT_MIN_SLIDES,
               max_edges: int = DEFAULT_MAX_EDGES) -> Dict:
    """Infer the implicit column grid from slide facts.

    Returns ``{"edges": {"left": [...], "right": [...], "center": [...]},
    "tolerance_in": float}`` with edge positions in inches (unrounded --
    the schema writer rounds). Families with no supported cluster come
    back as empty lists rather than being omitted, so consumers can rely
    on the keys.
    """
    if not slides:
        raise ValueError("infer_grid requires at least one slide")
    if min_slides < 1:
        raise ValueError(f"min_slides must be >= 1: {min_slides!r}")
    if max_edges < 1:
        raise ValueError(f"max_edges must be >= 1: {max_edges!r}")
    spans = _shape_spans(slides)
    left_lines = _grid_lines([(left, key) for left, _, key in spans],
                             tolerance_in, min_slides, max_edges)
    right_lines = _grid_lines([(right, key) for _, right, key in spans],
                              tolerance_in, min_slides, max_edges)
    boundaries = _column_boundaries(left_lines, right_lines, tolerance_in)
    center_lines = _grid_lines(
        _center_points(spans, boundaries, tolerance_in),
        tolerance_in, min_slides, max_edges)
    return {
        "edges": {"left": left_lines, "right": right_lines,
                  "center": center_lines},
        "tolerance_in": tolerance_in,
    }


def distance_to_nearest(value: float, grid_lines: Sequence[float]) -> float:
    """Distance from ``value`` to the nearest grid line (for lint use)."""
    if not grid_lines:
        raise ValueError("grid_lines must not be empty")
    return min(abs(value - line) for line in grid_lines)
