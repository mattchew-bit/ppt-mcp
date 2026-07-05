"""Tests for ``utils/style_roles.py`` -- the shared learn/apply role map.

The learn side (``utils.profile_extract._shape_role``) and the apply
side (``utils.style_apply._role_of``) must classify placeholder types
identically. Substring matching once let ``SUBTITLE (4)`` and
``VERTICAL_TITLE`` vote into the TITLE typography on the learn side
while the apply side mapped them to no role -- these tests pin the
shared exact-match table on both sides.
"""

from types import SimpleNamespace

import pytest
from pptx.enum.shapes import PP_PLACEHOLDER

from tests.conftest import load_corpus_truth, skip_if_house_corpus_missing
from utils.profile_extract import _shape_role, _typography_section
from utils.style_apply import _role_of
from utils.style_roles import ph_type_label_name, placeholder_role

# ------------------------------------------------------------- unit map


@pytest.mark.parametrize("name,role", [
    ("TITLE", "title"),
    ("CENTER_TITLE", "title"),
    ("BODY", "body"),
    ("OBJECT", "body"),
    ("FOOTER", "footer"),
    ("SLIDE_NUMBER", "footer"),
    ("DATE", "footer"),
])
def test_mapped_placeholder_names(name, role):
    assert placeholder_role(name) == role


@pytest.mark.parametrize("name", [
    "SUBTITLE", "VERTICAL_TITLE", "VERTICAL_BODY", "VERTICAL_OBJECT",
    "PICTURE", "CHART", "TABLE", "MEDIA_CLIP", "MIXED", "no_such_type",
    None,
])
def test_unmapped_placeholder_names_have_no_role(name):
    assert placeholder_role(name) is None


def test_label_name_strips_python_pptx_suffix():
    assert ph_type_label_name("SUBTITLE (4)") == "SUBTITLE"
    assert ph_type_label_name("TITLE (1)") == "TITLE"
    assert ph_type_label_name("MIXED (-2)") == "MIXED"
    assert ph_type_label_name(None) is None
    assert ph_type_label_name("") is None


# ---------------------------------------------------- learn/apply symmetry


class _FakePlaceholderShape:
    """Just enough shape surface for ``_role_of``."""

    def __init__(self, ph_type):
        self.is_placeholder = True
        self.placeholder_format = SimpleNamespace(type=ph_type)


@pytest.mark.parametrize("member", list(PP_PLACEHOLDER),
                         ids=lambda member: member.name)
def test_learn_and_apply_sides_agree_for_every_placeholder_type(member):
    """Regression: the two sides once disagreed on SUBTITLE and the
    VERTICAL_* types. The learn side reads the resolver's serialized
    label; the apply side reads the live enum -- both must land on the
    same role for every placeholder type python-pptx can report."""
    learn_role = placeholder_role(ph_type_label_name(str(member)))
    apply_role = _role_of(_FakePlaceholderShape(member))
    assert learn_role == apply_role, member.name


def test_role_of_non_placeholder_is_none():
    shape = SimpleNamespace(is_placeholder=False, placeholder_format=None)
    assert _role_of(shape) is None


def test_role_of_tolerates_none_placeholder_type():
    assert _role_of(_FakePlaceholderShape(None)) is None


# ------------------------------------------- corpus SubtitleBox regression


def _retag_subtitle_boxes(slides):
    """New slide facts with every corpus SubtitleBox re-tagged as the
    SUBTITLE placeholder PowerPoint's title layout would emit."""
    return [
        {**slide, "shapes": [
            ({**shape, "is_placeholder": True, "ph_type": "SUBTITLE (4)"}
             if shape.get("name") == "SubtitleBox" else shape)
            for shape in slide["shapes"]
        ]}
        for slide in slides
    ]


@skip_if_house_corpus_missing()
def test_subtitle_placeholder_gets_no_learn_role(house_facts):
    """The corpus SubtitleBox (20pt accent3, one per title slide), in
    its placeholder form, must not be counted as a title on the learn
    side -- ``"TITLE" in "SUBTITLE (4)"`` once made it vote."""
    slides = _retag_subtitle_boxes(house_facts["slides"])
    subtitle_shapes = [
        (shape, slide) for slide in slides for shape in slide["shapes"]
        if shape.get("ph_type") == "SUBTITLE (4)"
    ]
    assert len(subtitle_shapes) == 5  # one per corpus title slide
    for shape, slide in subtitle_shapes:
        assert _shape_role(shape, slide["height_pt"]) is None


@skip_if_house_corpus_missing()
def test_subtitle_placeholder_does_not_pollute_title_typography(house_facts):
    """Learned TITLE typography stays at the seeded master style even
    when the subtitles are placeholders (they carry a different font,
    size and color; any leakage would surface in these counters)."""
    slides = _retag_subtitle_boxes(house_facts["slides"])
    section, _ = _typography_section(slides)
    truth = load_corpus_truth()["typography"]["title"]
    assert section["title"]["font"] == {"value": truth["font"]}
    assert section["title"]["size"] == {"value": truth["size_pt"],
                                        "unit": "pt"}
    assert section["title"]["color"] == {
        "value": f"#{truth['color'].lstrip('#').upper()}"}
