"""Shared placeholder-type -> typography-role mapping (learn + apply).

The profile builder (``utils.profile_extract``) learns typography rules
per role from resolved slide facts; the apply engine
(``utils.style_apply``) writes those rules back to placeholders. Both
sides MUST agree on which placeholder types belong to which role, or
values learned from one population of shapes get applied to a different
one. This module is the single source of truth for that mapping.

Matching is EXACT on the ``PP_PLACEHOLDER`` member name. The learn side
once used substring matching (``"TITLE" in ph_type``), which let
``SUBTITLE (4)`` and ``VERTICAL_TITLE`` vote into the TITLE typography
while the apply side mapped them to no role at all -- a silent
learn/apply asymmetry. Deliberately unmapped (role ``None``): SUBTITLE
(display furniture with its own styling, never restyled as a title),
the VERTICAL_* variants (east-asian layout constructs the house profile
does not model), and every media/graphic placeholder type.
"""

from typing import Dict, Optional

#: ``PP_PLACEHOLDER`` member name -> typography role. Exact match only;
#: anything absent from this table has no role on either side.
ROLE_BY_PLACEHOLDER_NAME: Dict[str, str] = {
    "TITLE": "title",
    "CENTER_TITLE": "title",
    "BODY": "body",
    "OBJECT": "body",
    "FOOTER": "footer",
    "SLIDE_NUMBER": "footer",
    "DATE": "footer",
}


def placeholder_role(name: Optional[str]) -> Optional[str]:
    """Typography role for a ``PP_PLACEHOLDER`` member NAME, else None.

    ``name`` is the bare enum member name (``"TITLE"``, ``"SUBTITLE"``,
    ...); use :func:`ph_type_label_name` first when starting from the
    resolver's serialized label.
    """
    if name is None:
        return None
    return ROLE_BY_PLACEHOLDER_NAME.get(name)


def ph_type_label_name(label: Optional[str]) -> Optional[str]:
    """Member name out of a resolver ``ph_type`` label.

    ``utils.resolve_analysis`` serializes placeholder types as
    ``str(PP_PLACEHOLDER.X)``, which python-pptx renders as
    ``"NAME (value)"`` (e.g. ``"SUBTITLE (4)"`` -> ``"SUBTITLE"``).
    Returns ``None`` for ``None``/empty labels (non-placeholders).
    """
    if not label:
        return None
    return label.split(" (", 1)[0]
