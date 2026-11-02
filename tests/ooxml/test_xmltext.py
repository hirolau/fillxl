"""The text helpers, on the markup a real workbook contains.

Callers pass whole elements as well as start tags, so an attribute lookup that
scans the whole element can pick up a child's attribute — for example the
``t`` of a nested ``<f>`` instead of the cell's own type.
"""

from __future__ import annotations

from fillxl._ooxml import xmltext


def test_an_attribute_named_in_a_child_element_is_not_mistaken_for_the_parents() -> None:
    """Only the start tag of the element being asked about is searched."""
    cell = '<c r="A1"><f t="array" ref="A1:A2">X</f></c>'

    assert xmltext.attribute(cell, "t") is None
    assert xmltext.attribute(cell, "r") == "A1"
