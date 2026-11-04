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


def test_a_single_quoted_attribute_is_read_like_a_double_quoted_one() -> None:
    """XML allows either quote style, and some tools write the single one."""
    assert xmltext.attribute("<c r='B5' s='2'/>", "s") == "2"


def test_replacing_a_single_quoted_attribute_keeps_its_quote_style() -> None:
    """Rewriting a value must not introduce a second attribute with the same name."""
    assert xmltext.set_attribute("<c s='2'/>", "s", "3") == "<c s='3'/>"


def test_a_value_holding_the_other_quote_stays_valid_in_a_single_quoted_attribute() -> None:
    """Escaping the quote character in use is what keeps the tag parseable."""
    assert xmltext.set_attribute("<c s='2'/>", "s", "it's") == "<c s='it&apos;s'/>"


def test_adding_an_attribute_to_a_single_quoted_tag_uses_double_quotes() -> None:
    """A new attribute cannot change the quote style of the ones already there."""
    assert xmltext.set_attribute("<c r='B5'/>", "s", "3") == "<c r='B5' s=\"3\"/>"


def test_a_single_quoted_attribute_is_removed_with_its_quotes() -> None:
    assert xmltext.remove_attribute("<c r='B5' t='s'/>", "t") == "<c r='B5'/>"
