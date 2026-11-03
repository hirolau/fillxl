"""Text-level XML editing for OOXML parts.

A surgical editor keeps every byte it does not have to change.  Parsing a part
with :mod:`xml.etree.ElementTree` and writing it back renames namespace
prefixes and drops namespace declarations that only ``mc:Ignorable`` refers to
(that attribute lists *prefixes*), which makes Excel refuse the workbook.  The
parts this engine edits are therefore changed as text: element and attribute
surgery on the original string.

The helpers here are deliberately small and conservative: they find elements by
name (with any namespace prefix) and edit attributes without touching anything
around them.
"""

from __future__ import annotations

import re

_NAME = r"[A-Za-z_][\w.\-]*"
_ANY_PREFIX = rf"(?:{_NAME}:)?"

_CHARACTER_REFERENCE = re.compile(r"&(?:(#\d+)|#x([0-9A-Fa-f]+)|(lt|gt|amp|quot|apos));")
_NAMED_ENTITIES = {"lt": "<", "gt": ">", "amp": "&", "quot": '"', "apos": "'"}


def unescape(text: str) -> str:
    """The text of *text* with XML character and entity references decoded.

    Examples:
        >>> unescape("Region &amp; Area")
        'Region & Area'
        >>> unescape("A&#10;B")
        'A\\nB'

    """

    def replace(match: re.Match[str]) -> str:
        decimal, hexadecimal, named = match.groups()
        if decimal is not None:
            return chr(int(decimal[1:]))
        if hexadecimal is not None:
            return chr(int(hexadecimal, 16))
        return _NAMED_ENTITIES[named]

    return _CHARACTER_REFERENCE.sub(replace, text)


def escape_text(text: str) -> str:
    """*text* escaped for use as XML element content.

    A carriage return is escaped: XML parsers normalise a literal ``\\r`` (and
    a ``\\r\\n`` pair) to ``\\n`` when they read the part, which would silently
    change the user's text.

    Examples:
        >>> escape_text("a < b & c")
        'a &lt; b &amp; c'
        >>> escape_text("a\\r\\nb")
        'a&#13;\\nb'

    """
    return (
        text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\r", "&#13;")
    )


def escape_attribute(text: str, quote: str = '"') -> str:
    """*text* escaped for use as an XML attribute value delimited by *quote*.

    Examples:
        >>> escape_attribute('say "hi" & bye')
        'say &quot;hi&quot; &amp; bye'
        >>> escape_attribute("it's", quote="'")
        'it&apos;s'

    """
    escaped = escape_text(text).replace('"', "&quot;")
    return escaped.replace("'", "&apos;") if quote == "'" else escaped


def any_prefix(name: str) -> str:
    """A regular expression fragment matching *name* with any namespace prefix.

    Examples:
        >>> re.fullmatch(any_prefix("row"), "x:row") is not None
        True
        >>> re.fullmatch(any_prefix("row"), "row") is not None
        True

    """
    return rf"{_ANY_PREFIX}{name}"


def element_re(name: str) -> re.Pattern[str]:
    """A pattern matching a whole *name* element, self-closing or not.

    Examples:
        >>> element_re("f").search("<c><f>A1</f></c>").group(0)
        '<f>A1</f>'
        >>> element_re("f").search("<c:f/>").group(0)
        '<c:f/>'

    """
    tag = any_prefix(name)
    return re.compile(rf"<{tag}\b[^>]*?(?:/>|>.*?</{tag}>)", re.S)


def start_tag_re(name: str) -> re.Pattern[str]:
    """A pattern matching the start tag of *name* (with or without a prefix)."""
    return re.compile(rf"<{any_prefix(name)}\b[^>]*>")


def start_tag(xml: str, name: str) -> str | None:
    """The start tag of the first *name* element in *xml*, or ``None``.

    Examples:
        >>> start_tag('<c:f ref="A1:B2"/>', "f")
        '<c:f ref="A1:B2"/>'
        >>> start_tag('<table ref="A1:B2"/>', "row") is None
        True

    """
    match = start_tag_re(name).search(xml)
    return None if match is None else match.group(0)


def _attribute_re(name: str) -> re.Pattern[str]:
    """A pattern matching attribute *name* with either quote style.

    Groups: 1 is the whitespace and name up to the opening quote, ``quote`` is
    the quote character, and ``value`` is the raw (still escaped) value.
    """
    return re.compile(
        rf"""(\s{re.escape(name)}\s*=\s*)(?P<quote>["'])(?P<value>(?:(?!(?P=quote)).)*)(?P=quote)""",
        re.S,
    )


def attribute(tag: str, name: str) -> str | None:
    """The unescaped value of attribute *name* on the start tag *tag*.

    Only the start tag is searched: callers pass whole elements too, and a
    nested element may carry an attribute of the same name.

    Examples:
        >>> attribute('<c r="B5" s="2"/>', "s")
        '2'
        >>> attribute('<c r="B5"/>', "s") is None
        True
        >>> attribute('<c r="A1"><f t="array">X</f></c>', "t") is None
        True
        >>> attribute("<c s='2'/>", "s")
        '2'

    """
    opening = tag[: tag.find(">") + 1]
    match = _attribute_re(name).search(opening)
    if match is None:
        return None
    return unescape(match.group("value"))


def set_attribute(tag: str, name: str, value: str) -> str:
    """*tag* with attribute *name* set to *value*, added when it is missing.

    Examples:
        >>> set_attribute('<c r="B5"/>', "s", "2")
        '<c r="B5" s="2"/>'
        >>> set_attribute('<c r="B5" s="2"/>', "s", "3")
        '<c r="B5" s="3"/>'
        >>> set_attribute("<c s='2'/>", "s", "3")
        "<c s='3'/>"

    """
    pattern = _attribute_re(name)
    existing = pattern.search(tag)
    if existing is not None:
        quote = existing.group("quote")
        escaped = escape_attribute(value, quote)
        replacement = f"{existing.group(1)}{quote}{escaped}{quote}"
        return f"{tag[: existing.start()]}{replacement}{tag[existing.end() :]}"
    escaped = escape_attribute(value)
    split = len(tag) - (2 if tag.endswith("/>") else 1)
    prefix = tag[:split].rstrip()
    return f'{prefix} {name}="{escaped}"{tag[split:]}'


def remove_attribute(tag: str, name: str) -> str:
    """*tag* without attribute *name*.

    Examples:
        >>> remove_attribute('<c r="B5" t="s"/>', "t")
        '<c r="B5"/>'
        >>> remove_attribute("<c r='B5' t='s'/>", "t")
        "<c r='B5'/>"

    """
    return _attribute_re(name).sub("", tag, count=1)


def element_text(xml: str, name: str) -> str | None:
    """The unescaped character content of the first *name* element in *xml*.

    ``None`` is returned when the element is missing, self-closing or empty.

    Examples:
        >>> element_text("<t>hello</t>", "t")
        'hello'
        >>> element_text("<a><b>x</b></a>", "b")
        'x'
        >>> element_text("<a><b/></a>", "b") is None
        True

    """
    match = element_re(name).search(xml)
    if match is None:
        return None
    element = match.group(0)
    opening_end = element.find(">")
    closing_start = element.rfind("</")
    if closing_start < 0:
        return None
    inner = element[opening_end + 1 : closing_start]
    return unescape(inner) if inner else None


def set_tag_attribute(element: str, name: str, value: str) -> str:
    """*element* with attribute *name* set on its **start** tag.

    :func:`set_attribute` works on a start tag; passing a whole element would
    append the attribute to its closing tag instead, which corrupts the part.
    This helper makes that mistake impossible.

    Examples:
        >>> set_tag_attribute('<c><v>1</v></c>', "s", "2")
        '<c s="2"><v>1</v></c>'
        >>> set_tag_attribute('<c r="A1"/>', "s", "2")
        '<c r="A1" s="2"/>'

    """
    opening_end = element.find(">")
    if opening_end < 0:  # pragma: no cover - an element always has a start tag
        return element
    opening = set_attribute(element[: opening_end + 1], name, value)
    return f"{opening}{element[opening_end + 1 :]}"


def replace_element_text(xml: str, name: str, text: str) -> str:
    """*xml* with the character content of its first *name* element replaced.

    A self-closing element is returned unchanged: there is no content to
    replace, and turning ``<f/>`` into ``<f>x</f>`` would be a structural edit
    this helper does not model.

    Examples:
        >>> replace_element_text("<c><f>A1</f></c>", "f", "B2")
        '<c><f>B2</f></c>'
        >>> replace_element_text("<c><f/></c>", "f", "B2")
        '<c><f/></c>'

    """
    match = element_re(name).search(xml)
    if match is None:
        return xml
    element = match.group(0)
    opening_end = element.find(">")
    closing_start = element.rfind("</")
    if closing_start < 0 or closing_start < opening_end:
        return xml
    replacement = element[: opening_end + 1] + escape_text(text) + element[closing_start:]
    return f"{xml[: match.start()]}{replacement}{xml[match.end() :]}"


def replace_start_tag(xml: str, name: str, tag: str) -> str:
    """*xml* with the start tag of its first *name* element replaced by *tag*."""
    match = start_tag_re(name).search(xml)
    if match is None:
        return xml
    return f"{xml[: match.start()]}{tag}{xml[match.end() :]}"
