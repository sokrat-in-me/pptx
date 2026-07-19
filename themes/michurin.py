"""Apply the MICHURIN PowerPoint theme from cherkizovo-presentations.md."""

from __future__ import annotations

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

from themes.cherkizovo import OFFICE_COLOR_SCHEME_HEX, THEME_NAME, TITLE_FONT

_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NSMAP = {"a": _DRAWING_NS}


def _set_scheme_color(parent: etree._Element, hex_value: str) -> None:
    for child in list(parent):
        parent.remove(child)
    srgb = etree.SubElement(parent, f"{{{_DRAWING_NS}}}srgbClr")
    srgb.set("val", hex_value)


def apply_michurin_theme(prs: Presentation) -> None:
    """Apply MICHURIN color scheme, fonts, and hyperlink colors to a presentation."""
    theme_part = prs.slide_master.part.part_related_by(RT.THEME)
    theme = etree.fromstring(theme_part.blob)
    theme.set("name", THEME_NAME)

    for clr_scheme in theme.xpath("//a:clrScheme", namespaces=_NSMAP):
        clr_scheme.set("name", THEME_NAME)
        for slot, hex_value in OFFICE_COLOR_SCHEME_HEX.items():
            for element in clr_scheme.xpath(f"a:{slot}", namespaces=_NSMAP):
                _set_scheme_color(element, hex_value)

    for font_scheme in theme.xpath("//a:fontScheme", namespaces=_NSMAP):
        font_scheme.set("name", THEME_NAME)
        for path in ("majorFont", "minorFont"):
            for latin in font_scheme.xpath(f"a:{path}/a:latin", namespaces=_NSMAP):
                latin.set("typeface", TITLE_FONT)

    theme_part._blob = etree.tostring(theme, encoding="UTF-8", xml_declaration=True)
