"""MICHURIN theme constants and PowerPoint theme application.

Style guide: ai/rules/cherkizovo-presentations.md
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

_REPO_ROOT = Path(__file__).resolve().parents[1]
STYLE_GUIDE_PATH = _REPO_ROOT / "ai" / "rules" / "cherkizovo-presentations.md"

THEME_NAME = "МИЧУРИН"

# Slide format (section 5)
SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5

# Fonts (sections 1, 4)
TITLE_FONT = "Verdana"
TABLE_FONT = "Calibri"
PAGE_FONT = "Calibri"

# Typography, pt (section 4)
TITLE_SIZE = 18
TABLE_HEADER_SIZE = 12
TABLE_BODY_SIZE = 11
PAGE_NUMBER_SIZE = 9

# Colors, RGB (sections 3, 6)
COLORS = {
    "brand_red": (215, 0, 54),          # #D70036
    "dark_red": (175, 24, 46),          # #AF182E — заголовки таблиц
    "dark_burgundy": (87, 12, 23),      # #570C17
    "dark_gray": (69, 69, 71),          # #454547
    "title": (131, 18, 35),             # #831223
    "key": (119, 119, 122),             # #77777A
    "critical": (255, 0, 0),            # #FF0000
    "alt_row": (248, 231, 232),         # #F8E7E8
    "white": (255, 255, 255),
    "black": (0, 0, 0),
    "page_number": (154, 160, 166),     # #9AA0A6
    "green_accent": (25, 150, 70),      # #199646
    "salmon": (240, 146, 159),          # #F0929F
    "lime": (138, 238, 141),            # #8AEE8D
}

# Office theme color slots (section 3) — for ppt/theme/theme1.xml
_OFFICE_SLOT_COLORS = {
    "dk1": "black",
    "lt1": "white",
    "dk2": "dark_gray",
    "lt2": "alt_row",
    "accent1": "brand_red",
    "accent2": "dark_red",
    "accent3": "green_accent",
    "accent4": "title",
    "accent5": "salmon",
    "accent6": "lime",
    "hlink": "key",
    "folHlink": "dark_gray",
}


def rgb_hex(name: str) -> str:
    red, green, blue = COLORS[name]
    return f"{red:02X}{green:02X}{blue:02X}"


OFFICE_COLOR_SCHEME_HEX = {
    slot: rgb_hex(color_name) for slot, color_name in _OFFICE_SLOT_COLORS.items()
}

# Layout, inches (section 5)
TABLE_LEFT_IN = 0.583
TITLE_LEFT_IN = TABLE_LEFT_IN
TABLE_TOP_IN = 1.127
TABLE_WIDTH_IN = 12.483
SIDE_MARGIN_IN = 0.267
TITLE_TOP_IN = 0.361
LOGO_TOP_IN = 0.112
LOGO_LEFT_IN = 10.255
LOGO_WIDTH_IN = 2.65
MIN_DATA_ROW_HEIGHT_IN = 0.742

# Derived layout
MAX_TABLE_HEIGHT_IN = SLIDE_HEIGHT_IN - TABLE_TOP_IN - SIDE_MARGIN_IN
HEADER_ROW_HEIGHT_IN = 0.808
MAX_DATA_HEIGHT_IN = MAX_TABLE_HEIGHT_IN - HEADER_ROW_HEIGHT_IN

# Column widths, inches (section 5)
COLUMN_WIDTHS_IN = [0.526, 0.944, 4.051, 3.785, 1.306, 1.872]
TASK_COL_WIDTH_IN = COLUMN_WIDTHS_IN[2]
VALUE_COL_WIDTH_IN = COLUMN_WIDTHS_IN[3]
STATUS_COL_WIDTH_IN = COLUMN_WIDTHS_IN[5]

# Row-height estimation
LINE_HEIGHT_IN = TABLE_BODY_SIZE / 72 * 1.15
CHARS_PER_INCH = 11

# Table columns (section 6)
TABLE_COLUMNS = [
    "№пп",
    "Ключ",
    "Задача",
    "Ценность",
    "Предварительная\nоценка\nреализации,\nтыс.руб.",
    "Статус",
]

SLIDE_TITLE_TEMPLATE = "Дефекты и разработки по Пилоту ({page}/{total})"

# Yandex Tracker links for task keys
TRACKER_BASE_URL = "https://tracker.yandex.ru/"

# CSV filters (section 7)
CSV_INCLUDE_TEXT = "осипова"
CSV_EXCLUDE_TEXT = "важно"

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
