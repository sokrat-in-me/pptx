#!/usr/bin/env python3
"""Generate RTR launch schedule presentation from CSV.

Style and layout: ai/rules/cherkizovo-presentations.md
Theme: themes/cherkizovo.py
"""

from __future__ import annotations

import argparse
import csv
import math
import sys
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

ROOT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT_DIR))

from themes.cherkizovo import (  # noqa: E402
    COLORS,
    LOGO_LEFT_IN,
    LOGO_TOP_IN,
    LOGO_WIDTH_IN,
    SIDE_MARGIN_IN,
    SLIDE_HEIGHT_IN,
    SLIDE_WIDTH_IN,
    TABLE_LEFT_IN,
    TABLE_WIDTH_IN,
    TITLE_FONT,
    TITLE_LEFT_IN,
    TITLE_SIZE,
    TITLE_TOP_IN,
    apply_michurin_theme,
)

ASSETS_DIR = ROOT_DIR / "assets"
LOGO_PATH = ASSETS_DIR / "logo_0.png"

MONTH_LABELS = {
    "Jan": "01",
    "Feb": "02",
    "Mar": "03",
    "Apr": "04",
    "May": "05",
    "Jun": "06",
    "Jul": "07",
    "Aug": "08",
    "Sep": "09",
    "Oct": "10",
    "Nov": "11",
    "Dec": "12",
}

LABEL_COLUMNS = [
    ("Задача / направление", 3.55),
    ("Сист.", 0.4),
]

EXCLUDED_TIMELINE = {("2026", "May"), ("2026", "Jun")}

MAX_GANTT_SLIDES = 2
TABLE_TOP_IN = 1.05
YEAR_HEADER_HEIGHT_IN = 0.28
MONTH_HEADER_HEIGHT_IN = 0.28
MIN_ROW_HEIGHT_IN = 0.2

# Светлая палитра для диаграммы Ганта (на базе корпоративных оттенков МИЧУРИН)
LIGHT_COLORS = {
    "title_slide_bg": (253, 245, 246),   # #FDF5F6
    "title": (131, 18, 35),            # #831223
    "subtitle": (119, 119, 122),        # #77777A
    "header": (245, 180, 188),         # #F5B4BC — заголовки колонок
    "year_header": (252, 220, 224),    # #FCDCE0 — строка годов
    "subsection": (250, 228, 231),     # #FAE4E7 — подзаголовки секций
    "mark": (255, 210, 216),           # #FFD2D8 — ячейки графика
    "mark_text": (150, 40, 55),        # #962837
    "header_text": (100, 30, 40),      # #641E28
    "alt_row": (252, 248, 249),        # #FCF8F9
    "text": (50, 50, 50),              # #323232
}


def rgb(name: str) -> RGBColor:
    if name in LIGHT_COLORS:
        red, green, blue = LIGHT_COLORS[name]
    else:
        red, green, blue = COLORS[name]
    return RGBColor(red, green, blue)


@dataclass
class TimelineColumn:
    index: int
    year: str
    month: str

    @property
    def key(self) -> str:
        return f"{self.year}-{self.month}"

    @property
    def label(self) -> str:
        month_num = MONTH_LABELS.get(self.month, self.month[:3])
        year_short = self.year[-2:] if self.year else ""
        return f"{month_num}.{year_short}"


@dataclass
class ScheduleRow:
    kind: str  # section | subsection | task
    contour: str = ""
    block: str = ""
    task: str = ""
    system: str = ""
    marks: dict[str, str] = field(default_factory=dict)
    owner: str = ""


@dataclass
class YearSpan:
    year: str
    columns: list[TimelineColumn] = field(default_factory=list)


@dataclass
class SlideGroup:
    title: str
    rows: list[ScheduleRow]


def load_schedule(path: Path) -> tuple[list[TimelineColumn], list[ScheduleRow]]:
    with path.open(encoding="utf-8-sig") as handle:
        rows = list(csv.reader(handle))

    years = rows[1]
    months = rows[2]
    timeline: list[TimelineColumn] = []
    for index in range(8, len(months)):
        year = years[index].strip() if index < len(years) else ""
        month = months[index].strip() if index < len(months) else ""
        if not year and not month:
            continue
        if (year, month) in EXCLUDED_TIMELINE:
            continue
        timeline.append(TimelineColumn(index=index, year=year, month=month))

    records: list[ScheduleRow] = []
    for row in rows[3:]:
        contour = row[0].strip() if len(row) > 0 else ""
        block = row[1].strip() if len(row) > 1 else ""
        task = row[2].strip() if len(row) > 2 else ""
        system = row[6].strip() if len(row) > 6 else ""

        if not any((contour, block, task, system)):
            continue

        marks = {
            column.key: row[column.index].strip()
            for column in timeline
            if column.index < len(row) and row[column.index].strip()
        }

        if task and not contour and not block and not marks and not system:
            if task.isupper() or task.startswith("MTD") or task.startswith("RTR"):
                records.append(ScheduleRow(kind="section", task=task))
                continue
            if task in {"Себестоимость", "Бюджетный контроль", "Онлайн себестоимость"}:
                records.append(ScheduleRow(kind="subsection", task=task))
                continue
            if "Казначейство" in task or "Трансфертное" in task:
                records.append(ScheduleRow(kind="subsection", task=task))
                continue

        if contour and not task:
            records.append(ScheduleRow(kind="task", contour=contour, owner=contour))
            continue

        if not task:
            continue

        records.append(
            ScheduleRow(
                kind="task",
                contour=contour,
                block=block,
                task=task,
                system=system,
                marks=marks,
                owner=contour,
            )
        )

    return timeline, records


def build_year_spans(timeline: list[TimelineColumn]) -> list[YearSpan]:
    spans: list[YearSpan] = []
    for column in timeline:
        if spans and spans[-1].year == column.year:
            spans[-1].columns.append(column)
        else:
            spans.append(YearSpan(year=column.year, columns=[column]))
    return spans


def task_display_text(record: ScheduleRow) -> str:
    if record.task:
        if record.contour and not record.task.startswith(record.contour):
            return f"{record.contour}: {record.task}"
        return record.task
    return record.contour


def flatten_display_rows(records: list[ScheduleRow]) -> list[ScheduleRow]:
    """Convert section headers into inline subsection rows for a compact table."""
    display_rows: list[ScheduleRow] = []
    for record in records:
        if record.kind == "section":
            display_rows.append(ScheduleRow(kind="subsection", task=record.task))
        else:
            display_rows.append(record)
    return display_rows


def group_slides(records: list[ScheduleRow]) -> list[SlideGroup]:
    display_rows = flatten_display_rows(records)
    if not display_rows:
        return []

    chunk_size = math.ceil(len(display_rows) / MAX_GANTT_SLIDES)
    groups: list[SlideGroup] = []
    for index in range(0, len(display_rows), chunk_size):
        chunk = display_rows[index : index + chunk_size]
        page = len(groups) + 1
        groups.append(SlideGroup(title=f"График запуска ({page}/{MAX_GANTT_SLIDES})", rows=chunk))
    return groups[:MAX_GANTT_SLIDES]


def set_cell_border(cell, *, color: str = "E8C4C8", width: str = "6350") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    for edge in ("lnL", "lnR", "lnT", "lnB"):
        tag = qn(f"a:{edge}")
        element = tc_pr.find(tag)
        if element is None:
            element = OxmlElement(f"a:{edge}")
            tc_pr.append(element)
        element.set("w", width)
        element.set("cap", "flat")
        element.set("cmpd", "sng")
        element.set("algn", "ctr")
        fill = element.find(qn("a:solidFill"))
        if fill is None:
            fill = OxmlElement("a:solidFill")
            element.append(fill)
        srgb = fill.find(qn("a:srgbClr"))
        if srgb is None:
            srgb = OxmlElement("a:srgbClr")
            fill.append(srgb)
        srgb.set("val", color)


def style_table_borders(table) -> None:
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell)


def set_run_font(
    run,
    *,
    font_name: str,
    size: int,
    bold: bool = False,
    color: RGBColor | None = None,
) -> None:
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def set_cell(
    cell,
    text: str,
    *,
    font_name: str = "Calibri",
    bold: bool = False,
    size: int = 9,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    fill: RGBColor | None = None,
    font_color: RGBColor | None = None,
    valign=MSO_ANCHOR.MIDDLE,
) -> None:
    cell.text = ""
    cell.vertical_anchor = valign
    cell.margin_left = Pt(2)
    cell.margin_right = Pt(2)
    cell.margin_top = Pt(1)
    cell.margin_bottom = Pt(1)
    cell.text_frame.word_wrap = True
    paragraph = cell.text_frame.paragraphs[0]
    paragraph.alignment = align
    paragraph.space_before = Pt(0)
    paragraph.space_after = Pt(0)
    paragraph.line_spacing = 1.0
    run = paragraph.add_run()
    run.text = text
    set_run_font(
        run,
        font_name=font_name,
        size=size,
        bold=bold,
        color=font_color or rgb("text"),
    )
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill


def add_logo(slide) -> None:
    if LOGO_PATH.exists():
        slide.shapes.add_picture(
            str(LOGO_PATH),
            Inches(LOGO_LEFT_IN),
            Inches(LOGO_TOP_IN),
            width=Inches(LOGO_WIDTH_IN),
        )


def set_textbox(
    text_frame,
    text: str,
    *,
    font_name: str,
    size: int,
    bold: bool = False,
    color: RGBColor | None = None,
    align: PP_ALIGN = PP_ALIGN.LEFT,
) -> None:
    text_frame.clear()
    paragraph = text_frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    set_run_font(
        run,
        font_name=font_name,
        size=size,
        bold=bold,
        color=color or rgb("text"),
    )


def estimate_row_height(record: ScheduleRow) -> float:
    if record.kind == "subsection":
        return 0.24
    text_len = max(len(record.task), max((len(value) for value in record.marks.values()), default=0))
    if text_len > 80:
        return 0.34
    if text_len > 45:
        return 0.28
    return MIN_ROW_HEIGHT_IN


def merge_header_cells(table, row_index: int, start_col: int, end_col: int) -> None:
    if end_col <= start_col:
        return
    table.cell(row_index, start_col).merge(table.cell(row_index, end_col))


def build_gantt_slide(
    prs: Presentation,
    *,
    group: SlideGroup,
    timeline: list[TimelineColumn],
    year_spans: list[YearSpan],
    page_index: int,
    total_pages: int,
) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_logo(slide)

    title_box = slide.shapes.add_textbox(
        Inches(TITLE_LEFT_IN),
        Inches(TITLE_TOP_IN),
        Inches(9.5),
        Inches(0.45),
    )
    set_textbox(
        title_box.text_frame,
        group.title,
        font_name=TITLE_FONT,
        size=TITLE_SIZE,
        bold=True,
        color=rgb("title"),
    )

    label_widths = [width for _, width in LABEL_COLUMNS]
    month_width = (TABLE_WIDTH_IN - sum(label_widths)) / max(1, len(timeline))
    column_widths = [Inches(width) for width in label_widths] + [
        Inches(month_width) for _ in timeline
    ]

    row_heights = [estimate_row_height(record) for record in group.rows]
    header_height = YEAR_HEADER_HEIGHT_IN + MONTH_HEADER_HEIGHT_IN
    table_height = Inches(header_height + sum(row_heights))

    table_shape = slide.shapes.add_table(
        len(group.rows) + 2,
        len(LABEL_COLUMNS) + len(timeline),
        Inches(TABLE_LEFT_IN),
        Inches(TABLE_TOP_IN),
        Inches(TABLE_WIDTH_IN),
        table_height,
    )
    table = table_shape.table

    for index, width in enumerate(column_widths):
        table.columns[index].width = width

    table.rows[0].height = Inches(YEAR_HEADER_HEIGHT_IN)
    table.rows[1].height = Inches(MONTH_HEADER_HEIGHT_IN)
    for row_index, row_height in enumerate(row_heights, start=2):
        table.rows[row_index].height = Inches(row_height)

    label_col_count = len(LABEL_COLUMNS)
    timeline_col_count = len(timeline)

    for column_index, (title, _) in enumerate(LABEL_COLUMNS):
        set_cell(
            table.cell(0, column_index),
            title,
            font_name="Calibri",
            bold=True,
            size=7,
            align=PP_ALIGN.CENTER,
            fill=rgb("dark_burgundy"),
            font_color=rgb("white"),
        )
        table.cell(0, column_index).merge(table.cell(1, column_index))

    timeline_start = label_col_count
    for span in year_spans:
        start_col = timeline_start + sum(
            len(previous.columns) for previous in year_spans[: year_spans.index(span)]
        )
        end_col = start_col + len(span.columns) - 1
        set_cell(
            table.cell(0, start_col),
            span.year,
            font_name="Calibri",
            bold=True,
            size=7,
            align=PP_ALIGN.CENTER,
            fill=rgb("dark_burgundy"),
            font_color=rgb("white"),
        )
        merge_header_cells(table, 0, start_col, end_col)
        for offset, column in enumerate(span.columns):
            set_cell(
                table.cell(1, start_col + offset),
                column.label,
                font_name="Calibri",
                bold=True,
                size=6,
                align=PP_ALIGN.CENTER,
                fill=rgb("dark_burgundy"),
                font_color=rgb("white"),
            )

    for row_index, record in enumerate(group.rows, start=2):
        if record.kind == "subsection":
            for column_index in range(label_col_count + timeline_col_count):
                set_cell(
                    table.cell(row_index, column_index),
                    record.task if column_index == 0 else "",
                    font_name=TITLE_FONT,
                    bold=True,
                    size=8,
                    fill=rgb("subsection"),
                    font_color=rgb("title"),
                )
            continue

        fill = rgb("alt_row") if (row_index - 1) % 2 == 0 else rgb("white")
        set_cell(
            table.cell(row_index, 0),
            task_display_text(record),
            size=6,
            fill=fill,
            font_color=rgb("text"),
        )
        set_cell(
            table.cell(row_index, 1),
            record.system,
            size=6,
            fill=fill,
            align=PP_ALIGN.CENTER,
            font_color=rgb("text"),
        )

        for column_index, column in enumerate(timeline, start=label_col_count):
            mark = record.marks.get(column.key, "")
            if mark:
                set_cell(
                    table.cell(row_index, column_index),
                    mark,
                    size=5,
                    align=PP_ALIGN.CENTER,
                    fill=rgb("mark"),
                    font_color=rgb("mark_text"),
                )
            else:
                set_cell(
                    table.cell(row_index, column_index),
                    "",
                    size=5,
                    fill=fill,
                )

    style_table_borders(table)


def build_presentation(
    timeline: list[TimelineColumn],
    records: list[ScheduleRow],
    output_path: Path,
) -> int:
    groups = group_slides(records)
    year_spans = build_year_spans(timeline)
    total_pages = len(groups)

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH_IN)
    prs.slide_height = Inches(SLIDE_HEIGHT_IN)
    apply_michurin_theme(prs)

    for page_index, group in enumerate(groups, start=1):
        build_gantt_slide(
            prs,
            group=group,
            timeline=timeline,
            year_spans=year_spans,
            page_index=page_index,
            total_pages=total_pages,
        )

    prs.save(output_path)
    return len(groups)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--csv",
        type=Path,
        default=Path("data/RTR-launch-schedule.csv"),
        help="Path to launch schedule CSV",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("График_запуска_RTR.pptx"),
        help="Path to generated presentation",
    )
    args = parser.parse_args()

    timeline, records = load_schedule(args.csv)
    slide_count = build_presentation(timeline, records, args.output)
    print(
        f"Created {args.output} with {len(records)} rows across "
        f"{slide_count} gantt slides"
    )


if __name__ == "__main__":
    main()
