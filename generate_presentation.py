#!/usr/bin/env python3
"""Generate pilot defects presentation from RTR CSV data using the PDF template."""

from __future__ import annotations

import argparse
import csv
import math
import re
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

ROWS_PER_SLIDE = 7
ASSETS_DIR = Path(__file__).resolve().parent / "assets"
LOGO_PATH = ASSETS_DIR / "logo_0.png"


@dataclass(frozen=True)
class PdfTheme:
    """Colors and fonts extracted from the updated PDF template."""

    title_font = "Verdana"
    body_font = "Calibri"

    title_rgb = RGBColor(0x83, 0x12, 0x23)
    text_rgb = RGBColor(0x00, 0x00, 0x00)
    key_rgb = RGBColor(0x77, 0x77, 0x7A)
    accent_rgb = RGBColor(0xFF, 0x00, 0x00)
    header_bg_rgb = RGBColor(0xAF, 0x18, 0x2E)
    header_fg_rgb = RGBColor(0xFF, 0xFF, 0xFF)
    alt_row_rgb = RGBColor(0xF8, 0xE7, 0xE8)
    white_rgb = RGBColor(0xFF, 0xFF, 0xFF)

    title_size = 18
    header_size = 12
    body_size = 11


THEME = PdfTheme()
COLUMNS = [
    "№пп",
    "Ключ",
    "Задача",
    "Ценность",
    "Предварительная\nоценка\nреализации,\nтыс.руб.",
    "Статус",
]
COLUMN_WIDTHS = [
    Inches(0.45),
    Inches(0.85),
    Inches(3.75),
    Inches(3.55),
    Inches(1.2),
    Inches(1.65),
]


def load_csv(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig") as handle:
        rows = list(csv.reader(handle))

    header = rows[5]
    index = {name.strip(): position for position, name in enumerate(header)}

    def get(row: list[str], name: str, default: str = "") -> str:
        position = index.get(name)
        if position is None or position >= len(row):
            return default
        return (row[position] or "").strip()

    records: list[dict[str, str]] = []
    for row in rows[6:]:
        team = get(row, "Команда")
        if not team:
            continue
        records.append(
            {
                "team": team,
                "defect": get(row, "Дефект"),
                "task_key": get(row, "Ключ задач в ЯТ (ссылка)"),
                "description": get(row, "Описание задачи (текстовое поле)"),
                "priority": get(row, "Приоритет"),
                "release": get(row, "Релиз"),
                "chtz_status": get(row, "Статус ЧТЗ"),
                "development": get(row, "Разработка"),
                "comments": get(row, "Корректировки/комментарии") or get(row, "Комментарии"),
                "cost": get(row, "Стоимость разработки, тыс.руб."),
                "value": get(row, "value"),
                "ppk_critical": get(row, "критичны к запуску ППК"),
            }
        )
    return records


def is_critical(record: dict[str, str]) -> bool:
    priority = record["priority"].lower()
    ppk = record["ppk_critical"].lower()
    comments = record["comments"].lower()
    return (
        "критич" in priority
        or "критич" in ppk
        or "критич" in comments
        or "важно" in comments
    )


def value_text(record: dict[str, str]) -> str:
    if record["value"]:
        return record["value"]
    comments = record["comments"]
    if comments and len(comments) > 35 and not comments.lower().startswith("ок"):
        return comments
    return ""


def status_text(record: dict[str, str]) -> str:
    development = record["development"]
    release = record["release"]
    chtz_status = record["chtz_status"]
    comments = record["comments"]
    analyst = record.get("analyst", "")

    lowered = development.lower()
    if "беринг" in lowered:
        if any(word in lowered for word in ("разработ", "код", "dev")):
            return "в разработке у БерингПро"
        return "на оценке у БерингПро"
    if "базис" in lowered:
        return "в разработке у Базис"
    if "1с-перспектива" in lowered or "перспектив" in lowered:
        return "на оценке у 1С-Перспектива"
    if development:
        return development

    parts: list[str] = []
    if chtz_status:
        parts.append(chtz_status)
    if release:
        parts.append(f"{release} релиз")
    if comments and len(comments) <= 50 and "критич" not in comments.lower():
        parts.append(comments)
    if parts:
        return ", ".join(parts)
    return "в работе"


def task_parts(record: dict[str, str]) -> list[tuple[str, bool]]:
    defect = record["defect"] or "Задача"
    key = record["task_key"] or "—"
    description = record["description"]
    text = f"{key}: {defect} - {description}"
    parts: list[tuple[str, bool]] = [(text, False)]

    lowered = f"{description} {record['comments']}".lower()
    if "упп" in lowered or "реализован" in lowered:
        parts.append((" – реализовано в УПП", True))
    return parts


def cost_text(record: dict[str, str]) -> str:
    cost = record["cost"]
    if cost in {"", "0", "#N/A"}:
        return ""
    return cost


def set_cell_border(cell, *, color: str = "000000", width: str = "6350") -> None:
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
        color=color or THEME.text_rgb,
    )


def set_cell_runs(
    cell,
    runs: list[tuple[str, bool, RGBColor | None]],
    *,
    size: int = THEME.body_size,
    align=PP_ALIGN.LEFT,
    fill: RGBColor | None = None,
) -> None:
    cell.text = ""
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Pt(4)
    cell.margin_right = Pt(4)
    cell.margin_top = Pt(2)
    cell.margin_bottom = Pt(2)
    paragraph = cell.text_frame.paragraphs[0]
    paragraph.alignment = align
    for text, bold, color in runs:
        if not text:
            continue
        run = paragraph.add_run()
        run.text = text
        set_run_font(
            run,
            font_name=THEME.body_font,
            size=size,
            bold=bold,
            color=color or THEME.text_rgb,
        )
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill


def set_cell(
    cell,
    text: str,
    *,
    font_name: str = THEME.body_font,
    bold: bool = False,
    size: int = THEME.body_size,
    align=PP_ALIGN.LEFT,
    fill: RGBColor | None = None,
    font_color: RGBColor | None = None,
) -> None:
    set_cell_runs(
        cell,
        [(str(text), bold, font_color)],
        size=size,
        align=align,
        fill=fill,
    )
    for run in cell.text_frame.paragraphs[0].runs:
        run.font.name = font_name


def add_logo(slide) -> None:
    if not LOGO_PATH.exists():
        return
    slide.shapes.add_picture(str(LOGO_PATH), Inches(10.55), Inches(0.08), width=Inches(2.35))


def add_slide_title(slide, page_index: int, total_pages: int) -> None:
    title_box = slide.shapes.add_textbox(Inches(0.35), Inches(0.12), Inches(9.5), Inches(0.55))
    set_textbox(
        title_box.text_frame,
        f"Дефекты и разработки по Пилоту ({page_index}/{total_pages})",
        font_name=THEME.title_font,
        size=THEME.title_size,
        bold=True,
        color=THEME.title_rgb,
    )


def build_presentation(records: list[dict[str, str]], output_path: Path) -> None:
    total_pages = max(1, math.ceil(len(records) / ROWS_PER_SLIDE))

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    for page_index, chunk_start in enumerate(range(0, len(records), ROWS_PER_SLIDE), start=1):
        chunk = records[chunk_start : chunk_start + ROWS_PER_SLIDE]
        slide = prs.slides.add_slide(blank)
        add_logo(slide)
        add_slide_title(slide, page_index, total_pages)

        table_shape = slide.shapes.add_table(
            len(chunk) + 1,
            len(COLUMNS),
            Inches(0.35),
            Inches(0.75),
            sum(COLUMN_WIDTHS),
            Inches(6.35),
        )
        table = table_shape.table
        for index, width in enumerate(COLUMN_WIDTHS):
            table.columns[index].width = width

        for column_index, title in enumerate(COLUMNS):
            set_cell(
                table.cell(0, column_index),
                title,
                font_name=THEME.body_font,
                bold=True,
                size=THEME.header_size,
                align=PP_ALIGN.CENTER,
                fill=THEME.header_bg_rgb,
                font_color=THEME.header_fg_rgb,
            )

        for row_index, record in enumerate(chunk, start=1):
            row_number = chunk_start + row_index
            fill = THEME.alt_row_rgb if row_index % 2 == 0 else THEME.white_rgb
            critical = is_critical(record)
            accent = THEME.accent_rgb if critical else THEME.text_rgb

            set_cell(
                table.cell(row_index, 0),
                str(row_number),
                size=THEME.body_size,
                align=PP_ALIGN.CENTER,
                fill=fill,
            )
            set_cell(
                table.cell(row_index, 1),
                record["task_key"] or "—",
                size=THEME.body_size,
                fill=fill,
                font_color=THEME.key_rgb,
            )

            task_runs = [(text, bold, THEME.text_rgb) for text, bold in task_parts(record)]
            set_cell_runs(table.cell(row_index, 2), task_runs, fill=fill)

            value = value_text(record)
            if len(value) > 320:
                value = value[:317] + "..."
            set_cell(table.cell(row_index, 3), value, fill=fill)

            set_cell(
                table.cell(row_index, 4),
                cost_text(record),
                bold=critical and bool(cost_text(record)),
                align=PP_ALIGN.CENTER,
                fill=fill,
                font_color=accent if critical and cost_text(record) else THEME.text_rgb,
            )

            status = status_text(record)
            status_runs: list[tuple[str, bool, RGBColor | None]] = [(status, False, THEME.text_rgb)]
            if critical:
                status_runs.append(("\nКритичная!", True, THEME.accent_rgb))
            set_cell_runs(
                table.cell(row_index, 5),
                status_runs,
                align=PP_ALIGN.LEFT,
                fill=fill,
            )

        style_table_borders(table)

    prs.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--csv",
        type=Path,
        default=Path("data/RTR-status.csv"),
        help="Path to source CSV file",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("Дефекты_и_разработки_по_Пилоту.pptx"),
        help="Path to generated presentation",
    )
    args = parser.parse_args()

    records = load_csv(args.csv)
    build_presentation(records, args.output)
    pages = max(1, math.ceil(len(records) / ROWS_PER_SLIDE))
    print(f"Created {args.output} with {len(records)} tasks across {pages} slides")


if __name__ == "__main__":
    main()
