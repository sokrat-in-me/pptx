#!/usr/bin/env python3
"""Generate ERP reporting critique presentation from RTR CSV data."""

from __future__ import annotations

import argparse
import csv
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

SEGMENT_ORDER = [
    "Выращивание",
    "Птицепереработка",
    "Кормопроизводство",
    "Все сегменты",
    "Прочее",
]
ROWS_PER_SLIDE = 8
ASSETS_DIR = Path(__file__).resolve().parent / "assets"
LOGO_PATH = ASSETS_DIR / "logo_0.png"


@dataclass(frozen=True)
class PdfTheme:
    """Colors and fonts extracted from the PDF template."""

    title_font = "Verdana"
    body_font = "Verdana"
    page_font = "Calibri"

    title_rgb = RGBColor(0x1E, 0x20, 0x24)
    text_rgb = RGBColor(0x00, 0x00, 0x00)
    accent_rgb = RGBColor(0xFF, 0x00, 0x00)
    page_number_rgb = RGBColor(0x9A, 0xA0, 0xA6)
    header_bg_rgb = RGBColor(0xAE, 0x17, 0x2D)
    header_fg_rgb = RGBColor(0xFF, 0xFF, 0xFF)
    segment_bg_rgb = RGBColor(0xF0, 0xF0, 0xF0)
    white_rgb = RGBColor(0xFF, 0xFF, 0xFF)

    title_size = 20
    header_size = 8
    body_size = 11
    segment_size = 8
    footnote_size = 10
    page_number_size = 9


THEME = PdfTheme()


def parse_num(value: str) -> float | None:
    cleaned = re.sub(r"[^0-9.,]", "", value or "").replace(",", ".")
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def is_modification(defect: str) -> bool:
    return defect in {
        "Ошибка функциональности",
        "HotFix-расширение",
        "Консультация",
        "Задача",
        "Дефект интеграции",
        "Доработка роли",
        "Замечание к роли",
        "Ошибка данных миграции",
    }


def is_new_report(defect: str) -> bool:
    return defect in {"Новое требование", "Разработка", "разработка", "Требование"}


def map_segment(row: dict[str, str]) -> str:
    team = row["team"].lower()
    desc = row["description"].lower()
    complexity = row["complexity"].lower()

    if "кормопроизвод" in team:
        return "Кормопроизводство"
    if "переработ" in team or any(
        keyword in desc
        for keyword in ("убой", "переработ", "гп по sku", "выпущенной продукции")
    ):
        return "Птицепереработка"
    if "птицевод" in team or "выращиван" in desc or "птичник" in desc:
        return "Выращивание"
    if "все сегмент" in desc:
        return "Все сегменты"
    if "корм" in desc:
        return "Кормопроизводство"
    if "индейка" in complexity or "курочка" in complexity:
        if any(keyword in desc for keyword in ("убой", "себестоим", "переработ", "гп")):
            return "Птицепереработка"
        return "Выращивание"
    if "свинка" in complexity:
        return "Птицепереработка"
    return "Прочее"


def labor_text(row: dict[str, str]) -> str:
    analyst = parse_num(row["analyst_hours"])
    developer = parse_num(row["developer_hours"])
    parts: list[str] = []
    if analyst is not None:
        parts.append(str(int(analyst)) if analyst == int(analyst) else str(analyst))
    if developer is not None:
        parts.append(str(int(developer)) if developer == int(developer) else str(developer))
    if parts:
        return f"{parts[0]}+{parts[1]} часов" if len(parts) == 2 else f"{parts[0]} часов"
    return row["cost"]


def proposal_text(row: dict[str, str]) -> str:
    comments = row["comments"].lower()
    if "отказ" in comments:
        return "Отказ"
    if "non-erp" in comments or "non erp" in comments:
        return "Передать в non-ERP"
    if "mtd" in comments and "корм" in comments:
        return "Передать в MTD-Кормопроизводство"

    parts: list[str] = []
    if "подрядчик" in comments:
        parts.append("Отдаем в разработку подрядчикам")
    if row["release"]:
        parts.append(f"Релиз {row['release']}")
    if row["blocker"].lower() == "да":
        parts.append("Блокирует тираж")
    if row["task_key"]:
        parts.append(row["task_key"])
    if row["team"].startswith("MTD"):
        parts.append(row["team"])
    if row["comments"] and len(row["comments"]) < 120:
        parts.append(row["comments"])
    return "; ".join(dict.fromkeys(parts)) or "В работе"


def proposal_is_accent(text: str) -> bool:
    lowered = text.lower()
    return lowered.startswith("уточнить") or "важно" in lowered


def load_csv(path: Path) -> tuple[list[tuple[str, str]], list[dict[str, str]]]:
    with path.open(encoding="utf-8-sig") as handle:
        rows = list(csv.reader(handle))

    header = rows[5]
    index = {name.strip(): position for position, name in enumerate(header)}

    def get(row: list[str], name: str, default: str = "") -> str:
        position = index.get(name)
        if position is None or position >= len(row):
            return default
        return (row[position] or "").strip()

    release_info = []
    for row in rows[1:5]:
        release = row[8].strip() if len(row) > 8 else ""
        deadline = row[9].strip() if len(row) > 9 else ""
        if release:
            release_info.append((release, deadline))

    records: list[dict[str, str]] = []
    for row in rows[6:]:
        team = get(row, "Команда")
        if not team:
            continue

        description = get(row, "Описание задачи (текстовое поле)")
        project_solution = get(row, "Проектное решение")
        if not (
            "отчет" in description.lower()
            or "отчёт" in description.lower()
            or "Отчетность" in project_solution
            or "отчет" in project_solution.lower()
        ):
            continue

        records.append(
            {
                "team": team,
                "defect": get(row, "Дефект"),
                "task_key": get(row, "Ключ задач в ЯТ (ссылка)"),
                "description": description,
                "complexity": get(row, "Сложность"),
                "release": get(row, "Релиз"),
                "blocker": get(row, "Блокирует старт тиража? (Да/Нет)"),
                "comments": get(row, "Корректировки/комментарии") or get(row, "Комментарии"),
                "analyst_hours": get(row, "Оценака аналитики, ч"),
                "developer_hours": get(row, "Оценка разработки, ч"),
                "cost": get(row, "Стоимость разработки, тыс.руб."),
            }
        )

    return release_info, records


def set_cell_border(cell, *, color: str = "000000", width: str = "12700") -> None:
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
    italic: bool = False,
    color: RGBColor | None = None,
) -> None:
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color


def set_textbox(
    text_frame,
    text: str,
    *,
    font_name: str,
    size: int,
    bold: bool = False,
    italic: bool = False,
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
        italic=italic,
        color=color or THEME.text_rgb,
    )


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
    cell.text = str(text)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Pt(4)
    cell.margin_right = Pt(4)
    cell.margin_top = Pt(2)
    cell.margin_bottom = Pt(2)
    for paragraph in cell.text_frame.paragraphs:
        paragraph.alignment = align
        for run in paragraph.runs:
            set_run_font(
                run,
                font_name=font_name,
                size=size,
                bold=bold,
                color=font_color or THEME.text_rgb,
            )
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill


def add_logo(slide) -> None:
    if not LOGO_PATH.exists():
        return
    slide.shapes.add_picture(
        str(LOGO_PATH),
        Inches(10.55),
        Inches(0.12),
        width=Inches(2.35),
    )


def add_slide_title(slide, title: str) -> None:
    title_box = slide.shapes.add_textbox(Inches(0.35), Inches(0.12), Inches(9.8), Inches(0.7))
    set_textbox(
        title_box.text_frame,
        title,
        font_name=THEME.title_font,
        size=THEME.title_size,
        bold=True,
        color=THEME.title_rgb,
    )


def add_page_number(slide, number: int) -> None:
    number_box = slide.shapes.add_textbox(Inches(12.45), Inches(7.0), Inches(0.55), Inches(0.3))
    set_textbox(
        number_box.text_frame,
        str(number),
        font_name=THEME.page_font,
        size=THEME.page_number_size,
        color=THEME.page_number_rgb,
        align=PP_ALIGN.RIGHT,
    )


def add_footnotes(slide) -> None:
    notes = slide.shapes.add_textbox(Inches(0.35), Inches(6.95), Inches(10.5), Inches(0.35))
    set_textbox(
        notes.text_frame,
        "*  оценка по Пилоту\n** трудозатраты: аналитик + разработчик",
        font_name=THEME.body_font,
        size=THEME.footnote_size,
        color=THEME.text_rgb,
    )


def build_presentation(
    release_info: list[tuple[str, str]],
    records: list[dict[str, str]],
    output_path: Path,
) -> None:
    by_segment: dict[str, list[dict[str, str]]] = defaultdict(list)
    for record in records:
        by_segment[map_segment(record)].append(record)

    ordered_segments = [segment for segment in SEGMENT_ORDER if segment in by_segment]
    for segment in sorted(by_segment):
        if segment not in ordered_segments:
            ordered_segments.append(segment)

    flat_records: list[dict[str, str]] = []
    for segment in ordered_segments:
        flat_records.extend(by_segment[segment])

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    title_slide = prs.slides.add_slide(blank)
    add_logo(title_slide)
    title_box = title_slide.shapes.add_textbox(Inches(0.35), Inches(2.0), Inches(8.5), Inches(1.0))
    set_textbox(
        title_box.text_frame,
        "Критика отчетности ERP\nпо сегментам",
        font_name=THEME.title_font,
        size=THEME.title_size,
        bold=True,
        color=THEME.title_rgb,
    )

    subtitle_box = title_slide.shapes.add_textbox(Inches(0.35), Inches(3.2), Inches(9.0), Inches(0.5))
    set_textbox(
        subtitle_box.text_frame,
        "Статус пилота и план доработок (данные на 14.07)",
        font_name=THEME.body_font,
        size=THEME.body_size,
        color=THEME.text_rgb,
    )

    if release_info:
        release_box = title_slide.shapes.add_textbox(Inches(0.35), Inches(4.0), Inches(8.5), Inches(1.8))
        release_frame = release_box.text_frame
        release_frame.clear()
        for index, (release, deadline) in enumerate(release_info):
            paragraph = release_frame.paragraphs[0] if index == 0 else release_frame.add_paragraph()
            paragraph.alignment = PP_ALIGN.LEFT
            run = paragraph.add_run()
            run.text = (
                f"{release}: дедлайн согласования ЧТЗ — {deadline}" if deadline else release
            )
            set_run_font(
                run,
                font_name=THEME.body_font,
                size=THEME.body_size,
                color=THEME.text_rgb,
            )

    columns = [
        "Сегмент",
        "Доработка\nотчета",
        "Новый\nотчет",
        "Описание",
        "Трудоемкость\n, ЧЧ/мес*",
        "Предложение",
    ]
    widths = [Inches(1.35), Inches(0.75), Inches(0.75), Inches(5.2), Inches(1.15), Inches(3.45)]

    slide_number = 1
    for chunk_start in range(0, len(flat_records), ROWS_PER_SLIDE):
        chunk = flat_records[chunk_start : chunk_start + ROWS_PER_SLIDE]
        slide = prs.slides.add_slide(blank)
        add_logo(slide)
        add_slide_title(slide, "Критика отчетности ERP по сегментам")

        table_shape = slide.shapes.add_table(
            len(chunk) + 1,
            len(columns),
            Inches(0.35),
            Inches(0.85),
            sum(widths),
            Inches(5.85),
        )
        table = table_shape.table
        for index, width in enumerate(widths):
            table.columns[index].width = width

        for column_index, title in enumerate(columns):
            set_cell(
                table.cell(0, column_index),
                title,
                font_name=THEME.title_font,
                bold=True,
                size=THEME.header_size,
                align=PP_ALIGN.CENTER,
                fill=THEME.header_bg_rgb,
                font_color=THEME.header_fg_rgb,
            )

        for row_index, record in enumerate(chunk, start=1):
            description = record["description"]
            if len(description) > 260:
                description = description[:257] + "..."
            proposal = proposal_text(record)
            values = [
                (map_segment(record), 0, True, False, THEME.segment_bg_rgb, THEME.text_rgb),
                ("V" if is_modification(record["defect"]) else "", 1, True, False, THEME.white_rgb, THEME.text_rgb),
                ("V" if is_new_report(record["defect"]) else "", 2, True, False, THEME.white_rgb, THEME.text_rgb),
                (description, 3, False, False, THEME.white_rgb, THEME.text_rgb),
                (labor_text(record), 4, True, False, THEME.white_rgb, THEME.text_rgb),
                (
                    proposal,
                    5,
                    False,
                    proposal_is_accent(proposal),
                    THEME.white_rgb,
                    THEME.accent_rgb if proposal_is_accent(proposal) else THEME.text_rgb,
                ),
            ]
            for value, column_index, bold, _, fill, color in values:
                set_cell(
                    table.cell(row_index, column_index),
                    value,
                    font_name=THEME.title_font if column_index == 0 else THEME.body_font,
                    size=THEME.segment_size if column_index == 0 else THEME.body_size,
                    bold=bold,
                    align=PP_ALIGN.CENTER if column_index in (1, 2, 4) else PP_ALIGN.LEFT,
                    fill=fill,
                    font_color=color,
                )

        style_table_borders(table)
        add_footnotes(slide)
        add_page_number(slide, slide_number)
        slide_number += 1

    summary_slide = prs.slides.add_slide(blank)
    add_logo(summary_slide)
    add_slide_title(summary_slide, "Сводка по сегментам")

    summary_columns = [
        "Сегмент",
        "Кол-во задач",
        "Доработка",
        "Новый отчет",
        "Часы (аналитика+разработка)",
    ]
    summary_widths = [Inches(2.5), Inches(1.5), Inches(1.5), Inches(1.5), Inches(3.5)]
    summary_table = summary_slide.shapes.add_table(
        len(ordered_segments) + 1,
        len(summary_columns),
        Inches(1.5),
        Inches(1.2),
        sum(summary_widths),
        Inches(3.5),
    ).table
    for index, width in enumerate(summary_widths):
        summary_table.columns[index].width = width

    for column_index, title in enumerate(summary_columns):
        set_cell(
            summary_table.cell(0, column_index),
            title,
            font_name=THEME.title_font,
            bold=True,
            size=THEME.header_size,
            align=PP_ALIGN.CENTER,
            fill=THEME.header_bg_rgb,
            font_color=THEME.header_fg_rgb,
        )

    for row_index, segment in enumerate(ordered_segments, start=1):
        items = by_segment[segment]
        modifications = sum(1 for item in items if is_modification(item["defect"]))
        new_reports = sum(1 for item in items if is_new_report(item["defect"]))
        hours = sum(
            (parse_num(item["analyst_hours"]) or 0) + (parse_num(item["developer_hours"]) or 0)
            for item in items
        )
        values = [segment, len(items), modifications, new_reports, int(hours)]
        for column_index, value in enumerate(values):
            fill = THEME.segment_bg_rgb if column_index == 0 else THEME.white_rgb
            set_cell(
                summary_table.cell(row_index, column_index),
                value,
                font_name=THEME.title_font if column_index == 0 else THEME.body_font,
                size=THEME.segment_size if column_index == 0 else THEME.body_size,
                bold=column_index == 0,
                align=PP_ALIGN.CENTER,
                fill=fill,
            )

    style_table_borders(summary_table)
    add_page_number(summary_slide, slide_number)
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
        default=Path("Критика_отчетности_ERP_по_сегментам.pptx"),
        help="Path to generated presentation",
    )
    args = parser.parse_args()

    release_info, records = load_csv(args.csv)
    build_presentation(release_info, records, args.output)
    print(f"Created {args.output} with {len(records)} reporting tasks")


if __name__ == "__main__":
    main()
