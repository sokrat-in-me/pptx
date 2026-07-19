#!/usr/bin/env python3
"""Generate ERP reporting critique presentation from RTR CSV data."""

from __future__ import annotations

import argparse
import csv
import re
from collections import defaultdict
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

SEGMENT_ORDER = [
    "Выращивание",
    "Птицепереработка",
    "Кормопроизводство",
    "Все сегменты",
    "Прочее",
]
ROWS_PER_SLIDE = 8


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


def set_cell(
    cell,
    text: str,
    *,
    bold: bool = False,
    size: int = 9,
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
            run.font.size = Pt(size)
            run.font.name = "Calibri"
            run.font.bold = bold
            if font_color is not None:
                run.font.color.rgb = font_color
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill


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

    title_color = RGBColor(31, 78, 121)
    header_bg = RGBColor(31, 78, 121)
    header_fg = RGBColor(255, 255, 255)
    alt_row = RGBColor(238, 245, 251)

    title_slide = prs.slides.add_slide(blank)
    title_box = title_slide.shapes.add_textbox(Inches(0.5), Inches(2.2), Inches(12.3), Inches(1.2))
    title_paragraph = title_box.text_frame.paragraphs[0]
    title_paragraph.text = "Критика отчетности ERP по сегментам"
    title_paragraph.alignment = PP_ALIGN.CENTER
    title_paragraph.font.size = Pt(36)
    title_paragraph.font.bold = True
    title_paragraph.font.name = "Calibri"
    title_paragraph.font.color.rgb = title_color

    subtitle_box = title_slide.shapes.add_textbox(Inches(0.5), Inches(3.5), Inches(12.3), Inches(0.8))
    subtitle = subtitle_box.text_frame.paragraphs[0]
    subtitle.text = "Статус пилота и план доработок (данные на 14.07)"
    subtitle.alignment = PP_ALIGN.CENTER
    subtitle.font.size = Pt(18)
    subtitle.font.name = "Calibri"
    subtitle.font.color.rgb = RGBColor(68, 68, 68)

    if release_info:
        release_box = title_slide.shapes.add_textbox(Inches(1.2), Inches(4.5), Inches(10.9), Inches(1.8))
        release_frame = release_box.text_frame
        for index, (release, deadline) in enumerate(release_info):
            paragraph = release_frame.paragraphs[0] if index == 0 else release_frame.add_paragraph()
            paragraph.text = (
                f"{release}: дедлайн согласования ЧТЗ — {deadline}" if deadline else release
            )
            paragraph.font.size = Pt(14)
            paragraph.font.name = "Calibri"
            paragraph.alignment = PP_ALIGN.CENTER

    columns = [
        "Сегмент",
        "Доработка\nотчета",
        "Новый\nотчет",
        "Описание",
        "Трудоемкость\n(часы)*",
        "Предложение",
    ]
    widths = [Inches(1.35), Inches(0.75), Inches(0.75), Inches(5.2), Inches(1.15), Inches(3.45)]

    slide_number = 1
    for chunk_start in range(0, len(flat_records), ROWS_PER_SLIDE):
        chunk = flat_records[chunk_start : chunk_start + ROWS_PER_SLIDE]
        slide = prs.slides.add_slide(blank)

        header = slide.shapes.add_textbox(Inches(0.4), Inches(0.15), Inches(10), Inches(0.45))
        header_paragraph = header.text_frame.paragraphs[0]
        header_paragraph.text = "Критика отчетности ERP по сегментам"
        header_paragraph.font.size = Pt(20)
        header_paragraph.font.bold = True
        header_paragraph.font.name = "Calibri"
        header_paragraph.font.color.rgb = title_color

        number_box = slide.shapes.add_textbox(Inches(12.3), Inches(0.15), Inches(0.7), Inches(0.45))
        number_paragraph = number_box.text_frame.paragraphs[0]
        number_paragraph.text = str(slide_number)
        number_paragraph.alignment = PP_ALIGN.RIGHT
        number_paragraph.font.size = Pt(18)
        number_paragraph.font.bold = True
        number_paragraph.font.name = "Calibri"

        table_shape = slide.shapes.add_table(
            len(chunk) + 1,
            len(columns),
            Inches(0.35),
            Inches(0.65),
            sum(widths),
            Inches(5.7),
        )
        table = table_shape.table
        for index, width in enumerate(widths):
            table.columns[index].width = width

        for column_index, title in enumerate(columns):
            set_cell(
                table.cell(0, column_index),
                title,
                bold=True,
                size=9,
                align=PP_ALIGN.CENTER,
                fill=header_bg,
                font_color=header_fg,
            )

        for row_index, record in enumerate(chunk, start=1):
            description = record["description"]
            if len(description) > 260:
                description = description[:257] + "..."
            values = [
                map_segment(record),
                "V" if is_modification(record["defect"]) else "",
                "V" if is_new_report(record["defect"]) else "",
                description,
                labor_text(record),
                proposal_text(record),
            ]
            fill = alt_row if row_index % 2 == 0 else None
            for column_index, value in enumerate(values):
                set_cell(
                    table.cell(row_index, column_index),
                    value,
                    size=8,
                    align=PP_ALIGN.CENTER if column_index in (1, 2, 4) else PP_ALIGN.LEFT,
                    fill=fill,
                )

        notes = slide.shapes.add_textbox(Inches(0.4), Inches(6.55), Inches(12.5), Inches(0.7))
        notes_paragraph = notes.text_frame.paragraphs[0]
        notes_paragraph.text = "* оценка по Пилоту    ** трудозатраты: аналитик + разработчик"
        notes_paragraph.font.size = Pt(10)
        notes_paragraph.font.italic = True
        notes_paragraph.font.name = "Calibri"
        notes_paragraph.font.color.rgb = RGBColor(102, 102, 102)

        slide_number += 1

    summary_slide = prs.slides.add_slide(blank)
    summary_header = summary_slide.shapes.add_textbox(Inches(0.4), Inches(0.2), Inches(12), Inches(0.5))
    summary_header.text_frame.paragraphs[0].text = "Сводка по сегментам"
    summary_header.text_frame.paragraphs[0].font.size = Pt(24)
    summary_header.text_frame.paragraphs[0].font.bold = True
    summary_header.text_frame.paragraphs[0].font.color.rgb = title_color

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
            bold=True,
            size=11,
            align=PP_ALIGN.CENTER,
            fill=header_bg,
            font_color=header_fg,
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
            set_cell(
                summary_table.cell(row_index, column_index),
                value,
                size=11,
                align=PP_ALIGN.CENTER,
                fill=alt_row if row_index % 2 == 0 else None,
            )

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
