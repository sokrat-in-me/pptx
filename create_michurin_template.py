#!/usr/bin/env python3
"""Create assets/MICHURIN.pptx — PowerPoint template with the MICHURIN theme."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

ROOT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT_DIR))

from themes.cherkizovo import (  # noqa: E402
    COLORS,
    LOGO_LEFT_IN,
    LOGO_TOP_IN,
    LOGO_WIDTH_IN,
    SLIDE_HEIGHT_IN,
    SLIDE_WIDTH_IN,
    TEMPLATE_PATH,
    THEME_NAME,
    TITLE_FONT,
    TITLE_SIZE,
)
from themes.michurin import apply_michurin_theme  # noqa: E402

ASSETS_DIR = Path(__file__).resolve().parent / "assets"
LOGO_PATH = ASSETS_DIR / "logo_0.png"


def rgb(name: str) -> RGBColor:
    red, green, blue = COLORS[name]
    return RGBColor(red, green, blue)


def add_title_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[0])

    if LOGO_PATH.exists():
        slide.shapes.add_picture(
            str(LOGO_PATH),
            Inches(LOGO_LEFT_IN),
            Inches(LOGO_TOP_IN),
            width=Inches(LOGO_WIDTH_IN),
        )

    title_shape = slide.shapes.title
    title_shape.text = THEME_NAME
    title_frame = title_shape.text_frame
    title_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
    title_font = title_frame.paragraphs[0].font
    title_font.name = TITLE_FONT
    title_font.size = Pt(44)
    title_font.bold = True
    title_font.color.rgb = rgb("title")

    if len(slide.placeholders) > 1:
        subtitle = slide.placeholders[1]
        subtitle.text = "Корпоративный шаблон Черкизово"
        subtitle_font = subtitle.text_frame.paragraphs[0].font
        subtitle_font.name = TITLE_FONT
        subtitle_font.size = Pt(20)
        subtitle_font.color.rgb = rgb("dark_gray")


def add_palette_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    if LOGO_PATH.exists():
        slide.shapes.add_picture(
            str(LOGO_PATH),
            Inches(LOGO_LEFT_IN),
            Inches(LOGO_TOP_IN),
            width=Inches(LOGO_WIDTH_IN),
        )

    title_box = slide.shapes.add_textbox(Inches(0.583), Inches(0.361), Inches(8.5), Inches(0.6))
    title_run = title_box.text_frame.paragraphs[0].add_run()
    title_run.text = "Цветовая палитра"
    title_run.font.name = TITLE_FONT
    title_run.font.size = Pt(TITLE_SIZE)
    title_run.font.bold = True
    title_run.font.color.rgb = rgb("title")

    swatches = [
        ("Красный бренд", "brand_red"),
        ("Тёмно-красный", "dark_red"),
        ("Заголовок", "title"),
        ("Зелёный акцент", "green_accent"),
        ("Ключ / ссылка", "key"),
        ("Чередование строк", "alt_row"),
    ]
    left = Inches(0.583)
    top = Inches(1.2)
    width = Inches(2.0)
    height = Inches(0.55)
    gap = Inches(0.15)

    for index, (label, color_name) in enumerate(swatches):
        row = index // 3
        col = index % 3
        shape = slide.shapes.add_shape(
            1,
            left + col * (width + gap),
            top + row * (height + gap),
            width,
            height,
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(color_name)
        shape.line.color.rgb = rgb("dark_gray")
        label_box = slide.shapes.add_textbox(
            left + col * (width + gap),
            top + row * (height + gap) + height,
            width,
            Inches(0.3),
        )
        label_run = label_box.text_frame.paragraphs[0].add_run()
        label_run.text = label
        label_run.font.name = TITLE_FONT
        label_run.font.size = Pt(10)
        label_run.font.color.rgb = rgb("dark_gray")


def create_template(output_path: Path) -> None:
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH_IN)
    prs.slide_height = Inches(SLIDE_HEIGHT_IN)
    apply_michurin_theme(prs)
    add_title_slide(prs)
    add_palette_slide(prs)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output",
        type=Path,
        default=TEMPLATE_PATH,
        help="Path to generated MICHURIN template",
    )
    args = parser.parse_args()
    create_template(args.output)
    print(f"Created {args.output}")


if __name__ == "__main__":
    main()
