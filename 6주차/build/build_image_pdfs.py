#!/usr/bin/env python3
"""Build presentation and A4 PDFs from the verified week 6 slide renders."""

from pathlib import Path
import re

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas


WEEK_DIR = Path(__file__).resolve().parents[1]
SLIDE_DIR = WEEK_DIR / "rendered-v2"
OUTPUT_DIR = WEEK_DIR / "output"
PRESENTATION_PDF = OUTPUT_DIR / "전자기학_6주차_정전기에너지_도체_라플라스방정식.pdf"
A4_PDF = OUTPUT_DIR / "전자기학_6주차_정전기에너지_도체_라플라스방정식_A4인쇄용.pdf"


def slide_number(path: Path) -> int:
    match = re.search(r"(\d+)$", path.stem)
    if not match:
        raise ValueError(f"Cannot determine slide number from {path.name}")
    return int(match.group(1))


def slide_paths() -> list[Path]:
    paths = sorted(SLIDE_DIR.glob("slide-*.png"), key=slide_number)
    if not paths:
        raise FileNotFoundError(f"No rendered slides found in {SLIDE_DIR}")
    return paths


def build_presentation_pdf(paths: list[Path]) -> None:
    page_width, page_height = 1600.0, 900.0
    pdf = canvas.Canvas(str(PRESENTATION_PDF), pagesize=(page_width, page_height))
    for path in paths:
        pdf.drawImage(ImageReader(str(path)), 0, 0, page_width, page_height)
        pdf.showPage()
    pdf.save()


def build_a4_pdf(paths: list[Path]) -> None:
    page_width, page_height = landscape(A4)
    margin = 8.0 * 72.0 / 25.4
    available_width = page_width - 2 * margin
    available_height = page_height - 2 * margin
    draw_width = available_width
    draw_height = draw_width / (16.0 / 9.0)
    if draw_height > available_height:
        draw_height = available_height
        draw_width = draw_height * (16.0 / 9.0)
    x = (page_width - draw_width) / 2.0
    y = (page_height - draw_height) / 2.0
    pdf = canvas.Canvas(str(A4_PDF), pagesize=(page_width, page_height))
    for path in paths:
        pdf.drawImage(ImageReader(str(path)), x, y, draw_width, draw_height)
        pdf.showPage()
    pdf.save()


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    paths = slide_paths()
    build_presentation_pdf(paths)
    build_a4_pdf(paths)
    print(f"Created {PRESENTATION_PDF} ({len(paths)} pages)")
    print(f"Created {A4_PDF} ({len(paths)} pages)")


if __name__ == "__main__":
    main()
