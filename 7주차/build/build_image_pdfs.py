from pathlib import Path

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[1]
RENDERED = ROOT / "rendered-v4"
OUTPUT = ROOT / "output"
SLIDES = sorted(RENDERED.glob("slide-*.png"), key=lambda p: int(p.stem.split("-")[1]))


def write_pdf(path: Path, page_size: tuple[float, float], margin: float = 0) -> None:
    page_w, page_h = page_size
    pdf = canvas.Canvas(str(path), pagesize=page_size, pageCompression=1)
    for slide in SLIDES:
        image = ImageReader(str(slide))
        source_w, source_h = image.getSize()
        scale = min((page_w - 2 * margin) / source_w, (page_h - 2 * margin) / source_h)
        draw_w, draw_h = source_w * scale, source_h * scale
        x = (page_w - draw_w) / 2
        y = (page_h - draw_h) / 2
        pdf.setFillColorRGB(1, 1, 1)
        pdf.rect(0, 0, page_w, page_h, fill=1, stroke=0)
        pdf.drawImage(image, x, y, width=draw_w, height=draw_h, preserveAspectRatio=True, mask="auto")
        pdf.showPage()
    pdf.save()


if len(SLIDES) != 36:
    raise RuntimeError(f"Expected 36 rendered slides, found {len(SLIDES)}")

OUTPUT.mkdir(parents=True, exist_ok=True)
write_pdf(
    OUTPUT / "전자기학_7주차_영상법과_변수분리법.pdf",
    (13.3333 * 72, 7.5 * 72),
)
write_pdf(
    OUTPUT / "전자기학_7주차_영상법과_변수분리법_A4인쇄용.pdf",
    landscape(A4),
    margin=8 / 25.4 * 72,
)
