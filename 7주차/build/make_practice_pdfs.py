"""렌더된 PNG(2560×1440)로 이미지 기반 16:9 발표용 PDF와 A4 인쇄용 PDF를 만든다."""
import sys, glob
from pathlib import Path
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
rendered, out16, outA4 = Path(sys.argv[1]), sys.argv[2], sys.argv[3]
slides = sorted(rendered.glob("p-*.png"), key=lambda p: int(p.stem.split("-")[1]))
def write(path, size, margin=0):
    W, H = size; c = canvas.Canvas(path, pagesize=size, pageCompression=1)
    for sl in slides:
        im = ImageReader(str(sl)); w, h = im.getSize()
        k = min((W - 2*margin)/w, (H - 2*margin)/h); dw, dh = w*k, h*k
        c.setFillColorRGB(1, 1, 1); c.rect(0, 0, W, H, fill=1, stroke=0)
        c.drawImage(im, (W-dw)/2, (H-dh)/2, dw, dh, mask="auto"); c.showPage()
    c.save()
write(out16, (13.3333*72, 7.5*72))
write(outA4, landscape(A4), margin=8/25.4*72)
print(len(slides), "pages")
