#!/usr/bin/env python3
"""6주차 연습 발표자료: 검증된 2560x1440 렌더에서 발표용 16:9 PDF와 A4 인쇄용 PDF 생성 (지침 §6.2–6.3)."""
import glob, re, sys
from pathlib import Path
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
src = Path(sys.argv[1]); out = Path(sys.argv[2]); stem = sys.argv[3]
paths = sorted(src.glob("slide-*.png"), key=lambda p: int(re.search(r"(\d+)$", p.stem).group(1)))
c = canvas.Canvas(str(out / f"{stem}.pdf"), pagesize=(1600, 900))
for p in paths: c.drawImage(ImageReader(str(p)), 0, 0, 1600, 900); c.showPage()
c.save()
W, H = landscape(A4); m = 8 * 72 / 25.4; dw = W - 2 * m; dh = dw * 9 / 16
if dh > H - 2 * m: dh = H - 2 * m; dw = dh * 16 / 9
c = canvas.Canvas(str(out / f"{stem}_A4인쇄용.pdf"), pagesize=(W, H))
for p in paths: c.drawImage(ImageReader(str(p)), (W - dw) / 2, (H - dh) / 2, dw, dh); c.showPage()
c.save()
print(len(paths), "pages")
