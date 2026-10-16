# -*- coding: utf-8 -*-
"""PDF转图片再转回PDF，完全不依赖字体，任何电脑打开效果一致。"""
import os, glob, fitz

BASE = os.path.dirname(os.path.abspath(__file__))
DST = os.path.join(BASE, "已转")
os.makedirs(DST, exist_ok=True)

ZOOM = 300 / 72.0  # 300dpi

for f in sorted(glob.glob(os.path.join(BASE, "*.pdf"))):
    doc = fitz.open(f)
    out = fitz.open()
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(ZOOM, ZOOM), alpha=False)
        np = out.new_page(width=page.rect.width, height=page.rect.height)
        np.insert_image(np.rect, pixmap=pix)
    name = os.path.basename(f)
    out.save(os.path.join(DST, name), deflate=True, garbage=4)
    out.close()
    doc.close()
    print("OK", name)

print("完成，输出在:", DST)