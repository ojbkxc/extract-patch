# -*- coding: utf-8 -*-
"""从同目录所有 PDF 提取金额相关字形，合并生成 金额.ttf。
¥ 取自 MicrosoftYaHei，0-9/,/. 取自 FZKTJW--GB1-0，
两者 unitsPerEm 不同，¥ 字形按比例缩放后注入数字字体。"""
import io, glob, os, sys
import pymupdf
from fontTools.ttLib import TTFont
from fontTools.subset import Subsetter, Options
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.pens.transformPen import TransformPen
from fontTools.misc.transform import Transform

SRC = os.path.dirname(os.path.abspath(__file__))
DIGIT_CHARS = {0x30,0x31,0x32,0x33,0x34,0x35,0x36,0x37,0x38,0x39,0x2c,0x2e}  # 0-9 , .
YEN = 0xa5
FONT_DIGITS = 'FZKTJW--GB1-0'
FONT_YEN = 'MicrosoftYaHei'
OUT_NAME = '金额.ttf'


def extract_embedded_ttf(pdf_path, base_font_name):
    """从 PDF 中提取指定 BaseFont 的嵌入 TrueType 子集字节。"""
    doc = pymupdf.open(pdf_path)
    try:
        for xref in range(doc.xref_length()):
            try:
                obj = doc.xref_object(xref)
            except Exception:
                continue
            if '/FontDescriptor' not in obj:
                continue
            try:
                name, ext, typ, buf = doc.extract_font(xref)
            except Exception:
                continue
            if not buf or ext != 'ttf':
                continue
            bf = name.split('+', 1)[1] if '+' in name else name
            if bf == base_font_name:
                return buf
    finally:
        doc.close()
    return None


def subset_font(buf, unicodes):
    """子集化字体，只保留指定 Unicode 字符。"""
    f = TTFont(io.BytesIO(buf))
    opts = Options()
    opts.layout_features = []
    opts.name_IDs = ['*']
    opts.notdef_outline = True
    sub = Subsetter(opts)
    sub.populate(unicodes=unicodes)
    sub.subset(f)
    out = io.BytesIO()
    f.save(out)
    return out.getvalue()


def find_source(pdfs, base_font, required_chars):
    """在 PDF 列表中找第一个包含全部 required_chars 的嵌入子集。"""
    for p in pdfs:
        buf = extract_embedded_ttf(p, base_font)
        if not buf:
            continue
        try:
            cmap = TTFont(io.BytesIO(buf)).getBestCmap()
        except Exception:
            continue
        if all(c in cmap for c in required_chars):
            return p, buf
    return None, None


def inject_glyph(base, src, codepoint, glyph_name):
    """把 src 中 codepoint 的字形按 unitsPerEm 比例缩放后注入 base。"""
    upem_base = base['head'].unitsPerEm
    upem_src = src['head'].unitsPerEm
    scale = upem_base / upem_src

    src_cmap = src.getBestCmap()
    src_gname = src_cmap[codepoint]

    pen = TTGlyphPen(base.getGlyphSet())
    tpen = TransformPen(pen, Transform().scale(scale))
    src.getGlyphSet()[src_gname].draw(tpen)
    new_glyph = pen.glyph()

    if glyph_name in base['glyf'].glyphs:
        glyph_name += '_x'
    base['glyf'][glyph_name] = new_glyph

    src_advance = src['hmtx'].metrics[src_gname][0]
    base['hmtx'].metrics[glyph_name] = (round(src_advance * scale), 0)

    for sub in base['cmap'].tables:
        if sub.isUnicode():
            sub.cmap[codepoint] = glyph_name

    extra = getattr(base['post'], 'extraNames', None)
    if extra is not None and glyph_name not in extra:
        extra.append(glyph_name)
    base['post'].glyphOrder = base.getGlyphOrder() + [glyph_name]
    base['maxp'].numGlyphs = len(base['glyf'].glyphs)
    return scale


def main():
    pdfs = sorted(glob.glob(os.path.join(SRC, '*.pdf')))
    if not pdfs:
        print('未找到 PDF 文件'); sys.exit(1)
    print(f'扫描 {len(pdfs)} 个 PDF ...')

    p_digits, buf_digits = find_source(pdfs, FONT_DIGITS, DIGIT_CHARS)
    if not buf_digits:
        print(f'未找到含完整数字/标点的 {FONT_DIGITS} 子集'); sys.exit(1)
    print(f'数字字体来源: {os.path.basename(p_digits)}')

    p_yen, buf_yen = find_source(pdfs, FONT_YEN, {YEN})
    if not buf_yen:
        print(f'未找到含 ¥ 的 {FONT_YEN} 子集'); sys.exit(1)
    print(f'¥ 字体来源:   {os.path.basename(p_yen)}')

    font_digits = subset_font(buf_digits, DIGIT_CHARS)
    font_yen = subset_font(buf_yen, {YEN})

    base = TTFont(io.BytesIO(font_digits))
    src = TTFont(io.BytesIO(font_yen))
    scale = inject_glyph(base, src, YEN, 'yen')
    print(f'unitsPerEm: 数字={base["head"].unitsPerEm} ¥源={src["head"].unitsPerEm} 缩放={scale}')

    out_path = os.path.join(SRC, OUT_NAME)
    base.save(out_path)

    verify = TTFont(out_path)
    chars = sorted(chr(c) for c in verify.getBestCmap())
    print(f'生成 {OUT_NAME}: {os.path.getsize(out_path)} 字节, 字符={chars}')


if __name__ == '__main__':
    main()