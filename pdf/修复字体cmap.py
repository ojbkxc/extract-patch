"""
修复PDF嵌入子集TTF缺失cmap表导致跨阅读器乱码。
原理：从PDF ToUnicode CMap提取 Unicode→CID 映射(CIDToGIDMap=Identity即CID=GID=glyph索引)，
给缺cmap的子集TTF补上cmap表后写回FontFile2。不换字体不转图片。
"""
import fitz, re, io, os, glob
from fontTools.ttLib import TTFont, newTable
from fontTools.ttLib.tables._c_m_a_p import CmapSubtable


def parse_tounicode(data):
    """解析ToUnicode CMap流，返回 {CID: unicode}"""
    text = data.decode('latin-1', errors='replace')
    mapping = {}
    for m in re.finditer(r'beginbfchar\s*(.*?)\s*endbfchar', text, re.DOTALL):
        for pair in re.finditer(r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', m.group(1)):
            cid = int(pair.group(1), 16)
            uni = int(pair.group(2), 16)
            mapping[cid] = uni
    for m in re.finditer(r'beginbfrange\s*(.*?)\s*endbfrange', text, re.DOTALL):
        for r in re.finditer(r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', m.group(1)):
            start = int(r.group(1), 16)
            end = int(r.group(2), 16)
            uni_start = int(r.group(3), 16)
            for cid in range(start, end + 1):
                mapping[cid] = uni_start + (cid - start)
    return mapping


def find_ref(obj, key):
    m = re.search(rf'/{key}\s+(\d+)\s+0\s+R', obj)
    return int(m.group(1)) if m else None


def fix_pdf(src, dst):
    doc = fitz.open(src)
    n = doc.xref_length()
    fixed = []

    for xref in range(1, n):
        try:
            obj = doc.xref_object(xref, compressed=False)
        except Exception:
            continue
        if '/Type0' not in obj:
            continue

        tu_xref = find_ref(obj, 'ToUnicode')
        m_df = re.search(r'/DescendantFonts\s*\[\s*(\d+)\s+0\s+R\s*\]', obj)
        if not tu_xref or not m_df:
            continue

        try:
            tu_data = doc.xref_stream(tu_xref)
        except Exception:
            continue
        cid2uni = parse_tounicode(tu_data)
        if not cid2uni:
            continue

        cid_xref = int(m_df.group(1))
        cid_obj = doc.xref_object(cid_xref, compressed=False)

        if '/CIDFontType2' not in cid_obj:
            continue

        fd_xref = find_ref(cid_obj, 'FontDescriptor')
        if not fd_xref:
            continue
        fd_obj = doc.xref_object(fd_xref, compressed=False)

        ff_xref = find_ref(fd_obj, 'FontFile2')
        if not ff_xref:
            continue

        try:
            ff_data = doc.xref_stream(ff_xref)
        except Exception:
            continue

        font = TTFont(io.BytesIO(ff_data), lazy=False)
        if 'cmap' in font:
            continue

        glyph_order = font.getGlyphOrder()
        num_glyphs = len(glyph_order)
        uni2gid = {}
        for cid, uni in cid2uni.items():
            if 0 < cid < num_glyphs and uni > 0:
                uni2gid[uni] = glyph_order[cid]

        if not uni2gid:
            continue

        subtable = CmapSubtable.newSubtable(4)
        subtable.platformID = 3
        subtable.platEncID = 1
        subtable.language = 0
        subtable.cmap = uni2gid

        cmap = newTable('cmap')
        cmap.tableVersion = 0
        cmap.tables = [subtable]
        font['cmap'] = cmap

        out = io.BytesIO()
        font.save(out)
        new_data = out.getvalue()
        doc.update_stream(ff_xref, new_data)

        nm_m = re.search(r'/FontName\s+/([^\s/#]+)', fd_obj)
        nm = nm_m.group(1) if nm_m else '?'
        fixed.append(f'{nm}: +cmap({len(uni2gid)}条)')

    if fixed:
        doc.save(dst, deflate=True, garbage=4, clean=True)
        print(f'  修复 {len(fixed)} 个字体: {fixed}')
    else:
        print('  无需修复(都有cmap)')
    doc.close()


SRC = r'D:\Users\Documents\给老马-2026年三季度完税（原件）\新建文件夹 (2)'
DST = os.path.join(SRC, '修复字体cmap版')
os.makedirs(DST, exist_ok=True)

for f in sorted(glob.glob(os.path.join(SRC, '*.pdf'))):
    name = os.path.basename(f)
    print(f'处理: {name}')
    fix_pdf(f, os.path.join(DST, name))

print('\n全部完成，输出在:', DST)