# -*- coding: utf-8 -*-
"""JPG 完税证明票号 + 原凭证号替换脚本（墨迹精准定位版）。

读取 jpg.xlsx 的 L 列(新票号)和 M 列(新原凭证号)，
按 D 列姓名匹配 JPG 文件，用 OCR 粗定位 + ink_box 精确定位墨迹，
擦除旧文字 + 渲染新文字（SimHei，字号从墨迹高度反推），保存为 *_fixed2.jpg。

优先处理 *_fixed.jpg（2换所有金额.py 的输出）；若不存在则处理原始 *.jpg。

数据来源：jpg.xlsx（D列=姓名, L列=新票号, M列=新原凭证号）
字体：simhei.ttf（黑体），字号从原图墨迹高度反推
"""
import os
import re
import sys

import numpy as np
import openpyxl
from PIL import Image, ImageDraw, ImageFont

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

SRC = os.path.dirname(os.path.abspath(__file__))
FONT_HEI = r'C:\Windows\Fonts\simhei.ttf'
FONT_AMOUNT = os.path.join(SRC, '金额.ttf')
XLSX = os.path.join(SRC, 'jpg.xlsx')

INK_TH = 150
BG = (255, 255, 255)
FG = (0, 0, 0)
_PAD = 24

_SIMHEI_SIZE_MAP = {
    6: 9, 7: 10, 8: 12, 9: 13, 10: 14, 11: 16, 12: 17,
    13: 19, 14: 20, 15: 21, 16: 23, 17: 24, 18: 25,
}


def guess_font_size(ink_height):
    if ink_height in _SIMHEI_SIZE_MAP:
        return _SIMHEI_SIZE_MAP[ink_height]
    if ink_height <= 6:
        return 9
    if ink_height >= 18:
        return 25
    return _SIMHEI_SIZE_MAP.get(ink_height, 14)


def load_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    data = {}
    for row in range(2, ws.max_row + 1):
        name = ws.cell(row, 4).value
        if not name:
            continue
        ticket = ws.cell(row, 12).value
        orig = ws.cell(row, 13).value
        data[str(name)] = {
            'ticket': str(ticket) if ticket else None,
            'orig': str(orig) if orig else None,
        }
    return data


_rapid_engine = None


def get_rapid_engine():
    global _rapid_engine
    if _rapid_engine is not None:
        return _rapid_engine
    try:
        from rapidocr import RapidOCR
    except ImportError:
        try:
            from rapidocr_onnxruntime import RapidOCR
        except ImportError:
            sys.exit('缺少 rapidocr，请安装：\n  pip install rapidocr-onnxruntime')
    _rapid_engine = RapidOCR()
    return _rapid_engine


def ocr_boxes(img_path):
    eng = get_rapid_engine()
    out = eng(img_path)
    if out is None:
        return [], []
    if hasattr(out, 'txts'):
        return list(out.txts), list(out.boxes)
    result = out[0] if isinstance(out, tuple) else out
    if not result:
        return [], []
    txts, boxes = [], []
    for item in result:
        if len(item) >= 2:
            boxes.append(item[0])
            txts.append(item[1])
    return txts, boxes


def box_to_rect(box):
    xs = [float(p[0]) for p in box]
    ys = [float(p[1]) for p in box]
    return (int(min(xs)), int(min(ys)), int(max(xs)), int(max(ys)))


def ink_box(rgb, x1, y1, x2, y2):
    reg = rgb[y1:y2, x1:x2]
    ys, xs = np.where(reg.max(axis=2) < INK_TH)
    if len(xs) == 0:
        return None
    return (x1 + int(xs.min()), y1 + int(ys.min()),
            x1 + int(xs.max()), y1 + int(ys.max()))


def ink_runs(rgb, x1, y1, x2, y2, gap=4):
    m = rgb[y1:y2, x1:x2].max(axis=2) < INK_TH
    idx = np.where(m.any(axis=0))[0]
    if len(idx) == 0:
        return []
    runs, s, p = [], idx[0], idx[0]
    for c in idx[1:]:
        if c - p > gap:
            runs.append((s, p))
            s = c
        p = c
    runs.append((s, p))
    out = []
    for a, b in runs:
        ys = np.where(m[:, a:b + 1].any(axis=1))[0]
        out.append((x1 + int(a), y1 + int(ys.min()),
                    x1 + int(b), y1 + int(ys.max())))
    return out


def text_mask(font, text):
    w = int(font.getlength(text)) + _PAD * 2 + 8
    asc, desc = font.getmetrics()
    h = asc + desc + _PAD * 2 + 8
    canvas = Image.new('L', (max(w, 8), max(h, 8)), 255)
    ImageDraw.Draw(canvas).text((_PAD, _PAD), text, fill=0, font=font)
    g = np.array(canvas)
    ys, xs = np.where(g < INK_TH)
    bb = (int(xs.min()), int(ys.min()), int(xs.max()), int(ys.max()))
    return Image.fromarray((255 - g).astype(np.uint8)), bb


def draw_left(img, font, text, left_x, center_y):
    mask, (x0, y0, x1, y1) = text_mask(font, text)
    ax = left_x - x0
    ay = int(round(center_y - (y0 + y1) / 2.0))
    img.paste(FG, (int(round(ax)), ay), mask)


def erase(draw, box):
    x1, y1, x2, y2 = box
    draw.rectangle([x1 - 1, y1 - 1, x2 + 1, y2 + 1], fill=BG)


def find_ocr_box(txts, boxes, target_num):
    if not target_num:
        return None
    for i, txt in enumerate(txts):
        digits = re.sub(r'\D', '', txt)
        if digits == target_num:
            return box_to_rect(boxes[i])
    blocks = []
    for i, txt in enumerate(txts):
        digits = re.sub(r'\D', '', txt)
        if digits and digits in target_num and len(digits) >= 4:
            blocks.append(box_to_rect(boxes[i]))
    if blocks:
        all_xs, all_ys = [], []
        for x1, y1, x2, y2 in blocks:
            all_xs.extend([x1, x2])
            all_ys.extend([y1, y2])
        return (min(all_xs), min(all_ys), max(all_xs), max(all_ys))
    return None


def replace_number(img, draw, rgb, ocr_rect, old_num, new_num, label, has_label=False, font_path=None, fixed_size=None):
    """在 ocr_rect 范围内精确定位数字墨迹，擦除+渲染新数字。
    has_label=True 时用 ink_runs 排除左侧 'No.' 标签段，合并剩余段为数字 bbox。
    y 范围限制在 OCR box 上部，避免包含下方其他内容。
    font_path/fixed_size 指定时用该字体和字号，否则用 SimHei + guess_font_size。
    """
    x1, y1, x2, y2 = ocr_rect
    ry1 = y1 - 1
    ry2 = y2 - 2

    if has_label:
        runs = ink_runs(rgb, x1 - 2, ry1, x2 + 2, ry2, gap=3)
        if len(runs) < 2:
            ib = ink_box(rgb, x1 - 2, ry1, x2 + 2, ry2)
        else:
            digit_runs = runs[1:]
            all_xs = [r[0] for r in digit_runs] + [r[2] for r in digit_runs]
            all_ys = [r[1] for r in digit_runs] + [r[3] for r in digit_runs]
            ib = (min(all_xs), min(all_ys), max(all_xs), max(all_ys))
    else:
        ib = ink_box(rgb, x1 - 2, ry1, x2 + 2, ry2)

    if ib is None:
        print(f'  {label}: 未找到墨迹')
        return False

    ink_h = ib[3] - ib[1] + 1
    ink_w = ib[2] - ib[0] + 1
    if fixed_size:
        font_size = fixed_size
    else:
        font_size = guess_font_size(ink_h)
    fp = font_path if font_path else FONT_HEI
    font = ImageFont.truetype(fp, font_size)

    _, (tx0, ty0, tx1, ty1) = text_mask(font, new_num)
    new_w = tx1 - tx0 + 1
    erase_box = (ib[0], ib[1], max(ib[2], ib[0] + new_w - 1), ib[3])

    cy = (ib[1] + ib[3]) / 2.0
    erase(draw, erase_box)
    draw_left(img, font, new_num, ib[0], cy)

    print(f'  {label}: {old_num} -> {new_num}')
    print(f'    墨迹=({ib[0]},{ib[1]},{ib[2]},{ib[3]}) 高={ink_h} 宽={ink_w} 字号={font_size}')
    return True


def process_image(img_path, info, out_path, force=False):
    img = Image.open(img_path).convert('RGB')
    rgb = np.array(img)
    draw = ImageDraw.Draw(img)

    txts, boxes = ocr_boxes(img_path)
    if not txts:
        print('  OCR 无结果，跳过')
        return False

    all_text = '\n'.join(txts)
    changed = False

    old_ticket = None
    m = re.search(r'No\.?\s*[:：]?\s*(\d{15,20})', all_text)
    if m:
        old_ticket = m.group(1)

    new_ticket = info.get('ticket')
    if new_ticket and old_ticket and (old_ticket != new_ticket or force):
        ocr_rect = find_ocr_box(txts, boxes, old_ticket)
        if ocr_rect:
            changed |= replace_number(img, draw, rgb, ocr_rect,
                                       old_ticket, new_ticket, '票号', has_label=True)
        else:
            print(f'  票号: 未找到 OCR 位置')
    elif new_ticket and old_ticket == new_ticket:
        print(f'  票号: 已一致，无需变更')
    else:
        print(f'  票号: 无数据 (old={old_ticket}, new={new_ticket})')

    m_tax = re.search(r'识别号\s*[:：]?\s*(\d{17}[\dXx])', all_text)
    tax_id = m_tax.group(1) if m_tax else None
    old_orig = None
    for n in re.findall(r'\d{18}', all_text):
        if n != old_ticket and n != tax_id:
            old_orig = n
            break

    new_orig = info.get('orig')
    if new_orig and old_orig and (old_orig != new_orig or force):
        ocr_rect = find_ocr_box(txts, boxes, old_orig)
        if ocr_rect:
            changed |= replace_number(img, draw, rgb, ocr_rect,
                                       old_orig, new_orig, '原凭证号')
        else:
            print(f'  原凭证号: 未找到 OCR 位置')
    elif new_orig and old_orig == new_orig:
        print(f'  原凭证号: 已一致，无需变更')
    else:
        print(f'  原凭证号: 无数据 (old={old_orig}, new={new_orig})')

    img.save(out_path, 'JPEG', quality=95)
    print(f'  已保存: {os.path.basename(out_path)}')
    return True


def extract_name(jpg):
    parts = jpg.replace('.jpg', '').replace('_fixed', '').replace('_fixed2', '').split('_')
    for p in parts:
        p = p.strip()
        if '.' in p:
            p = p.split('.')[-1].strip()
        if len(p) >= 2 and not any(c.isdigit() for c in p) and p != '完税证明':
            return p
    return None


def main():
    force = '--force' in sys.argv
    if not os.path.exists(XLSX):
        print(f'找不到 {XLSX}')
        return

    data = load_xlsx(XLSX)
    print(f'读取 {len(data)} 条数据' + (' [强制重渲染]' if force else ''))

    fixed = sorted(f for f in os.listdir(SRC) if f.endswith('_fixed.jpg'))
    if fixed:
        jpgs = fixed
        out_suffix = '_fixed2.jpg'
        print(f'找到 {len(jpgs)} 张已替换金额的图片(*_fixed.jpg)')
    else:
        jpgs = sorted(f for f in os.listdir(SRC)
                      if f.endswith('.jpg') and '_fixed' not in f)
        out_suffix = '_fixed.jpg'
        print(f'未找到 *_fixed.jpg，处理原始图片 {len(jpgs)} 张')

    success = 0
    for jpg in jpgs:
        name = extract_name(jpg)
        if not name or name not in data:
            print(f'跳过 {jpg}：无法匹配人名 (提取到: {name})')
            continue

        info = data[name]
        print(f'\n处理: {jpg}')
        print(f'  姓名: {name}')

        in_path = os.path.join(SRC, jpg)
        base = jpg
        for s in ('_fixed2.jpg', '_fixed.jpg', '.jpg'):
            if base.endswith(s):
                base = base[:-len(s)]
                break
        out_path = os.path.join(SRC, base + out_suffix)
        try:
            if process_image(in_path, info, out_path, force=force):
                success += 1
        except Exception as e:
            print(f'  错误: {e}')
            import traceback
            traceback.print_exc()

    print(f'\n完成: {success}/{len(jpgs)}')


if __name__ == '__main__':
    main()