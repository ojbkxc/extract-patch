#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""从完税证明 PDF + 季度汇总表(xlsx) 生成 pdf.xlsx（含二维码解析与改写）。

用法：
    python build_pdf_xlsx.py [--pdf-dir DIR] [--q3 FILE] [--out FILE] [--seed N] [-v]

默认值：
    --pdf-dir   脚本所在目录（存放 *.pdf）
    --q3        在 --pdf-dir 及其上级目录自动查找 *季度*.xlsx；命中多个时需显式指定
    --out       <pdf-dir>/pdf.xlsx
    --seed      票号后三位随机种子（同种子 + 同输入 => 结果可复现），默认 20261013

依赖：openpyxl / pymupdf / Pillow / zxing-cpp
    pip install openpyxl pymupdf pillow zxing-cpp

依次完成的步骤：
    1. 扫描 --pdf-dir 下 *.pdf，按文件名开头的数字排序，序号即该数字
    2. 逐份抽取：企业名称、纳税人识别号、票号(No●)、实缴金额(¥)、二维码原文
    3. 法人姓名：按税号从 --q3 表 B/C 列匹配
    4. N/Q/R/S/T 五列：按 (企业名称, 税号) 匹配 --q3 表，写入值并完整复制格式
       （字体/填充/边框/对齐/数字格式/列宽），落在 F~J 列
    5. 二维码第 5 段（金额）改写为 F 列（所得税）值，保留两位小数
    6. 票号后三位随机重排，每位数字取自二维码中票号之前出现过的数字；
       若票号前的字段里没有任何数字（如区划码为 null），回退为票号自身前 15 位
       出现过的数字；两者都没有数字时才保持原值

安全约束：
    * 只读输入、只写 --out，输入文件永不被修改
    * 先写 .tmp 再 os.replace 原子替换；任何异常都会清理 .tmp，
      已存在的 --out 保持原样 —— 执行失败不影响正常文件输出
    * 依赖缺失 / 无 PDF / 二维码识别失败 => 报错退出，不产生输出文件
"""
from __future__ import annotations

import argparse
import os
import random
import re
import sys
import glob
from copy import copy

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEADERS = ['序号', '企业名称', '税号', '法人 姓名', '二维码的值']
Q3_SRC_COLS = (14, 17, 18, 19, 20)          # N Q R S T  ->  F G H I J
COL_WIDTHS = {'A': 8.0, 'B': 42.0, 'C': 24.0, 'D': 12.0, 'E': 60.0}
HEADER_H, DATA_H = 32.4, 28.8
FILL_HEADER, FILL_ODD, FILL_EVEN = '1F4E79', 'FFFFFF', 'F7F9FC'
SIDE = Side(style='thin', color='D9DEE7')
BORDER = Border(left=SIDE, right=SIDE, top=SIDE, bottom=SIDE)
FONT_NAME = '宋体'


def die(msg):
    sys.exit('[中止] ' + msg)


def load_fitz():
    try:
        import pymupdf as fitz
    except ImportError:
        try:
            import fitz
        except ImportError:
            die('缺少依赖 pymupdf，请先执行: pip install pymupdf')
    return fitz


def file_order_key(name):
    m = re.match(r'\s*(\d+)', name)
    return (0, int(m.group(1))) if m else (1, name)


def copy_style(src, dst):
    """整体复制单元格格式（字体/填充/边框/对齐/保护/数字格式）。"""
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.protection = copy(src.protection)
    dst.number_format = src.number_format


def find_q3(pdf_dir, explicit):
    if explicit:
        if not os.path.isfile(explicit):
            die('指定的 --q3 文件不存在: %s' % explicit)
        return os.path.abspath(explicit)
    pats = (os.path.join(pdf_dir, '*季度*.xlsx'),
            os.path.join(os.path.dirname(pdf_dir), '*季度*.xlsx'))
    hits = sorted({p for pat in pats for p in glob.glob(pat)
                   if os.path.basename(p) != 'pdf.xlsx'})
    if len(hits) == 1:
        return os.path.abspath(hits[0])
    if not hits:
        die('未找到季度汇总表，请用 --q3 指定')
    die('找到多个候选汇总表，请用 --q3 指定其一：\n  ' + '\n  '.join(hits))


def extract_rows(pdf_dir, warns):
    fitz = load_fitz()
    try:
        from PIL import Image
        import zxingcpp
    except ImportError:
        die('缺少依赖 Pillow / zxing-cpp，请先执行: pip install pillow zxing-cpp')

    files = [f for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')]
    if not files:
        die('目录中没有 PDF 文件: %s' % pdf_dir)
    files.sort(key=file_order_key)

    rows = []
    for f in files:
        path = os.path.join(pdf_dir, f)
        doc = fitz.open(path)
        text = '\n'.join(pg.get_text() for pg in doc)
        pix = doc[0].get_pixmap(dpi=300)
        img = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
        codes = zxingcpp.read_barcodes(img)
        qr = codes[0].text if codes else None
        if not qr:
            die('%s 未能识别到二维码，已终止（不产生输出文件）' % f)

        pdf_name = None
        m = re.search(r'纳税人名称\s*[:：]?\s*([^\n]+)', text)
        if m:
            pdf_name = m.group(1).strip()
            cuts = [x for x in (pdf_name.find('（'), pdf_name.find('(')) if x > 0]
            if cuts:
                pdf_name = pdf_name[:min(cuts)]
        stem = re.sub(r'^\s*\d+\s*[.、,，]?\s*', '', os.path.splitext(f)[0]).strip()
        name = pdf_name or stem
        if pdf_name and stem and pdf_name != stem:
            if stem.startswith(pdf_name) or pdf_name.startswith(stem):
                name = pdf_name if len(pdf_name) > len(stem) else stem
            else:
                warns.append('%s：PDF 名称「%s」与文件名「%s」不一致，采用 PDF 名称'
                             % (f, pdf_name, stem))

        m = re.search(r'纳税人识别号\s*[:：]?\s*([0-9A-Za-z]{15,20})', text)
        tid = m.group(1).upper() if m else None
        m = re.search(r'No[^\d\n]{0,4}(\d{15,25})', text)
        ticket = m.group(1) if m else None
        m = re.search(r'[¥￥]\s*([\d,]+\.\d{2})', text)
        amount = m.group(1).replace(',', '') if m else None
        if not tid:
            warns.append('%s：未提取到纳税人识别号' % f)
        if not ticket:
            warns.append('%s：未提取到票号(No●)' % f)

        rows.append(dict(file=f, name=name, tid=tid, ticket=ticket,
                         amount=amount, qr=qr))
    return rows


def read_q3(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[0]
    by_key, by_tid = {}, {}
    for r in range(2, ws.max_row + 1):
        b, c = ws.cell(r, 2).value, ws.cell(r, 3).value
        if b in (None, '') or c in (None, ''):
            continue
        by_key.setdefault((str(b).strip(), str(c).strip()), r)
        by_tid.setdefault(str(c).strip(), r)
    return ws, by_key, by_tid


def build(pdf_dir, q3_path, out_path, seed, verbose):
    warns = []
    rows = extract_rows(pdf_dir, warns)
    q3_ws, by_key, by_tid = read_q3(q3_path)

    wb = openpyxl.Workbook()
    sh = wb.active
    sh.title = 'Sheet1'

    for i, h in enumerate(HEADERS):
        c = sh.cell(1, i + 1)
        c.value = h
        c.font = Font(name=FONT_NAME, size=11, bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor=FILL_HEADER)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = BORDER
    sh.row_dimensions[1].height = HEADER_H

    unmatched = []
    for i, row in enumerate(rows):
        r = i + 2
        vals = [i + 1, row['name'], row['tid'], None, row['qr']]
        src_r = by_key.get((str(row['name']).strip(), str(row['tid']).strip()))
        if src_r is None:
            src_r = by_tid.get(str(row['tid']).strip())
        if src_r is None:
            unmatched.append(row['name'])
            warns.append('第 %d 行「%s」在汇总表中未匹配到，F~J 列留空' % (i + 1, row['name']))
        else:
            vals[3] = q3_ws.cell(src_r, 4).value      # 法人 姓名

        for cidx, v in enumerate(vals):
            c = sh.cell(r, cidx + 1)
            c.value = v
            c.font = Font(name=FONT_NAME, size=11)
            c.fill = PatternFill('solid', fgColor=FILL_ODD if i % 2 == 0 else FILL_EVEN)
            c.border = BORDER
            c.alignment = Alignment(
                horizontal='left' if cidx in (1, 4) else 'center', vertical='center')
        sh.cell(r, 3).number_format = '@'
        sh.cell(r, 5).number_format = '@'
        sh.row_dimensions[r].height = DATA_H

        row['src_r'] = src_r

    for j, sc in enumerate(Q3_SRC_COLS):
        src = q3_ws.cell(1, sc)
        dst = sh.cell(1, 6 + j)
        dst.value = src.value
        copy_style(src, dst)
        w = q3_ws.column_dimensions[get_column_letter(sc)].width
        if w:
            sh.column_dimensions[get_column_letter(6 + j)].width = w
    copy_style(sh.cell(1, 7), sh.cell(1, 11))
    sh.cell(1, 11).value = None
    for i, row in enumerate(rows):
        if row['src_r'] is None:
            continue
        for j, sc in enumerate(Q3_SRC_COLS):
            src = q3_ws.cell(row['src_r'], sc)
            dst = sh.cell(i + 2, 6 + j)
            dst.value = src.value
            copy_style(src, dst)
    for i, row in enumerate(rows):
        r = i + 2
        if row['src_r'] is not None:
            copy_style(sh.cell(r, 7), sh.cell(r, 11))
        else:
            kc = sh.cell(r, 11)
            kc.font = Font(name=FONT_NAME, size=11)
            kc.fill = PatternFill('solid', fgColor=FILL_ODD if i % 2 == 0 else FILL_EVEN)
            kc.border = BORDER
            kc.alignment = Alignment(horizontal='center', vertical='center')
        sh.cell(r, 11).value = None
    for col, w in COL_WIDTHS.items():
        sh.column_dimensions[col].width = w
    sh.freeze_panes = 'A2'

    changed, skip_ticket = 0, []
    for i, row in enumerate(rows):
        r = i + 2
        parts = str(sh.cell(r, 5).value).split(',')
        old = parts[2] if len(parts) > 2 else None
        fval = sh.cell(r, 6).value
        if len(parts) >= 5 and isinstance(fval, (int, float)):
            parts[4] = '%.2f' % float(fval)
        elif len(parts) >= 5:
            warns.append('第 %d 行无 F 列值，二维码金额段未改写' % (i + 1))
        pool = sorted({ch for ch in (parts[0] + parts[1]) if ch.isdigit()}) if len(parts) >= 2 else []
        fallback = False
        if not pool and old and len(old) > 3:
            # 票号前面的字段中没有数字（如区划码字段为 null）：回退用票号自身前面的数字
            pool = sorted({ch for ch in old[:-3] if ch.isdigit()})
            fallback = bool(pool)
        if old and len(old) > 3 and pool:
            rng = random.Random('%s|%s|%s' % (seed, row['tid'], old))
            cands = [a + b + c for a in pool for b in pool for c in pool if a + b + c != old[-3:]]
            parts[2] = old[:-3] + rng.choice(cands)
            changed += 1
            if verbose:
                print('  第 %-3d 行 票号后三位 %s -> %s  (可用数字 %s%s)'
                      % (i + 1, old[-3:], parts[2][-3:], ''.join(pool),
                         '，票号前无数字，回退用票号自身前面的数字' if fallback else ''))
        else:
            skip_ticket.append(i + 1)
            warns.append('第 %d 行无法确定可用数字，票号后三位未变更' % (i + 1))
        sh.cell(r, 5).value = ','.join(parts)

    tmp = out_path + '.tmp'
    try:
        wb.save(tmp)
        os.replace(tmp, out_path)
    except BaseException:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        raise

    print('输出文件: %s' % out_path)
    print('数据行: %d | 记录格式: %d | 票号后三位变更: %d | 未变更: %s'
          % (len(rows), len(rows) - len(unmatched), changed,
             '、'.join(map(str, skip_ticket)) or '无'))
    if unmatched:
        print('未匹配汇总表: %s' % '、'.join(unmatched))
    for w in warns:
        print('提示: ' + w)
    return len(rows)


def main():
    ap = argparse.ArgumentParser(description='完税证明 PDF + 季度汇总表 -> pdf.xlsx')
    ap.add_argument('--pdf-dir', default=None, help='存放 PDF 的目录（默认脚本所在目录）')
    ap.add_argument('--q3', default=None, help='季度汇总表 xlsx（默认自动查找 *季度*.xlsx）')
    ap.add_argument('--out', default=None, help='输出文件（默认 <pdf-dir>/pdf.xlsx）')
    ap.add_argument('--seed', type=int, default=20261013, help='票号随机种子，默认 20261013')
    ap.add_argument('-v', '--verbose', action='store_true', help='打印每行票号变更明细')
    a = ap.parse_args()

    pdf_dir = os.path.abspath(a.pdf_dir or os.path.dirname(os.path.abspath(__file__)))
    if not os.path.isdir(pdf_dir):
        die('目录不存在: %s' % pdf_dir)
    q3 = find_q3(pdf_dir, a.q3)
    out = os.path.abspath(a.out or os.path.join(pdf_dir, 'pdf.xlsx'))
    build(pdf_dir, q3, out, a.seed, a.verbose)


if __name__ == '__main__':
    main()
