#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""从完税证明 JPG + 季度汇总表(xlsx) 生成 jpg.xlsx。

用法：
    python 1读取jpg创建表格.py [--jpg-dir DIR] [--q3 FILE] [--out FILE] [-v]

默认值：
    --jpg-dir  脚本所在目录（存放 *.jpg）
    --q3       在 --jpg-dir 及其上级目录自动查找 *季度*.xlsx；命中多个时需显式指定
    --out      <jpg-dir>/jpg.xlsx

依赖：openpyxl
    pip install openpyxl

依次完成的步骤：
    1. 扫描 --jpg-dir 下 *.jpg，按文件名开头的数字排序，序号即该数字
    2. 从文件名提取法人姓名、身份证号（格式：日期_姓名_身份证号_完税证明_序号.jpg）
    3. 按身份证号从 --q3 表 D/E 列匹配汇总行
    4. 企业名称、税号、法人姓名：取自汇总表 B/C/D 列
    5. N/Q/R/S/T 五列：按匹配行复制值并完整复制格式
       （字体/填充/边框/对齐/数字格式/列宽），落在 F~J 列
    6. E 列「二维码的值」留空（JPG 二维码为加密内容，不解析）

安全约束：
    * 只读输入、只写 --out，输入文件永不被修改
    * 先写 .tmp 再 os.replace 原子替换；任何异常都会清理 .tmp，
      已存在的 --out 保持原样 —— 执行失败不影响正常文件输出
    * 依赖缺失 / 无 JPG => 报错退出，不产生输出文件
"""
from __future__ import annotations

import argparse
import os
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


def find_q3(jpg_dir, explicit):
    if explicit:
        if not os.path.isfile(explicit):
            die('指定的 --q3 文件不存在: %s' % explicit)
        return os.path.abspath(explicit)
    cur_pat = os.path.join(jpg_dir, '*季度*.xlsx')
    cur_hits = sorted(p for p in glob.glob(cur_pat)
                      if os.path.basename(p) != 'jpg.xlsx')
    if len(cur_hits) == 1:
        return os.path.abspath(cur_hits[0])
    if len(cur_hits) > 1:
        die('当前目录找到多个候选汇总表，请用 --q3 指定其一：\n  ' + '\n  '.join(cur_hits))
    up_hits = sorted(p for p in glob.glob(os.path.join(os.path.dirname(jpg_dir), '*季度*.xlsx'))
                     if os.path.basename(p) != 'jpg.xlsx')
    if len(up_hits) == 1:
        return os.path.abspath(up_hits[0])
    if not up_hits:
        die('未找到季度汇总表，请用 --q3 指定')
    die('上级目录找到多个候选汇总表，请用 --q3 指定其一：\n  ' + '\n  '.join(up_hits))


def parse_jpg_name(fname, warns):
    """从文件名提取法人姓名和身份证号。

    格式：[序号. ]日期_姓名_身份证号_完税证明_x.jpg
    身份证号为 18 位（末位可能为 X），优先用正则定位，再取其前一段为姓名。
    """
    stem = re.sub(r'^\s*\d+\s*[.、,，]?\s*', '', os.path.splitext(fname)[0]).strip()
    m = re.search(r'(\d{17}[\dXx])', stem)
    if not m:
        warns.append('%s：未能从文件名提取身份证号' % fname)
        parts = [p for p in stem.split('_') if p]
        name = parts[1] if len(parts) > 1 else stem
        return name, None
    idcard = m.group(1).upper()
    before = stem[:m.start()].rstrip('_')
    seg = [p for p in before.split('_') if p]
    name = seg[-1] if seg else None
    if not name:
        warns.append('%s：未能从文件名提取法人姓名' % fname)
    return name, idcard


def extract_rows(jpg_dir, warns):
    files = [f for f in os.listdir(jpg_dir) if f.lower().endswith('.jpg')]
    if not files:
        die('目录中没有 JPG 文件: %s' % jpg_dir)
    files.sort(key=file_order_key)

    rows = []
    for f in files:
        name, idcard = parse_jpg_name(f, warns)
        rows.append(dict(file=f, name=name, idcard=idcard))
    return rows


def num_to_chinese(num):
    """数字转人民币中文大写，返回去掉开头'人'的字符串。
    例: 4341.7 -> '民币肆仟叁佰肆拾壹圆柒角整'
    下游 2换所有金额.py 会自动补'人'，故此处去'人'避免重复。
    """
    digits = '零壹贰叁肆伍陆柒捌玖'
    n = int(round(num * 100))
    integer_part = n // 100
    fen = n % 100
    jiao = fen // 10
    fen_unit = fen % 10

    if integer_part == 0:
        int_str = '零'
    else:
        groups = []
        t = integer_part
        while t > 0:
            groups.append(t % 10000)
            t //= 10000
        big = ['', '万', '亿', '万亿']
        int_str = ''
        for gi in range(len(groups) - 1, -1, -1):
            g = groups[gi]
            if g == 0:
                continue
            s = ''
            for i in range(3, -1, -1):
                d = g // (10 ** i) % 10
                if d == 0:
                    if s and not s.endswith('零'):
                        s += '零'
                else:
                    s += digits[d] + ['', '拾', '佰', '仟'][i]
            s = s.rstrip('零')
            int_str += s + big[gi]
        int_str = int_str.rstrip('零')

    result = '人民币' + int_str + '圆'
    if jiao == 0 and fen_unit == 0:
        result += '整'
    else:
        if jiao == 0:
            result += '零'
        else:
            result += digits[jiao] + '角'
        if fen_unit == 0:
            result += '整'
        else:
            result += digits[fen_unit] + '分'
    return result[1:]


def read_q3(path):
    """读取汇总表，按 (法人姓名, 身份证号) 和 身份证号 建索引。

    汇总表列：B=企业名称 C=税号 D=法人姓名 E=身份证号
    """
    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[0]
    by_key, by_idcard = {}, {}
    for r in range(2, ws.max_row + 1):
        d, e = ws.cell(r, 4).value, ws.cell(r, 5).value
        if d in (None, '') or e in (None, ''):
            continue
        by_key.setdefault((str(d).strip(), str(e).strip()), r)
        by_idcard.setdefault(str(e).strip(), r)
    return ws, by_key, by_idcard


def build(jpg_dir, q3_path, out_path, verbose):
    warns = []
    rows = extract_rows(jpg_dir, warns)
    q3_ws, by_key, by_idcard = read_q3(q3_path)

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
        src_r = by_key.get((str(row['name']).strip(), str(row['idcard']).strip())) \
            if row['name'] and row['idcard'] else None
        if src_r is None and row['idcard']:
            src_r = by_idcard.get(str(row['idcard']).strip())

        if src_r is None:
            unmatched.append(row['name'])
            warns.append('第 %d 行「%s」在汇总表中未匹配到，企业名称/税号/法人留空，F~J 列留空'
                         % (i + 1, row['name']))
            vals = [i + 1, row['name'], row['idcard'], None, None]
        else:
            vals = [
                i + 1,
                q3_ws.cell(src_r, 2).value,      # 企业名称
                q3_ws.cell(src_r, 3).value,      # 税号
                q3_ws.cell(src_r, 4).value,      # 法人 姓名
                None,                            # 二维码的值（留空）
            ]

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
    # 补齐 H/I/J 列：大写金额与数字金额均取自汇总表 N 列（所得税）
    for i, row in enumerate(rows):
        if row['src_r'] is None:
            continue
        r = i + 2
        amt = q3_ws.cell(row['src_r'], 14).value  # N 列 = 所得税
        if amt is None:
            continue
        amt_f = float(amt)
        if sh.cell(r, 9).value is None:    # I 列 金额1
            sh.cell(r, 9).value = amt_f
        if sh.cell(r, 10).value is None:   # J 列 金额2
            sh.cell(r, 10).value = amt_f
        if sh.cell(r, 8).value is None:    # H 列 中文大写（去'人'）
            sh.cell(r, 8).value = num_to_chinese(amt_f)
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
    print('数据行: %d | 已匹配汇总表: %d | 未匹配: %d'
          % (len(rows), len(rows) - len(unmatched), len(unmatched)))
    if unmatched:
        print('未匹配汇总表: %s' % '、'.join(str(x) for x in unmatched))
    for w in warns:
        print('提示: ' + w)
    return len(rows)


def main():
    ap = argparse.ArgumentParser(description='完税证明 JPG + 季度汇总表 -> jpg.xlsx')
    ap.add_argument('--jpg-dir', default=None, help='存放 JPG 的目录（默认脚本所在目录）')
    ap.add_argument('--q3', default=None, help='季度汇总表 xlsx（默认自动查找 *季度*.xlsx）')
    ap.add_argument('--out', default=None, help='输出文件（默认 <jpg-dir>/jpg.xlsx）')
    ap.add_argument('-v', '--verbose', action='store_true', help='打印详细提示')
    a = ap.parse_args()

    jpg_dir = os.path.abspath(a.jpg_dir or os.path.dirname(os.path.abspath(__file__)))
    if not os.path.isdir(jpg_dir):
        die('目录不存在: %s' % jpg_dir)
    q3 = find_q3(jpg_dir, a.q3)
    out = os.path.abspath(a.out or os.path.join(jpg_dir, 'jpg.xlsx'))
    build(jpg_dir, q3, out, a.verbose)


if __name__ == '__main__':
    main()