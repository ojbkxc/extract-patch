#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""从完税证明 JPG (RapidOCR 识别) + 季度汇总表(xlsx) 生成 jpg.xlsx。

用法：
    python 1读取jpg创建表格.py [--jpg-dir DIR] [--q3 FILE] [--out FILE] [-v]

默认值：
    --jpg-dir  脚本所在目录（存放 *.jpg）
    --q3       在 --jpg-dir 及其上级目录自动查找 *季度*.xlsx；命中多个时需显式指定
    --out      <jpg-dir>/jpg.xlsx

依赖（Python 3.11）：
    pip install rapidocr-onnxruntime==1.4.4 "onnxruntime==1.20.1" "opencv-python==4.10.0.84" openpyxl

OCR 识别策略：
    1. 用 RapidOCR 识别每张 JPG 的全部文字
    2. 检测标题「完税证明」——识别到标题说明 OCR 成功，继续提取字段
    3. 从 OCR 文本解析：纳税人识别号(18位)、纳税人名称、凭证号(No.)、实缴金额(¥)
    4. 用纳税人识别号匹配汇总表 E 列（身份证号）

依次完成的步骤：
    1. 扫描 --jpg-dir 下 *.jpg，按文件名开头的数字排序
    2. RapidOCR 识别每张图片，提取纳税人识别号/名称/凭证号/金额
    3. 按识别号从 --q3 表 D/E 列匹配汇总行
    4. 企业名称、税号、法人姓名：取自汇总表 B/C/D 列
    5. N/Q/R/S/T 五列：按匹配行复制值并完整复制格式，落在 F~J 列
    6. E 列「二维码的值」留空；K 列格式与 G 列一致

安全约束：
    * 只读输入、只写 --out，输入文件永不被修改
    * 先写 .tmp 再 os.replace 原子替换；任何异常都会清理 .tmp
    * 依赖缺失 / 无 JPG => 报错退出，不产生输出文件
"""
from __future__ import annotations

import argparse
import glob
import os
import re
import sys
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
TITLE_KEYWORD = '完税证明'
_LABEL_KEYS = ('识别号', '凭证号', '税种', '品目', '所属', '金额', '日期',
               '机关', '填发', '原凭', '入（', '实缴', '合计', '票人', '盖章',
               '备注', '保管', '专用', 'No', '¥', '￥', '人民币', '圆', '角', '分',
               '中华', '和国', '税', '证', '明', '填')


def die(msg):
    sys.exit('[中止] ' + msg)


def file_order_key(name):
    m = re.match(r'\s*(\d+)', name)
    return (0, int(m.group(1))) if m else (1, name)


def copy_style(src, dst):
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


_rapid_engine = None


def get_rapid_engine():
    global _rapid_engine
    if _rapid_engine is not None:
        return _rapid_engine
    try:
        from rapidocr_onnxruntime import RapidOCR
    except ImportError:
        die('缺少 rapidocr-onnxruntime，请安装：\n'
            '  pip install rapidocr-onnxruntime==1.4.4 "onnxruntime==1.20.1" "opencv-python==4.10.0.84"')
    _rapid_engine = RapidOCR()
    return _rapid_engine


def ocr_image(img_path, warns):
    eng = get_rapid_engine()
    try:
        result, elapse = eng(img_path)
    except Exception as e:
        warns.append('%s：OCR 识别异常 %s' % (os.path.basename(img_path), str(e)[:80]))
        return {}
    if not result:
        warns.append('%s：OCR 无结果' % os.path.basename(img_path))
        return {}

    texts = [item[1] for item in result]
    all_text = '\n'.join(texts)
    _title_frags = ('中华', '和国', '税收', '证', '明')
    has_title = sum(1 for k in _title_frags if k in all_text) >= 3

    taxpayer_id = None
    for i, t in enumerate(texts):
        if '识别号' in t:
            for j in range(i + 1, min(i + 4, len(texts))):
                m = re.search(r'(\d{17}[\dXx])', texts[j])
                if m:
                    taxpayer_id = m.group(1).upper()
                    break
            if taxpayer_id:
                break

    m_no = re.search(r'No\.?\s*[:：]?\s*(\d{15,20})', all_text)
    voucher_no = m_no.group(1) if m_no else None

    m_amt = re.search(r'[¥￥]\s*([\d,]+\.?\d*)', all_text)
    amount = m_amt.group(1).replace(',', '') if m_amt else None

    taxpayer_name = None
    for i, t in enumerate(texts):
        if '名称' in t and '税' in t:
            for j in range(i + 1, min(i + 4, len(texts))):
                nxt = texts[j].strip()
                if nxt and not any(k in nxt for k in _LABEL_KEYS):
                    taxpayer_name = nxt
                    break
            break

    return dict(has_title=has_title, taxpayer_name=taxpayer_name,
                taxpayer_id=taxpayer_id, voucher_no=voucher_no, amount=amount,
                _texts=texts)


def extract_rows(jpg_dir, warns, verbose):
    files = [f for f in os.listdir(jpg_dir) if f.lower().endswith('.jpg')]
    if not files:
        die('目录中没有 JPG 文件: %s' % jpg_dir)
    files.sort(key=file_order_key)

    rows = []
    for f in files:
        info = ocr_image(os.path.join(jpg_dir, f), warns)
        if verbose:
            print('OCR %s -> 标题=%s 名称=%s 识别号=%s 凭证号=%s 金额=%s'
                  % (f, info.get('has_title'), info.get('taxpayer_name'),
                     info.get('taxpayer_id'), info.get('voucher_no'), info.get('amount')))
        if not info.get('has_title'):
            warns.append('%s：未识别到标题「%s」，OCR 可能失败' % (f, TITLE_KEYWORD))
        rows.append(dict(file=f, **info))
    return rows


def read_q3(path):
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
    rows = extract_rows(jpg_dir, warns, verbose)
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
        name = row.get('taxpayer_name')
        idcard = row.get('taxpayer_id')
        src_r = by_key.get((str(name).strip(), str(idcard).strip())) \
            if name and idcard else None
        if src_r is None and idcard:
            src_r = by_idcard.get(str(idcard).strip())

        if src_r is None:
            unmatched.append(name or idcard or row['file'])
            warns.append('第 %d 行「%s」在汇总表中未匹配到，企业名称/税号/法人留空，F~J 列留空'
                         % (i + 1, name or idcard or row['file']))
            vals = [i + 1, name, idcard, None, None]
        else:
            vals = [
                i + 1,
                q3_ws.cell(src_r, 2).value,
                q3_ws.cell(src_r, 3).value,
                q3_ws.cell(src_r, 4).value,
                None,
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
    ap = argparse.ArgumentParser(description='完税证明 JPG(RapidOCR) + 季度汇总表 -> jpg.xlsx')
    ap.add_argument('--jpg-dir', default=None, help='存放 JPG 的目录（默认脚本所在目录）')
    ap.add_argument('--q3', default=None, help='季度汇总表 xlsx（默认自动查找 *季度*.xlsx）')
    ap.add_argument('--out', default=None, help='输出文件（默认 <jpg-dir>/jpg.xlsx）')
    ap.add_argument('-v', '--verbose', action='store_true', help='打印 OCR 详细结果')
    a = ap.parse_args()

    jpg_dir = os.path.abspath(a.jpg_dir or os.path.dirname(os.path.abspath(__file__)))
    if not os.path.isdir(jpg_dir):
        die('目录不存在: %s' % jpg_dir)
    q3 = find_q3(jpg_dir, a.q3)
    out = os.path.abspath(a.out or os.path.join(jpg_dir, 'jpg.xlsx'))
    build(jpg_dir, q3, out, a.verbose)


if __name__ == '__main__':
    main()