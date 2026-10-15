# -*- coding: utf-8 -*-
"""
数字金额替换脚本
替换两处数字金额：
  1. 表格中"实缴(退)金额"下方的数字（无¥，带千分位逗号）
  2. "金额合计"行右侧的 ¥数字（有¥，无逗号）
读取 pdf.xlsx 的 F列(所得税)，按 C列(税号) 匹配 PDF，
用 金额.ttf 替换，靠右对齐。
字号通过测量原PDF实际渲染高度后缩放，确保落地大小与原件一模一样。
输出到 "已完成" 文件夹。
"""
import os
import re
import sys
import glob

import openpyxl
import fitz

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX_NAME = "pdf.xlsx"
DONE_NAME = "已完成"
FONT_PATH = os.path.join(BASE, "金额.ttf")

RIGHT_EDGE = 714.75

TABLE_Y_MIN = 250
TABLE_Y_MAX = 270

TOTAL_Y_MIN = 395
TOTAL_Y_MAX = 420


def load_amount_map(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    result = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 6:
            continue
        tax_no = row[2]
        amount = row[5]
        if tax_no and amount is not None:
            result[str(tax_no).strip()] = float(amount)
    return result


def get_tax_id(doc):
    text = doc[0].get_text()
    m = re.search(r"纳税人识别号[:\s]*([A-Z0-9]{15,20})", text)
    if m:
        return m.group(1).strip()
    page = doc[0]
    seq = []
    for block in page.get_text("dict")["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                seq.append(span)
    label = None
    for s in seq:
        if s["text"].strip() == "纳税人识别号":
            label = s
            break
    if not label:
        return None
    lx1 = label["bbox"][2]
    ly = label["origin"][1]
    for s in seq:
        if s["origin"][0] > lx1 and abs(s["origin"][1] - ly) < 20:
            m = re.search(r"([0-9A-Za-z]{15,20})", s["text"])
            if m:
                return m.group(1).upper()
    return None


def find_spans(page):
    d = page.get_text("dict")
    table_span = None
    yen_span = None
    total_num_span = None
    for block in d["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                t = span["text"]
                y = span["bbox"][1]
                if TABLE_Y_MIN < y < TABLE_Y_MAX:
                    if re.match(r"^[\d,.]+$", t) and len(t) >= 3:
                        table_span = span
                if TOTAL_Y_MIN < y < TOTAL_Y_MAX:
                    if "¥" in t:
                        yen_span = span
                    elif re.match(r"^[\d,.]+$", t) and len(t) >= 3:
                        total_num_span = span
    return table_span, yen_span, total_num_span


def measure_height(fontfile, fontsize, text="0"):
    tmp = fitz.open()
    p = tmp.new_page()
    p.insert_text(point=(0, 50), text=text, fontname="F",
                  fontfile=fontfile, fontsize=fontsize)
    d = p.get_text("dict")
    for block in d["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                return span["bbox"][3] - span["bbox"][1]
    return fontsize


def calc_size(orig_span, fontfile):
    orig_size = orig_span["size"]
    orig_h = orig_span["bbox"][3] - orig_span["bbox"][1]
    new_h = measure_height(fontfile, orig_size)
    if new_h > 0:
        return orig_size * (orig_h / new_h)
    return orig_size


def replace_num_amount(doc, amount_val):
    page = doc[0]
    table_span, yen_span, total_num_span = find_spans(page)
    if not table_span and not yen_span and not total_num_span:
        return False

    font = fitz.Font(fontfile=FONT_PATH)

    table_size = None
    total_size = None
    table_origin_y = None
    total_origin_y = None

    if table_span:
        table_size = calc_size(table_span, FONT_PATH)
        table_origin_y = table_span["origin"][1]

    ref_span = total_num_span if total_num_span else yen_span
    if ref_span:
        total_size = calc_size(ref_span, FONT_PATH)
        total_origin_y = ref_span["origin"][1]

    rects = []
    if table_span:
        rects.append(fitz.Rect(table_span["bbox"]))
    if yen_span:
        rects.append(fitz.Rect(yen_span["bbox"]))
    if total_num_span:
        rects.append(fitz.Rect(total_num_span["bbox"]))

    for r in rects:
        page.add_redact_annot(r, fill=(1, 1, 1))
    page.apply_redactions()

    if table_span:
        table_text = "{:,.2f}".format(amount_val)
        w = font.text_length(table_text, fontsize=table_size)
        x = RIGHT_EDGE - w
        page.insert_text(
            point=(x, table_origin_y),
            text=table_text,
            fontname="AmtFont",
            fontfile=FONT_PATH,
            fontsize=table_size,
            color=(0, 0, 0),
        )

    if total_size:
        total_text = "¥%.2f" % amount_val
        w = font.text_length(total_text, fontsize=total_size)
        x = RIGHT_EDGE - w
        page.insert_text(
            point=(x, total_origin_y),
            text=total_text,
            fontname="AmtFont",
            fontfile=FONT_PATH,
            fontsize=total_size,
            color=(0, 0, 0),
        )
    return True


def main():
    xlsx_path = os.path.join(BASE, XLSX_NAME)
    if not os.path.exists(xlsx_path):
        print("找不到 %s" % XLSX_NAME)
        return

    amount_map = load_amount_map(xlsx_path)
    print("Excel 中共 %d 条税号->数字金额映射" % len(amount_map))

    done_dir = os.path.join(BASE, DONE_NAME)
    os.makedirs(done_dir, exist_ok=True)

    pdfs = sorted([f for f in os.listdir(BASE) if f.lower().endswith(".pdf")])
    ok = skip = fail = 0

    for name in pdfs:
        src = os.path.join(BASE, name)
        dst = os.path.join(done_dir, name)
        doc = fitz.open(src)

        tax_id = get_tax_id(doc)
        if not tax_id or tax_id not in amount_map:
            print("[跳过] %s: 税号 %s 未匹配" % (name, tax_id))
            skip += 1
            doc.close()
            continue

        amount_val = amount_map[tax_id]

        try:
            changed = replace_num_amount(doc, amount_val)
        except Exception as e:
            print("[失败] %s: %s" % (name, e))
            fail += 1
            doc.close()
            continue

        if not changed:
            print("[失败] %s: 未找到数字金额" % name)
            fail += 1
            doc.close()
            continue

        doc.subset_fonts()
        doc.save(dst, garbage=4, deflate=True, deflate_fonts=True)
        doc.close()
        print("[完成] %s: %.2f" % (name, amount_val))
        ok += 1

    print("\n总计: 完成 %d, 跳过 %d, 失败 %d, 共 %d 个PDF" % (ok, skip, fail, len(pdfs)))


if __name__ == "__main__":
    main()