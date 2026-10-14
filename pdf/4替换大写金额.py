# -*- coding: utf-8 -*-
"""
大写金额替换脚本
读取 pdf.xlsx 的 H列(中文金额-楷体)，按 C列(税号) 匹配 PDF，
用华文楷体 STKAITI.TTF 10.5pt 替换 PDF 中的大写金额文本。
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
FONT_PATH = os.path.join(BASE, "STKAITI.TTF")
if not os.path.exists(FONT_PATH):
    FONT_PATH = r"C:\Windows\Fonts\STKAITI.TTF"

AMOUNT_ORIGIN = (183.85, 412.78)
AMOUNT_SIZE = 10.5


def load_amount_map(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    result = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 8:
            continue
        tax_no = row[2]
        amount = row[7]
        if tax_no and amount:
            result[str(tax_no).strip()] = str(amount).strip()
    return result


def get_tax_id(doc):
    text = doc[0].get_text()
    m = re.search(r"纳税人识别号[:\s]*([A-Z0-9]{15,20})", text)
    return m.group(1).strip() if m else None


def find_amount_span(page):
    d = page.get_text("dict")
    for block in d["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                if "大写" in span["text"]:
                    return span
    return None


def replace_amount(doc, new_text):
    page = doc[0]
    span = find_amount_span(page)
    if not span:
        return False

    bbox = fitz.Rect(span["bbox"])
    page.add_redact_annot(bbox, fill=(1, 1, 1))
    page.apply_redactions()

    page.insert_font(fontname="STKAITI", fontfile=FONT_PATH)
    page.insert_text(
        point=AMOUNT_ORIGIN,
        text=new_text,
        fontname="STKAITI",
        fontsize=AMOUNT_SIZE,
        color=(0, 0, 0),
    )
    return True


def main():
    xlsx_path = os.path.join(BASE, XLSX_NAME)
    if not os.path.exists(xlsx_path):
        print("找不到 %s" % XLSX_NAME)
        return

    amount_map = load_amount_map(xlsx_path)
    print("Excel 中共 %d 条税号->金额映射" % len(amount_map))

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

        h_value = amount_map[tax_id]
        new_text = "（大写）人" + h_value

        try:
            changed = replace_amount(doc, new_text)
        except Exception as e:
            print("[失败] %s: %s" % (name, e))
            fail += 1
            doc.close()
            continue

        if not changed:
            print("[失败] %s: 未找到大写金额" % name)
            fail += 1
            doc.close()
            continue

        doc.subset_fonts()
        doc.save(dst, garbage=4, deflate=True, deflate_fonts=True)
        doc.close()
        print("[完成] %s: %s" % (name, new_text))
        ok += 1

    print("\n总计: 完成 %d, 跳过 %d, 失败 %d, 共 %d 个PDF" % (ok, skip, fail, len(pdfs)))


if __name__ == "__main__":
    main()