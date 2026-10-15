# -*- coding: utf-8 -*-
"""
纳税人识别号替换脚本
按 pdf.xlsx A列(序号) 与文件名开头数字匹配，读取 C列(税号) 作为目标值，
对比 PDF 中"纳税人识别号"字段，不一致的才用 纳税人识别号.ttf 重绘。
字号通过测量原 PDF 实际渲染高度后缩放，确保落地大小与原件一模一样。
输出到 "已完成" 文件夹。
"""
import os
import re
import sys

import openpyxl
import pymupdf as fitz

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX_NAME = "pdf.xlsx"
DONE_NAME = "已完成"
FONT_PATH = os.path.join(BASE, "纳税人识别号.ttf")
LABEL = "纳税人识别号"
TAXID_RE = re.compile(r"^[0-9A-Z]{18}$")


def file_index(name):
    m = re.match(r"\s*(\d+)", name)
    return int(m.group(1)) if m else None


def load_tax_map(xlsx_path):
    """序号(A列) -> 税号(C列)"""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    result = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 3:
            continue
        idx, tax = row[0], row[2]
        if idx is None or tax is None:
            continue
        try:
            result[int(idx)] = str(tax).strip().upper()
        except (TypeError, ValueError):
            continue
    return result


def find_taxid_span(page):
    """返回"纳税人识别号"标签下一段文本的 span（税号）"""
    seq = []
    for block in page.get_text("dict")["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                seq.append(span)
    for i, s in enumerate(seq):
        if s["text"].strip() == LABEL and i + 1 < len(seq):
            nxt = seq[i + 1]
            if TAXID_RE.match(nxt["text"].strip()):
                return nxt
    return None


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
                h = span["bbox"][3] - span["bbox"][1]
                tmp.close()
                return h
    tmp.close()
    return fontsize


def calc_size(orig_span, fontfile):
    orig_size = orig_span["size"]
    orig_h = orig_span["bbox"][3] - orig_span["bbox"][1]
    new_h = measure_height(fontfile, orig_size)
    if new_h > 0:
        return orig_size * (orig_h / new_h)
    return orig_size


def replace_taxid(doc, new_tax):
    """对比并替换税号。返回 (旧税号, 是否修改)。
    未找到税号返回 (None, False)。"""
    page = doc[0]
    span = find_taxid_span(page)
    if span is None:
        return None, False
    old = span["text"].strip().upper()
    if old == new_tax:
        return old, False

    size = calc_size(span, FONT_PATH)
    origin = span["origin"]
    rect = fitz.Rect(span["bbox"])
    page.add_redact_annot(rect, fill=(1, 1, 1))
    page.apply_redactions()
    page.insert_text(
        point=(origin[0], origin[1]),
        text=new_tax,
        fontname="TaxFont",
        fontfile=FONT_PATH,
        fontsize=size,
        color=(0, 0, 0),
    )
    return old, True


def main():
    xlsx_path = os.path.join(BASE, XLSX_NAME)
    if not os.path.exists(xlsx_path):
        print("找不到 %s" % XLSX_NAME)
        return
    if not os.path.exists(FONT_PATH):
        print("找不到 %s" % os.path.basename(FONT_PATH))
        return

    tax_map = load_tax_map(xlsx_path)
    print("Excel 中共 %d 条序号->税号映射" % len(tax_map))

    done_dir = os.path.join(BASE, DONE_NAME)
    os.makedirs(done_dir, exist_ok=True)

    pdfs = sorted([f for f in os.listdir(BASE) if f.lower().endswith(".pdf")])
    ok = skip = fail = 0

    for name in pdfs:
        src = os.path.join(BASE, name)
        dst = os.path.join(done_dir, name)
        fi = file_index(name)
        target = tax_map.get(fi)
        if target is None:
            print("[跳过] %s: 序号 %s 未匹配 xlsx" % (name, fi))
            skip += 1
            continue

        doc = fitz.open(src)
        try:
            old, changed = replace_taxid(doc, target)
        except Exception as e:
            print("[失败] %s: %s" % (name, e))
            fail += 1
            doc.close()
            continue

        if old is None:
            print("[失败] %s: 未找到纳税人识别号" % name)
            fail += 1
            doc.close()
            continue

        doc.subset_fonts()
        doc.save(dst, garbage=4, deflate=True, deflate_fonts=True)
        doc.close()

        if changed:
            print("[完成] %s: %s -> %s" % (name, old, target))
            ok += 1
        else:
            print("[跳过] %s: 税号已一致 %s" % (name, old))
            skip += 1

    print("\n总计: 修改 %d, 跳过 %d, 失败 %d, 共 %d 个PDF" % (ok, skip, fail, len(pdfs)))
    print("输出目录: %s" % done_dir)


if __name__ == "__main__":
    main()