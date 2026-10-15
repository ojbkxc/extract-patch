# -*- coding: utf-8 -*-
"""
纳税人名称替换脚本
按 pdf.xlsx A列(序号) 与文件名开头数字匹配，读取 B列(企业名称)，
追加"（个体工商户）"后缀，用华文楷体 STKAITI.TTF 10.5pt 替换 PDF 中的纳税人名称。

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
FONT_PATH = os.path.join(BASE, "STKAITI.TTF")
if not os.path.exists(FONT_PATH):
    FONT_PATH = r"C:\Windows\Fonts\STKAITI.TTF"
LABEL = "纳税人名称"
SUFFIX = "（个体工商户）"


def file_index(name):
    m = re.match(r"\s*(\d+)", name)
    return int(m.group(1)) if m else None


def load_name_map(xlsx_path):
    """序号(A列) -> 企业名称(B列)"""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    result = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 2:
            continue
        idx, name = row[0], row[1]
        if idx is None or name is None:
            continue
        try:
            result[int(idx)] = str(name).strip()
        except (TypeError, ValueError):
            continue
    return result


def find_name_spans(page):
    """返回"纳税人名称"标签后的企业名称 span 列表（可能跨多行）。

    名称区域在页面右侧(origin x > 400)，遇到左侧字段(x < 400)说明结束。
    """
    seq = []
    for block in page.get_text("dict")["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                seq.append(span)
    for i, s in enumerate(seq):
        if s["text"].strip() == LABEL:
            name_spans = []
            for j in range(i + 1, len(seq)):
                sp = seq[j]
                t = sp["text"].strip()
                if not t:
                    continue
                if sp["origin"][0] < 400:
                    break
                name_spans.append(sp)
            return name_spans if name_spans else None
    return None



def split_lines(text, font, fontsize, max_width):
    """按最大宽度拆分文本为多行，逐字符测量累积宽度。"""
    lines = []
    cur = ""
    for ch in text:
        test = cur + ch
        w = font.text_length(test, fontsize=fontsize)
        if w > max_width and cur:
            lines.append(cur)
            cur = ch
        else:
            cur = test
    if cur:
        lines.append(cur)
    return lines


def replace_name(doc, new_name_full):
    """替换纳税人名称。返回 (旧名称, 是否修改)。"""
    page = doc[0]
    spans = find_name_spans(page)
    if not spans:
        return None, False

    old_text = "".join(s["text"] for s in spans).strip()
    if old_text == new_name_full:
        return old_text, False

    size = 10.5
    origin0 = spans[0]["origin"]
    max_x = max(s["bbox"][2] for s in spans)
    avail_width = max_x - origin0[0]

    line_height = (spans[1]["origin"][1] - spans[0]["origin"][1]
                   if len(spans) > 1 else 11.45)

    for s in spans:
        rect = fitz.Rect(s["bbox"])
        page.add_redact_annot(rect, fill=(1, 1, 1))
    page.apply_redactions()

    page.insert_font(fontname="STKAITI", fontfile=FONT_PATH)

    font = fitz.Font(fontfile=FONT_PATH)
    lines = split_lines(new_name_full, font, size, avail_width)
    for k, line in enumerate(lines):
        y = origin0[1] + k * line_height
        page.insert_text(
            point=(origin0[0], y),
            text=line,
            fontname="STKAITI",
            fontsize=size,
            color=(0, 0, 0),
        )

    return old_text, True


def main():
    xlsx_path = os.path.join(BASE, XLSX_NAME)
    if not os.path.exists(xlsx_path):
        print("找不到 %s" % XLSX_NAME)
        return
    if not os.path.exists(FONT_PATH):
        print("找不到 STKAITI.TTF")
        return

    name_map = load_name_map(xlsx_path)
    print("Excel 中共 %d 条序号->企业名称映射" % len(name_map))

    done_dir = os.path.join(BASE, DONE_NAME)
    os.makedirs(done_dir, exist_ok=True)

    pdfs = sorted([f for f in os.listdir(BASE) if f.lower().endswith(".pdf")])
    ok = skip = fail = 0

    for name in pdfs:
        src = os.path.join(BASE, name)
        dst = os.path.join(done_dir, name)
        fi = file_index(name)
        target = name_map.get(fi)
        if target is None:
            print("[跳过] %s: 序号 %s 未匹配 xlsx" % (name, fi))
            skip += 1
            continue

        new_name_full = target + SUFFIX

        doc = fitz.open(src)
        try:
            old, changed = replace_name(doc, new_name_full)
        except Exception as e:
            print("[失败] %s: %s" % (name, e))
            fail += 1
            doc.close()
            continue

        if old is None:
            print("[失败] %s: 未找到纳税人名称" % name)
            fail += 1
            doc.close()
            continue

        doc.subset_fonts()
        doc.save(dst, garbage=4, deflate=True, deflate_fonts=True)
        doc.close()

        if changed:
            print("[完成] %s: %s -> %s" % (name, old, new_name_full))
            ok += 1
        else:
            print("[跳过] %s: 名称已一致 %s" % (name, old))
            skip += 1

    print("\n总计: 修改 %d, 跳过 %d, 失败 %d, 共 %d 个PDF" % (ok, skip, fail, len(pdfs)))
    print("输出目录: %s" % done_dir)


if __name__ == "__main__":
    main()