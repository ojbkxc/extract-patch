# -*- coding: utf-8 -*-
"""
完税证明 PDF 凭证号码修正工具
================================================================
根据 pdf.xlsx 中"二维码的值"列(E列)第 3 项解析目标完税凭证号码，
按纳税人识别号(C列)匹配同目录下的 PDF，修正 No● 后的凭证号码。

原理：直接修改 PDF 内容流中的号码 TJ 数组（就地替换数字字符），
不改动字体引用 / 字号 / 字距 / 位置，字体与原件 100% 一致。
结果另存到 "已完成" 文件夹，原件不修改。

用法：把本脚本放到 PDF 和 pdf.xlsx 同目录，运行 python fix_certno.py
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


def load_cert_map(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    result = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 5:
            continue
        tax_no = row[2]
        qr_value = row[4]
        if not tax_no or not qr_value:
            continue
        parts = [p.strip() for p in str(qr_value).split(",")]
        if len(parts) > 2 and parts[2]:
            result[str(tax_no).strip()] = parts[2]
    return result


def get_pdf_info(doc):
    text = doc[0].get_text("text")
    m = re.search(r"No●(\d+)", text)
    cur_no = m.group(1) if m else None
    m2 = re.search(r"纳税人识别号\s*\n?\s*([0-9A-Za-z]+)", text)
    tax_no = m2.group(1) if m2 else None
    if not tax_no:
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
        if label:
            lx1 = label["bbox"][2]
            ly = label["origin"][1]
            for s in seq:
                if s["origin"][0] > lx1 and abs(s["origin"][1] - ly) < 20:
                    m3 = re.search(r"([0-9A-Za-z]{15,20})", s["text"])
                    if m3:
                        tax_no = m3.group(1).upper()
                        break
    return cur_no, tax_no


def _replace_in_stream(stream, old_b, new_b):
    matches = []
    for m in re.finditer(rb"\[([^\]]*)\]\s*TJ", stream):
        body = m.group(1)
        segs = [(mm.start(1), mm.end(1), mm.group(1))
                for mm in re.finditer(rb"\(([^)]*)\)", body)]
        if not segs:
            continue
        plain = b"".join(s for _, _, s in segs)
        if plain != old_b:
            continue
        new_body = bytearray(body)
        pos = 0
        repl = []
        for s, e, seg in segs:
            n = len(seg)
            repl.append((s, e, new_b[pos:pos + n]))
            pos += n
        for s, e, nb in reversed(repl):
            new_body[s:e] = nb
        matches.append((m.start(1), m.end(1), bytes(new_body)))
    if not matches:
        return None
    buf = bytearray(stream)
    for s, e, nb in reversed(matches):
        buf[s:e] = nb
    return bytes(buf)


def replace_cert_no(doc, old_no, new_no):
    if len(old_no) != len(new_no):
        raise ValueError("新旧号码长度不同，无法就地替换")
    old_b = old_no.encode("latin-1")
    new_b = new_no.encode("latin-1")
    changed = False
    for page in doc:
        for xref in page.get_contents():
            stream = doc.xref_stream(xref)
            new_stream = _replace_in_stream(stream, old_b, new_b)
            if new_stream is not None:
                doc.update_stream(xref, new_stream)
                changed = True
    return changed


def main():
    xlsx_path = os.path.join(BASE, XLSX_NAME)
    if not os.path.exists(xlsx_path):
        print("找不到 %s，请把脚本放到 PDF 所在目录" % XLSX_NAME)
        return

    cert_map = load_cert_map(xlsx_path)
    done_dir = os.path.join(BASE, DONE_NAME)
    os.makedirs(done_dir, exist_ok=True)

    pdfs = sorted([f for f in os.listdir(BASE) if f.lower().endswith(".pdf")])
    ok = skip = fail = 0
    for name in pdfs:
        src = os.path.join(BASE, name)
        dst = os.path.join(done_dir, name)
        try:
            doc = fitz.open(src)
        except Exception as e:
            print("[失败] %s: 无法打开 %s" % (name, e))
            fail += 1
            continue

        cur_no, tax_no = get_pdf_info(doc)
        if not cur_no or not tax_no:
            print("[失败] %s: 未找到凭证号码或纳税人识别号" % name)
            fail += 1
            doc.close()
            continue

        target = cert_map.get(tax_no)
        if not target:
            print("[失败] %s: xlsx 中无匹配税号 %s" % (name, tax_no))
            fail += 1
            doc.close()
            continue

        if cur_no == target:
            doc.save(dst, garbage=4, deflate=True)
            doc.close()
            print("[跳过] %s: 号码已一致 %s" % (name, cur_no))
            skip += 1
            continue

        try:
            changed = replace_cert_no(doc, cur_no, target)
        except Exception as e:
            print("[失败] %s: %s" % (name, e))
            fail += 1
            doc.close()
            continue

        if not changed:
            print("[失败] %s: 内容流未找到号码 %s" % (name, cur_no))
            fail += 1
            doc.close()
            continue

        doc.save(dst, garbage=4, deflate=True)
        doc.close()
        print("[完成] %s: %s -> %s" % (name, cur_no, target))
        ok += 1

    print("\n总计: 完成 %d, 跳过 %d, 失败 %d, 共 %d 个 PDF" % (ok, skip, fail, len(pdfs)))
    print("输出目录: %s" % done_dir)


if __name__ == "__main__":
    main()