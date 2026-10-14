#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""extract-patch 项目依赖一键安装

用法：
    python install_deps.py
    python install_deps.py --python D:\\python\\Python313\\python.exe
    python install_deps.py --index-url https://pypi.tuna.tsinghua.edu.cn/simple
    python install_deps.py --skip-verify

依赖清单（从各脚本 import 归纳）：
    jpg/1读取ocr创建表格.py       openpyxl + RapidOCR
    jpg/0读取jpg文件名创建表格.py   openpyxl
    pdf/1读取pdf创建表格.py        openpyxl pymupdf pillow zxing-cpp
    pdf/2换票号.py                openpyxl pymupdf
    pdf/3换二维码.py               openpyxl qrcode pymupdf
    pdf/4替换大写金额.py           openpyxl pymupdf
    pdf/5替换两个数字金额.py        openpyxl pymupdf
    pdf/0生成金额.ttf.py           pymupdf fonttools

两条硬约束：
    1. onnxruntime 一律钉 1.20.1 —— 1.24 及以上的新版本在部分机器上
       import 即崩溃(0xC0000005)，实测确认过。
    2. 旧包 rapidocr-onnxruntime 在 PyPI 上限定 Python <3.13，
       所以 OCR 依赖按目标 Python 版本二选一：
         <=3.12 -> rapidocr-onnxruntime==1.4.4（与现有脚本 import 一致）
         >=3.13 -> rapidocr==3.10.0（包名不同，脚本需改 import，见结尾提示）

验证跑在【独立子进程】里：崩溃是硬性访问冲突，会把进程直接干掉，
不能让安装器自己 import onnxruntime。
"""

import argparse
import json
import os
import subprocess
import sys
import textwrap

# ---------------------------------------------------------------- 依赖清单

# 与 Python 版本无关的通用依赖（六个库，8 个脚本共用）
COMMON_SPECS = [
    "openpyxl",
    "pymupdf",
    "pillow",
    "zxing-cpp",
    "qrcode",
    "fonttools",
]

ONNXRUNTIME_PIN = "1.20.1"

# <=3.12：旧包，与项目现有 import 完全一致
LEGACY_OCR_SPECS = [
    "rapidocr-onnxruntime==1.4.4",
    "opencv-python==4.10.0.84",
]
# >=3.13：新包（旧包在该版本上没有可安装的发行版）
NEW_OCR_SPECS = [
    "rapidocr==3.10.0",
]

# onnxruntime 1.20.1 的 wheel 覆盖 cp310~cp313
SUPPORTED_MIN = (3, 10)
SUPPORTED_MAX = (3, 13)
SPLIT_AT = (3, 13)          # >= 此版本走新 OCR 包

COMMON_PATHS = [
    r"D:\python\Python313\python.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Python\Python311\python.exe"),
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Python\Python313\python.exe"),
    r"C:\Python313\python.exe",
    r"C:\Python311\python.exe",
]

PROBE_CODE = (
    "import sys, json;"
    'print("__PROBE__" + json.dumps({"exe": sys.executable,'
    '"ver": list(sys.version_info[:3])}))'
)

# 独立子进程验证：逐个导入项目所需的全部第三方库
CHECK_CODE = textwrap.dedent(
    """
    import json
    from importlib.metadata import version
    out = {}

    # onnxruntime 放最前：一旦本机不兼容，会立刻崩在这里（父进程能捕获）
    import onnxruntime as ort
    out["onnxruntime"] = ort.__version__

    import numpy as np
    import openpyxl
    import pymupdf
    import PIL
    import zxingcpp
    import qrcode
    import fontTools
    import fitz          # 项目里 4 个脚本用的是 fitz 别名
    out["fitz_alias"] = True

    for mod, pkg in (("openpyxl", "openpyxl"), ("pymupdf", "PyMuPDF"),
                     ("PIL", "pillow"), ("zxingcpp", "zxing-cpp"),
                     ("qrcode", "qrcode"), ("fontTools", "fonttools")):
        out[mod] = version(pkg)

    # 按实际安装情况选用 OCR 引擎，并跑一次空图推理
    try:
        from rapidocr_onnxruntime import RapidOCR
        out["ocr_pkg"] = "rapidocr-onnxruntime"
    except ImportError:
        from rapidocr import RapidOCR
        out["ocr_pkg"] = "rapidocr"
    engine = RapidOCR()
    res = engine(np.full((64, 256, 3), 255, dtype=np.uint8))
    out["ocr_boxes"] = len(getattr(res, "txts", None) or [])

    print("__CHECK__" + json.dumps(out, ensure_ascii=False))
    """
).strip()

ACCESS_VIOLATION_CODES = (3221225477, -1073741819)


# ---------------------------------------------------------------- 工具函数

def run(cmd, timeout=60):
    try:
        return subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        return None
    except OSError:
        return None


def probe(prefix):
    p = run(prefix + ["-c", PROBE_CODE])
    if p is None or p.returncode != 0:
        return None
    for line in p.stdout.splitlines():
        if line.startswith("__PROBE__"):
            try:
                return json.loads(line[len("__PROBE__"):])
            except ValueError:
                return None
    return None


def find_interpreter(explicit):
    candidates = []
    if explicit:
        candidates.append([explicit])
    else:
        candidates.append([sys.executable])
        candidates.append(["py", "-3.13"])
        candidates.append(["py", "-3.11"])
        candidates.extend([[p] for p in COMMON_PATHS])

    rejected = []
    for prefix in candidates:
        info = probe(prefix)
        if not info:
            continue
        ver3 = tuple(info["ver"])
        ver2 = ver3[:2]
        if SUPPORTED_MIN <= ver2 <= SUPPORTED_MAX:
            if prefix != [sys.executable]:
                print("    找到解释器：%s  (Python %d.%d.%d)"
                      % (info["exe"], ver3[0], ver3[1], ver3[2]))
            return info["exe"], ver3
        rejected.append((info["exe"], "%d.%d.%d" % ver3))

    print("")
    if rejected:
        print("[错误] 没有找到可用的 Python 3.10~3.13 解释器。")
        print("       探测到但不支持的：")
        for exe, tag in rejected:
            print("         - Python %s  %s" % (tag, exe))
    print("")
    print("为什么不支持 3.14：")
    print("    onnxruntime %s 没有 cp314 的 wheel（cp314 最早只从 1.24.1 开始），"
          % ONNXRUNTIME_PIN)
    print("    而 1.24.1 及以上的新版本在部分机器上 import 即崩溃。")
    print("对策：装 Python 3.13（项目里的 python\\python-3.13.16-amd64.exe），")
    print("      或用 --python 显式指定。")
    sys.exit(2)


def ensure_pip(py):
    if run([py, "-m", "pip", "--version"]).returncode == 0:
        return True
    print("    未检测到 pip，尝试用 ensurepip 引导 ...")
    run([py, "-m", "ensurepip", "--upgrade"], timeout=300)
    return run([py, "-m", "pip", "--version"]).returncode == 0


def do_install(py, specs, index_url):
    base = [py, "-m", "pip", "install",
            "--upgrade-strategy", "only-if-needed",
            "--no-warn-script-location"]
    if index_url:
        base += ["--index-url", index_url]

    p = run(base + specs, timeout=1800)
    if p is None:
        return False, "安装超时或被中断"
    if p.returncode == 0:
        return True, p.stdout

    low = p.stdout.lower()
    if "access is denied" in low or "permission denied" in low or "拒绝访问" in p.stdout:
        print("    权限不足，改用 --user 模式重试 ...")
        p2 = run(base + ["--user"] + specs, timeout=1800)
        if p2 is not None and p2.returncode == 0:
            return True, p2.stdout
        return False, p2.stdout if p2 else "重试失败"
    return False, p.stdout


def verify(py):
    print("    启动独立子进程验证（防止崩溃拖死本进程）...")
    p = run([py, "-X", "faulthandler", "-u", "-c", CHECK_CODE], timeout=900)
    if p is None:
        return None, "验证子进程超时（首次运行需联网下载 OCR 模型，可能网络不通）"
    if p.returncode == 0:
        for line in p.stdout.splitlines():
            if line.startswith("__CHECK__"):
                return True, json.loads(line[len("__CHECK__"):])
    if p.returncode in ACCESS_VIOLATION_CODES or p.returncode < 0:
        return False, ("验证子进程【崩溃】了(退出码 %s / 0xC0000005)。\n"
                       "        说明这台机器的 onnxruntime %s 无法运行；"
                       "请勿升级它，或改用 OpenVINO 后端。"
                       % (p.returncode, ONNXRUNTIME_PIN))
    tail = "\n".join([l for l in p.stdout.splitlines() if l.strip()][-12:])
    return False, "验证子进程报错(退出码 %s)：\n%s" % (p.returncode, tail)


# ---------------------------------------------------------------- 主流程

def main():
    ap = argparse.ArgumentParser(
        description="extract-patch 项目依赖安装")
    ap.add_argument("--python", metavar="EXE", default=None,
                    help="显式指定目标 Python 解释器路径")
    ap.add_argument("--index-url", metavar="URL", default=None,
                    help="pip 镜像源（内网或国内镜像）")
    ap.add_argument("--skip-verify", action="store_true",
                    help="跳过安装后的验证步骤")
    args = ap.parse_args()

    print("=" * 64)
    print(" extract-patch 依赖安装")
    print("=" * 64)
    print("")

    print("[1/4] 定位 Python 解释器（需 3.10~3.13）")
    exe, ver3 = find_interpreter(args.python)
    ver2 = ver3[:2]
    print("    使用：%s  (Python %d.%d.%d)" % (exe, ver3[0], ver3[1], ver3[2]))
    print("")

    if ver2 >= SPLIT_AT:
        ocr_specs = NEW_OCR_SPECS
        print("    目标为 3.13+：旧包 rapidocr-onnxruntime 装不上，改用新版 rapidocr")
    else:
        ocr_specs = LEGACY_OCR_SPECS
        print("    目标 <=3.12：使用旧包 rapidocr-onnxruntime（与脚本 import 一致）")
    specs = COMMON_SPECS + ["onnxruntime==%s" % ONNXRUNTIME_PIN] + ocr_specs
    print("")

    print("[2/4] 检查 pip")
    if not ensure_pip(exe):
        print("[错误] 该解释器没有可用的 pip，无法安装。")
        sys.exit(3)
    print("    pip 正常")
    print("")

    print("[3/4] 安装依赖（%d 个包）" % len(specs))
    for s in specs:
        print("        %s" % s)
    print("    pip install ...")
    ok, out = do_install(exe, specs, args.index_url)
    if not ok:
        print("[错误] 安装失败，pip 输出（末尾 20 行）：")
        for line in [l for l in out.splitlines() if l.strip()][-20:]:
            print("        " + line)
        print("")
        print("若为网络问题，可加 --index-url 指定镜像源重试。")
        sys.exit(4)
    print("    安装完成")
    print("")

    print("[4/4] 验证（8 个脚本所需的库逐个导入）")
    if args.skip_verify:
        print("    已跳过（--skip-verify）")
        return

    ok, data = verify(exe)
    if ok is not True:
        print("[错误] 验证未通过：%s" % data)
        sys.exit(5)

    print("    onnxruntime   %s" % data["onnxruntime"])
    for key in ("openpyxl", "pymupdf", "PIL", "zxingcpp", "qrcode", "fontTools"):
        print("    %-13s %s" % (key, data[key]))
    print("    fitz 别名      OK")
    print("    OCR 引擎       %s（空图推理通过，检出 %d 框）"
          % (data["ocr_pkg"], data["ocr_boxes"]))
    print("")
    print("=" * 64)
    print("[完成] 依赖就绪")
    print("=" * 64)
    print("")
    print("运行项目脚本：")
    print("    %s  pdf\\1读取pdf创建表格.py" % exe)
    print("")
    print("注意：")
    print("    * 绝对不要升级 onnxruntime（pip install -U onnxruntime 一执行就废）；")
    print("      装别的库请带上 --upgrade-strategy only-if-needed")
    if data["ocr_pkg"] == "rapidocr":
        print("")
        print("    * 【需改代码】jpg\\1读取ocr创建表格.py 目前写的是：")
        print("          from rapidocr_onnxruntime import RapidOCR")
        print("          result, elapse = eng(img_path)   # 旧 API 返回二元组")
        print("      而当前装的是新包 rapidocr，用法不同：")
        print("          from rapidocr import RapidOCR")
        print("          res = eng(img_path)              # 返回对象")
        print("          texts = list(res.txts)           # 已是纯文本")
        print("      不改的话该脚本会 ImportError，其余 7 个脚本不受影响。")
        print("      想保持零改动，就改用 Python 3.11/3.12 跑这个脚本。")


if __name__ == "__main__":
    main()
