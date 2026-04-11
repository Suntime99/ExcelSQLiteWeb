#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把 index.html 里的“最后一个内联 <script>...</script>”抽取为 app.js，并输出一个新的入口 HTML。

使用方式（在项目根目录执行）：
  python tools/split_index_inline_js.py --in index.html --out-html index.html --out-js app.js --backup index.inline.bak.html

说明：
1) 只处理“最后一个内联 script 标签”，因为当前工程只有一个大 script。
2) 抽取后，HTML 会把该 script 替换为：<script src="app.js"></script>
3) 适用于 WebView2 走 file:// 导航的场景（相对路径可用）。
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", default="index.html", help="输入 HTML 文件路径")
    ap.add_argument("--out-html", dest="out_html", default="index.html", help="输出 HTML 文件路径")
    ap.add_argument("--out-js", dest="out_js", default="app.js", help="输出 JS 文件路径")
    ap.add_argument("--backup", dest="backup", default="index.inline.bak.html", help="备份原 HTML 路径")
    args = ap.parse_args()

    inp = Path(args.inp)
    out_html = Path(args.out_html)
    out_js = Path(args.out_js)
    backup = Path(args.backup)

    html = inp.read_text(encoding="utf-8", errors="ignore")

    # 匹配所有内联 script（不带 src 属性的）
    # 注意：用一个较宽松的正则，因为脚本很大。
    scripts = list(
        re.finditer(r"<script(?![^>]*\\bsrc=)[^>]*>\\s*(.*?)\\s*</script>", html, flags=re.IGNORECASE | re.DOTALL)
    )
    # 上面的正则在某些 Python/regex 场景下会把 \\s 解释为“字面量 \\s”，导致无法匹配。
    # 这里做一次兜底：使用真正的 \s / \b。
    if not scripts:
        scripts = list(
            re.finditer(r"<script(?![^>]*\bsrc=)[^>]*>\s*(.*?)\s*</script>", html, flags=re.IGNORECASE | re.DOTALL)
        )
    if not scripts:
        raise SystemExit("未找到内联 <script>...</script>，无需拆分。")

    m = scripts[-1]
    js = m.group(1)

    # 写出 app.js（不做任何格式化/压缩，保持原样）
    out_js.write_text(js + ("\n" if not js.endswith("\n") else ""), encoding="utf-8")

    # 替换为外链脚本
    replacement = "<script src=\"{}\"></script>".format(out_js.name)
    new_html = html[: m.start()] + replacement + html[m.end() :]

    # 备份原 HTML（只在 out_html 与 inp 相同时才备份，避免覆盖误操作）
    if out_html.resolve() == inp.resolve():
        backup.write_text(html, encoding="utf-8")

    out_html.write_text(new_html, encoding="utf-8")
    print("OK")
    print("  HTML:", out_html)
    print("  JS  :", out_js)
    if out_html.resolve() == inp.resolve():
        print("  BAK :", backup)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
