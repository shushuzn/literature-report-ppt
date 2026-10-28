#!/usr/bin/env python
"""按内容配置生成汇报 PPT。

用法：
    python tools/build.py projects/<项目>/content.yaml
    python tools/build.py projects/<项目>/content.yaml -o 输出.pptx
    python tools/build.py projects/<项目>/content.yaml --pages    # 只看页面清单
"""
import argparse
import os
import sys

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from pptkit import build_deck, resolve_pages, Style   # noqa: E402

from _console import setup_console  # noqa: E402

setup_console()


def main():
    ap = argparse.ArgumentParser(description="按内容配置生成汇报 PPT")
    ap.add_argument("config", help="content.yaml 路径")
    ap.add_argument("-o", "--out", help="覆盖输出路径")
    ap.add_argument("--pages", action="store_true", help="只打印页面清单，不生成")
    args = ap.parse_args()

    cfg_path = os.path.abspath(args.config)
    with open(cfg_path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    base = os.path.join(ROOT, cfg.get("base_dir", "."))
    out = os.path.join(ROOT, args.out or cfg["out"])
    pages = resolve_pages(cfg)

    if args.pages:
        print(f"{os.path.basename(cfg_path)}：{len(pages)} 页")
        for i, p in enumerate(pages, 1):
            print(f"  P{i:<2} {p.get('type', '?'):<16} "
                  f"{p.get('title') or p.get('title_cn') or ''}")
        return 0

    deck = build_deck(cfg, base)
    deck.save(out)

    print(f"已生成 {out}")
    print(f"页数 {deck.page_count}")
    print("页面顺序：" + " / ".join(
        f"P{i}{p.get('type', '?')}" for i, p in enumerate(pages, 1)))
    st = deck.style
    # deck.audit 只在字号低于「该页下限」时才记，页级覆盖已在内。
    bad = deck.audit
    print(f"画布 {st.slide_w:g}×{st.slide_h:g}in，"
          f"正文 {st.body_pt:g}pt，标题 {st.title_pt:g}pt")
    per = {i: p.get("body_pt") for i, p in enumerate(pages, 1) if p.get("body_pt")}
    if per:
        print("单页字号  " + "  ".join(f"P{k}:{v:g}pt" for k, v in per.items()))
    if deck.auto_pt_log:
        print("自适应字号  " + "  ".join(
            f"P{pg}:{pt:g}pt（区间 {lo:g}–{hi:g}）"
            for pg, pt, lo, hi in deck.auto_pt_log))
    print(f"低于该页字号下限：{len(bad) if bad else '无'}")
    for pg, pt, txt in bad[:8]:
        print(f"    P{pg} {pt:g}pt {txt!r}")

    if deck.warnings:
        print()
        print("版面告警：")
        for w in deck.warnings:
            print("   ", w)

    return 1 if (bad or deck.warnings) else 0


if __name__ == "__main__":
    sys.exit(main())
