#!/usr/bin/env python
"""把 content.yaml 里的讲稿（notes）导成 Markdown，并估算每页用时。

    python tools/export_notes.py projects/energy-materials/content.yaml

输出到 content.yaml 里 out 的同级目录，文件名 讲稿.md。
页面顺序、页标题与 PPT 一致，便于对着成品排练。

字数按汉字计，西文与数字按半个汉字折算；语速取 220 字/分钟。
限时与冗余比例从 content.yaml 的 compliance 段读（limit_min / time_buffer），
没有设 limit_min 的项目只报数、不判超时。判定逻辑在 tools/note_budget.py。
讲稿里以〔可略〕开头的段落是备料多出来的展开段，现场时间紧可整段跳过。
"""
import argparse
import os
import re
import sys

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from note_budget import (  # noqa: E402
    DEFAULT_SPEED as SPEED, budget_of, count_chars, split_optional,
)


def page_title(p: dict) -> str:
    t = p.get("type")
    if t == "cover":
        return p.get("title_cn", "封面")
    if t == "section":
        return f"{p.get('index', '')} {p.get('title', '')}".strip()
    return p.get("title", "")


def main():
    ap = argparse.ArgumentParser(description="导出讲稿 Markdown")
    ap.add_argument("config")
    ap.add_argument("-o", "--out", help="输出路径，默认与 pptx 同目录的 讲稿.md")
    args = ap.parse_args()

    cfg_path = args.config if os.path.isabs(args.config) else os.path.join(ROOT, args.config)
    with open(cfg_path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    pages = cfg.get("deck") or []
    if not pages:
        print("该配置没有 deck 列表（旧五段式），无法导出讲稿。")
        return 1

    blocks, total, must_total = [], 0.0, 0.0
    for i, p in enumerate(pages, 1):
        note = (p.get("notes") or "").strip()
        must, opt = split_optional(note)
        n, m = count_chars(note), count_chars(must)
        total += n
        must_total += m
        blocks.append((i, page_title(p), note, n, m))

    out = args.out or os.path.join(os.path.dirname(cfg["out"]), "讲稿.md")
    out = out if os.path.isabs(out) else os.path.join(ROOT, out)

    b = budget_of(cfg)
    lines = ["# 讲稿", ""]
    lines.append(f"共 {len(pages)} 页、{total:.0f} 字，全程约 {total / SPEED:.1f} 分钟；"
                 f"跳过〔可略〕段为 {must_total:.0f} 字、约 {must_total / SPEED:.1f} 分钟。")
    if b:
        limit, buf, speed, cap, must_cap = b
        lines.append(f"限时 {limit:g} 分钟、冗余 {buf:.0%}：备料上限 {cap:.0f} 字，"
                     f"余量 {cap - total:.0f} 字（{(cap - total) / speed * 60:.0f} 秒）。")
    lines.append("")

    missing = []
    for i, title, note, n, m in blocks:
        sec = n / SPEED * 60 if n else 0
        lines.append(f"## P{i}　{title}")
        lines.append("")
        if not note:
            missing.append(i)
            lines.append("（本页未写讲稿）")
        else:
            tail = f"（必讲 {m:.0f} 字 / {m / SPEED * 60:.0f} 秒）" if m < n else ""
            lines.append(f"> 约 {n:.0f} 字 / {sec:.0f} 秒{tail}")
            lines.append("")
            lines.append(note)
        lines.append("")

    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"已导出 {out}")
    print(f"共 {total:.0f} 字，全程约 {total / SPEED:.1f} 分钟；"
          f"跳可略 {must_total:.0f} 字，约 {must_total / SPEED:.1f} 分钟")
    if b:
        limit, buf, speed, cap, must_cap = b
        print(f"限时 {limit:g} 分钟、冗余 {buf:.0%} → 备料上限 {cap:.0f} 字，"
              f"余量 {cap - total:+.0f} 字")
    if missing:
        print("未写讲稿的页：", "、".join(f"P{i}" for i in missing))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
