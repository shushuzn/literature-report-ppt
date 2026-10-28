#!/usr/bin/env python
"""开工前把该问的问清，并算出这次能写多少字。

    python tools/intake.py                  # 交互式，逐项问，最后生成配置骨架
    python tools/intake.py --calc --pages 5 --minutes 5 --body-pt 14

六个数决定一份稿子的容量，做之前必须知道：页数、演讲时长、正文字号、
画布比例、语速、备料冗余。任何一个靠猜，后面都要返工——页数定了才知道
一页能写多少，时长定了才知道讲稿写多少，字号定了才知道版面放得下多少。

本脚本把「推荐字数」算成具体数字，写进配置，让后面每一步都有据可依。
"""
import argparse
import math

from _console import setup_console  # noqa: E402

setup_console()

# 正文占版心的比例。图文并茂的页面，图上占掉一半，右栏再留卡片边距，
# 实际落字面积约为整页的四成；纯文字页能到六成。这是排版经验值，
# 不是物理上限——写超了 build 会报放不下，写少了文字区间会报内容单薄。
FILL_IMAGE = 0.35
FILL_TEXT = 0.60
FILL_COVER = 0.16

DEFAULT_SPEED = 220      # 字/分钟，正式讲稿的实测中位
LINE_H = 1.24            # 行高倍数


def capacity(slide_w, slide_h, body_pt, margin_x=0.42,
             body_top=1.10, body_bottom=7.08):
    """整页满排能放多少字（不算图、不算卡片边距）。"""
    inner_w = (slide_w - 2 * margin_x) * 72.0
    per_line = max(1, inner_w // body_pt)
    line_h_in = body_pt * LINE_H / 72.0
    lines = max(1, (body_bottom - body_top) // line_h_in)
    return int(per_line * lines)


def plan(pages, minutes, body_pt, buffer_, speed, slide_w, slide_h):
    """给出讲稿字数与每页正文字数。"""
    cap = capacity(slide_w, slide_h, body_pt)
    must = int(minutes * speed)                 # 必讲（跳过可略段也要讲完的）
    total = int(minutes * speed * (1 + buffer_))  # 备料上限
    rows = []
    for i in range(1, pages + 1):
        if i == 1:
            lo, hi = int(cap * FILL_COVER * 0.55), int(cap * FILL_COVER)
        else:
            lo, hi = int(cap * FILL_IMAGE * 0.55), int(cap * FILL_IMAGE)
        rows.append((i, lo, hi))
    return cap, must, total, rows


def yaml_snippet(pages, minutes, buffer_, speed, body_pt,
                 slide_w, slide_h, chars_rows):
    lines = [
        "style:",
        f"  body_pt: {body_pt}          # 正文字号",
        f"  min_pt: {body_pt}           # 判违规的下限，一般与正文字号同值",
        f"  slide_w: {slide_w}",
        f"  slide_h: {slide_h}",
        "",
        "compliance:",
        f"  max_pages: {pages}",
        f"  limit_min: {minutes}",
        f"  time_buffer: {buffer_}",
        f"  chars_per_min: {speed}",
        "  chars_per_page:",
        f"    default: [{chars_rows[-1][1]}, {chars_rows[-1][2]}]",
    ]
    for i, lo, hi in chars_rows[:-1]:
        lines.append(f"    {i}: [{lo}, {hi}]")
    return "\n".join(lines)


def ask(prompt, default):
    raw = input(f"{prompt}（默认 {default}）: ").strip()
    return raw or str(default)


def main():
    ap = argparse.ArgumentParser(description="开工参数与字数预算")
    ap.add_argument("--calc", action="store_true", help="只算不生成")
    ap.add_argument("--pages", type=int, default=5, help="页数上限")
    ap.add_argument("--minutes", type=float, default=5, help="演讲时长（分钟）")
    ap.add_argument("--buffer", type=float, default=0.4, help="备料冗余，加在时长上")
    ap.add_argument("--body-pt", type=float, default=20, help="正文字号")
    ap.add_argument("--speed", type=int, default=DEFAULT_SPEED, help="语速 字/分钟")
    ap.add_argument("--ratio", default="16:9", choices=["16:9", "4:3"], help="画布比例")
    args = ap.parse_args()

    if not args.calc:
        print("做这份 PPT 之前，先把六个问题定下来。直接回车用括号里的默认值。\n")
        args.pages = int(ask("总共几页（含封面）", args.pages))
        args.minutes = float(ask("演讲限时几分钟", args.minutes))
        args.buffer = float(ask("备料冗余（0.4 表示多备四成）", args.buffer))
        args.body_pt = float(ask("正文字号多少 pt", args.body_pt))
        args.speed = int(ask("语速，每分钟多少字", args.speed))
        args.ratio = ask("画布比例 16:9 还是 4:3", args.ratio)
        print()

    w, h = (13.333, 7.5) if args.ratio == "16:9" else (10.0, 7.5)
    cap, must, total, rows = plan(args.pages, args.minutes, args.body_pt,
                                  args.buffer, args.speed, w, h)

    print("=" * 66)
    print(f"页数 {args.pages}　时长 {args.minutes:g} 分钟　字号 {args.body_pt:g}pt　"
          f"{args.ratio}")
    print("=" * 66)
    print(f"整页满排      {cap} 字（{args.body_pt:g}pt，不计图与卡片边距）")
    print(f"讲稿必讲      ≤ {must} 字（{args.minutes:g} 分钟 × {args.speed}）")
    print(f"讲稿备料      ≤ {total} 字（含 {args.buffer:.0%} 冗余，写到 "
          f"{total / args.speed:.1f} 分钟）")
    print()
    print("每页正文建议：")
    for i, lo, hi in rows:
        tag = "封面" if i == 1 else "正文页"
        print(f"  P{i}　{tag}　{lo}–{hi} 字")
    print()
    print("按这个区间写，页面既不会空，也不会挤到放不下。")
    print("写完用 python tools/check.py <pptx> -c <config> 看实际值与区间差多少。")
    print()
    print("-" * 66)
    print("配置片段（贴进 content.yaml）：")
    print("-" * 66)
    print(yaml_snippet(args.pages, args.minutes, args.buffer, args.speed,
                       args.body_pt, w, h, rows))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
