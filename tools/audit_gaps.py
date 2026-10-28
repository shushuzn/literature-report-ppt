#!/usr/bin/env python
"""间隔诊断：量某一区域内每行文字的位置与行间距。

用于排查「同一栏内各段间隔不一致」这类问题——把渲染图里每行文字的实际
上下位置列出来，间隔是否均匀一目了然，不必靠眼睛判断。

用法：
    python tools/audit_gaps.py <preview目录>/slide2.png 6.18 1.05 12.95 6.60

参数：图片路径、区域左、上、右、下（英寸，相对整页）
"""
import argparse
import sys

import numpy as np
from PIL import Image

from _console import setup_console  # noqa: E402

setup_console()

DEFAULT_W = 13.333      # 未给 --slide-w 时按 16:9


def main():
    ap = argparse.ArgumentParser(description="行间距诊断")
    ap.add_argument("png")
    ap.add_argument("x0", type=float)
    ap.add_argument("y0", type=float)
    ap.add_argument("x1", type=float)
    ap.add_argument("y1", type=float)
    ap.add_argument("--thresh", type=int, default=170, help="灰度阈值")
    ap.add_argument("--slide-w", type=float, default=DEFAULT_W,
                    help="画布宽（英寸），4:3 时填 10")
    args = ap.parse_args()

    im = Image.open(args.png).convert("L")
    a = np.asarray(im, dtype=np.uint8)
    scale = im.width / args.slide_w

    px0, px1 = int(args.x0 * scale), int(args.x1 * scale)
    py0, py1 = int(args.y0 * scale), int(args.y1 * scale)
    dark = a[py0:py1, px0:px1] < args.thresh

    rows = np.where(dark.any(axis=1))[0]
    if rows.size == 0:
        print("该区域没有检出文字像素。")
        return 0

    lines, cur = [], [rows[0]]
    for r in rows[1:]:
        if r - cur[-1] <= 2:
            cur.append(r)
        else:
            lines.append((cur[0], cur[-1]))
            cur = [r]
    lines.append((cur[0], cur[-1]))

    print(f"{'行':>3} {'顶端(in)':>10} {'底端(in)':>10} {'行高(pt)':>9} {'与上一行间隔(pt)':>16}")
    prev = None
    gaps = []
    for i, (t, b) in enumerate(lines, 1):
        yt, yb = (t + py0) / scale, (b + py0) / scale
        g = ""
        if prev is not None:
            gv = (yt - prev) * 72
            gaps.append(gv)
            g = f"{gv:16.1f}"
        print(f"{i:>3} {yt:>10.3f} {yb:>10.3f} {(yb - yt) * 72:>9.1f} {g}")
        prev = yb

    if gaps:
        arr = np.array(gaps)
        print()
        print(f"间隔：最小 {arr.min():.1f}pt  最大 {arr.max():.1f}pt  "
              f"极差 {arr.max() - arr.min():.1f}pt")
        print("提示：段内行距约 6pt、段间距约 18pt 属正常。"
              "同类间隔若极差超过约 4pt，说明版面分配不均。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
