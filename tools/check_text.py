#!/usr/bin/env python
"""孤行预检：生成前扫出会断出「孤行」的文案。

中文长句换行后若末行只剩一两个字或标点，视觉上很突兀。
本脚本按**页面试排后的实际字号**逐页估算换行，报出末行占比过低的条目。

为什么字号要取自排版：开了字号自适应，同一段文字在 16pt 排三行、
18pt 排四行，末行剩几个字完全不同。按配置里的 `body_pt` 算，就会
出现「预检通过、成品掉单字」。这里直接调用 `pages.choose_body_pt`，
与生成时的决策同源。

可用宽度也按版式规则推（见 `pptkit/flow.py`）：栅格页取单元格内宽，
左图右文页取右栏内宽，满页文字取版心宽。换页数、换图都不必改这里。

用法：
    python tools/check_text.py projects/<项目>/content.yaml
"""
import argparse
import os
import sys

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from pptkit.flow import ORPHAN_RATIO, orphan_rows          # noqa: E402
from pptkit.pages import choose_body_pt                   # noqa: E402
from pptkit.style import Style                            # noqa: E402
from pptkit.template import resolve_pages                 # noqa: E402

THRESHOLD = ORPHAN_RATIO


def main():
    ap = argparse.ArgumentParser(description="孤行预检")
    ap.add_argument("config")
    args = ap.parse_args()

    with open(os.path.join(ROOT, args.config), encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    base = os.path.join(ROOT, cfg.get("base_dir", "."))
    pages = resolve_pages(cfg)
    st = Style.from_config(cfg.get("style"))

    print("=" * 74)
    print(f"孤行预检（末行占比 < {THRESHOLD:.2f} 视为孤行；字号取排版实际值）")
    print("=" * 74)

    rows = []
    for pg, p in enumerate(pages, 1):
        pt = float(p.get("body_pt") or st.body_pt)
        page_st = st.scaled(body_pt=pt, min_pt=min(st.min_pt, pt))
        t = p.get("type")
        if p.get("auto_pt", st.auto_pt) and t not in ("cover", "section"):
            pt, _, _ = choose_body_pt(p, base, page_st, t)
            page_st = page_st.scaled(body_pt=pt, min_pt=min(st.min_pt, pt))
        for tag, n, ratio, text in orphan_rows(base, p, page_st, pt):
            rows.append((pg, tag, ratio, n, text))

    if not rows:
        print("没有需要换行的长文本，无需检查。")
        return 0

    bad = [r for r in rows if r[2] < THRESHOLD]
    for pg, tag, ratio, n, text in sorted(rows, key=lambda r: r[2]):
        mark = "孤行" if ratio < THRESHOLD else "正常"
        print(f"[{mark}] P{pg} {tag[:16]:18s} {n}行 末行{ratio:.2f}  {text[:30]}")
    print()
    print(f"需改写：{len(bad)} 条")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
