#!/usr/bin/env python
"""合规核对：对生成的 pptx 逐项检查。

用法：
    python tools/check.py 汇报材料/生物塑料汇报_仿生分层膜.pptx
    python tools/check.py <pptx> -c projects/LEAFF/content.yaml   # 带上期望值

退出码：0 全部通过；1 存在 ERROR。
"""
import argparse
import os
import sys

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from pptkit import check, Style     # noqa: E402


def main():
    ap = argparse.ArgumentParser(description="汇报 PPT 合规核对")
    ap.add_argument("pptx")
    ap.add_argument("-c", "--config", help="content.yaml，用于读取期望值")
    args = ap.parse_args()

    # max_pages 不给默认值：页数上限属于「这次汇报的要求」，
    # 只能由 content.yaml 的 compliance 段声明，不在代码里写死。
    kw = dict(max_pages=None, require_image_pages=None, required_text=None,
              style=None, chars_per_page=None, min_pt_pages=None,
              cover_allow=None)
    if args.config:
        with open(args.config, encoding="utf-8") as f:
            cfg = yaml.safe_load(f)
        c = cfg.get("compliance", {})
        kw["max_pages"] = c.get("max_pages")
        kw["require_image_pages"] = c.get("require_image_pages")
        rt = c.get("required_text") or {}
        kw["required_text"] = {int(k): v for k, v in rt.items()}
        kw["chars_per_page"] = c.get("chars_per_page")
        kw["cover_allow"] = c.get("cover_allow")
        kw["style"] = Style.from_config(cfg.get("style"))
        # 页级字号覆盖：某页写了 body_pt，该页下限就按它算，不牵动其他页。
        base_min = kw["style"].min_pt
        kw["min_pt_pages"] = {
            i: float(pg.get("min_pt", min(base_min, float(pg["body_pt"]))))
            for i, pg in enumerate(cfg.get("deck") or [], 1)
            if pg.get("body_pt")}

    path = args.pptx if os.path.isabs(args.pptx) else os.path.join(ROOT, args.pptx)
    findings, stats = check(path, **kw)

    print("=" * 60)
    print("合规核对：", os.path.basename(path))
    print("=" * 60)
    print(f"页数        {stats['pages']}")
    print(f"文本 run    {stats['runs']}")
    print(f"字号分布    " + "  ".join(f"{k:g}pt×{v}" for k, v in
                                     sorted((k, v) for k, v in stats["sizes"].items()
                                            if k is not None)))
    print(f"各页配图    " + "  ".join(f"P{k}:{v}" for k, v in stats["images"].items()))
    print(f"各页字数    " + "  ".join(f"P{k}:{v}" for k, v in stats["chars"].items()))
    print()
    st = kw["style"] or Style()
    per = kw["min_pt_pages"] or {}
    extra = "，" + " ".join(f"P{k}≥{v:g}pt" for k, v in sorted(per.items())) if per else ""
    print(f"字体约定    中文 {st.cn_font} / 英文 {st.en_font} / 下限 {st.min_pt:g}pt{extra}")

    if not findings:
        print("\n全部通过。")
        return 0

    print()
    for f in findings:
        print(f"[{f['level']}] {f['item']}")
        print(f"        {f['detail']}")
    print(f"\n共 {len(findings)} 项未通过。")
    return 1


if __name__ == "__main__":
    sys.exit(main())
