#!/usr/bin/env python
"""综述门禁：查「发展过程」这条线是否真的立得住。

    python tools/review_check.py projects/energy-materials/content.yaml

研究论文的门禁（事实核对、讲稿预算、合规、版式）都在查「这一页做得对不对」，
它们默认这份 deck 在讲一件事。综述不是：它要交的是**一个领域怎么走到今天**。
这件事没法靠任何单项数值检查兜住——页数够、字数够、图也配了，可全篇只讲
了最近三年，那就是把综述做成了研究论文。

本脚本只查三件事，都是研究论文不需要、综述缺了就不成立的：

  1. 有没有时间轴（`require_types`）
  2. 时间轴立不立得住：节点数、年代跨度、每轨年代顺序
  3. 每个节点是否写了「解决了什么」（head + desc 都在）
  4. 有没有串到研究论文的口吻（「本文提出」「本工作」——综述讲领域，不讲一篇）

触发条件：`content.yaml` 里有 `review:` 段。没有就跳过，不算失败——
研究论文型项目不受这条约束。

    review:
      min_nodes: 6            # 时间轴节点下限（所有轨合计）
      min_span_years: 50      # 时间轴年代跨度下限
      require_types: [timeline]
      forbid_voices: [本文, 本研究, 该研究, 该论文, 本工作]
"""
import argparse
import os
import re
import sys

import yaml

YEAR = re.compile(r"(1[89]\d{2}|20\d{2})")


def strings_of(node, skip=("compliance", "style", "review", "notes")):
    """递归取可见文字；讲稿（notes）不参与口吻检查。"""
    if isinstance(node, str):
        yield node
    elif isinstance(node, list):
        for x in node:
            yield from strings_of(x, skip)
    elif isinstance(node, dict):
        for k, v in node.items():
            if k in skip:
                continue
            yield from strings_of(v, skip)


def years_of(text):
    return sorted({int(y) for y in YEAR.findall(str(text or ""))})


def timeline_nodes(pages):
    """返回 [(轨名, [节点...]), ...]"""
    out = []
    for pg in pages:
        if pg.get("type") != "timeline":
            continue
        for tr in pg.get("tracks") or []:
            out.append((tr.get("name") or "（未命名轨）", tr.get("nodes") or []))
    return out


def main():
    ap = argparse.ArgumentParser(description="综述门禁")
    ap.add_argument("config")
    args = ap.parse_args()

    with open(args.config, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    rv = cfg.get("review")
    pages = cfg.get("deck") or []
    name = os.path.basename(os.path.dirname(os.path.abspath(args.config)))

    print("=" * 70)
    print(f"综述门禁：{name}")
    print("=" * 70)
    if not rv:
        print("没有 review 段，按研究论文处理，本关不适用（跳过）。")
        return 0

    min_nodes = int(rv.get("min_nodes", 6))
    min_span = int(rv.get("min_span_years", 50))
    require = rv.get("require_types") or ["timeline"]
    voices = rv.get("forbid_voices") or ["本文", "本研究", "该研究", "该论文", "本工作"]

    bad = []

    have = [p.get("type") for p in pages]
    for t in require:
        if t not in have:
            bad.append(f"缺 {t} 页：综述要讲发展过程，没有时间轴就只剩现状综述")

    tracks = timeline_nodes(pages)
    total, span_lo, span_hi = 0, None, None
    for tname, nodes in tracks:
        yrs = []
        for nd in nodes:
            ys = years_of(nd.get("year"))
            total += 1
            if not ys:
                bad.append(f"轨「{tname}」节点「{nd.get('head') or nd.get('desc')}」没有年代")
                continue
            yrs.append(min(ys))
            span_lo = min(ys) if span_lo is None else min(span_lo, min(ys))
            span_hi = max(ys) if span_hi is None else max(span_hi, max(ys))
            if not nd.get("head"):
                bad.append(f"{ys[-1]} 节点缺事件名（head）：只写年份读者不知道发生了什么")
            if not nd.get("desc"):
                bad.append(f"{ys[-1]} 节点缺说明（desc）：综述要讲清这一步解决了什么瓶颈")
        for a, b in zip(yrs, yrs[1:]):
            if b < a:
                bad.append(f"轨「{tname}」年代倒退：{a} 之后是 {b}")

    if tracks:
        span = (span_hi - span_lo) if span_lo is not None else 0
        print(f"时间轴　　　{len(tracks)} 条轨 / {total} 个节点")
        print(f"年代跨度　　{span_lo}–{span_hi}（{span} 年，下限 {min_span}）")
        if total < min_nodes:
            bad.append(f"时间轴只有 {total} 个节点，下限 {min_nodes}："
                       f"节点太少，讲不出「怎么走到今天」")
        if span < min_span:
            bad.append(f"时间轴跨度 {span} 年，下限 {min_span} 年："
                       f"只覆盖近若干年，这是现状而不是发展过程")
    else:
        print("时间轴　　　（未找到 timeline 页）")

    # 单篇口吻：综述讲领域，写「本文提出」就把读者拉回一篇论文里
    hits = []
    for s in strings_of(cfg):
        for w in voices:
            if w in s:
                hits.append((w, s[:34]))
    if hits:
        for w, s in hits[:8]:
            bad.append(f"出现单篇口吻「{w}」：{s}")
        if len(hits) > 8:
            bad.append(f"单篇口吻共 {len(hits)} 处（只列前 8）")

    print()
    if bad:
        print(f"{len(bad)} 项不合格：")
        for b in bad:
            print(f"  · {b}")
        return 1
    print("发展过程这条线立得住：有时间轴、跨度够、每个节点写明了解决了什么。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
