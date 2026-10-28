#!/usr/bin/env python
"""事实核对：把 content.yaml 里的每个数字拿回原文比对。

    python tools/fact_check.py projects/<项目>/content.yaml
    python tools/fact_check.py projects/<项目>/content.yaml --src <原文.txt>

为什么要有这一步：PPT 上的数字一旦写错，讲的时候就是错的，事后没人复核。
本脚本把配置里的全部数值 token 抽出来，逐个到原文里找；找不到的报出来，
要么改数，要么进白名单（影响因子、分区这类论文里没有的外部信息）。

比对按**完整数值**做，不做子串匹配：配置里的 0.7 不会因为原文有 10.7 就算
找到，30 也不会藏进 300。原文写成 0.70、1,700 这类等价写法仍算命中。

白名单文件：与 content.yaml 同级的 `_src/facts_allow.txt`，一行一个数值，
`#` 后可写原因。白名单只豁免「原文没有但确实正确」的外部数据，
不能拿它盖住写错的数。
"""
import argparse
import os
import re
import sys
from decimal import Decimal, InvalidOperation

import yaml

from _console import setup_console  # noqa: E402

setup_console()

NUM = re.compile(r"\d+(?:\.\d+)?")


def norm(s: str) -> str:
    """把千分位、全角负号、Unicode 减号统一，避免 1,700 与 1700 对不上。"""
    s = s.replace(",", "").replace("，", "")
    for ch in ("−", "–", "—", "－"):
        s = s.replace(ch, "-")
    return s


def _number(s: str):
    """数值字符串 → Decimal；不是数值则返回 None。

    用 Decimal 而不是 float：0.7 与 0.70、1700 与 1700.0 应当算同一个值，
    而浮点数会把「看起来相等」和「真相等」搅在一起。
    """
    try:
        return Decimal(s)
    except (InvalidOperation, ValueError):
        return None


def numbers_in(text: str):
    """原文里出现的全部数值（按完整 token，不按子串）。"""
    return {Decimal(m.group()) for m in NUM.finditer(norm(text))}


def strings_of(node):
    """递归取出配置里所有字符串（跳过 compliance 与 style 这类判据字段）。"""
    if isinstance(node, str):
        yield node
    elif isinstance(node, list):
        for x in node:
            yield from strings_of(x)
    elif isinstance(node, dict):
        for k, v in node.items():
            if k in ("compliance", "style", "notes"):
                continue
            yield from strings_of(v)


def tokens_in(text: str):
    """按出现顺序给出 (数值, 所在句)。"""
    out = []
    for m in NUM.finditer(text):
        s = max(0, m.start() - 18)
        out.append((m.group(), text[s:m.end() + 18].strip()))
    return out


def load_allow(path):
    if not os.path.exists(path):
        return set(), {}
    allow, why = set(), {}
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            num, _, reason = line.partition("#")
            num = num.strip()
            if num:
                allow.add(num)
                why[num] = reason.strip()
    return allow, why


def main():
    ap = argparse.ArgumentParser(description="配置数值与原文交叉核对")
    ap.add_argument("config")
    ap.add_argument("--src", help="论文正文文本；默认找 <项目>/_src/paper.txt")
    args = ap.parse_args()

    cfg_path = args.config
    with open(cfg_path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    proj = os.path.dirname(os.path.abspath(cfg_path))
    src_path = args.src
    if not src_path:
        # 抽取工具产物名字不一：paper.txt / paper_clean.txt / paper.md 都认
        for cand in ("paper.txt", "paper_clean.txt", "paper.md"):
            p = os.path.join(proj, "_src", cand)
            if os.path.exists(p):
                src_path = p
                break
    if not src_path or not os.path.exists(src_path):
        print("本项目未附原文，事实核对无法执行，跳过（不算失败）。")
        print("有原文时把 PDF 抽成 _src/paper.txt 再跑，这一关才有意义。")
        return 0

    with open(src_path, encoding="utf-8", errors="replace") as f:
        src = norm(f.read())
    body_numbers = numbers_in(src)

    allow, why = load_allow(os.path.join(proj, "_src", "facts_allow.txt"))
    allow_numbers = {_number(x) for x in allow}
    allow_numbers.discard(None)

    seen, ok, miss, skipped = set(), [], [], []
    for s in strings_of(cfg):
        for num, ctx in tokens_in(s):
            key = (num, ctx[:20])
            if key in seen:
                continue
            seen.add(key)
            val = _number(num)
            if num in allow or (val is not None and val in allow_numbers):
                skipped.append((num, ctx))
            elif val is not None and val in body_numbers:
                ok.append((num, ctx))
            else:
                miss.append((num, ctx))

    print("=" * 70)
    print(f"事实核对：{os.path.basename(proj)}　原文 {os.path.basename(src_path)}")
    print("=" * 70)
    print(f"原文已有　{len(ok)} 项")
    if skipped:
        print(f"白名单　　{len(skipped)} 项")
        for num, ctx in skipped[:6]:
            print(f"  {num:>8}  {why.get(num, '')[:40]}　（{ctx[:26]}）")
    if miss:
        print(f"\n原文未找到　{len(miss)} 项，逐个核：")
        for num, ctx in miss:
            print(f"  {num:>8}　{ctx}")
        print("\n处理办法：改数、或把外部数据（影响因子/分区/换算值）写进 "
              "_src/facts_allow.txt 并注明来源。")
        return 1
    print("\n全部数值都能在原文找到出处。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
