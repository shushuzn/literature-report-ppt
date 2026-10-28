#!/usr/bin/env python
"""讲稿时长门禁：按汇报限时与冗余比例算字数上限，超了就不许交。

    python tools/note_budget.py projects/<项目>/content.yaml

配置写在 content.yaml 的 compliance 段：

    compliance:
      limit_min: 5          # 汇报限时（分钟），不写则不检查时长
      time_buffer: 0.4      # 冗余比例，默认 0.4
      chars_per_min: 220    # 语速，默认 220

预算 = limit_min × chars_per_min × (1 + time_buffer)。
5 分钟、冗余 40% → 1540 字。冗余是**加在限时上**的备料：先按 5 分钟把话讲透，
再多备四成的展开段，现场时间紧就跳过。按 5 分钟写满 1100 字、不留余量，
现场一翻页一停顿就超时；只写 3 分钟的量，内容又撑不起一次汇报。

可略段写在讲稿里，以「〔可略〕」开头，到空行为止：

    〔可略〕顺带说一句，孔径分布还测了 0.7 到 1.2 纳米这一段……

门禁同时算两笔账：
    必讲（去掉〔可略〕）不得超 limit_min，保证跳过可略段能压回限时；
    全文不得超 limit_min × (1 + time_buffer)，保证备料没有灌水。
"""
import argparse
import os
import re
import sys

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from pptkit.template import resolve_pages  # noqa: E402

from _console import setup_console  # noqa: E402

setup_console()

DEFAULT_BUFFER = 0.4
DEFAULT_SPEED = 220.0
OPT_MARK = "〔可略〕"

CJK = re.compile(r"[\u4e00-\u9fff]")
ALNUM = re.compile(r"[A-Za-z0-9]")


def count_chars(text):
    """汉字算 1，西文字母与数字算 0.5（读一个英文词约等于半个汉字的时间）。"""
    return len(CJK.findall(text)) + len(ALNUM.findall(text)) * 0.5


def split_optional(note):
    """拆成 (必讲, 可略)。〔可略〕开头的一段算可略，跳到段尾。"""
    must, opt, in_opt = [], [], False
    for line in (note or "").splitlines():
        s = line.strip()
        if s.startswith(OPT_MARK):
            in_opt = True
            opt.append(s[len(OPT_MARK):])
            continue
        if not s:
            in_opt = False
            continue
        (opt if in_opt else must).append(s)
    return "".join(must), "".join(opt)


def budget_of(cfg):
    """返回 (限时, 冗余, 语速, 备料上限字数, 必讲上限字数)；未设 limit_min 返回 None。"""
    c = cfg.get("compliance") or {}
    limit = c.get("limit_min")
    if not limit:
        return None
    buf = float(c.get("time_buffer", DEFAULT_BUFFER))
    speed = float(c.get("chars_per_min", DEFAULT_SPEED))
    limit = float(limit)
    return limit, buf, speed, limit * speed * (1 + buf), limit * speed


def measure(cfg):
    """返回 [(页号, 标题, 必讲字数, 全文含可略字数)]。

    全文按原始 notes 计，与 export_notes 报的数字一致；〔可略〕标记本身
    也算进可略，它确实要占讲稿版面。
    """
    out = []
    for i, p in enumerate(resolve_pages(cfg), 1):
        t = p.get("title_cn") or p.get("title") or p.get("type", "")
        note = (p.get("notes") or "").strip()
        must, _ = split_optional(note)
        out.append((i, t, count_chars(must), count_chars(note)))
    return out


def main():
    ap = argparse.ArgumentParser(description="讲稿时长门禁")
    ap.add_argument("config")
    ap.add_argument("--quiet", action="store_true", help="只报错，不列明细")
    args = ap.parse_args()

    path = args.config if os.path.isabs(args.config) else os.path.join(ROOT, args.config)
    with open(path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    b = budget_of(cfg)
    rows = measure(cfg)
    total = sum(t for _, _, _, t in rows)
    must_total = sum(m for _, _, m, _ in rows)
    opt_total = total - must_total

    if b is None:
        print(f"未设 compliance.limit_min，跳过时长门禁（当前 {total:.0f} 字）")
        return 0

    limit, buf, speed, cap, must_cap = b

    print(f"限时 {limit:g} 分钟、冗余 {buf:.0%} → 备料上限 {limit * (1 + buf):.1f} 分钟 "
          f"= {cap:.0f} 字（语速 {speed:.0f} 字/分）")
    print(f"讲稿 {total:.0f} 字：必讲 {must_total:.0f} 字 / {must_total / speed:.1f} 分钟"
          f"（上限 {must_cap:.0f} 字），"
          f"可略 {opt_total:.0f} 字 / {opt_total / speed:.1f} 分钟")

    if not args.quiet:
        per = cap / max(len(rows), 1)
        for i, title, m, t in rows:
            flag = "超" if t > per * 1.35 else "  "
            tail = f"（含可略 {t - m:.0f}）" if t > m else ""
            print(f"  {flag} P{i} {t:>4.0f} 字{tail}　{title}")

    rc = 0
    if total > cap:
        over = total - cap
        worst = sorted(rows, key=lambda r: -r[3])[:3]
        print()
        print(f"备料超出 {over:.0f} 字（约 {over / speed * 60:.0f} 秒）——上限是"
              f" {limit:g} 分钟加 {buf:.0%} 冗余，不是无限备。砍展开段：")
        for i, title, m, t in worst:
            print(f"  P{i} {t:.0f} 字（必讲 {m:.0f}）　{title}")
        rc = 1
    if must_total > must_cap:
        over = must_total - must_cap
        print()
        print(f"必讲超出 {over:.0f} 字（约 {over / speed * 60:.0f} 秒）：跳过全部〔可略〕"
              f"也压不回 {limit:g} 分钟。把这几页的次要数据改成可略段，或删一句：")
        for i, title, m, _t in sorted(rows, key=lambda r: -r[2])[:3]:
            print(f"  P{i} 必讲 {m:.0f} 字　{title}")
        rc = 1
    if rc == 0 and opt_total < cap * buf / (1 + buf) * 0.8:
        print(f"\n提示：可略段只有 {opt_total:.0f} 字，不足四成冗余，现场时间紧时没得删。")

    if rc == 0:
        left = cap - total
        print(f"余量 {left:.0f} 字（{left / speed * 60:.0f} 秒），通过；"
              f"全程 {total / speed:.1f} 分钟，跳可略 {must_total / speed:.1f} 分钟。")
    return rc


if __name__ == "__main__":
    sys.exit(main())
