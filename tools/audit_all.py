#!/usr/bin/env python
"""版式验收：对**每一页**、**每个维度**做量化检查。

曾经反复出现排版问题，共同原因是检查维度不全、且只查被投诉的那一页。
本脚本把版式验收系统化——八个维度全部对所有页跑一遍：

  1. 越界      所有元素是否都在页面内
  2. 重叠      非包含关系的元素是否互相压盖
  3. 图片贴合  图片外框与图片真实比例是否吻合（死区比例）
  4. 框内居中  面板内文字的上下边距是否对称（读 XML，不依赖渲染）
  5. 卡内留白  渲染图上量：上下留白是否够、是否对称，标题与正文是否分得开
  6. 孤行      末行是否只剩一两个字或标点
  7. 段间距    同一文本框内的段落间隔是否均匀
  8. 卡片间距  同栏相邻卡片的间距是否均匀
  9. 栏内填充  各栏底部是否留出大片空白

用法：
    python tools/audit_all.py 汇报材料/生物塑料汇报_仿生分层膜.pptx 汇报材料/preview
"""
import os
import sys
from collections import defaultdict

import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.util import Emu

# 画布与正文区位置在 main() 里按实际 pptx 与配置赋值，不预设 16:9。
# 换 4:3、换更大字号的项目时，下面的像素↔英寸换算才仍然成立。
SLIDE_W_IN = 13.333
SLIDE_H_IN = 7.5
BODY_TOP = 1.10

# 判定阈值
TOL_BOUNDS = 0.02       # 越界容差（英寸）
DEAD_MAX = 0.12         # 图片死区上限
CENTER_TOL_PT = 6       # 居中容差（磅）
ORPHAN_MIN = 0.18       # 末行占比下限
GAP_SPREAD_MAX = 6      # 同类间隔极差上限（磅）
BLANK_MAX = 0.35        # 各栏底部留白上限（英寸）
MIN_PAD_PT = 7.2        # 卡内文字到框线的最小留白（0.10in）
MIN_LEVEL_PT = 8.0      # 卡内标题与正文的间隔下限（须明显大于段内行距）


def inches(v):
    return Emu(v).inches


def shapes_of(slide):
    """返回 [(kind, x, y, w, h, shape)]；kind ∈ picture/panel/text/other"""
    out = []
    for sh in slide.shapes:
        st = str(sh.shape_type) if sh.shape_type is not None else ""
        x, y = inches(sh.left), inches(sh.top)
        w, h = inches(sh.width), inches(sh.height)
        if "PICTURE" in st:
            k = "picture"
        elif "AUTO_SHAPE" in st and w >= 1.0 and h >= 0.35:
            k = "panel"
        elif sh.has_text_frame and sh.text_frame.text.strip():
            k = "text"
        else:
            k = "other"
        out.append((k, x, y, w, h, sh))
    return out


def inter(a, b):
    x = max(0.0, min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0]))
    y = max(0.0, min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1]))
    return x * y


def contains(a, b):
    return (a[0] <= b[0] + 1e-6 and a[1] <= b[1] + 1e-6 and
            a[0] + a[2] >= b[0] + b[2] - 1e-6 and a[1] + a[3] >= b[1] + b[3] - 1e-6)


# ---------- 维度 1：越界 ----------
def check_bounds(items):
    bad = []
    for k, x, y, w, h, _ in items:
        if x < -TOL_BOUNDS or y < -TOL_BOUNDS or \
           x + w > SLIDE_W_IN + TOL_BOUNDS or y + h > SLIDE_H_IN + TOL_BOUNDS:
            bad.append(f"{k} 位于 ({x:.2f},{y:.2f}) 尺寸 {w:.2f}x{h:.2f} 超出页面")
    return bad


# ---------- 维度 2：重叠 ----------
def check_overlap(items):
    bad = []
    cand = [(k, x, y, w, h) for k, x, y, w, h, _ in items if k in ("panel", "text")]
    for i in range(len(cand)):
        for j in range(i + 1, len(cand)):
            a, b = cand[i], cand[j]
            ra, rb = a[1:], b[1:]
            if contains(ra, rb) or contains(rb, ra):
                continue                      # 包含关系属正常嵌套
            ov = inter(ra, rb)
            if ov > 0.01:
                bad.append(f"{a[0]}({ra[0]:.2f},{ra[1]:.2f}) 与 "
                           f"{b[0]}({rb[0]:.2f},{rb[1]:.2f}) 重叠 {ov:.2f}in²")
    return bad


# ---------- 维度 3：图片贴合 ----------
def check_image_fit(items):
    bad = []
    frames = [(x, y, w, h) for k, x, y, w, h, _ in items if k == "panel"]
    for k, x, y, w, h, sh in items:
        if k != "picture":
            continue
        host = min((f for f in frames if contains(f, (x, y, w, h))),
                   key=lambda f: f[2] * f[3], default=None)
        ref = host if host else (x, y, w, h)
        dead = 1.0 - (w * h) / (ref[2] * ref[3])
        if dead > DEAD_MAX:
            bad.append(f"图片 ({x:.2f},{y:.2f}) 死区 {dead:.0%}（外框 {ref[2]:.2f}x{ref[3]:.2f}）")
    return bad


# ---------- 维度 4：框内居中（读 XML） ----------
def check_panel_centering(items):
    bad = []
    panels = [(x, y, w, h) for k, x, y, w, h, _ in items if k == "panel"]
    for k, x, y, w, h, sh in items:
        if k != "text":
            continue
        host = min((f for f in panels if contains(f, (x, y, w, h))),
                   key=lambda f: f[2] * f[3], default=None)
        if not host:
            continue
        top = y - host[1]
        bottom = (host[1] + host[3]) - (y + h)
        if abs(top - bottom) * 72 > CENTER_TOL_PT:
            bad.append(f"面板({host[0]:.2f},{host[1]:.2f}) 内文字上{top*72:.0f}pt "
                       f"下{bottom*72:.0f}pt")
    return bad


# ---------- 维度 5/6/7：渲染图 ----------
def _bold_lead(sh):
    """文本框首段首个非空 run 是否加粗——卡片的标题行就是靠这个认出来的"""
    if not sh.has_text_frame:
        return False
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if r.text.strip():
                return bool(r.font.bold)
    return False


def render_metrics(png, items):
    """返回 (卡内留白问题, 孤行问题)。

    卡内留白这一项查三件事：
      · 上下留白是否对称（居中）——不对称说明文字在框里偏上或偏下
      · 上下留白是否够——小于 MIN_PAD_PT 就是文字顶到框线
      · 有标题的卡片，标题与正文是否分得开——间隔小于 MIN_LEVEL_PT
        时两行会糊成一块（曾出现只隔 3pt、比框内留白还小的情况）
    """
    im = Image.open(png).convert("L")
    a = np.asarray(im, dtype=np.uint8)
    scale = im.width / SLIDE_W_IN
    pad_bad, orphan_bad = [], []

    panels = [(x, y, w, h) for k, x, y, w, h, _ in items if k == "panel"]
    for (x, y, w, h) in panels:
        inner_has_pic = any(k == "picture" and contains((x, y, w, h), (px, py, pw, ph))
                            for k, px, py, pw, ph, _ in items)
        if inner_has_pic:
            continue
        box = a[int(y * scale):int((y + h) * scale),
                int(x * scale):int((x + w) * scale)]
        if box.size == 0:
            continue
        pad = 4
        inner = box[pad:-pad, pad:-pad]
        dark = inner < 170
        rows = np.where(dark.any(axis=1))[0]
        if rows.size == 0:
            continue

        top = int(rows[0])
        bottom = int(inner.shape[0] - 1 - rows[-1])
        tag = f"面板({x:.2f},{y:.2f})"
        if abs(top - bottom) / scale * 72 > CENTER_TOL_PT:
            pad_bad.append(f"{tag} 上{top}px 下{bottom}px，上下不居中")
        lean = min(top, bottom) / scale * 72
        if lean < MIN_PAD_PT:
            pad_bad.append(f"{tag} 内留白仅 {lean:.0f}pt（下限 {MIN_PAD_PT:.0f}pt），文字顶到框线")

        # 文本行分组，量组间隔
        gaps = np.where(np.diff(rows) > 2)[0]
        starts = np.r_[0, gaps + 1]
        if len(starts) >= 2:
            ends = np.r_[rows[starts[1:] - 1], rows[-1]]
            lev = [(rows[s] - ends[i]) / scale * 72
                   for i, s in enumerate(starts[1:])]
            titled = any(_bold_lead(sh) for k, tx, ty, tw, th, sh in items
                         if k == "text" and contains((x, y, w, h), (tx, ty, tw, th)))
            if titled and max(lev) < MIN_LEVEL_PT:
                pad_bad.append(f"{tag} 标题与正文仅隔 {max(lev):.1f}pt"
                               f"（下限 {MIN_LEVEL_PT:.0f}pt），层级分不开")

        # 孤行：末行的横向占比
        r0, r1 = rows[starts[-1]], rows[-1]
        cols = np.where(dark[r0:r1 + 1].any(axis=0))[0]
        if cols.size and (cols[-1] - cols[0] + 1) / inner.shape[1] < ORPHAN_MIN:
            orphan_bad.append(f"{tag} 末行占比 "
                              f"{(cols[-1]-cols[0]+1)/inner.shape[1]:.2f}")
    return pad_bad, orphan_bad


def _text_lines(dark, min_h=6):
    """把暗行聚成文本行。min_h 用于剔除边框、分隔线等细线。"""
    rows = np.where(dark.any(axis=1))[0]
    if rows.size == 0:
        return []
    bands, cur = [], [rows[0]]
    for r in rows[1:]:
        if r - cur[-1] <= 2:
            cur.append(r)
        else:
            bands.append((cur[0], cur[-1]))
            cur = [r]
    bands.append((cur[0], cur[-1]))
    return [(t, b) for t, b in bands if b - t + 1 >= min_h]


def check_gaps(png, items):
    """**同一个文本框内**的段落间隔是否均匀。

    关键：按单个文本框分组，不跨文本框取整栏。此前按栏取 min..max，把
    「卡内标题与正文的间隔」和「卡片与卡片之间的间距」混在一起比——卡内
    层级间隔一放大（3pt → 15pt），两类间隔就被当成一组，误报极差 40pt。
    卡片之间的间距由 check_card_gaps() 单独负责。

    适用对象是「一个文本框里排多个同类条目」的版面（如 P2 的字段列表）。
    """
    im = Image.open(png).convert("L")
    a = np.asarray(im, dtype=np.uint8)
    scale = im.width / SLIDE_W_IN

    bad = []
    for k, x, y, w, h, sh in items:
        if k != "text":
            continue
        band = a[int(y * scale):int((y + h) * scale),
                 int(x * scale):int((x + w) * scale)]
        if band.size == 0:
            continue
        lines = _text_lines(band < 170)
        if len(lines) < 3:
            continue
        gs = [(lines[i + 1][0] - lines[i][1]) / scale * 72
              for i in range(len(lines) - 1)]
        big = [g for g in gs if g > 12]        # 段间距（段内行距约 6pt）
        if len(big) >= 2:
            spread = max(big) - min(big)
            if spread > GAP_SPREAD_MAX:
                bad.append(f"文本框({x:.2f},{y:.2f}) 段间距 "
                           f"{[round(g) for g in big]} 极差 {spread:.0f}pt")
    return bad


def check_card_gaps(items):
    """同栏相邻卡片的间距是否均匀（读 XML，精确到 0.001in）。

    与 check_gaps 分工：那一项管文本框内部，本项管卡片之间。
    """
    cols = defaultdict(list)
    for k, x, y, w, h, sh in items:
        if k == "panel" and w >= 1.0 and h >= 0.35:
            cols[round(x, 1)].append((y, y + h))
    bad = []
    for x, spans in cols.items():
        if len(spans) < 3:                     # 少于 3 张卡只有 1 个间距，无从比较
            continue
        spans.sort()
        gs = [(spans[i + 1][0] - spans[i][1]) * 72 for i in range(len(spans) - 1)]
        spread = max(gs) - min(gs)
        if spread > GAP_SPREAD_MAX:
            bad.append(f"x≈{x:.2f} 卡片间距 {[round(g) for g in gs]} "
                       f"极差 {spread:.0f}pt")
    return bad


def check_column_fill(png, items):
    """栏内填充：**同一横排**并列的各栏，内容底端是否齐平。

    只比较顶端一致的栏——它们是并列关系（如左图右文）。
    顶端不同的栏属于纵向堆叠（如 P4 上排图片、下排表格），
    各自结束在不同高度是正常的，不能互相比较。
    """
    im = Image.open(png).convert("L")
    a = np.asarray(im, dtype=np.uint8)
    scale = im.width / SLIDE_W_IN

    cols = defaultdict(list)
    for k, x, y, w, h, sh in items:
        if k in ("panel", "text"):        # 图片在框内，由框代表
            cols[round(x, 1)].append((y, y + h, w))
    if len(cols) < 2:
        return []

    # 量各栏的内容底端
    bottoms, tops = {}, {}
    for x, spans in cols.items():
        top = min(s[0] for s in spans)
        bottom = max(s[1] for s in spans)
        widest = max(s[2] for s in spans)
        if bottom < BODY_TOP + 1.0:        # 页眉区，跳过
            continue
        band = a[int(top * scale):int(bottom * scale),
                 int(x * scale):int((x + widest) * scale)]
        rows = np.where((band < 170).any(axis=1))[0]
        if rows.size == 0:
            continue
        bottoms[x] = (rows[-1] + int(top * scale)) / scale
        tops[x] = top

    # 按顶端归组：只有同一横排的栏才互相比较
    bands = defaultdict(list)
    for x, top in tops.items():
        bands[round(top, 1)].append(x)

    bad = []
    for top, xs in bands.items():
        if len(xs) < 2:
            continue
        vals = {x: bottoms[x] for x in xs}
        lowest = max(vals.values())
        for x, ink in vals.items():
            if lowest - ink > BLANK_MAX:
                bad.append(f"x≈{x:.2f} 栏内容止于 {ink:.2f}in，"
                           f"同排另一栏到 {lowest:.2f}in，差 {lowest - ink:.2f}in")
    return bad


def check_cover_rule(png):
    """封面专用：分隔线上下留白应相等"""
    im = Image.open(png).convert("L")
    a = np.asarray(im, dtype=np.uint8)
    scale = im.width / SLIDE_W_IN
    dark = a < 170
    rows = np.where(dark.any(axis=1))[0]
    if rows.size == 0:
        return []
    bands, cur = [], [rows[0]]
    for r in rows[1:]:
        if r - cur[-1] <= 2:
            cur.append(r)
        else:
            bands.append((cur[0], cur[-1]))
            cur = [r]
    bands.append((cur[0], cur[-1]))

    text = [(t, b) for t, b in bands if b - t + 1 >= 6]
    rules = [(t, b) for t, b in bands if b - t + 1 < 6]
    if not rules or len(text) < 2:
        return []
    rt, rb = rules[-1]
    above = [b for _, b in text if b < rt]
    below = [t for t, _ in text if t > rb]
    if not above or not below:
        return []
    up = (rt - above[-1]) / scale * 72
    down = (below[0] - rb) / scale * 72
    if abs(up - down) > 8:
        return [f"分隔线上方 {up:.0f}pt / 下方 {down:.0f}pt，不对称"]
    return []


def _load_style(config):
    """读 content.yaml 的 style 段；没有配置就返回 None。"""
    if not config or not os.path.exists(config):
        return None
    import yaml
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    sys.path.insert(0, root)
    from pptkit.style import Style
    with open(config, encoding="utf-8") as f:
        return Style.from_config((yaml.safe_load(f) or {}).get("style"))


def check_preview_fresh(pptx, pdir, n):
    """检查预览图是否与稿子同步，返回 (缺失页, 过期页)。

    审计量的是像素。图比 pptx 旧，量出来的就是上一版版面——
    稿子改坏了也照样「通过」，这比不审计更糟。
    """
    t_ppt = os.path.getmtime(pptx)
    missing, stale = [], []
    for i in range(1, n + 1):
        png = os.path.join(pdir, f"slide{i}.png")
        if not os.path.exists(png):
            missing.append(i)
        elif os.path.getmtime(png) + 1 < t_ppt:      # 1 秒容差，避开同秒写入
            stale.append(i)
    return missing, stale


def main():
    global SLIDE_W_IN, SLIDE_H_IN, BODY_TOP
    pptx = sys.argv[1] if len(sys.argv) > 1 else "汇报材料/生物塑料汇报_仿生分层膜.pptx"
    pdir = sys.argv[2] if len(sys.argv) > 2 else "汇报材料/preview"
    config = sys.argv[3] if len(sys.argv) > 3 else None

    prs = Presentation(pptx)
    SLIDE_W_IN = Emu(prs.slide_width).inches
    SLIDE_H_IN = Emu(prs.slide_height).inches
    st = _load_style(config)
    if st is not None:
        BODY_TOP = st.body_top
    total = 0
    print("=" * 76)
    print("版式验收（全部页面 × 九个维度）")
    print("=" * 76)

    missing, stale = check_preview_fresh(pptx, pdir, len(prs.slides))
    if missing or stale:
        print()
        print("预览图与稿子不同步，本次审计不成立：")
        if missing:
            print("  缺图：第 " + "、".join(map(str, missing)) + " 页")
        if stale:
            print("  旧图：第 " + "、".join(map(str, stale)) + " 页（早于 pptx 写入时间）")
        print()
        print("先导出预览图再审计：")
        print(f"  tools/export_preview.ps1  -Ppt \"{pptx}\" -Out \"{pdir}\"")
        return 2

    for si, slide in enumerate(prs.slides, 1):
        items = shapes_of(slide)
        findings = []
        findings += [("越界", m) for m in check_bounds(items)]
        findings += [("重叠", m) for m in check_overlap(items)]
        findings += [("图片贴合", m) for m in check_image_fit(items)]
        findings += [("框内居中", m) for m in check_panel_centering(items)]

        png = os.path.join(pdir, f"slide{si}.png")
        if os.path.exists(png):
            p, o = render_metrics(png, items)
            findings += [("卡内留白", m) for m in p]
            findings += [("孤行", m) for m in o]
            findings += [("栏内填充", m) for m in check_column_fill(png, items)]
            findings += [("卡片间距", m) for m in check_card_gaps(items)]
            if si == 1:
                findings += [("封面分隔线", m) for m in check_cover_rule(png)]
            else:
                findings += [("段间距", m) for m in check_gaps(png, items)]

        status = "通过" if not findings else f"{len(findings)} 项异常"
        print(f"\n── 第 {si} 页  {status}")
        for dim, msg in findings:
            print(f"   [{dim}] {msg}")
            total += 1

    print()
    print("=" * 76)
    print("全部通过。" if total == 0 else f"共 {total} 项异常，需逐项修正。")
    print("=" * 76, flush=True)
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
