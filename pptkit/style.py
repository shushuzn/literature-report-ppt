"""样式：一个汇报长什么样，全在这个对象里。

此前这些参数是 `theme.py` 里的模块级常量，被 30 多处 `from .theme import X`
按值绑定，运行时改不动——换个字号要改代码，换所学校要改代码。现在改成
一个可构造的对象：引擎、版式、合规检查都从 `deck.style` 取，`content.yaml`
的 `style:` 段覆盖哪一项，就只变那一项。

三层归属，别混：

| 层 | 例子 | 谁定 |
|---|---|---|
| 这一次的硬约束 | 字号下限、指定字体、画布比例 | 当次要求，写 `content.yaml` 的 `style:` |
| 这一次的版面选择 | 标题 28 还是 32、卡片几条 | 做的人 |
| 引擎内部实现 | 卡片间距上下限、图片外框留白 | 这里给默认值，少数情况才调 |

下面的默认值只是**起点**，不是标准：某次要求 14pt 就是 14pt。
不写 `style:` 段时，行为与 2026-10 那次汇报逐像素一致。
"""
from dataclasses import dataclass, fields, replace
from typing import Optional

from pptx.dml.color import RGBColor


def _rgb(v):
    """'000000' / '#000000' / RGBColor → RGBColor"""
    if isinstance(v, RGBColor):
        return v
    s = str(v).lstrip("#").strip()
    if len(s) != 6:
        raise ValueError(f"颜色要写成 6 位十六进制，收到 {v!r}")
    return RGBColor.from_string(s.upper())


@dataclass(frozen=True)
class Style:
    # ---- 字体 ----
    cn_font: str = "微软雅黑"
    en_font: str = "Times New Roman"

    # ---- 字号（pt）----
    body_pt: float = 20.0        # 正文、卡片、字段、图注
    min_pt: float = 20.0         # 硬性下限，低于此值为违规
    title_pt: float = 30.0       # 页标题
    cover_title_pt: float = 38.0  # 封面主标题
    cover_pt: float = 20.0       # 封面其余文字
    section_no_pt: float = 84.0  # 章节页序号
    section_title_pt: float = 40.0  # 章节页标题
    compare_label_pt: float = 24.0  # 对照页栏目标签
    bullets_lead_pt: float = 24.0   # 要点页导语
    bullets_align: bool = True      # 要点页「名称列」对齐：说明文字从同一列起排
    bullets_hang_max: float = 0.42  # 名称列最多占版心宽度的比例，超出则退回自然排
    notes_pt: float = 14.0       # 演讲者备注（不投影，不受下限约束）

    # ---- 字号自适应（按页开关）----
    # 打开后，每一页在 [auto_pt_min, auto_pt_max] 里挑「内容刚好放得下」的
    # 最大字号：字少的页自动排大一点，撑满版心；字多的页退回小一档，不溢出。
    # 全局默认写在 style 段，单页写 `auto_pt: false` 可以关掉这一页。
    # 0 表示沿用默认：min 取 min_pt，max 取 body_pt + 4。
    auto_pt: bool = False
    auto_pt_min: float = 0.0
    auto_pt_max: float = 0.0
    auto_pt_step: float = 1.0    # 试探步长（pt）
    # 试探时不许把文字撑出孤行（末行只剩一两个字）。放得下不等于排得好：
    # 加一个字号常常刚好让末行掉两个字，此时退一档更好看。写 false 只求
    # 放得下、不管末行，留给人工判断。
    auto_pt_orphan: bool = True
    orphan_ratio: float = 0.18   # 末行占比低于此值算孤行

    # ---- 行高 ----
    line_h: float = 1.24          # 行高 / 字号
    # 行数估算的宽度折减系数：1.0 按理论宽度算，越小越保守（估得行数越多）。
    # 0.97 适合字号大、西文多的稿子；中文密排、字号较小时 0.94 更稳，
    # 否则会出现「估算一行、渲染两行」，文字顶到卡片框线。
    line_safety: float = 0.97

    # ---- 画布（英寸）----
    slide_w: float = 13.333
    slide_h: float = 7.5

    # ---- 版心 ----
    margin_x: float = 0.42
    header_y: float = 0.26
    rule_y: float = 0.88
    body_top: float = 1.10
    body_bottom: float = 7.08

    # ---- 卡片与栅格 ----
    pad_x: float = 0.22
    pad_v: float = 0.18
    card_gap_pt: float = 10.0    # 卡内标题与正文的段前距
    card_line_w: float = 0.75    # 卡片边框线宽
    grid_gap: float = 0.20       # 栅格列间距
    card_min_gap: float = 0.14   # 卡片竖排的最小间隔
    card_max_gap: float = 0.42   # 卡片竖排的最大间隔
    # 内容放不下时，是否自动收紧卡内留白与卡间距（收到下限仍放不下才告警）。
    # 默认开：留白是弹性量，内容多一页能多容纳两三行，不让字数为留白让路。
    # 关掉则放不下只告警、照 pad_v 原样画，用于要复刻旧版版面的场合。
    cards_fit: bool = True
    # 内容偏少时，把剩余栏高摊给卡片，每张最多加这么高（英寸）。
    # 卡片是为内容服务的框，不是填满栏高的容器：一行正文的卡被抻到 1.3in，
    # 框里就是大片空白。字号调小、行数变少时，这条决定卡会不会被吹起来。
    card_grow_max: float = 0.30
    # 栅格版式里每行卡片最多能长多高。行数少（2×2）时余量全落在这里：
    # 上限低于余量，页面底部就空一条；高到把三行字的卡抻成两倍高，框里
    # 又只剩留白。0.6 是「框内留白翻一倍」的量级。
    grid_grow_max: float = 0.60

    # ---- 左图右文 ----
    col_r_min: float = 6.00      # 右栏最小宽度
    col_gap: float = 0.21        # 左右栏间距
    cap_h: float = 0.32          # 图注行高
    cap_top_gap: float = 0.04    # 图注与图框的间距
    img_pad: float = 0.06        # 图片外框留白
    # 字段之间的段前距（pt）。下限是「读得清是两组」，上限是「别散成两页」：
    # 字段少时若按剩余高度均摊，间隔能涨到 32pt 以上，比行高（约 20pt）还
    # 大一倍，看着像每行各占一段。到了上限就把整栏居中，不再拉间距。
    field_gap: float = 14.0
    field_gap_max: float = 22.0

    # ---- 时间轴（综述的「发展过程」页）----
    tl_label_w: float = 1.30     # 每轨左侧的行名宽度（轨名坐在这里，像行标题）
    tl_tick: float = 0.075       # 节点方块边长，压在轴线上
    tl_rail_w: float = 0.012     # 轴线粗细
    tl_head_gap: float = 0.14    # 轴线到节点文字的距离
    tl_track_gap: float = 0.42   # 两条轨之间的最小间距
    # 两条轨之间的最大间距。内容少时余量会落进这个间隔：封顶太低，余量
    # 全挤到页底（两轨的时间轴下面空三分之一页）；封顶太高，两条线离得
    # 远，像两张图。这个值决定「整块撑多开」。
    tl_track_gap_max: float = 1.10

    # ---- 阶段带（同一条线的逐级演进）----
    stage_year_w: float = 1.62   # 年代列宽（「1990–2010」22pt 约 1.38in，留余量）
    stage_year_pt: float = 22.0  # 年代字号（比正文大一档，做左侧锚点）
    stage_pad: float = 0.26      # 年代列与竖线、竖线与正文的间距
    stage_spine_w: float = 0.012  # 竖线粗细
    stage_gap: float = 0.24      # 阶段之间的最小间距
    stage_gap_max: float = 0.75  # 阶段之间的最大间距（余量落在这里）

    # ---- 配色（默认黑白，不含色相）----
    ink: RGBColor = RGBColor(0x00, 0x00, 0x00)
    ink_2: RGBColor = RGBColor(0x40, 0x40, 0x40)
    rule: RGBColor = RGBColor(0x9A, 0x9A, 0x9A)
    edge: RGBColor = RGBColor(0xBF, 0xBF, 0xBF)
    fill: RGBColor = RGBColor(0xF2, 0xF2, 0xF2)
    paper: RGBColor = RGBColor(0xFF, 0xFF, 0xFF)

    # ---- 派生量 ----

    @property
    def lh(self) -> float:
        """正文单行高（英寸）。改 body_pt 时行高跟着变，不能另设一个常量。"""
        return self.body_pt * self.line_h / 72.0

    @property
    def lh_cover(self) -> float:
        """封面信息行高（英寸）。

        封面除主标题外都用 `cover_pt`，它和 `body_pt` 不是一个数：正文降到
        16pt 时封面仍按 20pt 排。行高若跟着 `lh`（按 body_pt 算），文本框
        就比字矮一截，字溢出到框外，分隔线上下的留白一边厚一边薄。
        """
        return self.cover_pt * self.line_h / 72.0

    @property
    def content_bottom(self) -> float:
        """带图注的栏可用到的底端"""
        return self.body_bottom - self.cap_h - 0.10

    @property
    def inner_w(self) -> float:
        """整页版心内宽（去掉左右页边距与卡片内边距）"""
        return self.slide_w - 2 * self.margin_x - 2 * self.pad_x

    # ---- 构造 ----

    @classmethod
    def from_config(cls, cfg: Optional[dict] = None) -> "Style":
        """从 content.yaml 的 style 段构造；只覆盖写了的项。"""
        raw = cfg or {}
        known = {f.name for f in fields(cls)}
        bad = set(raw) - known
        if bad:
            raise ValueError(
                f"style 段有未知字段：{'、'.join(sorted(bad))}。"
                f"可用字段见 docs/使用说明.md「样式怎么改」")
        kw = {}
        for k, v in raw.items():
            if k in ("ink", "ink_2", "rule", "edge", "fill", "paper"):
                kw[k] = _rgb(v)
            elif k in ("cn_font", "en_font"):
                kw[k] = str(v)
            elif k in ("cards_fit", "auto_pt", "auto_pt_orphan", "bullets_align"):
                kw[k] = bool(v)
            else:
                kw[k] = float(v)
        return replace(cls(), **kw)

    def scaled(self, **kw) -> "Style":
        """派生一份改了几个字段的样式，原对象不动。"""
        return replace(self, **kw)
