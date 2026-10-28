"""默认样式的别名。

实际生效的样式是 `pptkit.style.Style` 的实例，由 `content.yaml` 的 `style:`
段覆盖、挂在 `deck.style` 上。本文件只做两件事：

1. 给 `Style()` 的默认值起好记的名字，供文档与旧脚本引用；
2. 保持 `from pptkit.theme import MIN_PT` 这类老写法能用——这些名字是
   **默认值快照**，不随配置变化。需要「当前生效的字号」时请用
   `deck.style.body_pt`，不要读这里的常量。
"""
from .style import Style

DEFAULT_STYLE = Style()

# ---- 字体 ----
CN_FONT = DEFAULT_STYLE.cn_font
EN_FONT = DEFAULT_STYLE.en_font

# ---- 字号 ----
MIN_PT = DEFAULT_STYLE.min_pt
BODY_PT = DEFAULT_STYLE.body_pt
TITLE_PT = DEFAULT_STYLE.title_pt
COVER_TITLE_PT = DEFAULT_STYLE.cover_title_pt
COVER_PT = DEFAULT_STYLE.cover_pt
NOTES_PT = DEFAULT_STYLE.notes_pt

# ---- 文本度量系数 ----
LINE_H = DEFAULT_STYLE.line_h
CJK_W = 1.00    # 中日韩字符宽 / 字号
ASCII_W = 0.52  # 西文数字宽 / 字号（Times New Roman 均值；按字符类别细分的系数在 metrics 里）

# ---- 黑白配色（不含色相，R=G=B）----
INK = DEFAULT_STYLE.ink
INK_2 = DEFAULT_STYLE.ink_2
RULE = DEFAULT_STYLE.rule
EDGE = DEFAULT_STYLE.edge
FILL = DEFAULT_STYLE.fill
PAPER = DEFAULT_STYLE.paper

# ---- 画布（英寸，16:9）----
SLIDE_W = DEFAULT_STYLE.slide_w
SLIDE_H = DEFAULT_STYLE.slide_h

# ---- 版心 ----
MARGIN_X = DEFAULT_STYLE.margin_x
HEADER_Y = DEFAULT_STYLE.header_y
RULE_Y = DEFAULT_STYLE.rule_y
BODY_TOP = DEFAULT_STYLE.body_top
BODY_BOTTOM = DEFAULT_STYLE.body_bottom

# ---- 文字卡片 ----
PAD_X = DEFAULT_STYLE.pad_x
PAD_V = DEFAULT_STYLE.pad_v
CARD_GAP_PT = DEFAULT_STYLE.card_gap_pt
