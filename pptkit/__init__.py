"""PPT 引擎 —— 与具体汇报内容无关的通用能力。

模块划分：
    theme       字体、字号下限、黑白配色、画布尺寸
    metrics     文本度量（按字体估算换行数与块高）
    engine      Deck：页面、文本、矩形、图片框（自动贴合宽高比）
    pages       页面类型注册表（cover / section / fields_image / ...）
    template    deck 配置解析与组装（兼容早期五段式配置）
    compliance  合规核对（页数 / 字体 / 字号 / 配图 / 内容项）
"""

from .style import Style
from .theme import (
    CN_FONT, EN_FONT, MIN_PT, LINE_H, CJK_W, ASCII_W,
    INK, INK_2, RULE, FILL, PAPER,
    SLIDE_W, SLIDE_H, TITLE_PT, COVER_PT, COVER_TITLE_PT,
)
from .metrics import est_lines, block_h
from .engine import Deck
from .pages import PAGES, render
from .template import build_deck, build_report, resolve_pages
from .compliance import check

__all__ = [
    "Style", "CN_FONT", "EN_FONT", "MIN_PT", "LINE_H", "CJK_W", "ASCII_W",
    "INK", "INK_2", "RULE", "FILL", "PAPER",
    "SLIDE_W", "SLIDE_H", "TITLE_PT", "COVER_PT", "COVER_TITLE_PT",
    "est_lines", "block_h", "Deck", "check",
    "PAGES", "render", "build_deck", "build_report", "resolve_pages",
]
