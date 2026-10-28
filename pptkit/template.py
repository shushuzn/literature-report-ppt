"""Deck 组装：把内容配置解析成页面列表，再交给 pages 里的版式函数。

两种配置格式都支持：

1. **新格式（推荐）**——在 content.yaml 里显式写出页面序列：

       deck:
         - type: cover
           title_cn: ...
         - type: section
           no: "01"
           title: 研究背景
         - type: fields_image
           title: 文献信息
           image: figs/page1.png
           fields: [...]

   页数、顺序、版式全由这份列表决定，与具体文献无关。

2. **旧格式（LEAFF 等早期项目）**——只写 cover / literature / mechanism /
   performance / application 五段，由 `legacy_pages()` 自动转成上表的五页。
   保留它是为了老项目不必改配置，新项目不要再用。

不再有「五页」这个概念。页数上限由 `compliance.max_pages` 控制，
不写则取 deck 的实际页数。
"""
import yaml

from .engine import Deck
from .pages import render
from .style import Style

LEGACY_ORDER = ("cover", "literature", "mechanism", "performance", "application")


def _cards_from_legacy(section: dict):
    """mechanism 段：components 列表 + 可选 process（无标题的工艺卡）。"""
    cards = [{"name": comp["name"], "desc": comp["desc"]}
             for comp in section.get("components") or []]
    if section.get("process"):
        cards.append({"desc": section["process"]})
    out = dict(section)
    out.pop("components", None)
    out.pop("process", None)
    out["cards"] = cards
    out["type"] = "image_cards"
    return out


def legacy_pages(cfg: dict):
    """旧五段式配置 → deck 页面列表。缺哪段就少哪页，不强制五页。"""
    pages = []
    if "cover" in cfg:
        pages.append({"type": "cover", **cfg["cover"]})
    if "literature" in cfg:
        pages.append({"type": "fields_image", **cfg["literature"]})
    if "mechanism" in cfg:
        pages.append(_cards_from_legacy(cfg["mechanism"]))
    if "performance" in cfg:
        pages.append({"type": "image_grid_cards", **cfg["performance"]})
    if "application" in cfg:
        pages.append({"type": "image_cards", **cfg["application"]})
    return pages


def resolve_pages(cfg: dict):
    """取 deck 页面列表；新格式优先，回落到旧五段式。"""
    pages = cfg.get("deck")
    if pages:
        if not isinstance(pages, list):
            raise ValueError("deck 必须是列表，每项是一页的配置")
        return pages
    pages = legacy_pages(cfg)
    if not pages:
        raise ValueError(
            "content.yaml 里既没有 deck 列表，也没有 cover/literature/... 五段之一")
    return pages


def build_deck(cfg: dict, base_dir: str) -> Deck:
    """按页面列表生成完整 PPT。样式来自 content.yaml 的 style 段。"""
    deck = Deck(style=Style.from_config(cfg.get("style")))
    pages = resolve_pages(cfg)
    for i, page in enumerate(pages, 1):
        render(deck, page, base_dir, i)
    return deck


# 向后兼容：早期工具脚本与文档里用的是 build_report
build_report = build_deck


def load(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)
