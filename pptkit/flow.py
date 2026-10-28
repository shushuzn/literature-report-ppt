"""文字流检查：每页的文本块各占多宽、换几行、末行剩几个字。

单独成模块，是因为同一套判断要服务两处，而两处必须说同一句话：

1. `tools/check_text.py` 孤行预检——生成前拦下会断出「孤行」的文案；
2. `pages.choose_body_pt` 字号自适应——往上试探字号时不能把好好的
   三行文字撑成「四行、末行两个字」。

此前两处各算各的：预检按配置里的 `body_pt` 算，排版却按自适应后的
字号排。字号一涨，预检报通过，成品里照样掉单字。宽度算错只是差一点，
字号取错则是整页判错，所以宽度与字号都必须同源。

可用宽度按版式规则推：栅格页取单元格内宽，左图右文页取右栏内宽
（由图片比例决定），满页文字取版心宽。换页数、换图都不必改这里。
"""
from .metrics import est_lines, last_line_ratio

# 末行占比低于此值视为孤行。0.18 对应「一行按 20 个字算，末行不足
# 四个字」——三个字以下在投影上看就是吊着的一个点。
ORPHAN_RATIO = 0.18


def _cells(page, n_col, st):
    """栅格单元格内宽（去掉左右内边距）"""
    cw = (st.slide_w - 2 * st.margin_x - st.grid_gap * (n_col - 1)) / n_col
    return cw - 2 * st.pad_x


def _right_col(base, page, st):
    """左图右文页的右栏内宽；无图返回 None"""
    if not page.get("image"):
        return None
    # 延迟导入：pages 反过来要用本模块判孤行，放在文件头会成环。
    from .pages import column_split
    _, _, rw = column_split(base, page["image"], style=st)
    return rw - 2 * st.pad_x


def page_texts(base, page, st):
    """产出 (标签, 文本, 可用宽度英寸)。

    只收会换行的正文块：标题、眉标、年代这些短标签不可能断出孤行，
    收进来只会让报告变长。
    """
    out = []

    def add(tag, text, width):
        if text is None:
            return
        if isinstance(text, list):
            for t in text:
                add(tag, t, width)
            return
        out.append((tag, str(text), width))

    t = page.get("type")
    if t == "cover":
        add("封面 主标题", page.get("title_cn"), st.slide_w * 0.88)
        add("封面 副标题", page.get("title_en"), st.slide_w * 0.88)
    elif t == "section":
        add("章节 导语", page.get("lead"), st.slide_w * 0.88)
    elif t == "fields_image":
        w = _right_col(base, page, st) or st.inner_w
        for f in page.get("fields") or []:
            add(f"字段 {f['label']}", f.get("value"), w)
    elif t == "image_cards":
        w = _right_col(base, page, st) or st.inner_w
        for it in page.get("cards") or page.get("items") or []:
            if isinstance(it, str):
                add("卡片", it, w)
            else:
                add(f"卡片 {(it.get('name') or '')[:10]}",
                    it.get("desc") or it.get("body"), w)
    elif t == "image_grid_cards":
        w = _cells(page, int(page.get("n_col", 2)), st)
        for it in page.get("items") or []:
            name = it.get("name") or ""
            if it.get("data"):
                add(f"{name} 数据", f"{name}　{it['data']}", w)
            if it.get("cause"):
                add(f"{name} 成因", "成因：" + it["cause"], w)
            if it.get("body"):
                add(f"{name} 正文", it["body"], w)
    elif t in ("cards_grid", "cards"):
        w = _cells(page, int(page.get("columns", page.get("n_col", 1))), st)
        for it in page.get("cards") or page.get("items") or []:
            if isinstance(it, str):
                add("卡片", it, w)
            else:
                add(f"卡片 {(it.get('name') or '')[:10]}",
                    it.get("desc") or it.get("body"), w)
    elif t == "bullets":
        add("导语", page.get("lead"), st.inner_w)
        for it in page.get("items") or []:
            if isinstance(it, str):
                add("要点", it, st.inner_w - 0.3)
            else:
                add(f"要点 {(it.get('name') or '')[:10]}",
                    it.get("desc") or it.get("body"), st.inner_w - 0.3)
    elif t == "timeline":
        # 节点列宽由节点个数决定：轨数不同、节点数不同，宽度就不同。
        tracks = page.get("tracks") or []
        n = max(len(tr.get("nodes") or []) for tr in tracks) or 1
        lab_w = st.tl_label_w if any(tr.get("name") for tr in tracks) else 0.0
        x0 = st.margin_x + lab_w
        w = (st.slide_w - st.margin_x - x0) / n - 0.20
        for tr in tracks:
            for nd in tr.get("nodes") or []:
                add(f"{nd.get('year') or ''} 事件", nd.get("head"), w)
                add(f"{nd.get('year') or ''} 说明", nd.get("desc"), w)
    elif t == "stages":
        cx = st.margin_x + st.stage_year_w + st.stage_pad
        w = st.slide_w - st.margin_x - cx
        for r in page.get("stages") or []:
            add(f"{(r.get('year') or '')[:10]} 阶段", r.get("name"), w)
            add(f"{(r.get('year') or '')[:10]} 说明", r.get("desc"), w)
            add(f"{(r.get('year') or '')[:10]} 代价", r.get("cost"), w)
    return out


def orphan_rows(base, page, st, pt=None):
    """该页所有会换行的文本块及其末行占比，按占比升序。

    pt 为该页最终字号；不给则取 style.body_pt。
    """
    size = float(pt or st.body_pt)
    rows = []
    for tag, text, width in page_texts(base, page, st):
        if not text:
            continue
        n = est_lines(text, size, width, st.line_safety)
        if n < 2:
            continue
        rows.append((tag, n, last_line_ratio(text, size, width,
                                            st.line_safety), text))
    return sorted(rows, key=lambda r: r[2])


def has_orphan(base, page, st, pt=None):
    """该字号下是否存在孤行。字号自适应用它否掉「排得下但难看」的一档。"""
    limit = ORPHAN_RATIO if st.orphan_ratio <= 0 else st.orphan_ratio
    return any(r[2] < limit for r in orphan_rows(base, page, st, pt))
