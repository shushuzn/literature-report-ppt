"""页面类型注册表。

每一页由内容配置里的一条记录决定：

    deck:
      - type: fields_image
        title: 汇报文献信息
        image: figs/page1.png
        fields: [...]

`PAGES` 把类型名映射到绘制函数，`template.build_deck` 按配置顺序调用。
新增一种版式 = 在这里加一个函数 + 注册一个名字，不动其他页面。

设计边界（重要）：
  · 版式函数只读取 page 字典与 deck.style，不读全局状态、不读项目名
  · 所有纵向位置都由内容高度反推或由 distribute() 均分，
    不写死 y 值（封面与章节页除外，它们本来就该定死）
  · 版心、字号、配色一律来自 `deck.style`（`pptkit.style.Style`），
    页面里不出现魔法数字的颜色

可用类型：
  cover              封面（眉标 / 主题 / 英文副题 / 汇报人 / 指导老师 / 出处）
  section            章节过渡页（序号 + 章节名 + 导语）
  compare            左右两栏对照（旧范式 / 新范式、优点 / 代价）
  fields_image       左图 + 右字段列表（文献信息、参数表）
  image_cards        左图 + 右卡片竖排（机理、应用）
  image_grid_cards   上方一行图 + 下方卡片栅格（性能）
  cards_grid         纯卡片栅格（挑战、结论、展望），columns=1 即整栏竖排
  bullets            要点列表，item 可为字符串或 {name, desc}
  timeline           时间轴，若干条并行轨、每轨一串年代节点（综述的「发展过程」）
  stages             阶段带，同一条线逐级演进（年代 + 阶段名 + 说明）

研究论文与综述的差别落在版式上：研究论文讲「一篇工作做了什么」，
用卡片与左右对照；综述讲「这个领域怎么走到今天」，需要先后关系，
timeline / stages 就是为此加的两种。

任意页面都可加 `notes:`，内容写进 PowerPoint 的演讲者备注（讲稿）。
"""
from dataclasses import replace as _replace

from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches

from .engine import Deck
from .flow import has_orphan
from .metrics import est_lines, distribute, text_width


# ------------------------------------------------------------
# 公共小工具
# ------------------------------------------------------------
def _as_lines(v):
    """字段值可以是 str，也可以是 list[str]（需要手动断行时用）"""
    if v is None:
        return [""]
    return v if isinstance(v, list) else [v]


def column_split(base: str, image: str, r_min=None, style=None):
    """定左栏宽与右栏位置。

    左栏宽受两个条件约束，取较小者：

    1. **图片按真实比例铺满可用高度所需的宽度。** 左栏若写死宽度，比例偏宽
       的图会被宽度卡住、高度用不满，图下方留下大片空白（实测过 0.73 英寸）。
    2. **不能把右栏挤窄。** 只按图片算会出事：图越宽、左栏越宽，右栏文字被
       压缩。实测过右栏内宽从 390pt 掉到 370pt，同一句话就多出一个换行，
       工艺行断出只剩 2% 宽的末行。故给右栏留下限（`style.col_r_min`，
       默认 6.0in）——20pt 下约可容 20 个汉字，是文案不断行的底线。

    因第 2 条而被截短时，图片填不满高度，`image_framed` 会把图连图注在栏内
    垂直居中，上下均分余量。

    返回 (左栏宽, 右栏 x, 右栏宽)。
    """
    from .style import Style
    st = style or Style()
    r_min = st.col_r_min if r_min is None else r_min
    avail_h = (st.content_bottom - st.body_top) - st.cap_h - st.cap_top_gap
    iw, _ = Deck.image_size(f"{base}/{image}", 99.0, avail_h - 2 * st.img_pad)
    left_max = st.slide_w - 2 * st.margin_x - st.col_gap - r_min
    left_w = min(left_max, iw + 2 * st.img_pad)
    right_x = st.margin_x + left_w + st.col_gap
    return left_w, right_x, st.slide_w - st.margin_x - right_x


def _left_image(deck: Deck, page: dict, base: str):
    """左图右文的公共左半：图按比例贴合外框，返回 (右栏 x, 右栏宽)。"""
    st = deck.style
    left_w, rx, rw = column_split(base, page["image"], style=st)
    avail_h = (st.content_bottom - st.body_top) - st.cap_h - st.cap_top_gap
    deck.image_framed(f"{base}/{page['image']}", st.margin_x, st.body_top,
                      left_w, avail_h,
                      caption=page.get("caption"), caption_h=st.cap_h)
    return rx, rw


def _cards_of(page: dict):
    """卡片列表：支持 cards 或 items；元素可为 {name, desc} 或 {head, body}。"""
    raw = page.get("cards") or page.get("items") or []
    out = []
    for it in raw:
        if isinstance(it, str):
            out.append((None, it))
            continue
        head = it.get("name") or it.get("head")
        body = it.get("desc") or it.get("body") or ""
        out.append((head, body))
    return out


def _stack_cards(deck: Deck, page: dict, cards, x, w, top=None,
                 bottom=None, tag="卡片"):
    """在 [top, bottom] 内竖排若干卡片，间隔在 min/max 之间均分。

    内容偏少时间隔会顶到 max_gap 就不再长，底部于是空出一块（实测 0.9in，
    审计的第 9 维「栏内填充」会报）。此时把剩余高度摊给每张卡片，让卡片
    自己变高——内容在卡内垂直居中，看起来是留白变松，而不是底部缺一块。

    内容偏多时反过来：先收紧卡内上下留白，再收紧卡间距，都压到下限仍放
    不下才告警。此前一律用 `style.pad_v`（0.18in），四张卡光留白就占
    1.44in，多写两行正文就溢出——这正是「一页放不了多少字」的由来。
    留白是弹性量，字数是内容量，撑不下时该让留白退。
    """
    st = deck.style
    top = st.body_top if top is None else top
    bottom = st.content_bottom if bottom is None else bottom
    avail = bottom - top
    n = len(cards)

    # 下限 0.14in。0.12in 时渲染图实测文字离框线只有 7pt（字形上沿的空白
    # 不占版面，但视觉上就是贴线），版式审计会判「顶到框线」。留白可以收紧，
    # 不能收到贴着线——真放不下就该减条目，而不是继续压留白。
    pad_lo = min(st.pad_v, max(0.14, st.pad_v * 0.78))
    gap_lo = st.card_min_gap * 0.55
    fitted = None
    for pad_v in ((st.pad_v, st.pad_v * 0.8, pad_lo) if st.cards_fit
                  else (st.pad_v,)):
        heights = [deck.card_height(h, b, w, pad_v=pad_v) for h, b in cards]
        for gap in ((st.card_min_gap, gap_lo) if st.cards_fit
                    else (st.card_min_gap,)):
            need = sum(heights) + gap * max(0, n - 1)
            if need <= avail + 1e-6:
                fitted = (pad_v, gap, heights)
                break
        if fitted:
            break

    if fitted is None:
        # 放不下：开了 cards_fit 就压到下限再画；没开则按原样画、只告警，
        # 保证老项目版面不变。
        pad_v, gap = ((pad_lo, gap_lo) if st.cards_fit
                      else (st.pad_v, st.card_min_gap))
        heights = [deck.card_height(h, b, w, pad_v=pad_v) for h, b in cards]
        deck.check_fit(top, bottom, heights, gap, tag)
    else:
        pad_v, gap, heights = fitted

    leftover = avail - sum(heights) - gap * max(0, n - 1)
    if leftover > 0.06 and n:
        # 撑高有上限。一行正文的卡被抻到 1.3in，框里就剩大片空白，
        # 看着像「忘了填内容」。多出来的高度不再摊给卡片。
        grow = min(leftover / n, st.card_grow_max)
        heights = [h + grow for h in heights]
    ys = distribute(top, bottom, heights, min_gap=gap,
                    max_gap=max(st.card_max_gap, gap))
    for (head, body), y, h in zip(cards, ys, heights):
        deck.card(x, y, w, h, head, body, pad_v=pad_v)


# ------------------------------------------------------------
# cover —— 封面：主题、英文副题、分隔线、汇报人、指导老师
# ------------------------------------------------------------
def cover(deck: Deck, page: dict, base: str):
    """封面：眉标、主题、英文副题、分隔线、汇报人与导师。

    整块按内容总高反推起点并垂直居中。此前把标题起点写死在 y=2.25，
    一旦加了眉标或增减信息行，整块就会偏上或偏下——内容一变就得重调数字。

    起止 x 与宽度按画布比例取（原来的 0.8 / 1.6 英寸是 16:9 下算出来的，
    换 4:3 会偏），故写成 `slide_w × 0.06` 与 `slide_w × 0.88`。
    """
    st = deck.style
    deck.new_slide()
    x, w = st.slide_w * 0.06, st.slide_w * 0.88
    TITLE_H = 0.80
    KICKER_H = 0.34
    GAP_KICKER = 0.26        # 眉标与主标题之间
    GAP_TITLE_EN = 0.22      # 主标题与英文标题之间
    GAP_RULE = 0.26          # 分隔线上下留白（必须相等）

    kicker = (page.get("kicker") or "").strip()
    en_lines = [t for t in _as_lines(page.get("title_en")) if t]
    lead = []
    if page.get("presenter"):
        lead.append(f"汇报人：{page['presenter']}")
    if page.get("advisor"):
        lead.append(f"指导老师：{page['advisor']}")
    # roles：并列的署名组，如 [[PPT 制作, PPT 审核], [资料收集, 资料审核]]。
    # 每行内部首列右对齐、其余列左对齐，整行以页面中线居中；只有一项的行
    # 直接居中，不会因为落在「左列」而偏到一边。长署名按估算行数加高，
    # 不靠固定行高硬塞。
    roles = [[str(c).strip() for c in row if str(c).strip()]
             for row in (page.get("roles") or [])]
    roles = [r for r in roles if r]
    tail = [t for t in _as_lines(page.get("meta")) if t]

    # 封面用 cover_pt，行高也要按 cover_pt 算：用 body_pt 的行高，正文一降字号
    # （比如 16pt）文本框就比字矮，字溢出框外，分隔线上下的留白就不对称了。
    LINE_INFO = st.lh_cover + 0.10   # 封面信息行距：正文行距搬到封面偏挤，多给一口气
    GUTTER = 0.36                    # 并列署名之间的固定中缝
    ncol = max((len(r) for r in roles), default=0)
    col_w = min(2.9, (w - GUTTER * (ncol - 1)) / ncol) if ncol else 0.0
    role_heights = []
    for row in roles:
        n = max(est_lines(cell, st.cover_pt, col_w, st.line_safety) for cell in row)
        role_heights.append(max(st.lh_cover, n * st.lh_cover))

    total = TITLE_H
    if kicker:
        total += KICKER_H + GAP_KICKER
    if en_lines:
        total += GAP_TITLE_EN + len(en_lines) * st.lh_cover
    total += GAP_RULE + 0.014 + GAP_RULE
    info_heights = ([st.lh_cover] * len(lead) + role_heights
                    + [st.lh_cover] * len(tail))
    if info_heights:
        total += sum(info_heights) + 0.10 * (len(info_heights) - 1)

    # 版面的视觉重心略高于几何中线，故上移 0.08in
    y = (st.slide_h - total) / 2 - 0.08

    if kicker:
        tb = deck.text(x, y, w, KICKER_H, [(kicker, st.body_pt, False, st.ink_2)],
                       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # 字距 4pt：眉标靠疏排与小字号拉开层级，不靠再小的字号
        tb.text_frame.paragraphs[0].runs[0]._r.get_or_add_rPr().set('spc', '400')
        y += KICKER_H + GAP_KICKER

    deck.text(x, y, w, TITLE_H,
              [(page["title_cn"], st.cover_title_pt, True, st.ink)],
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    y += TITLE_H

    if en_lines:
        y += GAP_TITLE_EN
        deck.text(x, y, w, len(en_lines) * st.lh_cover,
                  [(t, st.cover_pt, False, st.ink_2) for t in en_lines],
                  align=PP_ALIGN.CENTER, spacing=1.0)
        y += len(en_lines) * st.lh_cover

    y += GAP_RULE
    deck.rect(st.slide_w / 2 - 1.72, y, 3.43, 0.014, fill=st.rule)
    y += 0.014 + GAP_RULE

    for t in lead:
        deck.text(x, y, w, st.lh_cover, [(t, st.cover_pt, False, st.ink)],
                  align=PP_ALIGN.CENTER)
        y += LINE_INFO

    for row, rh in zip(roles, role_heights):
        if len(row) == 1:
            deck.text(st.slide_w / 2 - col_w / 2, y, col_w, rh,
                      [(row[0], st.cover_pt, False, st.ink)],
                      align=PP_ALIGN.CENTER)
        else:
            row_w = col_w * len(row) + GUTTER * (len(row) - 1)
            left = st.slide_w / 2 - row_w / 2
            for j, cell in enumerate(row):
                cx = left + j * (col_w + GUTTER)
                align = PP_ALIGN.RIGHT if j == 0 else PP_ALIGN.LEFT
                deck.text(cx, y, col_w, rh,
                          [(cell, st.cover_pt, False, st.ink)], align=align)
        y += rh + 0.10

    for t in tail:
        deck.text(x, y, w, st.lh_cover, [(t, st.cover_pt, False, st.ink)],
                  align=PP_ALIGN.CENTER)
        y += LINE_INFO


# ------------------------------------------------------------
# section —— 章节过渡页
# ------------------------------------------------------------
def section(deck: Deck, page: dict, base: str):
    """序号在左、标题与导语在右，一条竖线分隔。

    过渡页不画页眉：它是呼吸，不是内容页。

    序号字段叫 index 而不是 no —— YAML 1.1 把裸键 no/yes/on/off 解析成
    布尔值，写成 `no: "01"` 会得到 {False: "01"}，序号会静默消失。
    """
    st = deck.style
    deck.new_slide()
    x0 = st.margin_x + 0.10
    NO_W = 2.20
    GAP_V = 0.34

    no = str(page.get("index") or page.get("no") or "").strip()
    BLOCK_TOP = 3.02          # 数字块与标题块共用同一顶端，垂直重心落在页面中线
    if no:
        deck.text(x0, BLOCK_TOP, NO_W, 1.30,
                  [(no, st.section_no_pt, True, st.ink_2)],
                  align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    tx = x0 + NO_W + 0.30
    tw = st.slide_w - st.margin_x - tx
    deck.rect(x0 + NO_W, BLOCK_TOP + 0.07, 0.016, 1.16, fill=st.rule)

    title = page["title"]
    deck.text(tx, BLOCK_TOP + 0.07, tw, 0.72,
              [(title, st.section_title_pt, True, st.ink)],
              anchor=MSO_ANCHOR.MIDDLE)
    lead = [t for t in _as_lines(page.get("lead")) if t]
    if lead:
        h = len(lead) * st.lh
        deck.text(tx, BLOCK_TOP + 0.07 + 0.72 + GAP_V, tw, h,
                  [(t, st.body_pt, False, st.ink_2) for t in lead], spacing=1.0)


# ------------------------------------------------------------
# fields_image —— 左图 + 右字段（标签加粗、与首行同行）
# ------------------------------------------------------------
def _label_column(fields, inner, style):
    """标签列宽（磅）。返回 (列宽 pt, 每个标签后补的全角空格数)。

    标签与值同行时，值的左端位置由标签长度决定。各字段各算各的，标签
    长短不同，值就落在不同的 x 上——「分区」只有两个字，值比上下行左移
    近半个字宽，整列看着是错开的。

    故先取最长标签的宽度作为公共列宽，短标签用全角空格补齐，所有字段的
    值都从同一个 x 起排，折行的次行也退到这一列。

    标签列宽占掉右栏太多时（超过 42%）放弃对齐：宁可值参差，也不能让
    值那一列窄到每行只剩几个字。此时退回逐字段自然宽度。
    """
    natural = [text_width(f["label"] + "　", style.body_pt) for f in fields]
    col_pt = max(natural)
    if inner <= 0 or col_pt / 72.0 > inner * 0.42:
        return None, [1] * len(fields)
    pads = [max(1, round((col_pt - text_width(f["label"], style.body_pt))
                         / style.body_pt))
            for f in fields]
    return col_pt, pads


def _field_paras(f, pad=1):
    """字段渲染为一个或多个段落；标签与首行值同行、加粗区分。

    `pad` 是标签后补的全角空格数（含原本的一个分隔空格），由
    `_label_column` 统一给出，用来把值推到同一条竖线上。
    """
    vals = _as_lines(f["value"])
    paras = [[(f["label"] + "　" * max(1, pad), True), (vals[0], False)]]
    paras += [[(v, False)] for v in vals[1:]]
    return paras


def _field_h(f, inner, style, pad=1, hang=0.0):
    """字段高度：标签与首行合排，其余值行各占一行。

    续行段落带 marL（悬挂缩进），可用宽度是 `inner - hang`，不是整栏宽；
    按整栏估算会少算行数，卡片排下去才发现文字溢出。这里有 `hang` 就
    按真实可用宽度算。
    """
    vals = _as_lines(f["value"])
    n = est_lines(f["label"] + "　" * max(1, pad) + vals[0], style.body_pt, inner,
                  style.line_safety)
    for v in vals[1:]:
        n += est_lines(v, style.body_pt, inner - hang, style.line_safety)
    return n * style.lh


def fields_image(deck: Deck, page: dict, base: str):
    st = deck.style
    deck.new_slide()
    deck.header(page["title"])
    rx, rw = _left_image(deck, page, base)

    # 右：字段。整栏一个文本框，字段间隔用段前距保证恒定。
    #
    # 间隔不再「按剩余高度均摊到底」：字段少时那样算出来能到 32pt，比行高
    # 大一倍，每行看着各占一段。改为封顶 `field_gap_max`，封顶后仍有余量
    # 就把整栏垂直居中——余量变成上下对称的页边空白，而不是行与行之间的
    # 空隙。字段多时反过来，间隔收到 `field_gap` 下限。
    fields = page["fields"]
    inner = rw - 2 * st.pad_x
    col_pt, pads = _label_column(fields, inner, st)
    hang = (col_pt / 72.0) if col_pt else 0.0
    heights = [_field_h(f, inner, st, pad=p, hang=hang)
               for f, p in zip(fields, pads)]
    avail = st.content_bottom - st.body_top
    gap_pt = st.field_gap
    if len(fields) > 1 and sum(heights) < avail:
        gap_pt = (avail - sum(heights)) / (len(fields) - 1) * 72.0
        gap_pt = max(st.field_gap, min(st.field_gap_max, gap_pt))
    tb = deck.groups_text(rx + st.pad_x, st.body_top, inner, avail,
                          [_field_paras(f, p) for f, p in zip(fields, pads)],
                          gap_pt=gap_pt, anchor=MSO_ANCHOR.MIDDLE)

    # 悬挂缩进：值折行时次行要退到标签右侧，与值的左端对齐，否则次行会
    # 顶到标签下面，标签与值的从属关系就断了。
    # 每行值都是一个独立段落，所以首段用 indent=-marL 让首行回到 0，
    # 其余段落 indent=0——它们直接从 marL 起排，正好对齐首行的值。
    # marL 全栏统一（col_pt）：不统一时短标签的值会左移，整列参差。
    marl = int(Inches(hang)) if col_pt else None
    pi = 0
    paras = tb.text_frame.paragraphs
    for f in fields:
        if marl is None:
            hang_f = int(Inches(text_width(f["label"] + "　", st.body_pt) / 72.0))
        else:
            hang_f = marl
        for k, _ in enumerate(_as_lines(f["value"])):
            pPr = paras[pi]._p.get_or_add_pPr()
            pPr.set("marL", str(hang_f))
            pPr.set("indent", str(-hang_f if k == 0 else 0))
            pi += 1

    deck.check_fit(st.body_top, st.content_bottom, heights, gap_pt / 72.0,
                   page.get("tag", "字段列表"))


# ------------------------------------------------------------
# image_cards —— 左图 + 右卡片竖排（机理 / 应用）
# ------------------------------------------------------------
def image_cards(deck: Deck, page: dict, base: str):
    deck.new_slide()
    deck.header(page["title"])
    rx, rw = _left_image(deck, page, base)
    _stack_cards(deck, page, _cards_of(page), rx, rw,
                 tag=page.get("tag", "条目"))


# ------------------------------------------------------------
# image_grid_cards —— 上方一行图 + 下方卡片栅格（性能）
# ------------------------------------------------------------
def image_grid_cards(deck: Deck, page: dict, base: str):
    st = deck.style
    deck.new_slide()
    deck.header(page["title"])

    items = page["items"]
    n_col, gap = int(page.get("n_col", 2)), st.grid_gap
    cw = (st.slide_w - 2 * st.margin_x - gap * (n_col - 1)) / n_col

    # 标题行把「名」与「数据」并排，中间用全角空格分隔，与字段页的
    # 标注写法一致。不用「｜」——竖线把一行切成两半，看着生硬。
    def head_of(p):
        name, data = p.get("name") or p.get("head"), p.get("data")
        return f"{name}　{data}" if data else name

    def body_of(p):
        if p.get("body"):
            return p["body"]
        cause = p.get("cause")
        return ("成因：" + cause) if cause else ""

    rows = [items[i:i + n_col] for i in range(0, len(items), n_col)]
    row_h = [max(deck.card_height(head_of(p), body_of(p), cw) for p in r)
             for r in rows]
    grid_h = sum(row_h) + gap * (len(rows) - 1)

    IMG_CAP_GAP = 0.30                                          # 图片与卡片的间距
    img_h = (st.body_bottom - st.body_top) - grid_h - st.cap_h - IMG_CAP_GAP
    if img_h < 1.2:
        deck.warnings.append(
            f"P{deck.page} 卡片区占高 {grid_h:.2f}in，留给图片仅 {img_h:.2f}in，"
            f"建议精简条目或文案")
        img_h = max(img_h, 0.8)

    # 图片与下方的卡片共用同一套栅格，中线才对得齐。此前图片按左/右栏定列、
    # 卡片按两等分定列，两套栅格差了 0.3in，同一列上下错开。
    deck.image_row(
        [(f"{base}/{im['path']}", im.get("caption"),
          st.margin_x + cw / 2 + i * (cw + gap), cw)
         for i, im in enumerate(page.get("images") or [])],
        st.body_top, img_h)

    # 卡片一律白底细边框，不做灰白相间的棋盘格——交替填充没有依据，
    # 只会显得零碎。
    y = st.body_top + (img_h + st.cap_h + IMG_CAP_GAP if page.get("images") else 0)
    for pairs, hh in zip(rows, row_h):
        for col, p in enumerate(pairs):
            x = st.margin_x + col * (cw + gap)
            deck.card(x, y, cw, hh, head_of(p), body_of(p))
        y += hh + gap


# ------------------------------------------------------------
# cards_grid —— 纯卡片栅格（无图）；columns=1 即整栏竖排
# ------------------------------------------------------------
def cards_grid(deck: Deck, page: dict, base: str):
    st = deck.style
    deck.new_slide()
    deck.header(page["title"])

    cards = _cards_of(page)
    n_col = int(page.get("columns", page.get("n_col", 1)))
    if n_col <= 1:
        _stack_cards(deck, page, cards, st.margin_x, st.slide_w - 2 * st.margin_x,
                     tag=page.get("tag", "卡片"))
        return

    gap = st.grid_gap
    cw = (st.slide_w - 2 * st.margin_x - gap * (n_col - 1)) / n_col
    rows = [cards[i:i + n_col] for i in range(0, len(cards), n_col)]
    heights = [max(deck.card_height(h, b, cw) for h, b in r) for r in rows]
    avail = st.body_bottom - st.body_top
    deck.check_fit(st.body_top, st.body_bottom, heights, gap,
                   page.get("tag", "卡片"))

    # 内容偏少时不能只把行间距撑开——两行卡片会被拉到页面上下两端，
    # 中间空出一大块（实测 2.5in）。正确做法是间距封顶 0.45in，
    # 余量摊给行高让卡片自己变高；卡片高度也封顶，剩下的整体居中。
    extra = avail - sum(heights) - gap * (len(rows) - 1)
    if extra > 0 and len(rows) > 1:
        gap = min(0.45, gap + extra / (len(rows) - 1))
        extra = avail - sum(heights) - gap * (len(rows) - 1)
    grow = min(max(0.0, extra) / len(rows), st.grid_grow_max) if rows else 0.0
    heights = [h + grow for h in heights]

    used = sum(heights) + gap * (len(rows) - 1)
    y = st.body_top + max(0.0, (avail - used) / 2)
    for r, hh in zip(rows, heights):
        for col, (head, body) in enumerate(r):
            deck.card(st.margin_x + col * (cw + gap), y, cw, hh, head, body)
        y += hh + gap


# ------------------------------------------------------------
# compare —— 左右两栏对照（旧范式 / 新范式、优点 / 代价）
# ------------------------------------------------------------
def compare(deck: Deck, page: dict, base: str):
    """两栏对照：每栏「小标签 → 栏目名 → 若干要点」，中间一条竖线分隔。

    不用两张卡片并排：卡片自带边框，两块相靠时会多出两道边线夹一条缝，
    看着像两个盒子；这里只留一条居中竖线，版面更静，对照关系也更直接。
    """
    st = deck.style
    deck.new_slide()
    deck.header(page["title"])

    GAP = 0.72
    cw = (st.slide_w - 2 * st.margin_x - GAP) / 2
    avail = st.body_bottom - st.body_top

    PAD = 0.26
    TAG_H, LABEL_H, GAP_LABEL = 0.34, 0.46, 0.24
    inner = cw - 2 * PAD
    item_gap = float(page.get("item_gap_pt", 14.0))

    # 先把两栏各自的内容高算出来，取大者作为整块高度，再垂直居中。
    # 逐栏各画各的会出现「左栏占满、右栏下半页空白」，两栏不平衡。
    blocks = []
    for col in (page["left"], page["right"]):
        paras, heights = [], []
        for p in col.get("points") or []:
            if isinstance(p, str):
                runs = [("· ", False), (p, False)]
            else:
                runs = [("· ", False)]
                if p.get("name"):
                    runs.append((p["name"] + "　", True))
                runs.append((p.get("desc") or p.get("body") or "", False))
            paras.append(runs)
            heights.append(
                est_lines("".join(t for t, _ in runs), st.body_pt, inner,
                          st.line_safety) * st.lh)
        head_h = (TAG_H if col.get("tag") else 0.0) + LABEL_H + GAP_LABEL
        body_h = sum(heights) + item_gap / 72.0 * max(0, len(heights) - 1)
        blocks.append((col, paras, heights, head_h + body_h, body_h))

    block_h = max(b[3] for b in blocks)
    y0 = st.body_top + max(0.0, (avail - block_h) / 2)

    # 竖线只画到内容底部：整条贯穿到底会把「没写完」的空档标出来
    deck.rect(st.margin_x + cw + GAP / 2 - 0.008, y0 + 0.06,
              0.016, max(0.6, block_h - 0.12), fill=st.rule)

    for i, (col, paras, heights, h, body_h) in enumerate(blocks):
        x = st.margin_x + (0 if i == 0 else cw + GAP) + PAD
        y = y0
        if col.get("tag"):
            deck.text(x, y, inner, TAG_H,
                      [(col["tag"], st.body_pt, False, st.ink_2)])
            y += TAG_H
        deck.text(x, y, inner, LABEL_H, [(col["label"], st.compare_label_pt, True, st.ink)],
                  anchor=MSO_ANCHOR.TOP)
        y += LABEL_H + GAP_LABEL

        tag = f"对照栏「{col['label']}」"
        if not paras:
            continue
        if not deck.check_fit(y, st.body_bottom, heights, item_gap / 72.0, tag):
            continue
        deck.groups_text(x, y, inner, body_h, [paras],
                         gap_pt=0.0, item_gap_pt=item_gap,
                         anchor=MSO_ANCHOR.MIDDLE)


# ------------------------------------------------------------
# bullets —— 要点列表；item 可为字符串或 {name, desc}
# ------------------------------------------------------------
def bullets(deck: Deck, page: dict, base: str):
    """整页一个文本框：导语一段，要点各自一段，用段前距拉开层级。

    不用每行一个文本框——那样段落间距只能靠估算行数定位，
    行数一估偏就出现「有的行挤、有的行空」。

    竖直位置由本函数自己算，不用文本框的 MIDDLE 锚点：锚点是让
    PowerPoint 按它自己的行高模型居中，盒子远高于文字时会系统性偏下
    （实测偏 0.4in，上下留白 1.7 : 1.0）。按度量模块算出的总高反推
    起点，上下留白才是真的相等。
    """
    st = deck.style
    deck.new_slide()
    deck.header(page["title"])

    inner = st.inner_w
    lead = [t for t in _as_lines(page.get("lead")) if t]
    lead_pt = float(page.get("lead_pt", st.bullets_lead_pt))
    prefix = page.get("marker", "· ")
    gap = float(page.get("gap_pt", 16.0)) / 72.0
    item_gap = float(page.get("item_gap_pt", 0.0)) / 72.0

    runs_list, heights = [], []
    items = page.get("items") or []

    # 名称列：四个要点的名称字数不一，说明若紧贴名称起排，换行后次行退回
    # 左边距，四行读起来像一坨。取最宽的名称作为公共列宽，说明从同一列起
    # 排，续行也停在该列——和 fields_image 的标签列是同一个道理。
    names = [it.get("name") for it in items if isinstance(it, dict) and it.get("name")]
    hang = 0.0
    if names and page.get("align_desc", st.bullets_align):
        hang = max(text_width(prefix + n + "　", st.body_pt) for n in names) / 72.0
        hang = min(hang, inner * st.bullets_hang_max)
    text_w = inner - hang

    for it in items:
        if isinstance(it, str):
            runs = [(prefix, False), (it, False)]
        else:
            runs = [(prefix, False)]
            if it.get("name"):
                runs.append((it["name"] + "　", True))
            runs.append((it.get("desc") or it.get("body") or "", False))
        runs_list.append(runs)
        heights.append(
            est_lines("".join(t for t, _ in runs), st.body_pt, text_w,
                      st.line_safety) * st.lh)

    lead_h = len(lead) * lead_pt * 1.24 / 72.0
    items_h = sum(heights) + item_gap * max(0, len(heights) - 1)
    total = lead_h + (gap if lead and heights else 0.0) + items_h
    avail = st.body_bottom - st.body_top

    deck.check_fit(st.body_top, st.body_bottom, heights, item_gap,
                   page.get("tag", "要点"))
    y = st.body_top + max(0.0, (avail - total) / 2)

    # 文本框高度就取估算值，再用 MIDDLE 锚点：估算若偏小，文字会上下
    # 各溢出一点，视觉仍是居中；反之亦然。这样居中对估算误差不敏感。
    if lead:
        deck.text(st.margin_x + st.pad_x, y, inner, lead_h,
                  [(t, lead_pt, False, st.ink_2) for t in lead], spacing=1.0,
                  anchor=MSO_ANCHOR.MIDDLE)
        y += lead_h + (gap if heights else 0.0)

    if heights:
        tb = deck.groups_text(st.margin_x + st.pad_x, y, inner, items_h, [runs_list],
                              gap_pt=0.0, item_gap_pt=item_gap * 72.0,
                              anchor=MSO_ANCHOR.MIDDLE)
        if hang > 0:
            # 悬挂缩进：首行（名称行）回到左边距，续行停在名称列的右缘。
            px = int(Inches(hang))
            for para in tb.text_frame.paragraphs:
                pPr = para._p.get_or_add_pPr()
                pPr.set("marL", str(px))
                pPr.set("indent", str(-px))


def center_block(ys, heights, top, bottom):
    """把已排好的一组块整体下移，使剩余空间上下均分。

    distribute 只保证「从 top 起、间隙不超过 max_gap」，块少时余量全堆在
    页底（两条轨的时间轴能空掉三分之一页）。这里把余量对半补到上下。
    """
    used = ys[-1] + heights[-1] - top
    slack = (bottom - top) - used
    if slack <= 0:
        return ys
    return [y + slack / 2 for y in ys]


# ------------------------------------------------------------
# timeline —— 时间轴：若干条并行轨，每轨一串年代节点
# ------------------------------------------------------------
def timeline(deck: Deck, page: dict, base: str):
    """时间轴：几条线并排往下铺，每条线上一串年代节点。

    综述要回答的是「怎么走到今天」，这件事卡片排不出来——卡片是并列的，
    发展是先后。轴线给先后感，并行的轨给「几条线互相喂养」的关系：
    认识线提出模型，材料线被它指方向，新的材料又反过来逼出新模型。

    节点的横坐标按**序号均分**，不按年份比例。1853 与 2006 相隔一百五十年，
    按真实比例排，19 世纪那几个节点会挤成一条竖线、近二十年的又挤在一起；
    均分是「讲得清」优先于「坐标准」，年代仍写在节点上，间距不等于时长
    这件事读者自己看得出来。

    轨名放在轴线左端，像行标题；没有轨名的页面不留那一列，节点占满版心。
    """
    st = deck.style
    deck.new_slide()
    deck.header(page["title"])

    tracks = page["tracks"]
    n = max(len(t.get("nodes") or []) for t in tracks) or 1
    lab_w = st.tl_label_w if any(t.get("name") for t in tracks) else 0.0
    x0 = st.margin_x + lab_w
    cw = (st.slide_w - st.margin_x - x0) / n
    inner = cw - 0.20          # 节点文字两侧各留 0.10in，避免贴着相邻列

    # 每条轨的高度取决于该轨最长的那个节点：年代一行，事件名与说明各若干行
    heights = []
    for tr in tracks:
        h = 0.0
        for nd in tr.get("nodes") or []:
            hh = st.lh
            hh += est_lines(nd.get("head") or "", st.body_pt, inner,
                            st.line_safety) * st.lh
            hh += est_lines(nd.get("desc") or "", st.body_pt, inner,
                            st.line_safety) * st.lh
            h = max(h, hh)
        heights.append(h + st.tl_head_gap)

    deck.check_fit(st.body_top, st.body_bottom, heights, st.tl_track_gap,
                   page.get("tag", "时间轴"))
    ys = distribute(st.body_top, st.body_bottom, heights,
                    min_gap=st.tl_track_gap,
                    max_gap=max(st.tl_track_gap, st.tl_track_gap_max))
    # distribute 从上往下排，间隙封顶后余量全落在页底：两条轨的页面会变成
    # 上半页密、下半页空。剩余空间对半分给上下，整块落到视觉中线。
    ys = center_block(ys, heights, st.body_top, st.body_bottom)

    for tr, y, h in zip(tracks, ys, heights):
        if lab_w:
            deck.text(st.margin_x, y - 0.10, lab_w - 0.14, 0.42,
                      [(tr["name"], st.body_pt, True, st.ink_2)],
                      align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        # 轴线压在节点方块的中线上，方块才像是「钉」在轴上
        deck.rect(x0, y - st.tl_rail_w / 2, st.slide_w - st.margin_x - x0,
                  st.tl_rail_w, fill=st.rule)
        for i, nd in enumerate(tr.get("nodes") or []):
            cx = x0 + cw * (i + 0.5)
            deck.rect(cx - st.tl_tick / 2, y - st.tl_tick / 2,
                      st.tl_tick, st.tl_tick, fill=st.ink)
            lines = [(nd.get("year") or "", st.body_pt, True, st.ink)]
            if nd.get("head"):
                lines.append((nd["head"], st.body_pt, True, st.ink))
            if nd.get("desc"):
                lines.append((nd["desc"], st.body_pt, False, st.ink_2))
            deck.text(x0 + cw * i + 0.10, y + st.tl_head_gap, inner,
                      h - st.tl_head_gap, lines)


# ------------------------------------------------------------
# stages —— 阶段带：同一条线逐级演进
# ------------------------------------------------------------
def stages(deck: Deck, page: dict, base: str):
    """阶段带：左边年代、右边阶段名与说明，一条竖线把各级串起来。

    与 timeline 的分工：timeline 看全景（几条线各走了几步），stages 看纵深
    （一条线每一步解决了什么）。综述的常见结论是「每一步都在还上一步欠的
    账」，所以每一级都要能写清「上一级留下了什么问题」——那正是 `desc`
    该写的东西，不是再列一遍材料名。

    竖线只画第一条到最末一条的跨度：拉满整页会把「没讲完」的空档标出来。
    """
    st = deck.style
    deck.new_slide()
    deck.header(page["title"])

    rows = page["stages"]
    spine_x = st.margin_x + st.stage_year_w + st.stage_pad / 2
    cx = spine_x + st.stage_pad / 2
    inner = st.slide_w - st.margin_x - cx

    heights = []
    for r in rows:
        n = est_lines(r.get("name") or "", st.body_pt, inner, st.line_safety)
        for key in ("desc", "cost"):
            v = r.get(key)
            if v:
                n += est_lines(v, st.body_pt, inner, st.line_safety)
        heights.append(n * st.lh)

    deck.check_fit(st.body_top, st.body_bottom, heights, st.stage_gap,
                   page.get("tag", "阶段带"))
    ys = distribute(st.body_top, st.body_bottom, heights,
                    min_gap=st.stage_gap,
                    max_gap=max(st.stage_gap, st.stage_gap_max))
    ys = center_block(ys, heights, st.body_top, st.body_bottom)

    # 竖线从第一级的中心到最后一级的中心
    if len(rows) > 1:
        top = ys[0] + heights[0] / 2
        bottom = ys[-1] + heights[-1] / 2
        deck.rect(spine_x, top, st.stage_spine_w, bottom - top, fill=st.rule)

    for r, y, h in zip(rows, ys, heights):
        deck.rect(spine_x - st.tl_tick / 2 + st.stage_spine_w / 2,
                  y + h / 2 - st.tl_tick / 2, st.tl_tick, st.tl_tick,
                  fill=st.ink)
        deck.text(st.margin_x, y, st.stage_year_w, h,
                  [(r.get("year") or "", st.stage_year_pt, True, st.ink_2)],
                  align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        lines = []
        if r.get("name"):
            lines.append((r["name"], st.body_pt, True, st.ink))
        if r.get("desc"):
            lines.append((r["desc"], st.body_pt, False, st.ink_2))
        if r.get("cost"):
            lines.append((r["cost"], st.body_pt, False, st.ink_2))
        deck.text(cx, y, inner, h, lines)


PAGES = dict(cover=cover, section=section, compare=compare,
             fields_image=fields_image, image_cards=image_cards,
             image_grid_cards=image_grid_cards, cards_grid=cards_grid,
             bullets=bullets, timeline=timeline, stages=stages)


def page_min_pt(page, base):
    """该页的字号下限。

    页级 `min_pt` 显式优先；只写了 `body_pt` 时，这个字号就是该页的下限——
    页面声明「这页按 22pt 排」，不能因为全篇下限是 16pt 就悄悄压回 16pt。
    页级覆盖本来就允许低于全篇值（见 docs/使用说明.md「按页改字号」）。
    """
    if page.get("min_pt") is not None:
        return float(page["min_pt"])
    if page.get("body_pt") is not None:
        return float(page["body_pt"])
    return float(base)


def page_style(style, page):
    """把页级 `body_pt` / `min_pt` 叠加到全篇样式上；只影响这一页。"""
    body = float(page["body_pt"]) if page.get("body_pt") is not None else style.body_pt
    return _replace(style, body_pt=body, min_pt=page_min_pt(page, style.min_pt))


def _fits(page, base_dir, style, type_name):
    """在临时 deck 上试排一页，返回该字号是否放得下。

    版式函数的纵向位置全由内容高度反推（见模块开头的设计边界），所以
    「换个字号重排一遍，看有没有告警」就是可靠的可行性判断，不需要每个
    版式再写一份 measure 函数——那种做法一旦版式改了就会与真实排版脱节。

    试排画在一个丢弃的 Presentation 上：幻灯片对象建完即弃，不影响成品。
    图片会被读一次（add_picture 要读文件），相对整体生成可以忽略。

    这里**不吞异常**：缺图片、字段写错这类问题在试排时就要原样抛出来。
    以前一律 `except Exception: return False`，会把「配置错了」说成
    「这个字号放不下」，结果是在错误的字号上重试到下限，真正的原因
    （路径拼错、字段缺失）反而被藏起来。
    """
    probe = Deck(style=style)
    PAGES[type_name](probe, page, base_dir)
    return not probe.warnings


def choose_body_pt(page, base_dir, style, type_name):
    """在 [auto_pt_min, auto_pt_max] 里挑这一页的正文字号，返回 (字号, 下限, 上限)。

    两条判据先后有别：
    1. 放得下——在丢弃的 deck 上试排，看有没有版面告警；
    2. 不制造孤行——字号一涨，原本三行的说明可能变成「四行、末行两字」。
       放得下而难看的一档要退掉，继续往下试。

    若区间内所有字号都放不下，返回下限照排，让 check_fit 把超出量报出来，
    由人去减内容，而不是在这里悄悄压到看不清的字号。

    若逐档都排得下、却逐档都有孤行（窄栏版式的常见情形），退回作者写的
    `body_pt`：这套内容在这个版式里就没有更松的排法，往上调只会把末行
    挤成两个字。宁可保持原样，也不要「自动」出一个更差的版面。

    deck 与预检共用本函数，报出来的字号才和成品一致。
    """
    hi = float(page.get("auto_pt_max") or style.auto_pt_max or style.body_pt + 4)
    lo = float(page.get("auto_pt_min") or style.auto_pt_min or style.min_pt)
    # 区间下限不能低过本页的字号下限：自动排版挑出来的字号必须能过合规核对。
    lo = max(lo, float(style.min_pt))
    hi = max(hi, lo)
    step = max(0.5, style.auto_pt_step)
    fit_only = None      # 第一个「放得下」的字号，做兜底
    p = hi
    while p >= lo - 1e-6:
        cand = _replace(style, body_pt=p, min_pt=min(style.min_pt, p))
        if _fits(page, base_dir, cand, type_name):
            if fit_only is None:
                fit_only = p
            if not (cand.auto_pt_orphan
                    and has_orphan(base_dir, page, cand, p)):
                return p, lo, hi
        p -= step
    base = min(max(float(style.body_pt), lo), hi)
    if fit_only is not None:
        cand = _replace(style, body_pt=base, min_pt=min(style.min_pt, base))
        if _fits(page, base_dir, cand, type_name):
            return base, lo, hi
    return (fit_only if fit_only is not None else lo), lo, hi


def render(deck, page, base_dir, index):
    """按 type 分发到对应版式函数，并写讲稿。

    单页可覆盖正文字号：`page.body_pt` 只影响这一页，不牵动全篇。
    页与页定位不同（封面要小字排信息、内页要密排正文），字号本就该按页定。

    字号自适应（`auto_pt`，全局开、单页可关）：字少的页自动排大撑满版心，
    字多的页退回小一档不溢出，且不许把末行挤成一两个字。具体判据见
    `choose_body_pt`——门禁与生成共用它，避免「预检说通过、成品掉单字」。
    """
    t = page.get("type")
    fn = PAGES.get(t)
    if fn is None:
        raise ValueError(f"未知页面类型 {t!r}，可用：{'、'.join(PAGES)}")
    prev = deck.style
    if page.get("body_pt") is not None or page.get("min_pt") is not None:
        deck.style = page_style(prev, page)
    try:
        st = deck.style
        # 封面与章节页的字号是各自一套（cover_pt / section_title_pt），
        # 换 body_pt 不改观感，试排只会白花时间。
        auto = page.get("auto_pt", st.auto_pt)
        if auto and t not in ("cover", "section"):
            chosen, lo, hi = choose_body_pt(page, base_dir, st, t)
            deck.style = _replace(st, body_pt=chosen,
                                  min_pt=min(st.min_pt, chosen))
            deck.auto_pt_log.append((index, chosen, lo, hi))
        fn(deck, page, base_dir)
    finally:
        deck.style = prev
    if page.get("notes"):
        deck.notes(page["notes"])
    return deck
