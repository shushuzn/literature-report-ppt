"""Deck：把内容画到幻灯片上。

坐标一律用「英寸（float）」，在创建形状时才换算为 EMU，
这样版面计算可以直接做加减乘除。

所有尺寸与字体从 `self.style` 取（`pptkit.style.Style`），本文件不再
出现写死的字号、边距与字体名。换一套样式即换一个 Style 实例。
"""
import os
from PIL import Image
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

from .style import Style
from .metrics import est_lines


def script_chunks(text):
    """按「西文 / 中文」把一段文字切成若干块。

    为什么必须切：整段塞进一个 run 时，若这段以数字开头（如
    「2026 年 10 月」「300–500 F/g 多取自…」），PowerPoint 会用西文字体
    去排后面的中文，落到 Times 的中文回退字（宋体），一行里出现两种字形。
    切成 run 后，纯中文块只含中文，PowerPoint 就按 a:ea 用微软雅黑。

    全角标点、箭头、破折号、℃ 等（U+2010–U+2BFF）归入中文块：
    它们在微软雅黑里按全角渲染，与文本度量的口径一致。
    """
    def wide(ch):
        o = ord(ch)
        return o > 0x2E80 or 0x2010 <= o <= 0x2BFF

    out, buf, cur = [], [], None
    for ch in text:
        w = wide(ch)
        if cur is None or w == cur:
            buf.append(ch)
        else:
            out.append("".join(buf))
            buf, cur = [ch], w
        cur = w
    if buf:
        out.append("".join(buf))
    return [x for x in out if x]


class Deck:
    def __init__(self, style: Style = None):
        self.style = style or Style()
        st = self.style
        self.prs = Presentation()
        self.prs.slide_width = Inches(st.slide_w)
        self.prs.slide_height = Inches(st.slide_h)
        self.blank = self.prs.slide_layouts[6]
        self.slide = None
        self.page = 0
        self.audit = []      # (页码, 字号, 文本) —— 供合规自检
        self.warnings = []   # 版面放不下等告警
        self.auto_pt_log = []  # (页码, 选定字号, 试探区间) —— 自适应字号的决策记录

    def check_fit(self, y0, y1, heights, min_gap, tag):
        """检查一组内容能否放进 [y0, y1]。放不下则记录告警。

        没有这道检查时，内容超出可用高度会静默溢出到页脚之外。
        """
        need = sum(heights) + min_gap * (len(heights) - 1 if heights else 0)
        avail = y1 - y0
        if need > avail + 1e-6:
            self.warnings.append(
                f"P{self.page} {tag}：内容需 {need:.2f}in，可用 {avail:.2f}in，"
                f"超出 {need - avail:.2f}in（约 {(need - avail) / self.style.lh:.0f} 行）")
            return False
        return True

    # ---------- 页面 ----------

    def new_slide(self):
        self.slide = self.prs.slides.add_slide(self.blank)
        self.page += 1
        return self.slide

    def header(self, title: str):
        """页标题 + 一条细分隔线。不含副标题。"""
        st = self.style
        self.text(st.margin_x, st.header_y, st.slide_w - 2 * st.margin_x, 0.52,
                  [(title, st.title_pt, True, st.ink)], anchor=MSO_ANCHOR.MIDDLE)
        self.rect(st.margin_x, st.rule_y, st.slide_w - 2 * st.margin_x, 0.017,
                  fill=st.rule)

    def notes(self, text):
        """写入演讲者备注（讲稿）。

        备注只给讲者看，不投影、不参与版面计算，故按 `style.notes_pt` 排
        （默认 14pt）；正文区仍受 `style.min_pt` 下限约束，两者互不影响。
        text 可为字符串（内部换行按段落拆开）或字符串列表（每项一段）。
        """
        st = self.style
        paras = text if isinstance(text, list) else [text]
        tf = self.slide.notes_slide.notes_text_frame
        tf.text = ""
        first = True
        for raw in paras:
            for line in str(raw).split("\n"):
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                r = p.add_run()
                r.text = line
                r.font.size = Pt(st.notes_pt)
                r.font.name = st.cn_font
                rPr = r._r.get_or_add_rPr()
                ea = rPr.find(qn('a:ea'))
                if ea is None:
                    ea = rPr.makeelement(qn('a:ea'), {})
                    rPr.append(ea)
                ea.set('typeface', st.cn_font)

    # ---------- 基本元素 ----------

    def _add_runs(self, p, text, size, bold, color):
        """把一个逻辑文本切成若干 run 写进段落，逐块设字体。

        latin 一律用 `style.en_font`，东亚字体一律 `style.cn_font`；块内只有
        中文时 PowerPoint 才真的会用 a:ea（见 script_chunks 的说明）。
        """
        st = self.style
        for chunk in script_chunks(text):
            r = p.add_run()
            r.text = chunk
            f = r.font
            f.size, f.bold, f.name = Pt(size), bold, st.en_font
            f.color.rgb = color
            rPr = r._r.get_or_add_rPr()
            ea = rPr.find(qn('a:ea'))
            if ea is None:
                ea = rPr.makeelement(qn('a:ea'), {})
                rPr.append(ea)
            ea.set('typeface', st.cn_font)
            # 语言标记决定用哪套断行规则。不写 lang 时 PowerPoint 按 en-US 断行，
            # 中文标点的禁则（不让「，）」跑到行首）不生效，渲染出来就是标点顶行首。
            rPr.set('lang', 'zh-CN')
            rPr.set('altLang', 'en-US')
        if size < st.min_pt:
            self.audit.append((self.page, size, text[:30]))

    def text(self, x, y, w, h, lines, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, spacing=1.0):
        """lines: [(文本, 字号, 是否加粗, 颜色), ...]，每项独立成段。

        两个要点：
        1. 内部边距全部归零。PowerPoint 文本框默认上下留 0.05 英寸、
           左右留 0.1 英寸，若不清零，版面计算出的宽度与实际可用宽度
           不符，文字会被挤到下一行、并整体下移。
        2. 西文与数字用 en_font，中文用 cn_font。python-pptx 的
           font.name 只作用于西文，中文必须在 a:rPr 下写 a:ea 节点，
           否则回退默认字体。
        """
        tb = self.slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        for i, (t, sz, bold, color) in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = spacing
            self._add_runs(p, t, sz, bold, color)
        return tb

    def text_runs(self, x, y, w, h, paras, align=PP_ALIGN.LEFT,
                  anchor=MSO_ANCHOR.TOP, spacing=1.0):
        """一段多 run。paras = [[(文本, 是否加粗), ...], ...]，每个内层列表是一段。

        用于「标签 + 值」同行但字重不同的排版。颜色统一按加粗与否取色。
        """
        st = self.style
        tb = self.slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        for i, runs in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = spacing
            for t, bold in runs:
                self._add_runs(p, t, st.body_pt, bold,
                               st.ink if bold else st.ink_2)
        return tb

    def groups_text(self, x, y, w, h, groups, gap_pt=14.0,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.0,
                    item_gap_pt=0.0, size=None):
        """整栏一个文本框，组间用段前距隔开。

        为什么这样做：若给每组单独定位（估算行数 → 分配高度 → 均分间隔），
        估算偏大的组会留下内部空隙，使视觉间隔忽大忽小。放进同一个文本框、
        用 paragraph.space_before 控制间隔，则间隔由 PowerPoint 保证恒定，
        与行数估算是否准确无关。

        groups: [[ [(文本, 加粗), ...], [(文本,加粗),...] ], ...]
                外层每组是一个字段，内层每项是一段，段内是若干 run。

        item_gap_pt 是组内相邻段的段前距，默认 0（保持原行为）。要点页
        需要把同一组里的条目拉开时用它——此前只有组间能设，条目之间
        只能靠行距，四条要点会挤成一块。

        size 为 None 时取 `style.body_pt`。
        """
        st = self.style
        size = st.body_pt if size is None else size
        tb = self.slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        first = True
        for gi, group in enumerate(groups):
            for pi, runs in enumerate(group):
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                p.alignment = align
                p.line_spacing = spacing
                if gi > 0 and pi == 0:
                    p.space_before = Pt(gap_pt)
                elif pi > 0 and item_gap_pt:
                    p.space_before = Pt(item_gap_pt)
                for t, bold in runs:
                    self._add_runs(p, t, size, bold,
                                   st.ink if bold else st.ink_2)
        return tb

    def card(self, x, y, w, height, head, body, pad_x=None, pad_v=None,
             fill=None, line=None, spacing=1.0, gap_pt=None):
        """带边框的卡片：标题 + 正文，整体垂直居中。

        三点讲究，缺一个就会显得「文字挤在框里」：

        1. 高度由内容反推（调用方用 card_height 计算），不写死。
        2. 标题与正文放进**同一个**文本框并设 MIDDLE 锚点，由 PowerPoint
           负责居中，避免「多个文本框 + 固定偏移」造成的上下边距不等。
        3. 标题与正文之间留 gap_pt 的段前距。此前两级只隔 3–5pt，比框内
           留白（10–12pt）还小，两行糊成一块；层级间隔必须大于段内行距，
           层级才立得住。
        """
        st = self.style
        pad_x = st.pad_x if pad_x is None else pad_x
        pad_v = st.pad_v if pad_v is None else pad_v
        gap_pt = st.card_gap_pt if gap_pt is None else gap_pt
        fill = st.paper if fill is None else fill
        line = st.edge if line is None else line
        self.rect(x, y, w, height, fill=fill, line=line, lw=st.card_line_w)
        groups = []
        if head:
            groups.append([[(head, True)]])
        groups.append([[(b, False)] for b in
                       (body if isinstance(body, list) else [body])])
        self.groups_text(x + pad_x, y + pad_v, w - 2 * pad_x, height - 2 * pad_v,
                         groups, gap_pt=gap_pt,
                         anchor=MSO_ANCHOR.MIDDLE, spacing=spacing)

    def card_height(self, head, body, width, pad_v=None, gap_pt=None):
        """卡片所需高度（英寸）：标题行 + 正文行 + 层级间隔 + 上下留白"""
        st = self.style
        pad_v = st.pad_v if pad_v is None else pad_v
        gap_pt = st.card_gap_pt if gap_pt is None else gap_pt
        inner = width - 2 * st.pad_x
        n = (est_lines(head, st.body_pt, inner, st.line_safety) if head else 0)
        for b in (body if isinstance(body, list) else [body]):
            n += est_lines(b, st.body_pt, inner, st.line_safety)
        gap = gap_pt / 72.0 if head else 0.0
        return n * st.lh + gap + 2 * pad_v

    def rect(self, x, y, w, h, fill=None, line=None, lw=1.0):
        sp = self.slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        if fill is None:
            sp.fill.background()
        else:
            sp.fill.solid()
            sp.fill.fore_color.rgb = fill
        if line is None:
            sp.line.fill.background()
        else:
            sp.line.color.rgb, sp.line.width = line, Pt(lw)
        sp.shadow.inherit = False
        return sp

    def panel(self, x, y, w, h, fill=None, line=None, lw=1.0):
        st = self.style
        return self.rect(x, y, w, h,
                         fill=st.paper if fill is None else fill,
                         line=st.rule if line is None else line, lw=lw)

    # ---------- 图片 ----------

    @staticmethod
    def image_size(path, max_w, max_h):
        """按真实宽高比缩放到不超过 max_w x max_h，返回 (w, h)。"""
        iw, ih = Image.open(path).size
        r = iw / ih
        w, h = max_w, max_w / r
        if h > max_h:
            h, w = max_h, max_h * r
        return w, h

    def image(self, path, x, y, w, h):
        """在 (x, y) 处放置尺寸恰为 w x h 的图片（调用方保证比例正确）"""
        return self.slide.shapes.add_picture(
            path, Inches(x), Inches(y), Inches(w), Inches(h))

    def image_framed(self, path, x, y, max_w, max_h, caption=None,
                     caption_h=None, pad=None, line=None):
        """画一个恰好贴合图片宽高比的外框，图片充满外框。

        两个要点：

        1. 外框尺寸由图片比例反推，不用固定尺寸——固定框会把图片缩在中间
           留下大片死区（曾出现 65%）。
        2. 若因宽度受限而填不满 max_h，图连图注在栏内**垂直居中**，上下均分
           余量。否则余量会全部堆在图下方，看起来就是「图太靠上、下面空一块」。

        返回 (框的 x, y, w, h, 本块总高)
        """
        st = self.style
        caption_h = st.cap_h if caption_h is None else caption_h
        pad = st.img_pad if pad is None else pad
        line = st.rule if line is None else line
        iw, ih = self.image_size(path, max_w - 2 * pad, max_h - 2 * pad)
        bw, bh = iw + 2 * pad, ih + 2 * pad
        bx = x + (max_w - bw) / 2          # 在可用宽度内居中
        total = bh + (caption_h + st.cap_top_gap if caption else 0.0)
        if total < max_h:                  # 填不满就把图块垂直居中
            y += (max_h - total) / 2
        self.rect(bx, y, bw, bh, fill=st.paper, line=line, lw=1.0)
        self.image(path, bx + pad, y + pad, iw, ih)
        if caption:
            if est_lines(caption, st.body_pt, bw, st.line_safety) > 1:
                # 图注只给了一行的高度，折行会压到正文区。宁可在生成时告警，
                # 也不要在成品里出现半截露在框外的字。
                self.warnings.append(
                    f"P{self.page} 图注超过一行（图框宽 {bw:.2f}in）："
                    f"「{caption[:18]}…」请缩短")
            self.text(bx, y + bh + st.cap_top_gap, bw, caption_h,
                      [(caption, st.body_pt, False, st.ink_2)],
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
        return bx, y, bw, bh, total

    def image_row(self, items, y, height, print_center=None):
        """一行图片：统一显示高度，顶对齐，说明文字对齐同一基线。

        items: [(path, caption, x_center, max_w), ...]
        height: 统一显示高（不含外框留白与说明）
        返回本行总高
        """
        st = self.style
        pad = st.img_pad
        row_bottom = y
        cap_y = None
        boxes = []
        for path, caption, cx, max_w in items:
            iw, ih = self.image_size(path, max_w - 2 * pad, height - 2 * pad)
            bw, bh = iw + 2 * pad, ih + 2 * pad
            bx = cx - bw / 2
            self.rect(bx, y, bw, bh, fill=st.paper, line=st.rule, lw=1.0)
            self.image(path, bx + pad, y + pad, iw, ih)
            boxes.append((bx, caption))
            row_bottom = max(row_bottom, y + bh)
            cap_y = max(cap_y or 0, y + bh + st.cap_top_gap)
        for bx, caption in boxes:
            if caption:
                # 图注宽度按剩余页宽收窄。写死宽度时，若图框靠右，文本框会
                # 越出页面右边界（实测 x+w = 13.65 > 13.33）。
                cap_w = min(5.6, st.slide_w - st.margin_x - bx)
                self.text(bx, cap_y, cap_w, st.cap_h,
                          [(caption, st.body_pt, False, st.ink_2)],
                          align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
        return (cap_y + st.cap_h if cap_y else row_bottom) - y

    # ---------- 输出 ----------

    def save(self, path):
        # 允许 `Deck().save("out.pptx")`：dirname 为空时不能拿 "" 去 makedirs。
        parent = os.path.dirname(os.path.abspath(path))
        os.makedirs(parent, exist_ok=True)
        self.prs.save(path)
        return path

    @property
    def page_count(self):
        return len(self.prs.slides._sldIdLst)
