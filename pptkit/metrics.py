"""文本度量：估算换行数与块高。

为什么需要它：字号提到 20pt 后，同一段文字在固定宽度下的实际换行数
远超凭经验估计。若盒子高度靠手填，就会出现文字溢出边框、相邻元素重叠。
这里按字符宽度累加估算，据此反推盒子高度。
"""
import math

from .style import Style

# 单行高度（英寸），默认样式下 20pt × 1.24 / 72。
# 这是**默认值快照**：实际版面请用 `deck.style.lh`（随 body_pt 变化）。
LH_IN = Style().lh

# 可用宽度折减系数。字体度量与实际渲染存在误差，宁可略高估行数
# （高估只会让盒子略松，低估会导致文字溢出边框）。
# 默认 0.97；`style.line_safety` 可按项目收紧。0.94 时实测能兜住
# 「估一行、渲染两行」的情况（465pt 文字排在 475pt 可用宽里仍折行）。
SAFETY = 0.97

# 全角字符宽 / 字号。中文字体在设计上就是一字一 em，与字号成正比，
# 与选哪个字体无关（等宽汉字是硬约束），故作为常量而非配置项。
CJK_W = 1.00

# 西文按字符类别取值（em 相对字号）。此前对西文统一取 0.52，
# 对以大小写字母为主的文本会高估约三成，导致行数多算、版面留白不均。
# 下面这组系数由渲染图实测标定（Times New Roman 20pt）。
_EM_DIGIT = 0.50    # 数字
_EM_UPPER = 0.60    # 大写字母
_EM_LOWER = 0.42    # 小写字母
_EM_SPACE = 0.25    # 空格
_EM_PUNCT = 0.32    # 其余西文标点


def _em(ch: str) -> float:
    o = ord(ch)
    if o > 0x2E80:          # 中日韩字符
        return CJK_W
    if 0x2010 <= o <= 0x2BFF:
        # 箭头、破折号、℃ 等符号。在微软雅黑里按全角渲染（宽 1 em），
        # 不能按西文标点算——否则会低估宽度，导致多出一个字的换行。
        return CJK_W
    if ch.isdigit():
        return _EM_DIGIT
    if ch.isupper():
        return _EM_UPPER
    if ch.islower():
        return _EM_LOWER
    if ch == " ":
        return _EM_SPACE
    return _EM_PUNCT


def text_width(text: str, size_pt: float) -> float:
    """估算一行文本的宽度（磅）"""
    return sum(size_pt * _em(ch) for ch in text)


def est_lines(text: str, size_pt: float, width_in: float, safety=None) -> int:
    """估算在给定宽度内占几行。safety 缺省用模块默认值。"""
    if not text:
        return 0
    avail = width_in * 72.0 * (SAFETY if safety is None else safety)
    return max(1, math.ceil(text_width(text, size_pt) / avail))


def block_h(text: str, size_pt: float, width_in: float, pad_in: float = 0.0,
            safety=None, line_h=None) -> float:
    """文本块所需高度（英寸）。

    行高倍数默认取 `Style.line_h`（1.24）。这里以前写了一个不存在的模块
    常量 LINE_H，函数一调用就 NameError——它被 `pptkit.__init__` 导出，
    属于公开 API，必须能真跑。
    """
    n = est_lines(text, size_pt, width_in, safety)
    lh = Style().line_h if line_h is None else line_h
    return n * size_pt * lh / 72.0 + pad_in


def last_line_ratio(text: str, size_pt: float, width_in: float, safety=None) -> float:
    """末行占整行宽度的比例。

    中文长句换行后，末行若只剩一两个字或一个标点，视觉上就是「孤行」，
    非常难看。生成前用本函数扫一遍，占比低于约 0.18 的文案应当改写或断行。
    """
    avail = width_in * 72.0 * (SAFETY if safety is None else safety)
    total = text_width(text, size_pt)
    if total <= avail:
        return 1.0
    lines = math.ceil(total / avail)
    return (total - (lines - 1) * avail) / avail


def distribute(y0: float, y1: float, heights, min_gap=0.10, max_gap=0.40):
    """把一组高度在 [y0, y1] 区间内均匀排布，返回各块顶端的 y 坐标。

    块高之和不足时，剩余空间平均分配到间隙（受 max_gap 限制）；
    超出时按 min_gap 紧排（由调用方负责缩减内容）。
    """
    hs = list(heights)
    n = len(hs)
    if n == 0:
        return []
    if n == 1:
        return [y0]
    total = sum(hs)
    gap = (y1 - y0 - total) / (n - 1)
    gap = max(min_gap, min(max_gap, gap))
    ys, y = [], y0
    for h in hs:
        ys.append(y)
        y += h + gap
    return ys
