"""合规核对：对生成的 pptx 做机器可验证的检查。

检查项：
    页数 / 字号下限 / 中文字体 / 英文字体 / 配图 / 必填文字项 / 每页字数

所有阈值都从参数来，本文件不含任何一次汇报的具体数字：
字号、字体取自 `Style`（`content.yaml` 的 `style:` 段），
页数、配图页、必填项、字数区间取自 `compliance:` 段。
"""
import re
from pptx import Presentation
from pptx.oxml.ns import qn

from .style import Style

LATIN = re.compile(r"[A-Za-z]")
# 东亚字体只对含中日韩文字/全角标点的 run 有意义。纯西文的 run 不带 a:ea
# 是正常的：PowerPoint 手动编辑后这些 run 往往没有 ea 节点，旧逻辑会把
# 它们全判成「中文字体错」。
CJK = re.compile(r"[\u3000-\u303f\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uff00-\uffef]")
_SPACE = re.compile(r"\s+")

# 封面不该出现的东西。期刊卷期页与 DOI 在第 2 页的「发表期刊」字段里已经
# 完整给出，封面再写一遍只是把注意力从题目上分走；「文献汇报」这类眉标是
# 会议登记表上的字段，投影时人人都在做汇报，写了不增加信息。
# 这几条是版式规则，与具体文献无关，所以写在代码里；确需例外时在
# content.yaml 的 compliance.cover_allow 里列出放行的片段。
COVER_FORBID = [
    (re.compile(r"doi", re.I), "DOI"),
    (re.compile(r"\d+\s*\(\s*\d+\s*\)\s*[,，]\s*\d+"), "卷(期), 页码"),
    (re.compile(r"\d+\s*:\s*\d+\s*[,，]"), "卷: 页码"),
    (re.compile(r"文\s*献\s*汇\s*报"), "「文献汇报」眉标"),
]


def _runs(prs):
    """遍历所有文本 run，产出 (页码, 文本, 字号, latin字体, ea字体)"""
    for si, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    t = r.text.strip()
                    if not t:
                        continue
                    rPr = r._r.find(qn('a:rPr'))
                    ea = rPr.find(qn('a:ea')) if rPr is not None else None
                    yield (si, t,
                           r.font.size.pt if r.font.size else None,
                           r.font.name,
                           ea.get('typeface') if ea is not None else None)


def page_chars(pptx_path):
    """每页可见字数（正文 + 标题 + 图注，不含演讲者备注）。

    按去空白后的字符数计：中文一字算一，西文一个字母算一。
    要按「汉字计 1、西文计 0.5」的口径量讲稿，用 tools/note_budget.py。
    """
    prs = Presentation(pptx_path)
    out = {}
    for si, s in enumerate(prs.slides, 1):
        n = 0
        for sh in s.shapes:
            if sh.has_text_frame:
                n += len(_SPACE.sub("", sh.text_frame.text))
        out[si] = n
    return out


def _range_for(spec, page_no):
    """从 chars_per_page 配置里取第 page_no 页的区间。

    支持两种写法：

        chars_per_page: [120, 260]              # 所有页同一区间
        chars_per_page:                         # 按页覆盖
          default: [120, 260]
          2: [200, 400]
    """
    if spec is None:
        return None
    if isinstance(spec, (list, tuple)):
        return float(spec[0]), float(spec[1])
    if isinstance(spec, dict):
        v = spec.get(page_no, spec.get("default"))
        return (float(v[0]), float(v[1])) if v else None
    return None


def check(pptx_path, max_pages=None, require_image_pages=None,
          required_text=None, style=None, chars_per_page=None,
          min_pt_pages=None, cover_allow=None):
    """返回 (findings, stats)。

    max_pages 为 None 时不校验页数上限——页数上限只从 content.yaml 的
    compliance 段来，不内置「五页」这类假设。

    chars_per_page 为每页可见字数的允许区间：低于下限说明这一页内容单薄，
    高于上限说明塞不下（会挤压或溢出）。不写则不查。

    min_pt_pages 为按页覆盖的字号下限 {页号: pt}：一页的正文字号可以与该页
    内容密度相配（封面小字、内页密排），判定时按页取，不用全篇最小值放过所有页。

    cover_allow 为封面放行片段列表：命中 COVER_FORBID 但确实要写的内容
    （比如综述本身就题为「文献汇报方法」），逐个列出来才放行。

    findings: [{'level': 'ERROR'|'WARN', 'item': str, 'detail': str}]
    stats:    {'pages': int, 'runs': int, 'sizes': {字号: 计数},
               'images': {页: 数}, 'chars': {页: 字数}}
    """
    st = style or Style()
    prs = Presentation(pptx_path)

    images = {}
    texts = {}
    for si, s in enumerate(prs.slides, 1):
        images[si] = sum(1 for sh in s.shapes if sh.__class__.__name__ == "Picture")
        texts[si] = " ".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame)

    sizes = {}
    bad_size, bad_cn, bad_en = [], [], []
    n_runs = 0
    for si, t, sz, latin, ea in _runs(prs):
        n_runs += 1
        sizes[sz] = sizes.get(sz, 0) + 1
        lo = (min_pt_pages or {}).get(si, st.min_pt)
        if sz is None or sz < lo:
            bad_size.append((si, sz, t[:34]))
        if CJK.search(t) and ea != st.cn_font:
            bad_cn.append((si, ea, t[:34]))
        if LATIN.search(t) and latin != st.en_font:
            bad_en.append((si, latin, t[:34]))

    f = []
    pages = len(prs.slides)

    if max_pages is not None and pages > max_pages:
        f.append(dict(level="ERROR", item=f"页数 <= {max_pages}",
                      detail=f"实际 {pages} 页"))
    if bad_size:
        limits = sorted(set((min_pt_pages or {}).values()) | {st.min_pt})
        label = "/".join(f"{v:g}" for v in limits)
        f.append(dict(level="ERROR", item=f"字号 >= {label}pt",
                      detail=f"违规 {len(bad_size)} 处，如 " +
                             "; ".join(f"P{a} {b}pt {c!r}" for a, b, c in bad_size[:4])))
    if bad_cn:
        f.append(dict(level="ERROR", item=f"中文字体 {st.cn_font}",
                      detail=f"违规 {len(bad_cn)} 处，如 " +
                             "; ".join(f"P{a} {b!r}" for a, b, _ in bad_cn[:4])))
    if bad_en:
        f.append(dict(level="ERROR", item=f"英文字体 {st.en_font}",
                      detail=f"违规 {len(bad_en)} 处，如 " +
                             "; ".join(f"P{a} {b!r}" for a, b, _ in bad_en[:4])))

    for pg in (require_image_pages or []):
        if images.get(pg, 0) < 1:
            f.append(dict(level="ERROR", item=f"P{pg} 需配图", detail="未检出图片"))

    for pg, keywords in (required_text or {}).items():
        body = texts.get(pg, "")
        missing = [k for k in keywords if k not in body]
        if missing:
            f.append(dict(level="ERROR", item=f"P{pg} 必填项",
                          detail="缺：" + "、".join(missing)))

    # 封面重复文献出处。只查第 1 页：内页写卷期页是对的，「发表期刊」字段
    # 本来就该完整；错的是封面又抄一遍。
    cover = texts.get(1, "")
    allow = [str(a) for a in (cover_allow or [])]
    hits = []
    for pat, name in COVER_FORBID:
        m = pat.search(cover)
        if not m:
            continue
        if any(a and a in cover for a in allow):
            continue
        hits.append(f"{name}（{m.group(0).strip()!r}）")
    if hits:
        f.append(dict(level="ERROR", item="P1 封面不重复文献出处",
                      detail="检出 " + "、".join(hits) +
                             "；这类信息放第 2 页「发表期刊」字段，封面删掉"))

    chars = page_chars(pptx_path)
    if chars_per_page is not None:
        for pg in range(1, pages + 1):
            rng = _range_for(chars_per_page, pg)
            if not rng:
                continue
            lo, hi = rng
            n = chars[pg]
            if n < lo:
                f.append(dict(level="ERROR", item=f"P{pg} 内容单薄",
                              detail=f"{n} 字，下限 {lo:.0f} 字，差 {lo - n:.0f} 字"))
            elif n > hi:
                f.append(dict(level="ERROR", item=f"P{pg} 内容过密",
                              detail=f"{n} 字，上限 {hi:.0f} 字，超 {n - hi:.0f} 字"))

    stats = dict(pages=pages, runs=n_runs, sizes=sizes,
                 images=images, chars=chars)
    return f, stats
