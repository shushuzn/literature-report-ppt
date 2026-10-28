#!/usr/bin/env python
"""文风体检：找出稿子里的 AI 腔和空标题。

    python tools/style_check.py projects/energy-materials/content.yaml
    python tools/style_check.py --all
    python tools/style_check.py projects/energy-materials/content.yaml --strict

为什么要有这一步
----------------
版面检查管得住「装不下」，管不住「读着假」。AI 写中文有一批固定的句式癖好
（对偶转折、三段排比、抒情式破折号、空泛过渡词），它们单看每句都通顺，
叠在一份稿子里就成了塑料味。这些毛病不靠肉眼一遍遍扫，靠规则表。

分两级
------
严重（exit 1）：模型指纹明显，必须改。
提示（exit 0）：可能成立，交作者判断，不阻断。

标题另有两层（见 check_title / check_subheads）
---------------------------------------
1. 标点：`，：；` 一律判严重。逗号两侧各是一个判断，等于把两页并成一句，
   这是「不是…而是」之外最像 AI 的毛病，且肉眼极难扫出来。
2. 判断：标题必须是能读通的主谓宾，不能是名词标签。

规则来源
--------
句式清单参考公开的去 AI 味实践（「不是……而是」家族、破折号、空泛过渡词、
三段排比、虚假范围），并按中文学术汇报的语体做了增删：学术文本里的
「显著」「相比」有实义，不列入黑名单；「而非」在对比实验里偶有需要，
仍按严重处理，因为它在汇报稿里几乎总能改成直陈。

标题规则来自 Minto《The Pyramid Principle》「Use statements not captions」，
以及 High Bridge Academy 五条 action title 规则的第 1 条「no fragments,
no colons」。
"""
import argparse
import glob
import os
import re
import sys

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ------------------------------------------------------------
# 规则表
# ------------------------------------------------------------
# (级别, 规则名, 正则, 说明)
# 级别 S = 严重，改不掉就别交；W = 提示，看语境。
RULES = [
    # ---- 对偶转折：一句里先否定再肯定，把两个可以并存的事写成二选一 ----
    ("S", "不是…而是",
     r"不(?:是|在|仅|只)[^。；！？\n]{0,40}?(?:而是|而在|，是)",
     "直接说后半句，前半句删掉"),
    ("S", "并非…而是", r"并非[^。；！？\n]{0,30}?而是", "同上"),
    ("S", "而非", r"而非", "改成「不是」会被规则一并针对，这里建议直陈："
                            "把「A 而非 B」写成「是 A」，需要对比时另起一句说 B"),
    ("S", "与其…不如", r"与其[^。；！？\n]{0,30}?不如", "说清取舍即可，不必写成格言"),
    ("S", "不仅…而且",
     r"不仅[^。；！？\n]{0,30}?(?:而且|更是|还)", "删掉递进壳子，两件事分开说"),

    # ---- 空转过渡词：删掉不影响语义 ----
    ("S", "空泛过渡",
     r"(?:值得注意的是|值得一提的是|需要指出的是|不难看出|不难发现|"
     r"众所周知|毋庸置疑|毫无疑问|总的来说|综上所述|总而言之|"
     r"归根结底|一言以蔽之|可以说|事实上|实际上)",
     "整词删除，直接从事实起句"),

    # ---- 收尾套话：把结论压成一句口号 ----
    ("S", "口号式收尾",
     r"(?:一句话|一句话总结|归结起来|概括起来|用一句话)",
     "结论要能落地，写成具体动作，不要压成标语"),

    # ---- 排版陋习：破折号当万能连接符 ----
    ("S", "破折号", r"——", "改冒号或句号；一页最多留一处，且不能用来插入解释"),

    # ---- 正反问句：口语设问，书面汇报里显得随意 ----
    ("W", "正反问",
     r"([\u4e00-\u9fa5])不\1",
     "「是不是 / 能不能 / 在不在」改成陈述句；真需要设问时用「是否」"),

    # ---- 虚假范围：两端不在同一维度 ----
    ("W", "从…到…",
     r"从[^，。；！？\n]{1,14}?到[^，。；！？\n]{1,14}",
     "两端须是同一维度上的取值；若只是「先讲 A 再讲 B」，直接分句写"),

    # ---- 程度副词堆砌 ----
    ("W", "副词堆砌",
     r"(?:极大地|显著地|大幅度地|充分地|深入地|有效地)",
     "换成具体幅度：涨了多少、差几倍"),

    # ---- 名词化空壳 ----
    ("W", "空壳动宾",
     r"(?:进行了[^，。]{0,6}的?(?:研究|分析|探讨|优化)|"
     r"开展了[^，。]{0,6}的?(?:工作|研究)|起到了[^，。]{0,8}作用)",
     "换成实义动词：测了、算了、比了"),
]

# 标题判定：只由名词堆成的是标签，含谓词或数值的是结论。
# 不做语法分析，用一个谓词字表——宁可漏报，不要误报到作者不信这个工具。
VERDICT_WORDS = re.compile(
    r"(?:写在|来自|取决|决定|限制|归到|落在|变成|转向|拉开|卡住|出现|"
    r"修正|改写|推翻|违背|偏离|"
    r"差|降|升|涨|掉|超过|低于|高于|达到|只有|仅|能|会|要|可|应|需|"
    r"让|把|使|被|给|堵|抬|压|多|少|倍|%|\d|有|打开|换成|落到|"
    r"是|为|不|已|未|衰减|保持|循环|翻|剩|变成|相比)")
# 「XX 信息」「XX 概况」是纯标签，这一条判严重
LABEL_TITLE = re.compile(r"^(?:.{0,14}(?:信息|概述|简介|介绍|概况|总结|小结))$")

# 口语与比喻：标题是投影出去的门面，不能带说书腔。
# 这套词表只用于标题（页标题、卡片名、阶段名），正文与讲稿允许口语。
TITLE_COLLOQUIAL = re.compile(
    r"(?:的账|卡在|管不住|装不下|挤不进|挤不进去|好看|做模具|长出|冲到|"
    r"翘起来|拎|搞定|一堆|说白了|天壤之别|翻番|猛增|自己|咱们|"
    r"走通|走完|跑了|干了|搞|挺|蛮|老是|其实)")

# 谓词标记：出现在标题里，说明这条标题在替作者宣布判断，而不是交代栏目。
#
# 为什么要在 label 模式下查这个：label 的意思就是「这一页讲哪一块」，
# 全篇都得是名词短语。混进一条「现有瓶颈集中在概念层面」，翻到那页语体
# 就断了——观众读到的是一句结论，而相邻页只给了栏目名，落差全落在这一页。
# 这不是语法错，是**整篇语体不一致**，只有横向比对才看得见，所以做成规则。
#
# 只收多字词，避免「应用」「相应」「反应」这类名词里的单字误伤。
LABEL_VERDICT = re.compile(
    r"(?:集中在|转向|下一步|反而|难以|尚未|已经|正在|导致|表明|说明|"
    r"证明|显示|取决于|决定|必须|应当|开始|使得|提高|实现|降低|提升|"
    r"引入|进入|回到|拉开|归到|落在)")


# 章节页与封面本来就用短标题：它们是呼吸，不承担结论
SKIP_TYPES = {"cover", "section"}

# ------------------------------------------------------------
# 数据密度：一张卡里塞几个数
# ------------------------------------------------------------
# 综述汇报与研究论文的要求正好相反。研究论文讲「我做了什么」，数据就是
# 内容，一页六张卡六个数是正常的。综述讲「别人得出了什么」，读者要的是
# 关系与走向——哪个更好、哪个换来了什么。一张卡里同时摆 300–500 F/g、
# 10 mg/cm²、100 μm、50–80% 四个数，听众既记不住也来不及比，反而不如
# 一句「保持率往往不足一半」。
#
# 所以这条规则只管综述项目（配置里有 review 段），且只统计**同一张卡**
# 里的数字个数：跨卡比较另算，见 check_page_density。
#
# 阈值取 3：一张卡三个数以内还能当锚点用（一个数量级、一个单位、一条
# 参照），超过就说明在罗列而不是在论证。年份与型号不算。
DATA_NUM = re.compile(r"\d+(?:\.\d+)?")
# 这些数是「这篇文献/这个阶段叫什么」的一部分，不参与密度统计
DATA_ALLOW = re.compile(r"(?:Fig\.|图|表\s*\d|ref\.|文献\s*\d)")
# 卡名与阶段名里带数是允许的（「0.7 nm 碳化物衍生碳」），只统计正文
DENSITY_MAX = 3


def _card_text(item):
    """取一张卡的正文：desc + cost（stages 专有）。"""
    out = []
    for k in ("desc", "cost"):
        v = item.get(k)
        if isinstance(v, str):
            out.append(v)
        elif isinstance(v, list):
            out.extend(x for x in v if isinstance(x, str))
    return "".join(out)


def check_data_density(cfg):
    """综述页里单张卡片的数值个数超限，报提示。

    研究论文项目不适用：那里的数字是内容本身，见函数头说明。
    这条规则默认关闭，在 content.yaml 里写 `data_density: true` 才生效。
    """
    if not cfg.get("review") or not cfg.get("data_density"):
        return []
    hits = []
    for i, page in enumerate(cfg.get("deck") or [], 1):
        for key in ("cards", "items", "stages"):
            for j, it in enumerate(page.get(key) or []):
                if not isinstance(it, dict):
                    continue
                txt = DATA_ALLOW.sub("", _card_text(it))
                nums = DATA_NUM.findall(txt)
                if len(nums) > DENSITY_MAX:
                    hits.append((
                        i, f"deck[{i - 1}].{key}[{j}]",
                        ("W", "卡片数据堆砌",
                         f"一张卡里 {len(nums)} 个数（{('、'.join(nums))}），"
                         "听众记不住也来不及比。综述页的数字要当锚点用，"
                         "留一两个关键的，其余改成比较式表述"
                         "（「换成实用电极后保持率往往不足一半」比一串"
                         "负载、厚度、掉幅更好懂）。数字细节放讲稿里讲，"
                         "版面留给关系与走向", " ".join(nums), " ".join(nums))))
    return hits


def check_page_density(cfg):
    """同一页各卡片的数字个数落差过大，报提示。

    这是横向问题：一页四张卡，三张零数据、一张四个数，翻过去就能感到
    语体变了——数据播报混在判断之间。    只看单卡通顺，看不出这一层。

    同样默认关闭，靠 content.yaml 的 `data_density: true` 打开。
    """
    if not cfg.get("review") or not cfg.get("data_density"):
        return []
    hits = []
    for i, page in enumerate(cfg.get("deck") or [], 1):
        counts = []
        for key in ("cards", "items", "stages"):
            for j, it in enumerate(page.get(key) or []):
                if not isinstance(it, dict):
                    continue
                txt = DATA_ALLOW.sub("", _card_text(it))
                counts.append((len(DATA_NUM.findall(txt)),
                               f"deck[{i - 1}].{key}[{j}]",
                               it.get("name") or ""))
        if len(counts) < 2:
            continue
        mx = max(c[0] for c in counts)
        mn = min(c[0] for c in counts)
        if mx >= 3 and mx - mn >= 3:
            worst = max(counts)
            hits.append((
                i, worst[1],
                ("W", "同页语体不齐",
                 f"同页各卡的数字个数是 {('、'.join(str(c[0]) for c in counts))}，"
                 f"落差 {mx - mn}。「{worst[2]}」这张在罗列数据，其余在给判断，"
                 "翻过去能感到语体变了。同一页要么都用比较式，要么都保留锚点",
                 worst[2], worst[2])))
    return hits



def walk_texts(node, path="deck"):
    """展开 YAML，产出 (路径, 文本)。只取会印到幻灯片或讲稿上的字符串。"""
    if isinstance(node, dict):
        for k, v in node.items():
            yield from walk_texts(v, f"{path}.{k}")
    elif isinstance(node, list):
        for i, v in enumerate(node):
            yield from walk_texts(v, f"{path}[{i}]")
    elif isinstance(node, str):
        yield path, node


def sentences(text):
    """按句读切句。讲稿里一句一义，切得开才能定位。"""
    return [s.strip() for s in re.split(r"[。；！？\n]+", text) if s.strip()]


def check_text(path, text):
    """对一段文字跑全部规则，返回命中列表 (级别, 规则, 说明, 原句)。"""
    hits = []
    for level, name, pattern, advice in RULES:
        rx = re.compile(pattern)
        for s in sentences(text):
            m = rx.search(s)
            if not m:
                continue
            # 「从…到…」两端若都是具体数值/单位，就是真范围，放过
            if name == "从…到…" and re.search(r"\d", m.group(0)):
                continue
            hits.append((level, name, advice, s, m.group(0)))
    return hits


def check_title(path, page, mode="thesis"):
    """标题检查。

    两层，所有页面都要过第一层：
      1. 标点——标题里不能有逗号/冒号/分号（严重），顿号只提示；
         口语比喻（严重，投影出去的门面不能带说书腔）。
      2. 判断——封面与章节页免检，其余按 `title_mode` 判：
         thesis（默认）要一句能读通的结论句；label 用任务名/提纲名，
         本来就不该有谓词——反过来还要查「有没有偷偷写成判断句」，
         否则一条判断句混在栏目名里，翻页时语体就断了。
    """
    t = page.get("title_cn") or page.get("title")
    if not t:
        return []
    hits = []
    m = re.search(r"[，,：:；;]", t)
    if m:
        hits.append(("S", "标题带逗号/冒号",
                     f"「{m.group(0)}」两侧是各自的判断，标题只留一个："
                     "把前半句压进主语，或整句改写成一个主谓宾", t, t))
    elif "、" in t:
        hits.append(("W", "标题并列名词",
                     "顿号堆名词，读者要等谓词落地；建议改成一个判断句", t, t))
    c = TITLE_COLLOQUIAL.search(t)
    if c:
        hits.append(("S", "标题口语",
                     f"「{c.group(0)}」是口语或比喻。汇报标题用书面语："
                     "「每级都要还上一级的账」→「每个阶段均有相应代价」", t, t))
    if page.get("type") in SKIP_TYPES:
        return hits
    if mode == "label":
        v = LABEL_VERDICT.search(t)
        if v:
            hits.append(("S", "标签模式写判断句",
                         f"「{v.group(0)}」把标题写成了判断，而全篇用的是栏目名"
                         "（label），翻到这一页语体会断。两选一：改成名词短语"
                         "（「现有瓶颈集中在概念层面」→「概念层面的四个瓶颈」），"
                         "或整篇改 title_mode: thesis", t, t))
        return hits
    if LABEL_TITLE.match(t):
        hits.append(("W", "栏目名标题",
                     "文献汇报可作栏目名；若这是在讲自己的研究，建议改成结论句", t, t))
    elif not VERDICT_WORDS.search(t):
        hits.append(("W", "名词短语标题",
                     "标题没给出判断；试着写成一句能读通的话（含动词或数值）；"
                     "文献汇报用任务名时，在 content.yaml 写 title_mode: label", t, t))
    return hits


def check_page_titles(cfg):
    mode = cfg.get("title_mode") or "thesis"
    hits = []
    for i, page in enumerate(cfg.get("deck") or [], 1):
        for h in check_title(f"deck[{i - 1}]", page, mode):
            hits.append((i, f"deck[{i - 1}].title", h))
    return hits


def check_subheads(cfg):
    """卡片名 / 要点名也是标题，同一条标点规则。

    只查 dict 里的 name：items 允许写成纯字符串，那是正文句子，
    逗号是正常行文，不能一并当标题判。冒号放行，EC:DMC 这类是化学写法。

    label 模式下另提示「小标题写成判断句」：卡片名比页标题松，短句也读得通，
    所以只报提示不判失败——但一页里四条卡片若都是判断句、下一页全是名词，
    翻过去就能感到语体在变，值得作者自己看一眼。
    """
    mode = cfg.get("title_mode") or "thesis"
    hits = []
    for i, page in enumerate(cfg.get("deck") or [], 1):
        # stages 的 name 也是投影上的小标题（阶段名），同一套判据。
        for key in ("cards", "items", "stages"):
            for j, it in enumerate(page.get(key) or []):
                if not isinstance(it, dict):
                    continue
                n = it.get("name")
                if not n:
                    continue
                m = re.search(r"[，,；;]", n)
                if m:
                    hits.append((
                        i, f"deck[{i - 1}].{key}[{j}].name",
                        ("S", "小标题带逗号",
                         f"「{m.group(0)}」把两个判断并排放，小标题只留一个；"
                         "删掉前半句，或改成一个主谓宾", n, n)))
                c = TITLE_COLLOQUIAL.search(n)
                if c:
                    hits.append((
                        i, f"deck[{i - 1}].{key}[{j}].name",
                        ("S", "小标题口语",
                         f"「{c.group(0)}」是口语或比喻。卡片名是投影上的标题，"
                         "用书面语：「孔细到装不下溶剂壳」→"
                         "「亚纳米孔内的部分去溶剂化」", n, n)))
                if mode == "label":
                    v = LABEL_VERDICT.search(n)
                    if v:
                        hits.append((
                            i, f"deck[{i - 1}].{key}[{j}].name",
                            ("W", "小标题写成判断句",
                             f"「{v.group(0)}」是谓词。同页其余卡片若是名词短语，"
                             "这一条会显眼；改成名词短语更齐（判断留给正文与讲稿）",
                             n, n)))
    return hits


def density_report(texts):
    """词频密度：同一虚词反复出现，是比单句更明显的机器味。"""
    joined = "".join(texts)
    n = max(1, len(joined))
    out = []
    for w, limit in (("关键", 400), ("决定", 500), ("本质", 800),
                     ("维度", 700), ("赋能", 3000), ("底层", 3000)):
        c = joined.count(w)
        if c and n / c < limit:
            out.append((w, c, round(n / c)))
    return out


def run_one(config, strict):
    with open(os.path.join(ROOT, config), encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    # projects/_xxx 是示例与骨架，文案本身是占位说明（「每条结论一句话」这类），
    # 拿它当交付稿挑文风没有意义。只报告，不阻断。
    demo = os.path.basename(os.path.dirname(config)).startswith("_")
    # 讲稿与版面文案分开统计：讲稿允许口语，但规则一样适用
    rows, texts = [], []
    for path, text in walk_texts(cfg, "deck"):
        if path.endswith(".type") or path.endswith(".image"):
            continue
        texts.append(text)
        for level, name, advice, sent, frag in check_text(path, text):
            rows.append((level, path, name, frag, sent, advice))

    for page_no, path, (level, name, advice, sent, frag) in check_page_titles(cfg):
        rows.append((level, path, name, frag, sent, advice, page_no))

    for page_no, path, (level, name, advice, sent, frag) in check_subheads(cfg):
        rows.append((level, path, name, frag, sent, advice, page_no))

    for page_no, path, (level, name, advice, sent, frag) in check_data_density(cfg):
        rows.append((level, path, name, frag, sent, advice, page_no))

    for page_no, path, (level, name, advice, sent, frag) in check_page_density(cfg):
        rows.append((level, path, name, frag, sent, advice, page_no))

    print(f"文风体检：{config}")
    print(f"  扫描 {len(texts)} 段文字")
    if not rows:
        print("  未发现 AI 腔句式。")
    else:
        sev = {}
        for r in rows:
            sev[r[0]] = sev.get(r[0], 0) + 1
        print(f"  命中 {len(rows)} 处：严重 {sev.get('S', 0)}，提示 {sev.get('W', 0)}")
        print()
        for r in sorted(rows, key=lambda x: (x[0] != "S", x[1])):
            level, path, name, frag, sent, advice = r[:6]
            mark = "严重" if level == "S" else "提示"
            print(f"  [{mark}] {name:<8} {path}")
            print(f"         命中：「{frag}」")
            if len(sent) > len(frag) + 2:
                print(f"         原句：{sent}")
            print(f"         建议：{advice}")

    dens = density_report(texts)
    if dens:
        print()
        print("  用词密度（平均每 N 字出现一次，越小越密）：")
        for w, c, every in dens:
            print(f"    「{w}」{c} 次，约每 {every} 字一次")

    bad = [r for r in rows if r[0] == "S"]
    print()
    if demo:
        print("  结论：示例项目，只报告不阻断。")
        return 0
    if bad:
        print(f"  结论：{len(bad)} 处严重问题需要改。")
        return 1
    if strict and rows:
        print("  结论：无严重问题，但有提示项（--strict 下视为未通过）。")
        return 1
    print("  结论：通过。")
    return 0


def main():
    ap = argparse.ArgumentParser(description="文风体检")
    ap.add_argument("config", nargs="?", help="content.yaml 路径")
    ap.add_argument("--all", action="store_true", help="检查 projects/*/content.yaml")
    ap.add_argument("--strict", action="store_true", help="提示项也判为不通过")
    args = ap.parse_args()
    if not args.all and not args.config:
        ap.error("需要给出 content.yaml，或使用 --all")

    configs = ([os.path.relpath(p, ROOT).replace("\\", "/")
                for p in sorted(glob.glob(os.path.join(ROOT, "projects", "*", "content.yaml")))]
               if args.all else [args.config])
    rc = 0
    for c in configs:
        rc |= run_one(c, args.strict)
        print()
    return rc


if __name__ == "__main__":
    sys.exit(main())
