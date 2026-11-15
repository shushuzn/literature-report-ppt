---
name: literature-report-ppt
description: 用 YAML 配置驱动生成研究生「文献汇报 / 课题汇报」PPT，并跑九道量化门禁（文风、孤行、综述、讲稿时长、事实核对、生成自检、预览导出、合规、版式九维审计）验收。当用户要把一篇论文或一个课题做成汇报幻灯片（.pptx），要求字体字号画布版心可配、每页字数有上下限、讲稿控时长、数字必须能回到论文原文、版面能按渲染图量像素验收时使用；也用于修改已有汇报稿的页面、文案或版式。适用于文献精读汇报、组会汇报、开题/中期/答辩预演、综述汇报。
---

# 文献汇报PPT

一页 = `content.yaml` 里的一条记录，`type` 选版式、字段填内容。换文献或换课题只改 YAML 与配图，不写代码。验收标准是**全页 × 全维度**，不挑页也不挑维度。

本仓库根目录即技能目录：克隆下来可直接按技能加载，下文用 `R` 表示仓库根。

## 目录

```
pptkit/                引擎：style / metrics / engine / pages / template / compliance
tools/                 门禁与生成脚本；export_preview.py 为无 PowerPoint 时的预览导出
projects/_template/    9 页可运行样例（新建项目从这里抄）
docs/                  页面类型、文风、合规、完整使用说明、汇报要求原文
agents/                技能面板配置（中文显示名）
```

## 环境

```bash
python3 -m pip install -r requirements.txt   # python-pptx / PyYAML / Pillow / numpy
python3 -m pip install pymupdf               # 仅从 PDF 抽图时需要
```

- 命令按 `python3` 写；Windows 上换成 `py`。
- 有 Microsoft PowerPoint（Windows）：`tools/verify.py` 一条命令跑完全部九道门禁。
- 只有 Linux/LibreOffice：预览用 `tools/export_preview.py`，审计照跑，但字体是替代字体（微软雅黑→Noto Sans CJK），断行可能与放映稿略有偏差，结论按「该渲染器下通过」看待。

## 第 0 步：先定六个数（不可跳过）

页数、演讲时长、正文字号、画布比例、语速、备料冗余决定一页能写多少字、讲稿写多少字。先定容量，再填内容。

```bash
cd $R && python3 tools/intake.py --calc --pages 6 --minutes 5 --body-pt 20 --ratio 16:9
```

把输出的 `style:` 与 `compliance:` 片段直接贴进 `content.yaml`。缺省值：语速 220 字/分、冗余 0.4（加在限时上，不是从限时里扣）、封面字数下限单独放宽。

## 第 1 步：建项目

```bash
mkdir -p $R/projects/<项目>/figs
cp $R/projects/_template/content.yaml $R/projects/<项目>/content.yaml
```

改两行：`out:`（生成路径，相对仓库根；也可用 `-o` 传绝对路径）、`base_dir:`（图片基准目录）。

## 第 2 步：抽图

图片清晰度：正文图宽 ≥ 1500 px，论文首页按 3 倍以上渲染。

```python
import pymupdf, os
d = pymupdf.open("论文.pdf")
outdir = "projects/<项目>/figs"; os.makedirs(outdir, exist_ok=True)
d[0].get_pixmap(matrix=pymupdf.Matrix(3.2, 3.2)).save(f"{outdir}/page1.png")
for pno in range(d.page_count):
    for i, info in enumerate(d[pno].get_images(full=True)):
        im = d.extract_image(info[0])
        if im["width"] < 300 or im["height"] < 220:
            continue
        open(f"{outdir}/p{pno+1}_{i}.{im['ext']}", "wb").write(im["image"])
```

同时把全文导成 `projects/<项目>/_src/paper.txt`，事实核对要用它逐个查数字出处。

## 第 3 步：填内容

`deck:` 是页面序列，**页序即列表顺序，页数即列表长度**。十种版式：

| type | 用途 | 关键字段 |
|---|---|---|
| `cover` | 封面 | `title_cn` `title_en` `presenter` `advisor` `roles` `meta` |
| `section` | 章节过渡 | `index` `title` `lead` |
| `fields_image` | 左图右字段 | `image` `caption` `fields[{label,value}]` |
| `image_cards` | 左图右卡片竖排 | `image` `cards/items[{name,desc}]` |
| `image_grid_cards` | 上图下卡 | `images[{path,caption}]` `items[{name,data,cause}]` |
| `cards_grid` | 卡片栅格 | `columns` `items[{name,desc}]` |
| `compare` | 两栏对照 | `left/right{tag,label,points}` |
| `bullets` | 要点列表 | `lead` `items`（字符串或 `{name,desc}`） |
| `timeline` | 时间轴 | 见 `docs/页面类型.md` |
| `stages` | 阶段/流程图 | 见 `docs/页面类型.md` |

字段级细节（必填项、字号、断行、页级覆盖）一律读 `docs/页面类型.md`，不要凭印象填。写文案前先读 `docs/文风规范.md`——AI 腔靠正则清单抓，不靠肉眼扫。

三条硬规则：

1. **标题按任务名写，不写判断句。** 文献汇报用 `title_mode: label`，每页标题照当次要求里的原话（论文基本信息 / 材料制备与各组分作用 / 主要性能与成因 / 主要应用），封面用论文题目中译。讲自己的研究删掉这行（默认 `thesis`，标题才写判断）。
2. **序号字段写 `index`，不要写 `no`。** YAML 会把裸键 `no`/`yes` 解析成布尔，序号静默消失。
3. **值里出现「英文冒号+空格」要加引号。** `DOI: 10.1038/…`、`时间: 14:30` 会被解析成键值对，报错信息离现场很远。

其余排版约束交给引擎：同一栏多段文字用段间距而非逐段定位、文本框内边距归零、图片外框按真实比例贴合、字号按字体度量反推高度。不要手写 python-pptx 代码绕过 `pptkit`，否则审计量到的和生成的对不上。

## 第 4 步：跑门禁

```bash
# Linux / 无 PowerPoint：先 --skip-audit 跑前八道
cd $R && python3 tools/verify.py projects/<项目>/content.yaml --skip-audit

# 导出预览图（图必须比 pptx 新，否则审计直接退出码 2）
python3 $R/tools/export_preview.py projects/<项目>/<成品>.pptx projects/<项目>/preview

# 版式九维审计
python3 $R/tools/audit_all.py projects/<项目>/<成品>.pptx projects/<项目>/preview
```

Windows + PowerPoint 下只需 `python3 tools/verify.py ...`，第九道会自动生成预览并审计。

单跑某一关（定位问题用）：`style_check.py` `check_text.py` `review_check.py` `note_budget.py` `fact_check.py` `build.py` `check.py` `audit_gaps.py`，参数都是 `content.yaml` 或 `pptx`。

门禁读法：

| 报什么 | 含义 | 怎么改 |
|---|---|---|
| 文风体检命中 | 对偶转折、口号收尾、标题带逗号或冒号、用词密度超标 | 改句式架构，不是换词 |
| 孤行（末行占比 < 0.18） | 断行后末行只剩一两个字 | 补字或改断行；预检偏乐观，0.18–0.22 要留余量 |
| 讲稿超时 | 必讲 > 限时，或全文 > 限时×(1+冗余) | 把整段标〔可略〕，或压正文 |
| 事实核对失败 | 某数字在 `_src/paper.txt` 找不到 | 改数，或写进 `_src/facts_allow.txt` 并注明来源 |
| 每页字数低于下限 | 内容单薄 | 补内容，别调下限 |
| 版式审计某维度 | 越界 / 重叠 / 图片死区 / 框内不居中 / 留白不足 / 段距极差 / 卡片间距 / 栏内填充 | 按提示的页号与数值调 `style` 或文案 |

## 第 5 步：交付

```bash
python3 $R/tools/export_notes.py projects/<项目>/content.yaml --md   # 讲稿逐页字数与用时
```

交付 `.pptx` 与讲稿文本，放在用户指定的项目目录。生成时用 `-o <目标路径>` 直接把成品写到用户目录，`projects/<项目>/` 里只留配置与配图。

## 常见坑（按出现频率排序）

1. **预览图是门禁输入，不是附件。** 改了 `pad_v`、重建 pptx 却没重导图，审计量的是上一版版面，照样「全部通过」。
2. **预检通过 ≠ 渲染通过。** `check_text.py` 按字符宽度估算，`audit_all.py` 量渲染像素，末行占比能差 1–3 个百分点。
3. **限时是最先该问的数。** 讲稿写完才发现超时，等于重写。
4. **别让「写死」把内容压扁。** 卡片留白默认弹性（`cards_fit: true`），内容多时先收留白再收卡间距；版面参数只在 `style:` 段改，别改代码常量。
5. **图注只能一行**（图注区固定 0.32 in），超一行会压到正文。
6. **单图页左栏宽按图片比例反算**，引擎已处理；不要手写左栏宽，会挤窄右栏导致多一个换行。

## 参考文件

- `docs/页面类型.md` — 十种版式的字段、字号、断行约定（填内容时逐页对照）
- `docs/文风规范.md` — AI 腔清单与标题写法（写文案前必读）
- `docs/合规核对.md` — 核什么、阈值从哪个字段读
- `docs/使用说明.md` — 完整操作步骤、阈值来源、样式段全部可配项（长，按需检索）
- `docs/汇报要求.md` — 某次作业的原文存档，了解「当次要求」长什么样
- `LICENSE` — MIT 许可
