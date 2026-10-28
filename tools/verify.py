#!/usr/bin/env python
"""全检：按顺序跑完所有质量门禁。

    python tools/verify.py projects/<项目>/content.yaml
    python tools/verify.py --all          # 跑 projects/*/content.yaml 全部项目

依次执行：
    1. 文风体检        对偶转折、口号收尾、空标题等 AI 腔与标签式标题
    2. 孤行预检        内容断行是否会产生孤儿字
    3. 综述门禁        时间轴在不在、跨度够不够、有没有串成研究论文口吻
    4. 讲稿时长        必讲不超限时，全文不超 限时×(1+冗余)
    5. 事实核对        配置里的每个数字能否在原文找到出处
    6. 生成 PPT        字号、版面溢出会在这一步报出
    7. 导出预览图      逐页 PNG，版式审计的输入
    8. 合规核对        页数 / 字体 / 字号 / 配图 / 必填项 / 每页字数
    9. 版式审计        面板内文字是否垂直居中、末行是否孤字、卡内留白

任一步失败即返回非零。第 7 步失败（无 PowerPoint）时第 9 步直接判失败，
不拿旧图顶替——图比稿子旧，审的就是上一版版面。
"""
import argparse
import glob
import os
import subprocess
import sys

from _console import setup_console  # noqa: E402

setup_console()

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PY = sys.executable


def preflight():
    """缺依赖时给出可直接执行的修复命令，而不是在子进程里抛 ImportError。"""
    missing = []
    for mod, pkg in (("pptx", "python-pptx"), ("yaml", "PyYAML"),
                     ("PIL", "Pillow"), ("numpy", "numpy")):
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)
    if missing:
        print("缺少依赖：" + "、".join(missing))
        print("修复：python -m pip install -r requirements.txt")
        return False
    return True


def run(title, args):
    print()
    print("#" * 74)
    print(f"# {title}")
    print("#" * 74, flush=True)
    return subprocess.call([PY] + args, cwd=ROOT)


def export_preview(pptx):
    """把稿子逐页导成 PNG，供版式审计量像素。

    这一步属于门禁，不是手工操作：审计读的是图，图不跟着稿子走就等于没审。
    非 Windows 或没装 PowerPoint 时返回 False，由调用方判失败而不是拿旧图硬审。
    """
    import shutil
    # Git Bash 的 PATH 里通常没有 powershell.exe，按安装路径兜底找一次。
    cands = [shutil.which("powershell"), shutil.which("pwsh"),
             os.path.join(os.environ.get("SystemRoot", r"C:\Windows"),
                          "System32", "WindowsPowerShell", "v1.0", "powershell.exe")]
    exe = next((c for c in cands if c and os.path.exists(c)), None)
    if not exe:
        print("未找到 PowerShell，导不出预览图。")
        return False
    path = os.path.join(ROOT, pptx)
    out = os.path.join(os.path.dirname(path), "preview")
    r = subprocess.run(
        [exe, "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-File", os.path.join(ROOT, "tools", "export_preview.ps1"),
         "-Ppt", path, "-Out", out],
        cwd=ROOT, capture_output=True, text=True,
        encoding="utf-8", errors="replace")
    text = ((r.stdout or "") + (r.stderr or "")).strip()
    if r.returncode != 0:
        print("预览导出失败：")
        print(text[-500:] or "（无输出）")
        return False
    print(text.splitlines()[-1] if text else "预览已导出")
    return True


def verify_one(config, skip_audit):
    """跑一个项目的全部门禁，返回 (退出码, 页数)。"""
    import yaml
    with open(os.path.join(ROOT, config), encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    pptx = cfg["out"]
    # 旧五段式配置没有 deck 列表，页数要按转换后的结果算
    sys.path.insert(0, ROOT)
    from pptkit.template import resolve_pages
    pages = len(resolve_pages(cfg))

    rc = 0
    rc |= run("1/9 文风体检", ["tools/style_check.py", config])
    rc |= run("2/9 孤行预检", ["tools/check_text.py", config])
    rc |= run("3/9 综述门禁（研究论文型项目自动跳过）", ["tools/review_check.py", config])
    rc |= run("4/9 讲稿时长（必讲 ≤ 限时，全文 ≤ 限时×(1+冗余)）", ["tools/note_budget.py", config])
    rc |= run("5/9 事实核对（数值逐个回原文找）", ["tools/fact_check.py", config])
    rc |= run("6/9 生成 PPT", ["tools/build.py", config])
    print()
    print("#" * 74)
    print("# 7/9 导出预览图（版式审计的输入，图必须比稿子新）")
    print("#" * 74, flush=True)
    exported = export_preview(pptx)
    rc |= run("8/9 合规核对", ["tools/check.py", pptx, "-c", config])
    if not skip_audit:
        if exported:
            rc |= run("9/9 版式验收（全页 × 九维度）",
                      ["tools/audit_all.py", pptx,
                       os.path.join(os.path.dirname(pptx), "preview"), config])
        else:
            print()
            print("没有与本稿同步的预览图，版式验收不成立，本步判失败。")
            rc |= 1
    return rc, pages


def main():
    ap = argparse.ArgumentParser(description="全检")
    ap.add_argument("config", nargs="?", help="content.yaml 路径")
    ap.add_argument("--all", action="store_true",
                    help="跑 projects/*/content.yaml 下的全部项目")
    ap.add_argument("--skip-audit", action="store_true",
                    help="跳过依赖预览图的版式审计")
    args = ap.parse_args()

    if not preflight():
        return 2
    if not args.all and not args.config:
        ap.error("需要给出 content.yaml，或使用 --all")

    configs = ([os.path.relpath(p, ROOT).replace("\\", "/")
                for p in sorted(glob.glob(os.path.join(ROOT, "projects", "*", "content.yaml")))]
               if args.all else [args.config])
    if not configs:
        print("projects/ 下没有找到 content.yaml")
        return 2

    results = []
    for cfg in configs:
        name = os.path.basename(os.path.dirname(cfg))
        print()
        print("=" * 74)
        print(f"项目：{name}　（{cfg}）")
        print("=" * 74, flush=True)
        try:
            rc, pages = verify_one(cfg, args.skip_audit)
        except FileNotFoundError:
            print(f"找不到 {cfg}")
            rc, pages = 1, 0
        results.append((name, rc, pages))

    print()
    print("=" * 74)
    print("汇总")
    print("=" * 74)
    for name, rc, pages in results:
        print(f"  {'通过' if rc == 0 else '未通过'}　{name:<22}{pages:>3} 页")
    bad = [n for n, rc, _ in results if rc]
    print()
    print("全检通过。" if not bad else "未通过：" + "、".join(bad))
    print("=" * 74)
    return 0 if not bad else 1


if __name__ == "__main__":
    sys.exit(main())
