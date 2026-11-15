#!/usr/bin/env python3
"""把 pptx 逐页导成 slide1.png / slide2.png ...（无 PowerPoint 时的替代方案）。

走 LibreOffice 渲染：pptx -> pdf -> pdftoppm。输出文件名与
engine/tools/audit_all.py 期望的 slide{i}.png 完全一致，可直接喂给版式审计。

用法：
    python scripts/export_preview.py <pptx> <输出目录> [--width 1600]

注意：
    · 只删除输出目录里本脚本命名的 slide*.png，不碰其他文件。
    · 本机没有 PowerPoint / 微软雅黑时，渲染用的是替代字体，断行位置可能与
      最终放映稿不同；审计通过只代表"在这个渲染器下通过"。
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile


def find_soffice():
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    return None


def slide_size(pptx_path):
    """读画布的长宽比，保证导出图不变形（审计要按像素换算英寸）。"""
    try:
        from pptx import Presentation
        prs = Presentation(pptx_path)
        return float(prs.slide_width), float(prs.slide_height)
    except Exception:
        return 13.333, 7.5


def convert_to_pdf(pptx_path, workdir):
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("未找到 LibreOffice（soffice），无法导出预览图")
    profile = os.path.join(workdir, "lo-profile")
    cmd = [
        soffice, "--headless", "--norestore", "--invisible",
        "-env:UserInstallation=file://" + profile.replace(os.sep, "/"),
        "--convert-to", "pdf", "--outdir", workdir, pptx_path,
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE,
                   stderr=subprocess.STDOUT, timeout=300)
    pdf_path = os.path.join(
        workdir, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError("LibreOffice 未产出 PDF：" + pdf_path)
    return pdf_path


def pdf_to_png(pdf_path, outdir, width, height):
    if not shutil.which("pdftoppm"):
        raise RuntimeError("未找到 pdftoppm（poppler-utils），无法把 PDF 转成 PNG")
    prefix = os.path.join(outdir, "_page")
    subprocess.run(
        ["pdftoppm", "-png", "-scale-to-x", str(width), "-scale-to-y", str(height),
         pdf_path, prefix],
        check=True, timeout=300)
    pages = sorted(glob.glob(prefix + "-*.png"))
    if not pages:
        raise RuntimeError("pdftoppm 没有产出任何页面")
    for index, page in enumerate(pages, 1):
        os.replace(page, os.path.join(outdir, "slide%d.png" % index))
    return len(pages)


def main():
    ap = argparse.ArgumentParser(description="pptx 逐页导出 PNG（LibreOffice 方案）")
    ap.add_argument("pptx", help="pptx 路径")
    ap.add_argument("outdir", help="输出目录")
    ap.add_argument("--width", type=int, default=1600, help="导出宽度像素（默认 1600）")
    args = ap.parse_args()

    pptx_path = os.path.abspath(args.pptx)
    outdir = os.path.abspath(args.outdir)
    if not os.path.exists(pptx_path):
        print("找不到 pptx：" + pptx_path)
        return 2
    os.makedirs(outdir, exist_ok=True)
    for stale in glob.glob(os.path.join(outdir, "slide*.png")):
        os.remove(stale)

    w_in, h_in = slide_size(pptx_path)
    height = int(round(args.width * h_in / w_in))

    workdir = tempfile.mkdtemp(prefix="ppt_preview_")
    try:
        pdf_path = convert_to_pdf(pptx_path, workdir)
        count = pdf_to_png(pdf_path, outdir, args.width, height)
    except subprocess.CalledProcessError as error:
        print("转换失败：" + str(error))
        return 1
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

    print("导出完成：%d 页 -> %s（%dx%d）" % (count, outdir, args.width, height))

    missing = []
    if shutil.which("pdftoppm") is None:
        missing.append("poppler-utils")
    if missing:
        print("提示：缺 " + "、".join(missing))
    return 0


if __name__ == "__main__":
    sys.exit(main())
