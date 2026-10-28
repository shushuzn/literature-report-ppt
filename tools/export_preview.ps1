# 导出 PPT 逐页预览图（每页 1600x900 PNG）
# 用法： powershell -File tools/export_preview.ps1
#        powershell -File tools/export_preview.ps1 -Ppt "路径.pptx" -Out "输出目录"

param(
    [string]$Ppt = "D:\OpenClaw\2026-10-08-18-15-50\汇报材料\生物塑料汇报_仿生分层膜.pptx",
    [string]$Out = "D:\OpenClaw\2026-10-08-18-15-50\汇报材料\preview"
)

if (-not (Test-Path $Ppt)) { Write-Error "找不到 PPT：$Ppt"; exit 1 }
if (-not (Test-Path $Out)) { New-Item -ItemType Directory -Path $Out -Force | Out-Null }

Get-ChildItem $Out -Filter *.png -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue

$app = New-Object -ComObject PowerPoint.Application
try {
    $pres = $app.Presentations.Open($Ppt, $true, $false, $false)
    # 导出尺寸按实际画布比例算：写死 1600x900 会把 4:3 的稿子拉变形，
    # 而版式审计是拿像素换算英寸的，一变形所有测量都错。
    $pxW = 1600
    $pxH = [int]($pxW * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
    $i = 1
    foreach ($sl in $pres.Slides) {
        $sl.Export((Join-Path $Out ("slide{0}.png" -f $i)), "PNG", $pxW, $pxH)
        $i++
    }
    $pres.Close()
    Write-Output "导出完成：$($i - 1) 页 -> $Out"
} finally {
    $app.Quit()
}
