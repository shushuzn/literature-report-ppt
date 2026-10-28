# 导出 PPT 逐页预览图（默认宽 1600px，高度按画布比例算）
# 用法： powershell -File tools/export_preview.ps1 -Ppt "路径.pptx" -Out "输出目录"
#
# 安全约定：
#   · 不接管、也不退出用户已经打开的 PowerPoint；只有本次新建的实例才 Quit。
#   · 只清理输出目录里本脚本命名的 slide*.png，不碰其他文件。
#   · 路径必须显式给出，脚本里不写任何个人目录。

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Ppt,
    [Parameter(Mandatory = $true)]
    [string]$Out,
    [int]$Width = 1600
)

$ErrorActionPreference = 'Stop'
# 被管道接走时按 UTF-8 输出，Git Bash / IDE 里中文才不会乱码。
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
$Ppt = (Resolve-Path -LiteralPath $Ppt).Path
$Out = [System.IO.Path]::GetFullPath($Out)
if (-not (Test-Path -LiteralPath $Out)) {
    New-Item -ItemType Directory -Path $Out -Force | Out-Null
}

# 只删本脚本上一轮生成的图。输出目录里用户自己放的 PNG 一律不动。
Get-ChildItem -LiteralPath $Out -Filter 'slide*.png' -File -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue

# 先看 PowerPoint 是否已在运行：已在运行就复用，结束时也不退出它。
$hadPpt = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
$app = $null
$owned = $false
$pres = $null
$openedByUs = $false
try {
    try {
        $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    }
    catch {
        $app = New-Object -ComObject PowerPoint.Application
        # 只有此前没有 PowerPoint 进程时，这个实例才归本次调用所有。
        $owned = -not $hadPpt
    }

    # 同一份稿子已经打开时直接复用它，结束时也不能替用户关掉。
    $existing = $null
    try {
        $existing = @($app.Presentations | Where-Object {
            $_.FullName -and $_.FullName -eq $Ppt
        }) | Select-Object -First 1
    }
    catch { }
    if ($null -ne $existing) {
        $pres = $existing
    }
    else {
        $pres = $app.Presentations.Open($Ppt, $true, $false, $false)
        $openedByUs = $true
    }

    # 导出尺寸按实际画布比例算：写死 1600x900 会把 4:3 的稿子拉变形，
    # 而版式审计是拿像素换算英寸的，一变形所有测量都错。
    $pxH = [int]($Width * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
    $i = 1
    foreach ($sl in $pres.Slides) {
        $sl.Export((Join-Path $Out ("slide{0}.png" -f $i)), "PNG", $Width, $pxH)
        $i++
    }
    Write-Output "导出完成：$($i - 1) 页 -> $Out"
}
finally {
    if ($openedByUs -and $null -ne $pres) {
        try { $pres.Close() } catch { }
    }
    if ($owned -and $null -ne $app) {
        try { $app.Quit() } catch { }
    }
}
