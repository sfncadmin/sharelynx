<#
.SYNOPSIS
    Generates the add-in PNG icons (Fluent-blue tile with a white link glyph
    from Segoe MDL2 Assets) into src\assets\.
#>
[CmdletBinding()]
param(
    [int[]]$Sizes = @(16, 25, 32, 48, 64, 80, 128)
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

$assetDir = Join-Path (Split-Path $PSScriptRoot -Parent) "src\assets"
if (-not (Test-Path $assetDir)) { New-Item -ItemType Directory -Path $assetDir | Out-Null }

foreach ($size in $Sizes) {
    $bmp = New-Object System.Drawing.Bitmap($size, $size)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.Clear([System.Drawing.Color]::Transparent)

    # rounded-rect background
    $radius = [Math]::Max(2, [int]($size / 5))
    $rectPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $radius * 2
    $rectPath.AddArc(0, 0, $d, $d, 180, 90)
    $rectPath.AddArc($size - $d - 1, 0, $d, $d, 270, 90)
    $rectPath.AddArc($size - $d - 1, $size - $d - 1, $d, $d, 0, 90)
    $rectPath.AddArc(0, $size - $d - 1, $d, $d, 90, 90)
    $rectPath.CloseFigure()
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 15, 108, 189))
    $g.FillPath($brush, $rectPath)

    # white link glyph (Segoe MDL2 Assets 0xE71B)
    $fontSize = [Math]::Max(8, $size * 0.55)
    $font = New-Object System.Drawing.Font("Segoe MDL2 Assets", $fontSize, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $glyph = [char]0xE71B
    $fmt = New-Object System.Drawing.StringFormat
    $fmt.Alignment = [System.Drawing.StringAlignment]::Center
    $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
    $white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $g.DrawString($glyph, $font, $white, (New-Object System.Drawing.RectangleF(0, 0, $size, $size)), $fmt)

    $outPath = Join-Path $assetDir "icon-$size.png"
    $bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose(); $font.Dispose(); $brush.Dispose(); $white.Dispose()
    Write-Host "Wrote $outPath"
}
