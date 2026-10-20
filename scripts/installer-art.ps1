# Draws the installer's side and header images from assets/icon.png: .\scripts\installer-art.ps1
Add-Type -AssemblyName System.Drawing
$root = Split-Path $PSScriptRoot
$out = Join-Path $root "src-tauri\installer"
New-Item -ItemType Directory -Force $out | Out-Null
$icon = [Drawing.Image]::FromFile((Join-Path $root "assets\icon.png"))

function New-Canvas($w, $h) {
  $bmp = New-Object Drawing.Bitmap $w, $h, ([Drawing.Imaging.PixelFormat]::Format24bppRgb)
  $g = [Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = "AntiAlias"
  $g.InterpolationMode = "HighQualityBicubic"
  $g.TextRenderingHint = "AntiAliasGridFit"
  return $bmp, $g
}
function Center($g, $text, $font, $brush, $w, $y) {
  $size = $g.MeasureString($text, $font)
  $g.DrawString($text, $font, $brush, [single](($w - $size.Width) / 2), [single]$y)
}

# Side image on the welcome and finish pages: the icon's gradient, the icon, name and tagline
$bmp, $g = New-Canvas 164 314
$rect = New-Object Drawing.Rectangle 0, 0, 164, 314
$bg = New-Object Drawing.Drawing2D.LinearGradientBrush $rect, ([Drawing.Color]::FromArgb(79, 141, 255)), ([Drawing.Color]::FromArgb(91, 91, 240)), 70
$g.FillRectangle($bg, $rect)
$g.DrawImage($icon, 38, 54, 88, 88)
$white = [Drawing.Brushes]::White
$soft = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(215, 255, 255, 255))
$faint = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(160, 255, 255, 255))
Center $g "Ribbonless" (New-Object Drawing.Font "Segoe UI Semibold", 15) $white 164 158
Center $g "Word files," (New-Object Drawing.Font "Segoe UI", 9) $soft 164 192
Center $g "without the ribbon." (New-Object Drawing.Font "Segoe UI", 9) $soft 164 208
Center $g "Free and open source" (New-Object Drawing.Font "Segoe UI", 7.5) $faint 164 284
$bmp.Save((Join-Path $out "sidebar.bmp"), [Drawing.Imaging.ImageFormat]::Bmp)

# Header strip on the other pages, shown top right on white
$bmp, $g = New-Canvas 150 57
$g.Clear([Drawing.Color]::White)
$g.DrawImage($icon, 103, 8, 41, 41)
$bmp.Save((Join-Path $out "header.bmp"), [Drawing.Imaging.ImageFormat]::Bmp)

$icon.Dispose()
"wrote $out\sidebar.bmp and header.bmp"
