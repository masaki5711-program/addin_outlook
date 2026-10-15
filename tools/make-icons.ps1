# アドインのアイコン（青地に白い封筒）を各サイズの PNG で作る
Add-Type -AssemblyName System.Drawing

$out = Join-Path $PSScriptRoot '..\assets'
New-Item -ItemType Directory -Force $out | Out-Null

foreach ($size in 16, 32, 64, 80, 128) {
  $bmp = New-Object System.Drawing.Bitmap $size, $size
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = 'AntiAlias'
  $g.Clear([System.Drawing.Color]::Transparent)

  $r = [Math]::Max(2, $size * 0.2)
  $path = New-Object System.Drawing.Drawing2D.GraphicsPath
  $path.AddArc(0, 0, $r, $r, 180, 90)
  $path.AddArc($size - $r - 1, 0, $r, $r, 270, 90)
  $path.AddArc($size - $r - 1, $size - $r - 1, $r, $r, 0, 90)
  $path.AddArc(0, $size - $r - 1, $r, $r, 90, 90)
  $path.CloseFigure()
  $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(15, 108, 189))), $path)

  $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::White), ([Math]::Max(1.2, $size / 14))
  $pen.LineJoin = 'Round'
  $x = $size * 0.2; $y = $size * 0.3; $w = $size * 0.6; $h = $size * 0.42
  $g.DrawRectangle($pen, [float]$x, [float]$y, [float]$w, [float]$h)
  $g.DrawLines($pen, [System.Drawing.PointF[]]@(
      (New-Object System.Drawing.PointF $x, $y),
      (New-Object System.Drawing.PointF ($x + $w / 2), ($y + $h * 0.6)),
      (New-Object System.Drawing.PointF ($x + $w), $y)))

  $bmp.Save((Join-Path $out "icon-$size.png"), [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose(); $bmp.Dispose()
}
