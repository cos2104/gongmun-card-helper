# Draws the default app icon (blue rounded square, white sheet with folded corner, text lines, red seal)
# and writes icon.ico with PNG frames 16/24/32/48/64/256.
Add-Type -AssemblyName System.Drawing
Set-Location $PSScriptRoot

function RoundRect([System.Drawing.RectangleF]$r, [float]$rad) {
  $p = New-Object System.Drawing.Drawing2D.GraphicsPath
  $d = $rad * 2
  $p.AddArc($r.X, $r.Y, $d, $d, 180, 90)
  $p.AddArc($r.Right - $d, $r.Y, $d, $d, 270, 90)
  $p.AddArc($r.Right - $d, $r.Bottom - $d, $d, $d, 0, 90)
  $p.AddArc($r.X, $r.Bottom - $d, $d, $d, 90, 90)
  $p.CloseFigure()
  return $p
}

function Draw([int]$n) {
  $bmp = New-Object System.Drawing.Bitmap $n, $n
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = 'AntiAlias'
  $g.Clear([System.Drawing.Color]::Transparent)
  $k = $n / 256.0
  # background
  $bg = RoundRect (New-Object System.Drawing.RectangleF (8 * $k), (8 * $k), (240 * $k), (240 * $k)) (52 * $k)
  $br = New-Object System.Drawing.Drawing2D.LinearGradientBrush (New-Object System.Drawing.PointF 0, 0), (New-Object System.Drawing.PointF 0, $n), ([System.Drawing.Color]::FromArgb(24, 120, 220)), ([System.Drawing.Color]::FromArgb(0, 84, 165))
  $g.FillPath($br, $bg)
  # sheet with folded corner
  $x = 66 * $k; $y = 44 * $k; $w = 124 * $k; $h = 168 * $k; $f = 38 * $k
  $sheet = New-Object System.Drawing.Drawing2D.GraphicsPath
  $sheet.AddLines([System.Drawing.PointF[]]@(
    (New-Object System.Drawing.PointF $x, $y),
    (New-Object System.Drawing.PointF ($x + $w - $f), $y),
    (New-Object System.Drawing.PointF ($x + $w), ($y + $f)),
    (New-Object System.Drawing.PointF ($x + $w), ($y + $h)),
    (New-Object System.Drawing.PointF $x, ($y + $h))))
  $sheet.CloseFigure()
  $g.FillPath([System.Drawing.Brushes]::White, $sheet)
  $fold = New-Object System.Drawing.Drawing2D.GraphicsPath
  $fold.AddLines([System.Drawing.PointF[]]@(
    (New-Object System.Drawing.PointF ($x + $w - $f), $y),
    (New-Object System.Drawing.PointF ($x + $w - $f), ($y + $f)),
    (New-Object System.Drawing.PointF ($x + $w), ($y + $f))))
  $fold.CloseFigure()
  $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(200, 222, 245))), $fold)
  # text lines
  if ($n -ge 24) {
    $line = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(150, 170, 195))
    $ly = $y + 58 * $k
    foreach ($len in 88, 88, 64) {
      $lp = RoundRect (New-Object System.Drawing.RectangleF ($x + 18 * $k), $ly, ($len * $k), (10 * $k)) (5 * $k)
      $g.FillPath($line, $lp); $ly += 24 * $k
    }
  }
  # red seal
  $sr = 30 * $k; $cx = $x + $w - 34 * $k; $cy = $y + $h - 34 * $k
  $g.FillEllipse((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(222, 48, 48))), ($cx - $sr), ($cy - $sr), (2 * $sr), (2 * $sr))
  if ($n -ge 32) { $g.FillEllipse([System.Drawing.Brushes]::White, ($cx - $sr * 0.45), ($cy - $sr * 0.45), ($sr * 0.9), ($sr * 0.9)); $g.FillEllipse((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(222, 48, 48))), ($cx - $sr * 0.28), ($cy - $sr * 0.28), ($sr * 0.56), ($sr * 0.56)) }
  $g.Dispose()
  return $bmp
}

$sizes = 16, 24, 32, 48, 64, 256
$pngs = @()
foreach ($s in $sizes) { $b = Draw $s; $ms = New-Object System.IO.MemoryStream; $b.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png); $pngs += , $ms.ToArray(); if ($s -eq 256) { $b.Save("$PSScriptRoot\icon.png", [System.Drawing.Imaging.ImageFormat]::Png) }; $b.Dispose() }
$out = New-Object System.IO.MemoryStream
$bw = New-Object System.IO.BinaryWriter $out
$bw.Write([UInt16]0); $bw.Write([UInt16]1); $bw.Write([UInt16]$sizes.Count)
$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
  $s = $sizes[$i]; $dim = if ($s -ge 256) { 0 } else { $s }
  $bw.Write([byte]$dim); $bw.Write([byte]$dim); $bw.Write([byte]0); $bw.Write([byte]0)
  $bw.Write([UInt16]1); $bw.Write([UInt16]32); $bw.Write([UInt32]$pngs[$i].Length); $bw.Write([UInt32]$offset)
  $offset += $pngs[$i].Length
}
foreach ($p in $pngs) { $bw.Write($p) }
$bw.Flush()
[System.IO.File]::WriteAllBytes("$PSScriptRoot\icon.ico", $out.ToArray())
"icon.ico " + $out.Length + " bytes"
