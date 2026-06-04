Add-Type -AssemblyName System.Drawing

$sourcePath = "C:\Users\Chaim\.cursor\projects\f-extention-giuer\assets\icon-source.png"
$outDir = "f:\extention-giuer\icons"

if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

$src = [System.Drawing.Image]::FromFile($sourcePath)

$side = [Math]::Min($src.Width, $src.Height)
$srcX = [int](($src.Width - $side) / 2)
$srcY = [int](($src.Height - $side) / 2)

$sizes = @(16, 32, 48, 128)

foreach ($size in $sizes) {
    $bmp = New-Object System.Drawing.Bitmap $size, $size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality

    $destRect = New-Object System.Drawing.Rectangle 0, 0, $size, $size
    $g.DrawImage($src, $destRect, $srcX, $srcY, $side, $side, [System.Drawing.GraphicsUnit]::Pixel)

    $outPath = Join-Path $outDir "icon-$size.png"
    $bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose()
    $bmp.Dispose()
    Write-Host "Saved $outPath"
}

$src.Dispose()
