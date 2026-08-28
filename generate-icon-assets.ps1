param(
  [string]$SourcePath = (Join-Path $PSScriptRoot 'icon-master-ai.png'),
  [string]$OutputDirectory = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

Add-Type -AssemblyName System.Drawing

$sourceFullPath = [System.IO.Path]::GetFullPath($SourcePath)
$outputFullPath = [System.IO.Path]::GetFullPath($OutputDirectory)
if (-not [System.IO.File]::Exists($sourceFullPath)) {
  throw "Icon master not found: $sourceFullPath"
}
[System.IO.Directory]::CreateDirectory($outputFullPath) | Out-Null

$source = New-Object System.Drawing.Bitmap($sourceFullPath)

function New-ResizedPngBytes {
  param([int]$Size)

  $bitmap = New-Object System.Drawing.Bitmap($Size, $Size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
  try {
    $graphics.Clear([System.Drawing.Color]::Transparent)
    $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
    $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $destination = New-Object System.Drawing.Rectangle(0, 0, $Size, $Size)
    $graphics.DrawImage($source, $destination, 0, 0, $source.Width, $source.Height, [System.Drawing.GraphicsUnit]::Pixel)
  }
  finally {
    $graphics.Dispose()
  }

  $stream = New-Object System.IO.MemoryStream
  try {
    $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
    return ,$stream.ToArray()
  }
  finally {
    $stream.Dispose()
    $bitmap.Dispose()
  }
}

try {
  $pngBySize = @{}
  foreach ($size in @(16, 32, 48, 64, 80, 256)) {
    $pngBySize[$size] = New-ResizedPngBytes -Size $size
  }

  foreach ($size in @(16, 32, 64, 80)) {
    $pngPath = Join-Path $outputFullPath ("icon{0}.png" -f $size)
    [System.IO.File]::WriteAllBytes($pngPath, $pngBySize[$size])
  }

  $icoSizes = @(16, 32, 48, 64, 256)
  $headerSize = 6 + (16 * $icoSizes.Count)
  $offset = $headerSize
  $iconStream = New-Object System.IO.MemoryStream
  $writer = New-Object System.IO.BinaryWriter($iconStream)
  try {
    $writer.Write([uint16]0)
    $writer.Write([uint16]1)
    $writer.Write([uint16]$icoSizes.Count)

    foreach ($size in $icoSizes) {
      $dimension = if ($size -eq 256) { [byte]0 } else { [byte]$size }
      $data = $pngBySize[$size]
      $writer.Write($dimension)
      $writer.Write($dimension)
      $writer.Write([byte]0)
      $writer.Write([byte]0)
      $writer.Write([uint16]1)
      $writer.Write([uint16]32)
      $writer.Write([uint32]$data.Length)
      $writer.Write([uint32]$offset)
      $offset += $data.Length
    }

    foreach ($size in $icoSizes) {
      $writer.Write([byte[]]$pngBySize[$size])
    }
    $writer.Flush()
    [System.IO.File]::WriteAllBytes((Join-Path $outputFullPath 'app.ico'), $iconStream.ToArray())
  }
  finally {
    $writer.Dispose()
    $iconStream.Dispose()
  }

  $pagesDirectory = Join-Path $outputFullPath 'pages'
  if ([System.IO.Directory]::Exists($pagesDirectory)) {
    foreach ($size in @(16, 32, 64, 80)) {
      [System.IO.File]::Copy(
        (Join-Path $outputFullPath ("icon{0}.png" -f $size)),
        (Join-Path $pagesDirectory ("icon{0}.png" -f $size)),
        $true
      )
    }
  }

  Write-Host 'Generated icon16.png, icon32.png, icon64.png, icon80.png, app.ico, and Pages copies.'
}
finally {
  $source.Dispose()
}
