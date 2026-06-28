Add-Type -AssemblyName System.Drawing

$outDir = "D:\TREA文件\.claude\scripts\extracted-icons"
$icoPath = "D:\TREA文件\.claude\scripts\claude-code.ico"
$tempJpg = "$outDir\claude-code-temp.jpg"

# Try a few prompts to find a good one
$prompts = @(
  "claude AI app icon, orange rounded square background, simple white letter A, minimalist modern logo, flat design, no text, app icon style, 1024x1024, clean",
  "claude code terminal icon, orange robot character with smiley face, simple cute mascot, rounded square, flat illustration, app icon, no text",
  "anthropic claude logo, orange coral gradient, white star sparkle, minimalist geometric, app icon style, no text"
)

$downloadedJpg = $null
foreach ($p in $prompts) {
  $url = "https://trae-api-cn.mchost.guru/api/ide/v1/text_to_image?prompt=" + [Uri]::EscapeDataString($p) + "&image_size=square_hd"
  try {
    Write-Host "Generating: $p"
    $req = [System.Net.HttpWebRequest]::Create($url)
    $req.Timeout = 90000
    $req.ReadWriteTimeout = 90000
    $req.UserAgent = "Mozilla/5.0"
    $resp = $req.GetResponse()
    $stream = $resp.GetResponseStream()
    $ms = New-Object System.IO.MemoryStream
    $stream.CopyTo($ms)
    $bytes = $ms.ToArray()
    [System.IO.File]::WriteAllBytes($tempJpg, $bytes)
    Write-Host "  OK, $($bytes.Length) bytes"
    $resp.Close()
    $downloadedJpg = $tempJpg
    break
  } catch {
    Write-Host "  Failed: $($_.Exception.Message)"
  }
}

if (-not $downloadedJpg) {
  Write-Host "All prompts failed"
  exit 1
}

# Now convert JPG to a multi-size ICO
# Use ImageMagick if available, otherwise build ICO manually
function ConvertTo-Icon {
  param([string]$SourceJpg, [string]$DestIco, [int[]]$Sizes = @(16, 32, 48, 64, 128, 256))

  Add-Type -AssemblyName System.Drawing

  $src = [System.Drawing.Image]::FromFile($SourceJpg)
  $ms = New-Object System.IO.MemoryStream

  # ICONDIR header
  $bw = New-Object System.IO.BinaryWriter $ms
  $bw.Write([uint16]0)        # reserved
  $bw.Write([uint16]1)        # type = icon
  $bw.Write([uint16]$Sizes.Count)

  # First, build all the image data and store positions
  $entries = @()
  $allImageData = @()
  $headerSize = 6 + 16 * $Sizes.Count
  $currentOffset = $headerSize

  foreach ($sz in $Sizes) {
    $bmp = New-Object System.Drawing.Bitmap $sz, $sz
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.DrawImage($src, 0, 0, $sz, $sz)
    $g.Dispose()

    # Encode as PNG
    $imgMs = New-Object System.IO.MemoryStream
    $bmp.Save($imgMs, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    $imgBytes = $imgMs.ToArray()
    $imgMs.Dispose()

    $allImageData += ,$imgBytes

    $wByte = if ($sz -ge 256) { 0 } else { [byte]$sz }
    $hByte = if ($sz -ge 256) { 0 } else { [byte]$sz }

    $entries += [pscustomobject]@{
      W = $wByte
      H = $hByte
      BPP = 32
      Size = $imgBytes.Length
      Offset = $currentOffset
    }

    $currentOffset += $imgBytes.Length
  }

  # Write directory entries
  foreach ($e in $entries) {
    $bw.Write([byte]$e.W)
    $bw.Write([byte]$e.H)
    $bw.Write([byte]0)  # color count
    $bw.Write([byte]0)  # reserved
    $bw.Write([uint16]1)  # planes
    $bw.Write([uint16]$e.BPP)
    $bw.Write([uint32]$e.Size)
    $bw.Write([uint32]$e.Offset)
  }

  # Write image data
  foreach ($imgBytes in $allImageData) {
    $bw.Write($imgBytes)
  }

  $bw.Flush()
  [System.IO.File]::WriteAllBytes($DestIco, $ms.ToArray())
  $ms.Dispose()
  $src.Dispose()

  Write-Host "Created ICO: $DestIco ($((Get-Item $DestIco).Length) bytes) with sizes: $($Sizes -join ', ')"
}

ConvertTo-Icon -SourceJpg $downloadedJpg -DestIco $icoPath
Write-Host "Done."
