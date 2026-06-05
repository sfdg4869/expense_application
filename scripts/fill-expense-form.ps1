param(
  [Parameter(Mandatory = $true)]
  [string]$ConfigPath
)

$ErrorActionPreference = "Stop"
chcp 65001 | Out-Null
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

. (Join-Path $PSScriptRoot "ExpenseCommon.ps1")

$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

$templatePath = $config.templatePath
$outputPath = $config.outputPath
$transactions = @($config.transactions)
$defaults = $config.defaults
$images = @($config.imagePaths | ForEach-Object { "$_" })

if (-not (Test-Path -LiteralPath $templatePath)) {
  throw "Template file not found: $templatePath"
}

$outputDir = Split-Path -Parent $outputPath
if ($outputDir -and -not (Test-Path -LiteralPath $outputDir)) {
  New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

Copy-Item -LiteralPath $templatePath -Destination $outputPath -Force

$excel = $null
$wb = $null
$wbClosed = $false

function Set-CellValue {
  param($Sheet, [int]$Row, [int]$Col, $Value, [switch]$AsString)
  $cell = $Sheet.Cells.Item($Row, $Col)
  if ($AsString) {
    $cell.Value2 = [string](As-Text $Value)
  }
  else {
    $cell.Value2 = (As-Number $Value)
  }
}

function Clear-UsageRows {
  param($Sheet, [int]$StartRow, [int]$EndRow)
  # 15·21열은 수식열 — 건드리지 않음
  $cols = @(1..14) + @(16..20)
  foreach ($r in $StartRow..$EndRow) {
    foreach ($c in $cols) {
      try { $Sheet.Cells.Item($r, $c).ClearContents() } catch {}
    }
  }
}

function Get-ColValue {
  param($Cols, [int]$Index)
  if ($Index -lt $Cols.Count) { return $Cols[$Index] }
  return $null
}

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.EnableEvents = $false
  $excel.Interactive = $false
  $excel.ScreenUpdating = $false
  $excel.AskToUpdateLinks = $false
  try { $excel.AutomationSecurity = 3 } catch {}

  $wb = $excel.Workbooks.Open($outputPath, $null, $false)
  $usageSheet = $wb.Worksheets.Item(1)

  Clear-UsageRows -Sheet $usageSheet -StartRow 3 -EndRow 320

  $textCardCols = @(1, 2, 3, 4, 5, 8, 9, 10, 11, 12, 13, 14)
  $numericCardCols = @(6, 7)
  $row = 3

  foreach ($tx in $transactions) {
    $cols = @($tx.cardColumns)

    foreach ($colNum in $numericCardCols) {
      $val = Get-ColValue -Cols $cols -Index ($colNum - 1)
      Set-CellValue -Sheet $usageSheet -Row $row -Col $colNum -Value $val
    }

    for ($c = 0; $c -lt 14; $c++) {
      $colNum = $c + 1
      if ($numericCardCols -contains $colNum) { continue }
      $val = Get-ColValue -Cols $cols -Index $c
      if ($textCardCols -contains $colNum) {
        Set-CellValue -Sheet $usageSheet -Row $row -Col $colNum -Value $val -AsString
      }
      else {
        Set-CellValue -Sheet $usageSheet -Row $row -Col $colNum -Value $val
      }
    }

    $personal = As-Number $tx.personalUseAmount
    $cardCol6 = As-Number (Get-ColValue -Cols $cols -Index 5)
    if ($cardCol6 -le 0) { $cardCol6 = Get-TransactionClaimAmount $tx }
    # 20열 경비신청 = min(청구, 12,000×인원), 21열 개인사용은 양식 수식
    $expenseAmount = $cardCol6 - $personal

    $baseUser = As-Text $tx.userName
    if ($baseUser -eq "") { $baseUser = As-Text $defaults.userName }
    $userName = Format-UserCell -UserName $baseUser -Companions (As-Text $tx.companions)
    $detail = As-Text $tx.detail
    if ($detail -eq "") { $detail = As-Text $defaults.detail }

    Set-CellValue -Sheet $usageSheet -Row $row -Col 16 -Value $defaults.account -AsString
    Set-CellValue -Sheet $usageSheet -Row $row -Col 17 -Value $defaults.client -AsString
    Set-CellValue -Sheet $usageSheet -Row $row -Col 18 -Value $userName -AsString
    Set-CellValue -Sheet $usageSheet -Row $row -Col 19 -Value $detail -AsString
    Set-CellValue -Sheet $usageSheet -Row $row -Col 20 -Value $expenseAmount
    $row++
  }

  $imgIndex = 0
  if ($images.Count -gt 0) {
    Add-Type -AssemblyName System.Drawing

    function Remove-ReceiptPictures($sheet) {
      for ($i = $sheet.Shapes.Count; $i -ge 1; $i--) {
        try {
          $shape = $sheet.Shapes.Item($i)
          if ($shape.Type -eq 13 -or $shape.Type -eq 11) { $shape.Delete() }
        } catch {}
      }
    }

    function Get-SlotBounds($sheet, [string]$addr) {
      # 병합 셀 때문에 Range.Width가 커지는 경우 방지 — 좌상·우하 셀 기준
      $rng = $sheet.Range($addr)
      $tl = $rng.Cells.Item(1, 1)
      $br = $rng.Cells.Item($rng.Rows.Count, $rng.Columns.Count)
      $left = [double]$tl.Left
      $top = [double]$tl.Top
      $w = ([double]$br.Left + [double]$br.Width) - $left
      $h = ([double]$br.Top + [double]$br.Height) - $top
      return @{ Left = $left; Top = $top; W = [Math]::Max(1, $w); H = [Math]::Max(1, $h) }
    }

    function Get-ReceiptImageSlots($sheet) {
      $slots = New-Object System.Collections.Generic.List[object]

      # 위 4칸: 잘 나오던 방식 그대로 (Range 전체 + cover)
      foreach ($addr in @("A3:D21", "E3:H21", "I3:L21", "M3:P21")) {
        $r = $sheet.Range($addr)
        $slots.Add(@{
          Left = [double]$r.Left
          Top = [double]$r.Top
          W = [Math]::Max(1, [double]$r.Width)
          H = [Math]::Max(1, [double]$r.Height)
          Fit = "cover"
        }) | Out-Null
      }

      # 아래 4칸만: 22~40행 실선 박스 + 병합셀 보정
      foreach ($addr in @("A22:D40", "E22:H40", "I22:L40", "M22:P40")) {
        $b = Get-SlotBounds $sheet $addr
        $slots.Add(@{
          Left = $b.Left
          Top = $b.Top
          W = $b.W
          H = $b.H
          Fit = "cover"
        }) | Out-Null
      }
      return $slots
    }

    function New-OrientationCorrectedBitmap([System.Drawing.Image]$img) {
      $orientation = 1
      try {
        if ($img.PropertyIdList -contains 274) {
          $orientation = [int]$img.GetPropertyItem(274).Value[0]
        }
      } catch {}

      $bmp = New-Object System.Drawing.Bitmap($img)
      switch ($orientation) {
        2 { [void]$bmp.RotateFlip([System.Drawing.RotateFlipType]::RotateNoneFlipX) }
        3 { [void]$bmp.RotateFlip([System.Drawing.RotateFlipType]::Rotate180FlipNone) }
        4 { [void]$bmp.RotateFlip([System.Drawing.RotateFlipType]::Rotate180FlipX) }
        5 { [void]$bmp.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipX) }
        6 { [void]$bmp.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipNone) }
        7 { [void]$bmp.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipX) }
        8 { [void]$bmp.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipNone) }
      }
      return $bmp
    }

    function New-ResizedImageForSlot([string]$srcPath, $slot) {
      $targetW = [Math]::Max(1, [int][Math]::Round($slot.W * 96.0 / 72.0))
      $targetH = [Math]::Max(1, [int][Math]::Round($slot.H * 96.0 / 72.0))
      $fit = [string](As-Text $slot.Fit)
      if ($fit -eq "") { $fit = "cover" }

      $src = [System.Drawing.Image]::FromFile($srcPath)
      $work = New-OrientationCorrectedBitmap $src
      $src.Dispose()
      $bmp = $null
      try {
        $bmp = New-Object System.Drawing.Bitmap($targetW, $targetH)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try {
          $g.Clear([System.Drawing.Color]::White)
          $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
          $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

          if ($fit -eq "contain") {
            $scale = [Math]::Min($targetW / $work.Width, $targetH / $work.Height) * 0.98
            $drawW = [Math]::Max(1, [int][Math]::Round($work.Width * $scale))
            $drawH = [Math]::Max(1, [int][Math]::Round($work.Height * $scale))
            $x = [int][Math]::Round(($targetW - $drawW) / 2)
            $y = [int][Math]::Round(($targetH - $drawH) / 2)
            $g.DrawImage($work, $x, $y, $drawW, $drawH)
          }
          else {
            $scale = [Math]::Max($targetW / $work.Width, $targetH / $work.Height)
            $srcX = ($work.Width * $scale - $targetW) / 2.0 / $scale
            $srcY = ($work.Height * $scale - $targetH) / 2.0 / $scale
            $srcW = $targetW / $scale
            $srcH = $targetH / $scale
            $dest = New-Object System.Drawing.Rectangle(0, 0, $targetW, $targetH)
            $source = New-Object System.Drawing.Rectangle(
              [int][Math]::Round($srcX),
              [int][Math]::Round($srcY),
              [Math]::Max(1, [int][Math]::Round($srcW)),
              [Math]::Max(1, [int][Math]::Round($srcH))
            )
            $g.DrawImage($work, $dest, $source, [System.Drawing.GraphicsUnit]::Pixel)
          }
        }
        finally { $g.Dispose() }

        $tempPath = Join-Path $env:TEMP ("expense_" + [Guid]::NewGuid().ToString("N") + ".jpg")
        $bmp.Save($tempPath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        return $tempPath
      }
      finally {
        if ($bmp) { $bmp.Dispose() }
        $work.Dispose()
      }
    }

    function Add-SlotPicture($sheet, [string]$imgPath, $slot) {
      if ($imgPath -eq "" -or -not (Test-Path -LiteralPath $imgPath)) { return }

      $tempPath = $null
      try {
        $tempPath = New-ResizedImageForSlot $imgPath $slot
        $pic = $sheet.Shapes.AddPicture($tempPath, $false, $true, $slot.Left, $slot.Top, $slot.W, $slot.H)
        $pic.LockAspectRatio = 0
        $pic.Placement = 1
      }
      finally {
        if ($tempPath -and (Test-Path -LiteralPath $tempPath)) {
          Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
      }
    }

    $templateReceiptIdx = 2
    $slotsPerSheet = 8
    $neededSheets = [int][Math]::Ceiling($images.Count / [double]$slotsPerSheet)

    # 템플릿에 지출증빙 시트는 1개뿐 — 9장부터는 첫 시트를 복사해 바로 뒤에 추가
    for ($copy = 1; $copy -lt $neededSheets; $copy++) {
      $insertAfterIdx = $templateReceiptIdx + $copy - 1
      $wb.Worksheets.Item($templateReceiptIdx).Copy([Type]::Missing, $wb.Worksheets.Item($insertAfterIdx))
    }

    for ($batch = 0; $batch -lt $neededSheets; $batch++) {
      $rs = $wb.Worksheets.Item($templateReceiptIdx + $batch)
      Remove-ReceiptPictures $rs

      $slots = Get-ReceiptImageSlots $rs
      for ($s = 0; $s -lt $slots.Count; $s++) {
        if ($imgIndex -ge $images.Count) { break }
        $imgPath = [string](As-Text $images[$imgIndex])
        try { Add-SlotPicture $rs $imgPath $slots[$s] } catch {}
        $imgIndex++
      }
    }
  }

  $wb.Save()
  $wb.Close($false)
  $wbClosed = $true
  Write-Output "OK: $outputPath"
}
catch {
  Write-Error $_.Exception.Message
  exit 1
}
finally {
  if ($wb -and -not $wbClosed) {
    try { $wb.Close($false) } catch {}
  }
  if ($wb) {
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) } catch {}
    $wb = $null
  }
  if ($excel) {
    try { $excel.DisplayAlerts = $false } catch {}
    try { $excel.Quit() } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch {}
    $excel = $null
  }
}