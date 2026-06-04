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

    function Get-ReceiptImageSlots($sheet) {
      # 템플릿 점선 칸: A3:D21 한 칸, E3:H21, I3:L21, M3:P21 / 아래줄 동일
      $addresses = @(
        "A3:D21", "E3:H21", "I3:L21", "M3:P21",
        "A22:D40", "E22:H40", "I22:L40", "M22:P40"
      )
      $slots = New-Object System.Collections.Generic.List[object]
      foreach ($addr in $addresses) {
        $r = $sheet.Range($addr)
        $slots.Add(@{
          Left = [double]$r.Left
          Top = [double]$r.Top
          W = [Math]::Max(1, [double]$r.Width)
          H = [Math]::Max(1, [double]$r.Height)
        }) | Out-Null
      }
      return $slots
    }

    function Get-ImageSizeInPoints([string]$path) {
      $img = [System.Drawing.Image]::FromFile($path)
      try {
        $w = [double]$img.Width * 72.0 / 96.0
        $h = [double]$img.Height * 72.0 / 96.0
        return @{ W = $w; H = $h }
      }
      finally { $img.Dispose() }
    }

    function Add-FitPicture($sheet, [string]$imgPath, $slot) {
      if ($imgPath -eq "" -or -not (Test-Path -LiteralPath $imgPath)) { return }

      $orig = Get-ImageSizeInPoints $imgPath
      if ($orig.W -le 0 -or $orig.H -le 0) { return }

      # A3:D21 칸(Left/Top/Width/Height) 안에 전부 들어가게 맞춤
      $scale = [Math]::Min($slot.W / $orig.W, $slot.H / $orig.H)
      $fitW = $orig.W * $scale
      $fitH = $orig.H * $scale

      $left = $slot.Left + (($slot.W - $fitW) / 2)
      $top = $slot.Top + (($slot.H - $fitH) / 2)
      if ($left -lt $slot.Left) { $left = $slot.Left }
      if ($top -lt $slot.Top) { $top = $slot.Top }
      if (($left + $fitW) -gt ($slot.Left + $slot.W)) { $left = $slot.Left + $slot.W - $fitW }
      if (($top + $fitH) -gt ($slot.Top + $slot.H)) { $top = $slot.Top + $slot.H - $fitH }

      $pic = $sheet.Shapes.AddPicture($imgPath, $false, $true, $left, $top, $fitW, $fitH)
      $pic.LockAspectRatio = -1
      $pic.Placement = 1
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
        try { Add-FitPicture $rs $imgPath $slots[$s] } catch {}
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