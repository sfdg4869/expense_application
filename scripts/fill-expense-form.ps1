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
    function Remove-ReceiptPictures($sheet) {
      for ($i = $sheet.Shapes.Count; $i -ge 1; $i--) {
        try {
          $shape = $sheet.Shapes.Item($i)
          if ($shape.Type -eq 13 -or $shape.Type -eq 11) { $shape.Delete() }
        } catch {}
      }
    }

    function Get-ReceiptImageSlots($sheet) {
      $boxes = @(
        $sheet.Range("A3:H21"),
        $sheet.Range("I3:P21"),
        $sheet.Range("A22:H40"),
        $sheet.Range("I22:P40")
      )
      $slots = New-Object System.Collections.Generic.List[object]
      foreach ($box in $boxes) {
        $halfW = $box.Width / 2
        $slots.Add(@{ Left = $box.Left; Top = $box.Top; W = $halfW; H = $box.Height }) | Out-Null
        $slots.Add(@{ Left = ($box.Left + $halfW); Top = $box.Top; W = $halfW; H = $box.Height }) | Out-Null
      }
      return $slots
    }

    function Add-CoverPicture($sheet, [string]$imgPath, $slot) {
      if ($imgPath -eq "" -or -not (Test-Path -LiteralPath $imgPath)) { return }
      $pic = $sheet.Shapes.AddPicture($imgPath, $false, $true, $slot.Left, $slot.Top, 10, 10)
      $scale = [Math]::Max($slot.W / $pic.Width, $slot.H / $pic.Height)
      $pic.Width = $pic.Width * $scale
      $pic.Height = $pic.Height * $scale
      $pic.Left = $slot.Left + (($slot.W - $pic.Width) / 2)
      $pic.Top = $slot.Top + (($slot.H - $pic.Height) / 2)
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
        try { Add-CoverPicture $rs $imgPath $slots[$s] } catch {}
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