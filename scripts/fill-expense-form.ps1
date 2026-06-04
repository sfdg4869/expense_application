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
    $expenseAmount = $cardCol6 - $personal

    $userName = As-Text $tx.userName
    if ($userName -eq "") { $userName = As-Text $defaults.userName }
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
  for ($si = 2; $si -le [Math]::Min(3, $wb.Worksheets.Count); $si++) {
    $rs = $wb.Worksheets.Item($si)
    for ($i = $rs.Shapes.Count; $i -ge 1; $i--) {
      try {
        $shape = $rs.Shapes.Item($i)
        if ($shape.Type -eq 13 -or $shape.Type -eq 11) { $shape.Delete() }
      } catch {}
    }

    $box1 = $rs.Range("A3:H21")
    $box2 = $rs.Range("I3:P21")
    $slots = @(
      @{ Left = $box1.Left; Top = $box1.Top; W = ($box1.Width / 2); H = $box1.Height },
      @{ Left = ($box1.Left + ($box1.Width / 2)); Top = $box1.Top; W = ($box1.Width / 2); H = $box1.Height },
      @{ Left = $box2.Left; Top = $box2.Top; W = ($box2.Width / 2); H = $box2.Height },
      @{ Left = ($box2.Left + ($box2.Width / 2)); Top = $box2.Top; W = ($box2.Width / 2); H = $box2.Height }
    )

    foreach ($slot in $slots) {
      if ($imgIndex -ge $images.Count) { break }
      $imgPath = [string](As-Text $images[$imgIndex])
      if ($imgPath -ne "" -and (Test-Path -LiteralPath $imgPath)) {
        try {
          $pic = $rs.Shapes.AddPicture($imgPath, $false, $true, $slot.Left, $slot.Top, 10, 10)
          $ratio = $pic.Width / [double]$pic.Height
          $w = [double]$slot.W
          $h = $w / $ratio
          if ($h -gt $slot.H) { $h = [double]$slot.H; $w = $h * $ratio }
          $pic.Width = $w
          $pic.Height = $h
          $pic.Left = $slot.Left + (($slot.W - $w) / 2)
          $pic.Top = $slot.Top + (($slot.H - $h) / 2)
        } catch {}
      }
      $imgIndex++
    }
    if ($imgIndex -ge $images.Count) { break }
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