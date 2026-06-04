param(
  [Parameter(Mandatory = $true)]
  [string]$StatementPath,
  [Parameter(Mandatory = $true)]
  [string]$OutputJsonPath
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "ExpenseCommon.ps1")

if (-not (Test-Path -LiteralPath $StatementPath)) {
  throw "Cannot find statement file"
}

$excel = $null
$wb = $null

function Get-CellText($cell) {
  $t = $cell.Text
  if ($t) { return "$t".Trim() }
  return (As-Text $cell.Value2)
}

function Get-CellNumber($cell) { return As-Number $cell.Value2 }

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.EnableEvents = $false

  $wb = $excel.Workbooks.Open($StatementPath, $null, $true)
  $ws = $wb.Worksheets.Item(1)

  $dataStartRow = 0
  for ($r = 1; $r -le 30; $r++) {
    $v = As-Text $ws.Cells.Item($r, 1).Value2
    if ($v -match '^\d{4}\.\d{2}\.\d{2}$') {
      $dataStartRow = $r
      break
    }
  }
  if ($dataStartRow -eq 0) { throw "No transaction rows found (expected date like 2026.04.01 in column A)" }

  $list = [System.Collections.Generic.List[object]]::new()
  $lastRow = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1

  for ($r = $dataStartRow; $r -le $lastRow; $r++) {
    $useDate = Get-CellText $ws.Cells.Item($r, 1)
    if ($useDate -notmatch '^\d{4}\.\d{2}\.\d{2}') { continue }

    $cols = @()
    for ($c = 1; $c -le 14; $c++) {
      $cell = $ws.Cells.Item($r, $c)
      if ($c -in 6, 7) { $cols += (Get-CellNumber $cell) }
      else { $cols += (Get-CellText $cell) }
    }

    $claim = Get-CellNumber $ws.Cells.Item($r, 7)
    if ($claim -eq 0) { $claim = Get-CellNumber $ws.Cells.Item($r, 6) }
    $merchant = Get-CellText $ws.Cells.Item($r, 10)
    if ($claim -eq 0 -and $merchant -eq "") { continue }

    $list.Add([ordered]@{
      cardColumns = $cols
      domesticClaimAmount = $claim
      personalUseAmount = 0
      headCount = 1
      companions = ""
      merchant = $merchant
      useDate = $useDate
    })
  }

  $wb.Close($false)
  $wb = $null

  if ($list.Count -eq 0) { throw "No transactions found in statement" }

  @{ transactions = $list } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputJsonPath -Encoding UTF8
  Write-Output $list.Count
}
finally {
  if ($wb) { try { $wb.Close($false) } catch {} }
  if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
