# 경비신청서 자동 작성 (PowerShell + Excel COM)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
chcp 65001 | Out-Null

$ScriptDir = $PSScriptRoot
. (Join-Path $ScriptDir "scripts\ExpenseCommon.ps1")

$settings = Load-Settings
$script:presetsConfig = Load-ExpensePresets
$script:statementPath = ""
$script:templatePath = ""
$script:transitPath = ""
$script:receiptFolder = ""
$script:receiptFiles = @()
$script:transactions = @()
$script:transitTransactions = @()

function Pick-File([string]$title, [string]$filter) {
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title = $title
  $dlg.Filter = $filter
  if ($dlg.ShowDialog() -eq "OK") { return $dlg.FileName }
  return $null
}

function Pick-Folder([string]$title) {
  $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
  $dlg.Description = $title
  if ($dlg.ShowDialog() -eq "OK") { return $dlg.SelectedPath }
  return $null
}

function Pick-ImageFiles([string]$title) {
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title = $title
  $dlg.Filter = "이미지 (*.jpg;*.jpeg;*.png;*.webp;*.heic)|*.jpg;*.jpeg;*.png;*.webp;*.heic"
  $dlg.Multiselect = $true
  if ($dlg.ShowDialog() -eq "OK") { return @($dlg.FileNames) }
  return @()
}

function Update-ReceiptDisplay([System.Windows.Forms.TextBox]$textBox) {
  $parts = @()
  if ($script:receiptFolder) { $parts += "폴더: $script:receiptFolder" }
  if ($script:receiptFiles.Count -gt 0) { $parts += "파일 $($script:receiptFiles.Count)개" }
  $textBox.Text = if ($parts.Count -gt 0) { $parts -join " | " } else { "" }
}

function Get-ReceiptImagePaths {
  $paths = New-Object System.Collections.Generic.List[string]
  if ($script:receiptFolder -and (Test-Path -LiteralPath $script:receiptFolder)) {
    Get-ChildItem -LiteralPath $script:receiptFolder -File |
      Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png|webp|heic)$' } |
      ForEach-Object { $paths.Add($_.FullName) | Out-Null }
  }
  foreach ($f in $script:receiptFiles) {
    if ($f -and (Test-Path -LiteralPath $f)) { $paths.Add([string]$f) | Out-Null }
  }
  return @(
    $paths |
      Select-Object -Unique |
      Sort-Object { [System.IO.Path]::GetFileName($_) }
  )
}

function Load-Transactions([string]$path) {
  $jsonPath = Join-Path $env:TEMP ("expense-tx-" + [guid]::NewGuid().ToString() + ".json")
  $readScript = Join-Path $ScriptDir "scripts\read-card-statement.ps1"
  $out = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $readScript -StatementPath $path -OutputJsonPath $jsonPath 2>&1
  if ($LASTEXITCODE -ne 0) { throw "명세서 읽기 실패: $out" }
  $data = Get-Content $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
  Remove-Item $jsonPath -Force -ErrorAction SilentlyContinue
  $list = @($data.transactions)
  foreach ($tx in $list) {
    Initialize-TransactionDefaults -Transaction $tx -DefaultAccount $tbAccount.Text -DefaultDetail $tbDetail.Text -DefaultMealLimit (Get-MealLimitPerPerson)
  }
  return $list
}

function Load-TransitTransactions([string]$path) {
  $jsonPath = Join-Path $env:TEMP ("expense-transit-" + [guid]::NewGuid().ToString() + ".json")
  $readScript = Join-Path $ScriptDir "scripts\read-transit-statement.ps1"
  $out = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $readScript -StatementPath $path -OutputJsonPath $jsonPath 2>&1
  if ($LASTEXITCODE -ne 0) { throw "교통비 명세서 읽기 실패: $out" }
  $data = Get-Content $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
  Remove-Item $jsonPath -Force -ErrorAction SilentlyContinue
  return @($data.transactions)
}

function Get-MealLimitPerPerson {
  $limit = As-Number $settings.mealLimitPerPerson
  if ($limit -le 0) { return 12000.0 }
  return $limit
}

function Get-DefaultRuleContext {
  return @{
    account = $tbAccount.Text
    detail = $tbDetail.Text
    mealLimit = Get-MealLimitPerPerson
  }
}

function Apply-RulesToTransaction {
  param($Transaction)
  $ctx = Get-DefaultRuleContext
  Initialize-TransactionDefaults -Transaction $Transaction -DefaultAccount $ctx.account -DefaultDetail $ctx.detail -DefaultMealLimit $ctx.mealLimit
  $applyMeal = Test-MealLimitApplies -Detail $Transaction.detail -Rule $Transaction.rule
  if (-not $Transaction.headCountManuallySet) {
    $Transaction.headCount = Get-HeadCountFromCompanions -Companions (As-Text $Transaction.companions) -ApplyMealLimit:$applyMeal
  }
  $hc = [int][Math]::Max(1, (As-Number $Transaction.headCount))
  $Transaction.headCount = $hc
  $amounts = Compute-TransactionAmounts `
    -Transaction $Transaction `
    -DefaultAccount $ctx.account `
    -DefaultDetail $ctx.detail `
    -DefaultMealLimit $ctx.mealLimit `
    -HeadCount $hc
  $Transaction.personalUseAmount = $amounts.personal
}

function Apply-RulesToAllTransactions {
  foreach ($tx in $script:transactions) { Apply-RulesToTransaction $tx }
}

function Recalc-RowAmounts {
  param([int]$RowIndex, [switch]$HeadCountManual)
  if ($RowIndex -lt 0 -or $RowIndex -ge $script:transactions.Count) { return }
  $tx = $script:transactions[$RowIndex]
  $ctx = Get-DefaultRuleContext

  Set-TxProperty $tx 'account' (As-Text $grid.Rows[$RowIndex].Cells["account"].Value)
  Set-TxProperty $tx 'detail' (As-Text $grid.Rows[$RowIndex].Cells["detail"].Value)
  $tx.companions = As-Text $grid.Rows[$RowIndex].Cells["companions"].Value
  Update-TransactionRuleFromType -Transaction $tx

  $applyMeal = Test-MealLimitApplies -Detail $tx.detail -Rule $tx.rule
  if (-not $applyMeal) { $tx.headCountManuallySet = $false }
  if ($HeadCountManual) {
    $tx.headCountManuallySet = $true
    $hc = [int][Math]::Max(1, (As-Number $grid.Rows[$RowIndex].Cells["headCount"].Value))
  }
  elseif (-not $tx.headCountManuallySet) {
    $hc = Get-HeadCountFromCompanions -Companions $tx.companions -ApplyMealLimit:$applyMeal
    $tx.headCount = $hc
    $grid.Rows[$RowIndex].Cells["headCount"].Value = $hc
  }
  else {
    $hc = [int][Math]::Max(1, (As-Number $grid.Rows[$RowIndex].Cells["headCount"].Value))
  }
  $tx.headCount = $hc

  $amounts = Compute-TransactionAmounts `
    -Transaction $tx `
    -DefaultAccount $ctx.account `
    -DefaultDetail $ctx.detail `
    -DefaultMealLimit $ctx.mealLimit `
    -HeadCount $hc
  $tx.personalUseAmount = $amounts.personal
  $grid.Rows[$RowIndex].Cells["personal"].Value = $amounts.personal
  $grid.Rows[$RowIndex].Cells["receiptRequired"].Value = if (Test-ReceiptRequired -Merchant $tx.merchant) { "필요" } else { "" }
}

function Recalc-AllFromGrid {
  for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
    if ($i -ge $script:transactions.Count) { break }
    Recalc-RowAmounts $i
  }
}

function Refresh-Grid {
  Apply-RulesToAllTransactions
  $grid.Rows.Clear()
  foreach ($tx in $script:transactions) {
    $hc = if ($tx.headCount) { [int]$tx.headCount } else { 1 }
    $comp = if ($tx.companions) { As-Text $tx.companions } else { "" }
    $claim = Get-TransactionClaimAmount $tx
    $acct = Get-TransactionAccount -Transaction $tx -DefaultAccount $tbAccount.Text
    $det = Get-TransactionDetail -Transaction $tx -DefaultDetail $tbDetail.Text
    $receiptFlag = if (Test-ReceiptRequired -Merchant $tx.merchant) { "필요" } else { "" }
    [void]$grid.Rows.Add($acct, $det, $tx.useDate, $tx.merchant, $claim, $hc, $comp, [double]$tx.personalUseAmount, $receiptFlag)
  }
}

function Apply-PresetToSelectedRows {
  param($Preset)
  if (-not $Preset) { return }
  $rows = @($grid.SelectedRows | Sort-Object { $_.Index })
  if ($rows.Count -eq 0 -and $grid.CurrentRow) { $rows = @($grid.CurrentRow) }
  foreach ($row in $rows) {
    $idx = $row.Index
    if ($idx -lt 0 -or $idx -ge $script:transactions.Count) { continue }
    Apply-PresetToTransaction -Transaction $script:transactions[$idx] -Preset $Preset
    $grid.Rows[$idx].Cells["account"].Value = As-Text $Preset.account
    $grid.Rows[$idx].Cells["detail"].Value = As-Text $Preset.detail
    $script:transactions[$idx].headCountManuallySet = $false
    Recalc-RowAmounts $idx
  }
}

function Apply-DefaultsToSelectedRows {
  $rows = @($grid.SelectedRows | Sort-Object { $_.Index })
  if ($rows.Count -eq 0 -and $grid.CurrentRow) { $rows = @($grid.CurrentRow) }
  foreach ($row in $rows) {
    $idx = $row.Index
    if ($idx -lt 0 -or $idx -ge $script:transactions.Count) { continue }
    Set-TxProperty $script:transactions[$idx] 'account' $tbAccount.Text
    Set-TxProperty $script:transactions[$idx] 'detail' $tbDetail.Text
    $grid.Rows[$idx].Cells["account"].Value = $tbAccount.Text
    $grid.Rows[$idx].Cells["detail"].Value = $tbDetail.Text
    $script:transactions[$idx].headCountManuallySet = $false
    Update-TransactionRuleFromType -Transaction $script:transactions[$idx]
    Recalc-RowAmounts $idx
  }
}

function Sync-GridToTransactions {
  Recalc-AllFromGrid
  for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
    if ($i -ge $script:transactions.Count) { break }
    $script:transactions[$i].account = As-Text $grid.Rows[$i].Cells["account"].Value
    $script:transactions[$i].detail = As-Text $grid.Rows[$i].Cells["detail"].Value
    $script:transactions[$i].headCount = [int][Math]::Max(1, (As-Number $grid.Rows[$i].Cells["headCount"].Value))
    $script:transactions[$i].companions = As-Text $grid.Rows[$i].Cells["companions"].Value
    $script:transactions[$i].personalUseAmount = [double]$grid.Rows[$i].Cells["personal"].Value
  }
}

function Update-ReceiptWarning {
  $required = Get-ReceiptRequiredCount -Transactions $script:transactions
  $images = @(Get-ReceiptImagePaths)
  if ($required -eq 0) { return "" }
  if ($images.Count -lt $required) {
    return "증빙 필요 $required건 · 영수증 $($images.Count)개 (부족할 수 있음)"
  }
  return "증빙 필요 $required건 · 영수증 $($images.Count)개"
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "경비신청서 만들기"
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Malgun Gothic", 10)
$form.AutoScroll = $true

$y = 15
function Add-Lbl([string]$t,[int]$py) { $l=New-Object System.Windows.Forms.Label; $l.Text=$t; $l.Location=New-Object System.Drawing.Point(15,$py); $l.AutoSize=$true; $form.Controls.Add($l) }

Add-Lbl "카드 사용내역 (.xls)" $y
$tbStatement = New-Object System.Windows.Forms.TextBox
$tbStatement.Location = New-Object System.Drawing.Point(15, ($y+22)); $tbStatement.Size = New-Object System.Drawing.Size(650, 24); $tbStatement.ReadOnly = $true
$form.Controls.Add($tbStatement)
$btnStmt = New-Object System.Windows.Forms.Button; $btnStmt.Text = "찾기"; $btnStmt.Location = New-Object System.Drawing.Point(675, ($y+20)); $btnStmt.Size = New-Object System.Drawing.Size(80, 28); $form.Controls.Add($btnStmt)
$y += 58

Add-Lbl "회사 양식 (.xls)" $y
$tbTemplate = New-Object System.Windows.Forms.TextBox
$tbTemplate.Location = New-Object System.Drawing.Point(15, ($y+22)); $tbTemplate.Size = New-Object System.Drawing.Size(650, 24); $tbTemplate.ReadOnly = $true
$form.Controls.Add($tbTemplate)
$btnTpl = New-Object System.Windows.Forms.Button; $btnTpl.Text = "찾기"; $btnTpl.Location = New-Object System.Drawing.Point(675, ($y+20)); $btnTpl.Size = New-Object System.Drawing.Size(80, 28); $form.Controls.Add($btnTpl)
$y += 58

Add-Lbl "후불교통비 내역 (.xls, 선택 — 택시/고속버스 제외)" $y
$tbTransit = New-Object System.Windows.Forms.TextBox
$tbTransit.Location = New-Object System.Drawing.Point(15, ($y+22)); $tbTransit.Size = New-Object System.Drawing.Size(650, 24); $tbTransit.ReadOnly = $true
$form.Controls.Add($tbTransit)
$btnTransit = New-Object System.Windows.Forms.Button; $btnTransit.Text = "찾기"; $btnTransit.Location = New-Object System.Drawing.Point(675, ($y+20)); $btnTransit.Size = New-Object System.Drawing.Size(80, 28); $form.Controls.Add($btnTransit)
$y += 58

Add-Lbl "영수증 (선택, 파일명 순 — 폴더 또는 여러 파일)" $y
$tbReceipts = New-Object System.Windows.Forms.TextBox
$tbReceipts.Location = New-Object System.Drawing.Point(15, ($y+22)); $tbReceipts.Size = New-Object System.Drawing.Size(555, 24); $tbReceipts.ReadOnly = $true
$form.Controls.Add($tbReceipts)
$btnRcptFolder = New-Object System.Windows.Forms.Button; $btnRcptFolder.Text = "폴더"; $btnRcptFolder.Location = New-Object System.Drawing.Point(580, ($y+20)); $btnRcptFolder.Size = New-Object System.Drawing.Size(80, 28); $form.Controls.Add($btnRcptFolder)
$btnRcptFiles = New-Object System.Windows.Forms.Button; $btnRcptFiles.Text = "파일"; $btnRcptFiles.Location = New-Object System.Drawing.Point(670, ($y+20)); $btnRcptFiles.Size = New-Object System.Drawing.Size(80, 28); $form.Controls.Add($btnRcptFiles)
$y += 58

Add-Lbl "기본값 (신규 행 · 선택 행 일괄 적용): 계정 | 거래처 | 사용자 | 업무상세" $y
$tbAccount = New-Object System.Windows.Forms.TextBox; $tbAccount.Text = $settings.account; $tbAccount.Location = New-Object System.Drawing.Point(15, ($y+22)); $tbAccount.Width = 100
$tbClient = New-Object System.Windows.Forms.TextBox; $tbClient.Text = $settings.client; $tbClient.Location = New-Object System.Drawing.Point(125, ($y+22)); $tbClient.Width = 100
$tbUser = New-Object System.Windows.Forms.TextBox; $tbUser.Text = $settings.userName; $tbUser.Location = New-Object System.Drawing.Point(235, ($y+22)); $tbUser.Width = 100
$tbDetail = New-Object System.Windows.Forms.TextBox; $tbDetail.Text = $settings.detail; $tbDetail.Location = New-Object System.Drawing.Point(345, ($y+22)); $tbDetail.Width = 120
$form.Controls.AddRange(@($tbAccount,$tbClient,$tbUser,$tbDetail))
$y += 58

Add-Lbl "저장 파일명: 부서 | 성명 (성명 = 사용자)" $y
$tbDepartment = New-Object System.Windows.Forms.TextBox; $tbDepartment.Text = $settings.department; $tbDepartment.Location = New-Object System.Drawing.Point(15, ($y+22)); $tbDepartment.Width = 200
$lblFileUser = New-Object System.Windows.Forms.Label; $lblFileUser.Location = New-Object System.Drawing.Point(225, ($y+24)); $lblFileUser.AutoSize = $true; $lblFileUser.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.AddRange(@($tbDepartment, $lblFileUser))
$y += 58

$tbUser.Add_TextChanged({
  $lblFileUser.Text = if ($tbUser.Text) { $tbUser.Text } else { "(사용자 입력)" }
})
$lblFileUser.Text = if ($tbUser.Text) { $tbUser.Text } else { "(사용자 입력)" }

Add-Lbl "경비 유형 프리셋 (행 선택 후 적용)" $y
$cboPreset = New-Object System.Windows.Forms.ComboBox
$cboPreset.Location = New-Object System.Drawing.Point(15, ($y + 22))
$cboPreset.Size = New-Object System.Drawing.Size(220, 28)
$cboPreset.DropDownStyle = "DropDownList"
foreach ($p in $script:presetsConfig.presets) {
  [void]$cboPreset.Items.Add((As-Text $p.label))
}
if ($cboPreset.Items.Count -gt 0) { $cboPreset.SelectedIndex = 0 }
$form.Controls.Add($cboPreset)
$btnApplyPreset = New-Object System.Windows.Forms.Button
$btnApplyPreset.Text = "선택 행에 적용"
$btnApplyPreset.Location = New-Object System.Drawing.Point(245, ($y + 20))
$btnApplyPreset.Size = New-Object System.Drawing.Size(120, 28)
$form.Controls.Add($btnApplyPreset)
$btnApplyDefaults = New-Object System.Windows.Forms.Button
$btnApplyDefaults.Text = "기본값 일괄 적용"
$btnApplyDefaults.Location = New-Object System.Drawing.Point(375, ($y + 20))
$btnApplyDefaults.Size = New-Object System.Drawing.Size(130, 28)
$form.Controls.Add($btnApplyDefaults)
$y += 58

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(15, $y)
$grid.Size = New-Object System.Drawing.Size(930, 280)
$grid.AllowUserToAddRows = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $true
$grid.AutoSizeColumnsMode = "Fill"
[void]$grid.Columns.Add("account", "계정")
[void]$grid.Columns.Add("detail", "업무상세")
[void]$grid.Columns.Add("useDate", "이용일")
[void]$grid.Columns.Add("merchant", "가맹점")
[void]$grid.Columns.Add("claim", "청구금액")
[void]$grid.Columns.Add("headCount", "인원")
[void]$grid.Columns.Add("companions", "동행자")
[void]$grid.Columns.Add("personal", "개인사용")
[void]$grid.Columns.Add("receiptRequired", "증빙")
$grid.Columns["useDate"].ReadOnly = $true
$grid.Columns["merchant"].ReadOnly = $true
$grid.Columns["claim"].ReadOnly = $true
$grid.Columns["personal"].ReadOnly = $true
$grid.Columns["receiptRequired"].ReadOnly = $true
$grid.Columns["headCount"].FillWeight = 35
$grid.Columns["companions"].FillWeight = 70
$grid.Columns["account"].FillWeight = 70
$grid.Columns["detail"].FillWeight = 70
$grid.Columns["receiptRequired"].FillWeight = 35
$form.Controls.Add($grid)
$y += 295

$grid.Add_CellEndEdit({
  param($sender, $e)
  if ($e.ColumnIndex -lt 0) { return }
  $colName = $grid.Columns[$e.ColumnIndex].Name
  if ($colName -eq "headCount") {
    Recalc-RowAmounts $e.RowIndex -HeadCountManual
  }
  elseif ($colName -in @("companions", "account", "detail")) {
    if ($colName -eq "companions" -and $e.RowIndex -lt $script:transactions.Count) {
      $script:transactions[$e.RowIndex].headCountManuallySet = $false
    }
    Recalc-RowAmounts $e.RowIndex
  }
})

$grid.Add_CellFormatting({
  param($sender, $e)
  if ($e.ColumnIndex -lt 0 -or $e.RowIndex -lt 0) { return }
  if ($grid.Columns[$e.ColumnIndex].Name -eq "receiptRequired") {
    if ((As-Text $e.Value) -eq "필요") {
      $e.CellStyle.ForeColor = [System.Drawing.Color]::DarkRed
      $e.CellStyle.Font = New-Object System.Drawing.Font($grid.Font, [System.Drawing.FontStyle]::Bold)
    }
  }
})

$btnApplyPreset.Add_Click({
  if ($cboPreset.SelectedIndex -lt 0) { return }
  $preset = $script:presetsConfig.presets[$cboPreset.SelectedIndex]
  Apply-PresetToSelectedRows -Preset $preset
})

$btnApplyDefaults.Add_Click({ Apply-DefaultsToSelectedRows })


$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(15, $y)
$lblStatus.Size = New-Object System.Drawing.Size(930, 22)
$form.Controls.Add($lblStatus)
$y += 30

$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = "완성 엑셀 만들기"
$btnGenerate.Location = New-Object System.Drawing.Point(15, $y)
$btnGenerate.Size = New-Object System.Drawing.Size(200, 36)
$form.Controls.Add($btnGenerate)

$script:summaryValueLabels = @{}
$script:summaryPanel = New-Object System.Windows.Forms.Panel
$script:summaryPanel.Location = New-Object System.Drawing.Point(230, $y)
$script:summaryPanel.BorderStyle = "FixedSingle"
$script:summaryPanel.BackColor = [System.Drawing.Color]::White
$script:summaryPanel.Visible = $false
$summaryRows = @(
  @{ key = "total"; label = "합계" },
  @{ key = "expense"; label = "경비신청금액" },
  @{ key = "personal"; label = "개인사용금액" }
)
$rowH = 26
$labelW = 120
$valueW = 158
$script:summaryPanel.Size = New-Object System.Drawing.Size(($labelW + $valueW), ($rowH * 3))
foreach ($i in 0..2) {
  $ry = $i * $rowH
  $lblSum = New-Object System.Windows.Forms.Label
  $lblSum.Text = $summaryRows[$i].label
  $lblSum.Location = New-Object System.Drawing.Point(0, $ry)
  $lblSum.Size = New-Object System.Drawing.Size($labelW, $rowH)
  $lblSum.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
  $lblSum.BackColor = [System.Drawing.Color]::FromArgb(242, 242, 242)
  $lblSum.BorderStyle = "FixedSingle"
  $valSum = New-Object System.Windows.Forms.TextBox
  $valSum.Text = ""
  $valSum.Location = New-Object System.Drawing.Point($labelW, $ry)
  $valSum.Size = New-Object System.Drawing.Size($valueW, $rowH)
  $valSum.ReadOnly = $true
  $valSum.BorderStyle = "FixedSingle"
  $valSum.BackColor = [System.Drawing.Color]::White
  $valSum.TextAlign = "Right"
  $valSum.TabStop = $false
  $script:summaryValueLabels[$summaryRows[$i].key] = $valSum
  $script:summaryPanel.Controls.AddRange(@($lblSum, $valSum))
}
$form.Controls.Add($script:summaryPanel)
$script:summaryPanel.BringToFront()

$y += 90
$form.ClientSize = New-Object System.Drawing.Size(960, [Math]::Max(740, $y))

function Update-SummaryPanel {
  $summary = Get-TransactionSummary -Transactions $script:transactions -TransitTransactions $script:transitTransactions
  $script:summaryValueLabels["total"].Text = Format-Amount $summary.total
  $script:summaryValueLabels["expense"].Text = Format-Amount $summary.expense
  $script:summaryValueLabels["personal"].Text = Format-Amount $summary.personal
  $script:summaryPanel.Visible = $true
  $script:summaryPanel.BringToFront()
  $form.Refresh()
}

function Clear-SummaryPanel {
  foreach ($key in @("total", "expense", "personal")) {
    $script:summaryValueLabels[$key].Text = "-"
  }
  $script:summaryPanel.Visible = $true
}
Clear-SummaryPanel

$btnStmt.Add_Click({
  $p = Pick-File "카드 사용내역" "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
  if (-not $p) { return }
  try {
    $script:statementPath = $p
    $tbStatement.Text = $p
    $script:transactions = Load-Transactions $p
    Refresh-Grid
    $warn = Update-ReceiptWarning
    $lblStatus.Text = "$($script:transactions.Count)건 불러옴" + $(if ($warn) { " · $warn" } else { "" })
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
  } catch {
    $lblStatus.Text = $_.Exception.Message
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkRed
  }
})

$btnTpl.Add_Click({
  $p = Pick-File "회사 양식" "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
  if ($p) { $script:templatePath = $p; $tbTemplate.Text = $p }
})

$btnTransit.Add_Click({
  $p = Pick-File "후불교통비 내역" "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
  if (-not $p) { return }
  try {
    $script:transitPath = $p
    $tbTransit.Text = $p
    $script:transitTransactions = Load-TransitTransactions $p
    $lblStatus.Text = "교통비 $($script:transitTransactions.Count)건 불러옴 (택시/고속버스 제외)"
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
  } catch {
    $lblStatus.Text = $_.Exception.Message
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkRed
  }
})

$btnRcptFolder.Add_Click({
  $p = Pick-Folder "영수증 폴더"
  if ($p) {
    $script:receiptFolder = $p
    Update-ReceiptDisplay $tbReceipts
    $warn = Update-ReceiptWarning
    if ($warn -and $script:transactions.Count -gt 0) {
      $lblStatus.Text = $warn
      $lblStatus.ForeColor = [System.Drawing.Color]::DarkOrange
    }
  }
})

$btnRcptFiles.Add_Click({
  $files = Pick-ImageFiles "영수증 이미지 (여러 개 선택 가능)"
  if ($files.Count -gt 0) {
    $script:receiptFiles = @($files)
    Update-ReceiptDisplay $tbReceipts
    $warn = Update-ReceiptWarning
    if ($warn -and $script:transactions.Count -gt 0) {
      $lblStatus.Text = $warn
      $lblStatus.ForeColor = [System.Drawing.Color]::DarkOrange
    }
  }
})

$btnGenerate.Add_Click({
  $lblStatus.Text = ""
  Clear-SummaryPanel
  if (-not $script:statementPath -or -not $script:templatePath) {
    $lblStatus.Text = "명세서와 회사 양식을 선택하세요."
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkRed
    return
  }
  if ($script:transactions.Count -eq 0) {
    try { $script:transactions = Load-Transactions $script:statementPath; Refresh-Grid } catch {
      $lblStatus.Text = $_.Exception.Message; return
    }
  }

  Sync-GridToTransactions

  $receiptWarn = Update-ReceiptWarning
  if ($receiptWarn) {
    $required = Get-ReceiptRequiredCount -Transactions $script:transactions
    $images = @(Get-ReceiptImagePaths)
    if ($images.Count -lt $required) {
      $ans = [System.Windows.Forms.MessageBox]::Show(
        "$receiptWarn`n`n영수증이 부족할 수 있습니다. 그래도 계속하시겠습니까?",
        "증빙 확인",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
      )
      if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }
  }

  $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
  $saveDlg.Filter = "Excel 97-2003 (*.xls)|*.xls"
  $saveDlg.InitialDirectory = Split-Path -Parent $script:templatePath
  $saveDlg.FileName = Get-OutputFileName `
    -TemplatePath $script:templatePath `
    -Department $tbDepartment.Text `
    -UserName $tbUser.Text
  if ($saveDlg.ShowDialog() -ne "OK") { return }

  $imagePaths = @(Get-ReceiptImagePaths)

  Save-Settings @{
    account            = $tbAccount.Text
    client             = $tbClient.Text
    department         = $tbDepartment.Text
    userName           = $tbUser.Text
    detail             = $tbDetail.Text
    mealLimitPerPerson = Get-MealLimitPerPerson
  }

  $job = @{
    templatePath = $script:templatePath
    outputPath = $saveDlg.FileName
    transactions = $script:transactions
    transitTransactions = $script:transitTransactions
    defaults = @{
      account = $tbAccount.Text
      client = $tbClient.Text
      userName = $tbUser.Text
      detail = $tbDetail.Text
      mealLimitPerPerson = Get-MealLimitPerPerson
    }
    imagePaths = $imagePaths
  }
  $jobPath = Join-Path $env:TEMP ("expense-job-" + [guid]::NewGuid().ToString() + ".json")
  $job | ConvertTo-Json -Depth 8 | Set-Content $jobPath -Encoding UTF8

  $fillScript = Join-Path $ScriptDir "scripts\fill-expense-form.ps1"
  $btnGenerate.Enabled = $false
  $lblStatus.Text = "작성 중... (Excel)"
  $form.Refresh()

  try {
    $fillOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $fillScript -ConfigPath $jobPath 2>&1
    if ($LASTEXITCODE -ne 0) {
      $msg = ($fillOut | Out-String).Trim()
      if ($msg -eq "") { $msg = "엑셀 작성 실패. 열려 있는 Excel을 모두 닫고 다시 시도하세요." }
      throw $msg
    }
    $lblStatus.Text = "저장됨: $($saveDlg.FileName)"
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
    Update-SummaryPanel
    [System.Windows.Forms.MessageBox]::Show("완료`n$($saveDlg.FileName)", "완료") | Out-Null
  } catch {
    $lblStatus.Text = $_.Exception.Message
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkRed
  } finally {
    $btnGenerate.Enabled = $true
    Remove-Item $jobPath -Force -ErrorAction SilentlyContinue
  }
})

[void]$form.ShowDialog()