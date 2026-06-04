# 경비신청서 자동 작성 (PowerShell + Excel COM)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
chcp 65001 | Out-Null

$ScriptDir = $PSScriptRoot
. (Join-Path $ScriptDir "scripts\ExpenseCommon.ps1")

$settings = Load-Settings
$script:statementPath = ""
$script:templatePath = ""
$script:receiptFolder = ""
$script:receiptFiles = @()
$script:transactions = @()

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
  return @($data.transactions)
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "경비신청서 만들기"
$form.Size = New-Object System.Drawing.Size(980, 760)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Malgun Gothic", 10)

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

Add-Lbl "영수증 (선택, 파일명 순 — 폴더 또는 여러 파일)" $y
$tbReceipts = New-Object System.Windows.Forms.TextBox
$tbReceipts.Location = New-Object System.Drawing.Point(15, ($y+22)); $tbReceipts.Size = New-Object System.Drawing.Size(555, 24); $tbReceipts.ReadOnly = $true
$form.Controls.Add($tbReceipts)
$btnRcptFolder = New-Object System.Windows.Forms.Button; $btnRcptFolder.Text = "폴더"; $btnRcptFolder.Location = New-Object System.Drawing.Point(580, ($y+20)); $btnRcptFolder.Size = New-Object System.Drawing.Size(80, 28); $form.Controls.Add($btnRcptFolder)
$btnRcptFiles = New-Object System.Windows.Forms.Button; $btnRcptFiles.Text = "파일"; $btnRcptFiles.Location = New-Object System.Drawing.Point(670, ($y+20)); $btnRcptFiles.Size = New-Object System.Drawing.Size(80, 28); $form.Controls.Add($btnRcptFiles)
$y += 58

Add-Lbl "엑셀 기본값: 계정 | 거래처 | 사용자 | 업무상세" $y
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

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(15, $y)
$grid.Size = New-Object System.Drawing.Size(930, 280)
$grid.AllowUserToAddRows = $false
$grid.AutoSizeColumnsMode = "Fill"
[void]$grid.Columns.Add("useDate", "이용일")
[void]$grid.Columns.Add("merchant", "가맹점")
[void]$grid.Columns.Add("claim", "청구금액")
[void]$grid.Columns.Add("headCount", "인원")
[void]$grid.Columns.Add("companions", "동행자")
[void]$grid.Columns.Add("personal", "개인사용")
$grid.Columns["useDate"].ReadOnly = $true
$grid.Columns["merchant"].ReadOnly = $true
$grid.Columns["claim"].ReadOnly = $true
$grid.Columns["headCount"].FillWeight = 40
$grid.Columns["companions"].FillWeight = 80
$form.Controls.Add($grid)
$y += 295

function Get-MealLimitPerPerson {
  $limit = As-Number $settings.mealLimitPerPerson
  if ($limit -le 0) { return 12000.0 }
  return $limit
}

function Apply-MealToAllTransactions {
  $apply = Test-MealLimitApplies $tbDetail.Text
  $limit = Get-MealLimitPerPerson
  foreach ($tx in $script:transactions) {
    if (-not $tx.PSObject.Properties['headCount']) {
      $tx | Add-Member -NotePropertyName headCount -NotePropertyValue 1 -Force
    }
    if (-not $tx.PSObject.Properties['companions']) {
      $tx | Add-Member -NotePropertyName companions -NotePropertyValue "" -Force
    }
    if (-not $tx.PSObject.Properties['headCountManuallySet']) {
      $tx | Add-Member -NotePropertyName headCountManuallySet -NotePropertyValue $false -Force
    }
    if (-not $tx.headCountManuallySet) {
      $tx.headCount = Get-HeadCountFromCompanions `
        -Companions (As-Text $tx.companions) `
        -ApplyMealLimit:$apply
    }
    $hc = [int][Math]::Max(1, (As-Number $tx.headCount))
    $tx.headCount = $hc
    $claim = Get-TransactionClaimAmount $tx
    $amounts = Compute-MealAmounts -ClaimAmount $claim -HeadCount $hc -LimitPerPerson $limit -ApplyMealLimit:$apply
    $tx.personalUseAmount = $amounts.personal
  }
}

function Recalc-RowMealPersonal([int]$rowIndex, [switch]$HeadCountManual) {
  if ($rowIndex -lt 0 -or $rowIndex -ge $script:transactions.Count) { return }
  $tx = $script:transactions[$rowIndex]
  $apply = Test-MealLimitApplies $tbDetail.Text
  $limit = Get-MealLimitPerPerson
  $claim = Get-TransactionClaimAmount $tx
  $tx.companions = As-Text $grid.Rows[$rowIndex].Cells["companions"].Value
  if ($HeadCountManual) {
    $tx.headCountManuallySet = $true
    $hc = [int][Math]::Max(1, (As-Number $grid.Rows[$rowIndex].Cells["headCount"].Value))
  }
  elseif (-not $tx.headCountManuallySet) {
    $hc = Get-HeadCountFromCompanions -Companions $tx.companions -ApplyMealLimit:$apply
    $tx.headCount = $hc
    $grid.Rows[$rowIndex].Cells["headCount"].Value = $hc
  }
  else {
    $hc = [int][Math]::Max(1, (As-Number $grid.Rows[$rowIndex].Cells["headCount"].Value))
  }
  $tx.headCount = $hc
  $amounts = Compute-MealAmounts -ClaimAmount $claim -HeadCount $hc -LimitPerPerson $limit -ApplyMealLimit:$apply
  $tx.personalUseAmount = $amounts.personal
  $grid.Rows[$rowIndex].Cells["personal"].Value = $amounts.personal
}

function Recalc-AllMealFromGrid {
  for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
    if ($i -ge $script:transactions.Count) { break }
    $tx = $script:transactions[$i]
    $tx.companions = As-Text $grid.Rows[$i].Cells["companions"].Value
    if ($tx.headCountManuallySet) {
      $tx.headCount = [int][Math]::Max(1, (As-Number $grid.Rows[$i].Cells["headCount"].Value))
    }
    else {
      $apply = Test-MealLimitApplies $tbDetail.Text
      $tx.headCount = Get-HeadCountFromCompanions -Companions $tx.companions -ApplyMealLimit:$apply
      $grid.Rows[$i].Cells["headCount"].Value = $tx.headCount
    }
    $apply = Test-MealLimitApplies $tbDetail.Text
    $limit = Get-MealLimitPerPerson
    $claim = Get-TransactionClaimAmount $tx
    $amounts = Compute-MealAmounts -ClaimAmount $claim -HeadCount $tx.headCount -LimitPerPerson $limit -ApplyMealLimit:$apply
    $tx.personalUseAmount = $amounts.personal
    $grid.Rows[$i].Cells["personal"].Value = $amounts.personal
  }
}

function Refresh-Grid {
  Apply-MealToAllTransactions
  $grid.Rows.Clear()
  foreach ($tx in $script:transactions) {
    $hc = if ($tx.headCount) { [int]$tx.headCount } else { 1 }
    $comp = if ($tx.companions) { As-Text $tx.companions } else { "" }
    $claim = Get-TransactionClaimAmount $tx
    [void]$grid.Rows.Add($tx.useDate, $tx.merchant, $claim, $hc, $comp, [double]$tx.personalUseAmount)
  }
}

$grid.Add_CellEndEdit({
  param($sender, $e)
  if ($e.ColumnIndex -lt 0) { return }
  $colName = $grid.Columns[$e.ColumnIndex].Name
  if ($colName -eq "headCount") {
    Recalc-RowMealPersonal $e.RowIndex -HeadCountManual
  }
  elseif ($colName -eq "companions") {
    if ($e.RowIndex -lt $script:transactions.Count) {
      $script:transactions[$e.RowIndex].headCountManuallySet = $false
    }
    Recalc-RowMealPersonal $e.RowIndex
  }
})

$tbDetail.Add_TextChanged({
  if ($script:transactions.Count -gt 0) { Refresh-Grid }
})

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(15, $y)
$lblStatus.Size = New-Object System.Drawing.Size(850, 50)
$form.Controls.Add($lblStatus)
$y += 55

$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = "완성 엑셀 만들기"
$btnGenerate.Location = New-Object System.Drawing.Point(15, $y)
$btnGenerate.Size = New-Object System.Drawing.Size(200, 36)
$form.Controls.Add($btnGenerate)

function Sync-GridToTransactions {
  Recalc-AllMealFromGrid
  for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
    if ($i -ge $script:transactions.Count) { break }
    $script:transactions[$i].headCount = [int][Math]::Max(1, (As-Number $grid.Rows[$i].Cells["headCount"].Value))
    $script:transactions[$i].companions = As-Text $grid.Rows[$i].Cells["companions"].Value
    $script:transactions[$i].personalUseAmount = [double]$grid.Rows[$i].Cells["personal"].Value
  }
}

$btnStmt.Add_Click({
  $p = Pick-File "카드 사용내역" "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
  if (-not $p) { return }
  try {
    $script:statementPath = $p
    $tbStatement.Text = $p
    $script:transactions = Load-Transactions $p
    Refresh-Grid
    $lblStatus.Text = "$($script:transactions.Count)건 불러옴"
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

$btnRcptFolder.Add_Click({
  $p = Pick-Folder "영수증 폴더"
  if ($p) {
    $script:receiptFolder = $p
    Update-ReceiptDisplay $tbReceipts
  }
})

$btnRcptFiles.Add_Click({
  $files = Pick-ImageFiles "영수증 이미지 (여러 개 선택 가능)"
  if ($files.Count -gt 0) {
    $script:receiptFiles = @($files)
    Update-ReceiptDisplay $tbReceipts
  }
})

$btnGenerate.Add_Click({
  $lblStatus.Text = ""
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
    defaults = @{ account=$tbAccount.Text; client=$tbClient.Text; userName=$tbUser.Text; detail=$tbDetail.Text }
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