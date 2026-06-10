function As-Text([object]$value) {
  if ($null -eq $value) { return "" }
  return ([System.Convert]::ToString($value, [System.Globalization.CultureInfo]::InvariantCulture)).Trim()
}

function As-Number([object]$value) {
  if ($null -eq $value) { return 0.0 }
  if ($value -is [double] -or $value -is [float] -or $value -is [decimal] -or $value -is [int] -or $value -is [long]) {
    return [double]$value
  }
  $s = As-Text $value
  if ($s -eq "") { return 0.0 }
  return [double]($s -replace ',', '')
}

function Get-SettingsPath() {
  $dir = Join-Path $env:APPDATA "ExpenseApp"
  if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  return Join-Path $dir "settings.json"
}

function Get-DefaultSettings() {
  return [ordered]@{
    account            = "복리후생비"
    client             = "엑셈"
    department         = "제품기술연구"
    userName           = "정경수"
    detail             = "야근식대"
    mealLimitPerPerson = 12000
  }
}

function Load-Settings() {
  $path = Get-SettingsPath
  if (-not (Test-Path $path)) { return Get-DefaultSettings }
  $raw = Get-Content $path -Raw -Encoding UTF8 | ConvertFrom-Json
  $d = Get-DefaultSettings
  foreach ($key in @("account", "client", "department", "userName", "detail", "mealLimitPerPerson")) {
    if ($raw.PSObject.Properties[$key]) {
      if ($key -eq "mealLimitPerPerson") { $d[$key] = As-Number $raw.$key }
      else { $d[$key] = As-Text $raw.$key }
    }
  }
  return $d
}

function Save-Settings($settings) {
  $obj = @{
    account            = As-Text $settings.account
    client             = As-Text $settings.client
    department         = As-Text $settings.department
    userName           = As-Text $settings.userName
    detail             = As-Text $settings.detail
    mealLimitPerPerson = As-Number $settings.mealLimitPerPerson
  }
  $obj | ConvertTo-Json | Set-Content -Path (Get-SettingsPath) -Encoding UTF8
}

function Get-PresetsConfigPath() {
  return Join-Path $PSScriptRoot "expense-presets.json"
}

function Load-ExpensePresets() {
  $path = Get-PresetsConfigPath
  if (-not (Test-Path -LiteralPath $path)) {
    return @{
      mealLimitDefault = 12000
      presets = @()
      receiptRequiredKeywords = @()
    }
  }
  $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
  return @{
    mealLimitDefault = if ($raw.mealLimitDefault) { As-Number $raw.mealLimitDefault } else { 12000.0 }
    presets = @($raw.presets)
    receiptRequiredKeywords = @($raw.receiptRequiredKeywords | ForEach-Object { As-Text $_ })
  }
}

function Find-ExpensePreset {
  param(
    [string]$Account,
    [string]$Detail,
    [string]$Label = ""
  )
  $cfg = Load-ExpensePresets
  $label = As-Text $Label
  if ($label -ne "") {
    $byLabel = @($cfg.presets | Where-Object { (As-Text $_.label) -eq $label })
    if ($byLabel.Count -gt 0) { return $byLabel[0] }
  }
  $account = As-Text $Account
  $detail = As-Text $Detail
  $match = @($cfg.presets | Where-Object {
    (As-Text $_.account) -eq $account -and (As-Text $_.detail) -eq $detail
  })
  if ($match.Count -gt 0) { return $match[0] }
  return $null
}

function Get-TransactionRule {
  param(
    $Transaction,
    [string]$DefaultAccount = "",
    [string]$DefaultDetail = "",
    [double]$DefaultMealLimit = 12000
  )
  $account = if ($Transaction -and $Transaction.account) { As-Text $Transaction.account } else { As-Text $DefaultAccount }
  $detail = if ($Transaction -and $Transaction.detail) { As-Text $Transaction.detail } else { As-Text $DefaultDetail }
  $preset = Find-ExpensePreset -Account $account -Detail $detail
  if ($preset) {
    $rule = As-Text $preset.rule
  }
  elseif ($detail -match '식대$') {
    $rule = "meal_per_person"
  }
  else {
    $rule = "full_claim"
  }

  $limit = $DefaultMealLimit
  if ($Transaction -and $Transaction.PSObject.Properties['mealLimit']) {
    $limit = As-Number $Transaction.mealLimit
  }
  elseif ($preset -and $preset.limit) {
    $limit = As-Number $preset.limit
  }
  if ($limit -le 0) { $limit = $DefaultMealLimit }

  return @{ rule = $rule; limit = $limit }
}

function Sanitize-FileNamePart([string]$text) {
  $t = As-Text $text
  foreach ($c in [IO.Path]::GetInvalidFileNameChars()) {
    $t = $t.Replace([string]$c, '')
  }
  if ($t -eq '') { return '이름없음' }
  return $t
}

function Get-OutputFileName {
  param(
    [string]$TemplatePath,
    [string]$Department,
    [string]$UserName
  )
  $base = [IO.Path]::GetFileNameWithoutExtension($TemplatePath)
  $ext = [IO.Path]::GetExtension($TemplatePath)
  $dept = Sanitize-FileNamePart $Department
  $user = Sanitize-FileNamePart $UserName
  $suffix = "_${dept}_${user}"
  $placeholder = '_부서_성명'

  if ($base.EndsWith($placeholder)) {
    $base = $base.Substring(0, $base.Length - $placeholder.Length) + $suffix
  }
  else {
    $base = $base -replace '_[^_]+_[^_]+$', $suffix
  }
  return "${base}${ext}"
}

function Test-MealLimitApplies {
  param(
    [string]$Detail = "",
    [string]$Rule = ""
  )
  if (As-Text $Rule -ne "") {
    return (As-Text $Rule) -eq "meal_per_person"
  }
  $preset = Find-ExpensePreset -Detail $Detail
  if ($preset -and (As-Text $preset.rule) -eq "meal_per_person") { return $true }
  return (As-Text $Detail) -match '식대$'
}

function Get-CompanionCount([string]$Companions) {
  $comp = As-Text $Companions
  if ($comp -eq "") { return 0 }
  $parts = @($comp -split '[,，、]' | ForEach-Object { As-Text $_ } | Where-Object { $_ -ne "" })
  return $parts.Count
}

function Get-TransactionClaimAmount($tx) {
  $cols = @($tx.cardColumns)
  if ($cols.Count -ge 6) {
    $c6 = As-Number $cols[5]
    if ($c6 -gt 0) { return $c6 }
  }
  return As-Number $tx.domesticClaimAmount
}

function Get-HeadCountFromCompanions {
  param(
    [string]$Companions,
    [switch]$ApplyMealLimit
  )
  if (-not $ApplyMealLimit) { return 1 }
  return [int][Math]::Max(1, 1 + (Get-CompanionCount $Companions))
}

function Compute-MealAmounts {
  param(
    [double]$ClaimAmount,
    [int]$HeadCount = 1,
    [double]$LimitPerPerson = 12000,
    [switch]$ApplyMealLimit
  )
  if (-not $ApplyMealLimit) {
    return @{ personal = 0.0; expense = $ClaimAmount }
  }
  $headCount = [Math]::Max(1, $HeadCount)
  $allowance = $LimitPerPerson * $headCount
  $expense = [Math]::Min($ClaimAmount, $allowance)
  $personal = [Math]::Max(0, $ClaimAmount - $allowance)
  return @{ personal = $personal; expense = $expense }
}

function Compute-TransactionAmounts {
  param(
    $Transaction,
    [string]$DefaultAccount = "",
    [string]$DefaultDetail = "",
    [double]$DefaultMealLimit = 12000,
    [int]$HeadCount = 1
  )
  $claim = Get-TransactionClaimAmount $Transaction
  $ruleInfo = Get-TransactionRule `
    -Transaction $Transaction `
    -DefaultAccount $DefaultAccount `
    -DefaultDetail $DefaultDetail `
    -DefaultMealLimit $DefaultMealLimit
  $rule = $ruleInfo.rule
  $limit = $ruleInfo.limit
  $hc = [int][Math]::Max(1, $HeadCount)

  switch ($rule) {
    "meal_per_person" {
      return Compute-MealAmounts -ClaimAmount $claim -HeadCount $hc -LimitPerPerson $limit -ApplyMealLimit
    }
    "full_personal" {
      return @{ personal = $claim; expense = 0.0 }
    }
    default {
      return @{ personal = 0.0; expense = $claim }
    }
  }
}

function Set-TxProperty {
  param(
    $Transaction,
    [string]$Name,
    $Value
  )
  $Transaction | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
}

function Update-TransactionRuleFromType {
  param($Transaction)
  $preset = Find-ExpensePreset `
    -Account (As-Text $Transaction.account) `
    -Detail (As-Text $Transaction.detail)
  if ($preset) {
    Set-TxProperty $Transaction 'rule' (As-Text $preset.rule)
    if ($preset.limit) { Set-TxProperty $Transaction 'mealLimit' (As-Number $preset.limit) }
    return
  }
  $det = As-Text $Transaction.detail
  if ($det -match '식대$') {
    Set-TxProperty $Transaction 'rule' 'meal_per_person'
  }
  else {
    Set-TxProperty $Transaction 'rule' 'full_claim'
  }
}

function Apply-PresetToTransaction {
  param(
    $Transaction,
    $Preset
  )
  if (-not $Preset) { return }
  Set-TxProperty $Transaction 'account' (As-Text $Preset.account)
  Set-TxProperty $Transaction 'detail' (As-Text $Preset.detail)
  Update-TransactionRuleFromType -Transaction $Transaction
}

function Initialize-TransactionDefaults {
  param(
    $Transaction,
    [string]$DefaultAccount,
    [string]$DefaultDetail,
    [double]$DefaultMealLimit = 12000
  )
  if (-not $Transaction.PSObject.Properties['account'] -or (As-Text $Transaction.account) -eq "") {
    $Transaction | Add-Member -NotePropertyName account -NotePropertyValue (As-Text $DefaultAccount) -Force
  }
  if (-not $Transaction.PSObject.Properties['detail'] -or (As-Text $Transaction.detail) -eq "") {
    $Transaction | Add-Member -NotePropertyName detail -NotePropertyValue (As-Text $DefaultDetail) -Force
  }
  if (-not $Transaction.PSObject.Properties['headCount']) {
    $Transaction | Add-Member -NotePropertyName headCount -NotePropertyValue 1 -Force
  }
  if (-not $Transaction.PSObject.Properties['companions']) {
    $Transaction | Add-Member -NotePropertyName companions -NotePropertyValue "" -Force
  }
  if (-not $Transaction.PSObject.Properties['headCountManuallySet']) {
    $Transaction | Add-Member -NotePropertyName headCountManuallySet -NotePropertyValue $false -Force
  }
  if (-not $Transaction.PSObject.Properties['personalUseAmount']) {
    $Transaction | Add-Member -NotePropertyName personalUseAmount -NotePropertyValue 0.0 -Force
  }
  Update-TransactionRuleFromType -Transaction $Transaction
}

function Test-ReceiptRequired {
  param([string]$Merchant)
  $cfg = Load-ExpensePresets
  $name = As-Text $Merchant
  if ($name -eq "") { return $false }
  foreach ($kw in $cfg.receiptRequiredKeywords) {
    if ($kw -ne "" -and $name -like "*$kw*") { return $true }
  }
  return $false
}

function Get-ReceiptRequiredCount {
  param($Transactions)
  $count = 0
  foreach ($tx in @($Transactions)) {
    if (Test-ReceiptRequired -Merchant (As-Text $tx.merchant)) { $count++ }
  }
  return $count
}

function Format-UserCell {
  param(
    [string]$UserName,
    [string]$Companions
  )
  $user = As-Text $UserName
  $comp = As-Text $Companions
  if ($comp -eq "") { return $user }
  return "$user, $comp"
}

function Format-Amount([double]$Amount) {
  return "{0:N0}" -f [Math]::Round($Amount, 0)
}

function Get-TransactionSummary {
  param(
    $Transactions,
    $TransitTransactions = @()
  )
  $total = 0.0
  $personal = 0.0
  foreach ($tx in @($Transactions)) {
    $claim = Get-TransactionClaimAmount $tx
    $p = As-Number $tx.personalUseAmount
    $total += $claim
    $personal += $p
  }
  foreach ($tx in @($TransitTransactions)) {
    $claim = Get-TransactionClaimAmount $tx
    $p = As-Number $tx.personalUseAmount
    $total += $claim
    $personal += $p
  }
  return @{
    total    = $total
    expense  = $total - $personal
    personal = $personal
  }
}

function Get-TransactionAccount {
  param($Transaction, [string]$DefaultAccount)
  $a = if ($Transaction.account) { As-Text $Transaction.account } else { "" }
  if ($a -eq "") { return As-Text $DefaultAccount }
  return $a
}

function Get-TransactionDetail {
  param($Transaction, [string]$DefaultDetail)
  $d = if ($Transaction.detail) { As-Text $Transaction.detail } else { "" }
  if ($d -eq "") { return As-Text $DefaultDetail }
  return $d
}
