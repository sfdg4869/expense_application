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

function Test-MealLimitApplies([string]$Detail) {
  return (As-Text $Detail) -eq "야근식대"
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
  param($Transactions)
  $total = 0.0
  $personal = 0.0
  foreach ($tx in @($Transactions)) {
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