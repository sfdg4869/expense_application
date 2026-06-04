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
  return [ordered]@{ account = "복리후생비"; client = "엑셈"; userName = "정경수"; detail = "야근식대" }
}

function Load-Settings() {
  $path = Get-SettingsPath
  if (-not (Test-Path $path)) { return Get-DefaultSettings }
  $raw = Get-Content $path -Raw -Encoding UTF8 | ConvertFrom-Json
  $d = Get-DefaultSettings
  foreach ($key in @("account","client","userName","detail")) {
    if ($raw.PSObject.Properties[$key]) { $d[$key] = As-Text $raw.$key }
  }
  return $d
}

function Save-Settings($settings) {
  $obj = @{ account = As-Text $settings.account; client = As-Text $settings.client; userName = As-Text $settings.userName; detail = As-Text $settings.detail }
  $obj | ConvertTo-Json | Set-Content -Path (Get-SettingsPath) -Encoding UTF8
}