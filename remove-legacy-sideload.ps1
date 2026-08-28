param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$addinId = '7c1e9a24-5d3f-4b8e-9c2a-0f6d1b4e8a5c'
$developerKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$excelOptionsKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
$installerStateKey = 'HKCU:\Software\DeepSeekExcelAssistant'

$removed = New-Object System.Collections.Generic.List[string]

if (Test-Path -LiteralPath $developerKey) {
  $developerProperties = Get-ItemProperty -LiteralPath $developerKey
  $developerEntry = $developerProperties.PSObject.Properties[$addinId]
  if ($null -ne $developerEntry) {
    Remove-ItemProperty -LiteralPath $developerKey -Name $addinId
    $removed.Add("WEF Developer entry: $addinId")
  }
}

if (Test-Path -LiteralPath $excelOptionsKey) {
  $options = Get-ItemProperty -LiteralPath $excelOptionsKey
  foreach ($property in $options.PSObject.Properties) {
    if ($property.Name -match '^OPEN\d*$' -and
        [string]$property.Value -match 'DeepSeekExcelAssistant-(?:Standalone-)?Autoload(?:-v[\d.]+)?\.xlsm') {
      Remove-ItemProperty -LiteralPath $excelOptionsKey -Name $property.Name
      $removed.Add("Excel startup value: $($property.Name)")
    }
  }
}

if (Test-Path -LiteralPath $installerStateKey) {
  foreach ($name in @('ExcelStartupValue', 'ExcelStartupArgument')) {
    $state = Get-ItemProperty -LiteralPath $installerStateKey
    if ($null -ne $state.PSObject.Properties[$name]) {
      Remove-ItemProperty -LiteralPath $installerStateKey -Name $name
      $removed.Add("Installer state: $name")
    }
  }
}

if ($removed.Count -eq 0) {
  Write-Host 'No legacy DeepSeek sideload registration was found.'
}
else {
  Write-Host 'Removed legacy DeepSeek sideload registration:'
  foreach ($entry in $removed) {
    Write-Host "  - $entry"
  }
}

Write-Host 'The Microsoft 365 admin deployment and installed files were not removed.'
