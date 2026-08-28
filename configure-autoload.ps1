param(
  [Parameter(Mandatory = $true)]
  [string]$InstallRoot,
  [string]$CarrierPath,
  [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'

$optionsKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
$trustedLocationKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\DeepSeekExcelAssistant'
$stateKey = 'HKCU:\Software\DeepSeekExcelAssistant'
$startupValueProperty = 'ExcelStartupValue'
$startupArgumentProperty = 'ExcelStartupArgument'

if ([string]::IsNullOrWhiteSpace($CarrierPath)) {
  $CarrierPath = Join-Path $InstallRoot 'carrier\DeepSeekExcelAssistant-Standalone-Autoload.xlsm'
}
$CarrierPath = [System.IO.Path]::GetFullPath($CarrierPath)
$carrierDirectory = [System.IO.Path]::GetDirectoryName($CarrierPath).TrimEnd('\') + '\'
# Force the hidden carrier to open read-only. Excel can then load the same
# carrier in multiple independent /x processes without showing a file-lock
# prompt or skipping the WebExtension activation.
$argumentValue = '/r "' + $CarrierPath + '"'

function Remove-StartupRegistration {
  if (-not (Test-Path -LiteralPath $stateKey)) { return }
  $state = Get-ItemProperty -LiteralPath $stateKey
  $valueName = $state.$startupValueProperty
  $storedArgument = $state.$startupArgumentProperty
  if ($valueName -and $storedArgument -and (Test-Path -LiteralPath $optionsKey)) {
    $options = Get-ItemProperty -LiteralPath $optionsKey
    if ($options.$valueName -eq $storedArgument) {
      Remove-ItemProperty -LiteralPath $optionsKey -Name $valueName -Force
    }
  }
  Remove-ItemProperty -LiteralPath $stateKey -Name $startupValueProperty -Force -ErrorAction SilentlyContinue
  Remove-ItemProperty -LiteralPath $stateKey -Name $startupArgumentProperty -Force -ErrorAction SilentlyContinue
}

Remove-StartupRegistration

if ($Uninstall) {
  Remove-Item -LiteralPath $trustedLocationKey -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $stateKey -Recurse -Force -ErrorAction SilentlyContinue
  Write-Host 'Removed ExcelAI startup registration.'
  exit 0
}

if (-not (Test-Path -LiteralPath $CarrierPath -PathType Leaf)) {
  throw "Startup carrier not found: $CarrierPath"
}

New-Item -Path $optionsKey -Force | Out-Null
New-Item -Path $trustedLocationKey -Force | Out-Null
New-ItemProperty -LiteralPath $trustedLocationKey -Name 'Path' -PropertyType String -Value $carrierDirectory -Force | Out-Null
New-ItemProperty -LiteralPath $trustedLocationKey -Name 'AllowSubfolders' -PropertyType DWord -Value 0 -Force | Out-Null
New-ItemProperty -LiteralPath $trustedLocationKey -Name 'Description' -PropertyType String -Value 'ExcelAI startup carrier' -Force | Out-Null

$registeredValueName = $null
for ($i = 0; $i -le 999; $i++) {
  $candidate = if ($i -eq 0) { 'OPEN' } else { "OPEN$i" }
  $options = Get-ItemProperty -LiteralPath $optionsKey
  $existing = $options.PSObject.Properties[$candidate]
  if (-not $existing) {
    New-ItemProperty -LiteralPath $optionsKey -Name $candidate -PropertyType String -Value $argumentValue -Force | Out-Null
    $registeredValueName = $candidate
    break
  }
  if ($existing.Value -eq $argumentValue) {
    $registeredValueName = $candidate
    break
  }
}

if (-not $registeredValueName) {
  throw 'No free Excel OPEN registry value was available.'
}

New-Item -Path $stateKey -Force | Out-Null
New-ItemProperty -LiteralPath $stateKey -Name $startupValueProperty -PropertyType String -Value $registeredValueName -Force | Out-Null
New-ItemProperty -LiteralPath $stateKey -Name $startupArgumentProperty -PropertyType String -Value $argumentValue -Force | Out-Null
Write-Host "Registered startup carrier with Excel value $registeredValueName."
