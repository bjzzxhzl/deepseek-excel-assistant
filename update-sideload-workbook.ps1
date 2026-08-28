param(
  [Parameter(Mandatory = $true)]
  [string]$SourcePath,

  [Parameter(Mandatory = $true)]
  [string]$OutputPath,

  [string]$ManifestPath = ''
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
  $ManifestPath = Join-Path $PSScriptRoot 'manifest.xml'
}

$source = (Resolve-Path -LiteralPath $SourcePath).Path
$manifestFile = (Resolve-Path -LiteralPath $ManifestPath).Path
[xml]$manifest = Get-Content -LiteralPath $manifestFile -Raw -Encoding UTF8
$addinId = [string]$manifest.OfficeApp.Id
$addinVersion = [string]$manifest.OfficeApp.Version
if ([string]::IsNullOrWhiteSpace($addinId) -or [string]::IsNullOrWhiteSpace($addinVersion)) {
  throw 'manifest.xml is missing Id or Version.'
}

$output = [IO.Path]::GetFullPath($OutputPath)
$outputDir = Split-Path -Parent $output
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
$working = Join-Path ([IO.Path]::GetTempPath()) (([IO.Path]::GetRandomFileName()) + '.xlsx')

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

try {
  Copy-Item -LiteralPath $source -Destination $working -Force
  $archive = [IO.Compression.ZipFile]::Open($working, [IO.Compression.ZipArchiveMode]::Update)
  try {
    $entryName = 'xl/webextensions/webextension1.xml'
    $entry = $archive.GetEntry($entryName)
    if ($null -eq $entry) { throw "Workbook entry not found: $entryName" }

    $reader = [IO.StreamReader]::new($entry.Open(), [Text.UTF8Encoding]::new($false), $true)
    try { [xml]$webExtension = $reader.ReadToEnd() } finally { $reader.Dispose() }

    $reference = $webExtension.webextension.reference
    if ($null -eq $reference) { throw 'The sideload workbook has no we:reference element.' }
    if ([string]$reference.id -ne $addinId) {
      throw "Workbook add-in Id '$($reference.id)' does not match manifest Id '$addinId'."
    }
    $reference.SetAttribute('version', $addinVersion)

    $settings = [Xml.XmlWriterSettings]::new()
    $settings.Encoding = [Text.UTF8Encoding]::new($false)
    $settings.Indent = $false
    $settings.OmitXmlDeclaration = $false
    $memory = [IO.MemoryStream]::new()
    $writer = [Xml.XmlWriter]::Create($memory, $settings)
    try { $webExtension.Save($writer) } finally { $writer.Dispose() }
    $bytes = $memory.ToArray()
    $memory.Dispose()

    $entry.Delete()
    $newEntry = $archive.CreateEntry($entryName, [IO.Compression.CompressionLevel]::Optimal)
    $stream = $newEntry.Open()
    try { $stream.Write($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
  }
  finally {
    $archive.Dispose()
  }

  Copy-Item -LiteralPath $working -Destination $output -Force
  Write-Host "Sideload workbook synchronized: Id=$addinId Version=$addinVersion"
  Write-Host "Output: $output"
}
finally {
  if (Test-Path -LiteralPath $working) { Remove-Item -LiteralPath $working -Force }
}
