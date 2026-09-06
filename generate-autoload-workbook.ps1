param(
  [string]$SourceWorkbook = (Join-Path $PSScriptRoot 'DeepSeekExcelAssistant-MacroSource.xlsm'),
  [string]$OutputPath = (Join-Path $PSScriptRoot 'DeepSeekExcelAssistant-Standalone-Autoload.xlsm'),
  [string]$AddinId = '1537f254-10aa-41d5-aed2-0a00b89da104',
  [string]$AddinVersion = '1.0.41.0'
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $SourceWorkbook -PathType Leaf)) {
  throw "Source workbook not found: $SourceWorkbook"
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$outputDirectory = Split-Path -Parent $OutputPath
if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
  New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

Copy-Item -LiteralPath $SourceWorkbook -Destination $OutputPath -Force

function Replace-ZipEntryText {
  param(
    [System.IO.Compression.ZipArchive]$Archive,
    [string]$EntryName,
    [scriptblock]$Transform
  )

  $entry = $Archive.GetEntry($EntryName)
  if (-not $entry) {
    throw "Required package entry not found: $EntryName"
  }

  $reader = [System.IO.StreamReader]::new($entry.Open())
  try { $original = $reader.ReadToEnd() }
  finally { $reader.Dispose() }

  $updated = & $Transform $original
  if ([string]::IsNullOrWhiteSpace($updated)) {
    throw "Transform produced empty content for: $EntryName"
  }

  $entry.Delete()
  $replacement = $Archive.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
  $writer = [System.IO.StreamWriter]::new($replacement.Open(), [System.Text.UTF8Encoding]::new($false))
  try { $writer.Write($updated) }
  finally { $writer.Dispose() }
}

$archive = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
  Replace-ZipEntryText $archive 'xl/workbook.xml' {
    param($xml)
    if ($xml -match '<workbookView\s') {
      if ($xml -match '<workbookView\b[^>]*\bvisibility="[^"]*"') {
        [regex]::Replace($xml, '(<workbookView\b[^>]*\b)visibility="[^"]*"', '${1}visibility="hidden"', 1)
      }
      else {
        [regex]::Replace($xml, '<workbookView\s', '<workbookView visibility="hidden" ', 1)
      }
    }
    else {
      throw 'Workbook package has no workbookView element.'
    }
  }

  Replace-ZipEntryText $archive 'xl/webextensions/taskpanes.xml' {
    param($xml)
    $xml = [regex]::Replace($xml, 'visibility="[01]"', 'visibility="1"')
    [regex]::Replace($xml, 'width="[0-9]+"', 'width="613"')
  }

  Replace-ZipEntryText $archive 'xl/webextensions/webextension1.xml' {
    param($xml)
    $xml = [regex]::Replace($xml, 'id="[0-9a-fA-F-]{36}"', "id=`"$AddinId`"", 1)
    [regex]::Replace($xml, 'version="[0-9]+(?:\.[0-9]+){3}"', "version=`"$AddinVersion`"")
  }
}
finally {
  $archive.Dispose()
}

Write-Host "Generated hidden macro-enabled autoload workbook: $OutputPath"
