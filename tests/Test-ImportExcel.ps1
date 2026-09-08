param(
  [string]$SamplePath = '\\servertest\Bus\Buoninfante\Asc\SINTESI\faxb2b.xls',
  [string]$AssemblyPath = (Join-Path $PSScriptRoot '..\TEST\ExcelDataReader.dll'),
  [string]$ExpectedSheet = 'faxb2b',
  [int]$MinimumColumns = 15
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $SamplePath)) {
  throw "File Excel non trovato: $SamplePath"
}
if (-not (Test-Path -LiteralPath $AssemblyPath)) {
  throw "ExcelDataReader.dll non trovata. Compilare prima la solution: $AssemblyPath"
}

Add-Type -Path (Resolve-Path -LiteralPath $AssemblyPath)
$stream = [System.IO.File]::Open(
  $SamplePath,
  [System.IO.FileMode]::Open,
  [System.IO.FileAccess]::Read,
  [System.IO.FileShare]::ReadWrite
)

try {
  $reader = [ExcelDataReader.ExcelReaderFactory]::CreateReader($stream)
  try {
    $found = $false
    do {
      if ([string]::Equals($reader.Name, $ExpectedSheet,
          [System.StringComparison]::OrdinalIgnoreCase)) {
        $found = $true
        break
      }
    } while ($reader.NextResult())

    if (-not $found) {
      throw "Foglio '$ExpectedSheet' non trovato."
    }
    if ($reader.FieldCount -lt $MinimumColumns) {
      throw "Il foglio contiene $($reader.FieldCount) colonne; richieste almeno $MinimumColumns."
    }

    $rows = 0
    while ($reader.Read()) { $rows++ }
    if ($rows -eq 0) {
      throw "Il foglio '$ExpectedSheet' non contiene righe."
    }

    [pscustomobject]@{
      Result = 'PASS'
      Sheet = $reader.Name
      Rows = $rows
      Columns = $reader.FieldCount
    }
  }
  finally {
    $reader.Dispose()
  }
}
finally {
  $stream.Dispose()
}
