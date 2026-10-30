param(
  [string]$WorkbookPath = (Join-Path $PSScriptRoot "service_rating_translation-config.xlsx"),
  [string]$OutputPath = (Join-Path $PSScriptRoot "translation-config.js")
)

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
  throw "Translation workbook not found: $WorkbookPath"
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $WorkbookPath))

try {
  $sharedStringsEntry = $archive.GetEntry("xl/sharedStrings.xml")
  $worksheetEntry = $archive.GetEntry("xl/worksheets/sheet1.xml")
  if ($null -eq $sharedStringsEntry -or $null -eq $worksheetEntry) {
    throw "The workbook must contain shared strings and a first worksheet."
  }

  $sharedStringsReader = [System.IO.StreamReader]::new($sharedStringsEntry.Open())
  $worksheetReader = [System.IO.StreamReader]::new($worksheetEntry.Open())
  try {
    [xml]$sharedStringsXml = $sharedStringsReader.ReadToEnd()
    [xml]$worksheetXml = $worksheetReader.ReadToEnd()
  }
  finally {
    $sharedStringsReader.Dispose()
    $worksheetReader.Dispose()
  }

  $sharedStrings = @($sharedStringsXml.sst.si | ForEach-Object {
    if ($null -ne $_.t) { return [string]$_.t }
    return (@($_.r | ForEach-Object { [string]$_.t }) -join "")
  })
  $translations = [ordered]@{ EN = [ordered]@{}; PH = [ordered]@{} }

  foreach ($row in @($worksheetXml.worksheet.sheetData.row)) {
    if ([int]$row.r -eq 1) { continue }
    $cells = @{}
    foreach ($cell in @($row.c)) {
      $column = ([string]$cell.r -replace "\d", "")
      $cells[$column] = if ([string]$cell.t -eq "s") { $sharedStrings[[int]$cell.v] } else { [string]$cell.v }
    }

    if ([string]::IsNullOrWhiteSpace($cells["A"])) { continue }
    $translations.EN[$cells["A"]] = $cells["B"]
    $translations.PH[$cells["A"]] = $cells["C"]
  }

  $content = "/* Generated from service_rating_translation-config.xlsx. */`nwindow.serviceRatingTranslationConfig = " + ($translations | ConvertTo-Json -Depth 3) + ";`n"
  Set-Content -LiteralPath $OutputPath -Value $content -Encoding utf8
}
finally {
  $archive.Dispose()
}
