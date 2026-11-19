<#
Windows integration smoke test. Needs desktop Microsoft Excel and Python with
pywin32 + pymupdf installed. Uses the synthetic 47-chart workbook in examples/.

This test intentionally exercises the native Windows clipboard (EMF export).
Other clipboard activity during the run can cause interference.
#>
param(
    [string]$Workbook = (Join-Path $PSScriptRoot "..\examples\Excel_Chart_Extraction_Stress_Test.xlsx"),
    [string]$Output = (Join-Path $PSScriptRoot "..\smoke-test-output"),
    [string]$Python = "py"
)

$ErrorActionPreference = "Stop"
$Source = (Resolve-Path $Workbook).Path
$Script = (Resolve-Path (Join-Path $PSScriptRoot "..\excel_charts_to_svg.py")).Path

Write-Host "Running native Excel smoke test on $Source"
if ($Python -eq "py") {
    & py -3 $Script $Source --engine excel --format both --show-excel -o $Output
} else {
    & $Python $Script $Source --engine excel --format both --show-excel -o $Output
}
if ($LASTEXITCODE -ne 0) {
    throw "Exporter failed with exit code $LASTEXITCODE. Review the manifest in $Output."
}

$ManifestPath = Join-Path $Output "manifest.json"
if (-not (Test-Path $ManifestPath)) { throw "Manifest missing: $ManifestPath" }
$Manifest = Get-Content -Raw -Encoding UTF8 $ManifestPath | ConvertFrom-Json
$Summary = $Manifest.summary
if ($Summary.workbooks -ne 1 -or $Summary.charts_enumerated -ne 47 -or
    $Summary.charts_exported -ne 47 -or $Summary.files_exported -ne 94 -or
    $Summary.errors -ne 0) {
    throw ("Smoke test incomplete: " + ($Summary | ConvertTo-Json -Compress))
}

$Exported = Get-ChildItem $Output -Recurse -File | Where-Object {
    $_.Extension -in @('.svg', '.emf')
}
if ($Exported.Count -lt 94) {
    throw "Expected at least 94 chart files, found $($Exported.Count)."
}

Write-Host "SUCCESS: 47 charts, 94 vector graphic files, zero extraction errors."
Write-Host "Manifest: $ManifestPath"
