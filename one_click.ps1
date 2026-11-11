$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "  ==================================================" -ForegroundColor Cyan
Write-Host "   TABLE WORKER - one click clean" -ForegroundColor Cyan
Write-Host "  ==================================================" -ForegroundColor Cyan
Write-Host ""

$Engine = Join-Path $Root "table_worker.py"
if (-not (Test-Path $Engine)) {
    Write-Host "  [ERROR] table_worker.py not found next to this script." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

$Py = "C:\Users\geng\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"
if (-not (Test-Path $Py)) { $Py = "python" }

$InputFile = $args | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $InputFile) {
    $InputFile = Join-Path $Root "samples\dirty-table.xlsx"
    Write-Host "  [HINT] No file given. Using the built-in sample." -ForegroundColor Yellow
    Write-Host "         Usage: drag an .xlsx or .csv file onto this file." -ForegroundColor Yellow
    Write-Host ""
}

if (-not (Test-Path $InputFile)) {
    Write-Host "  [ERROR] File not found: $InputFile" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

$Stamp = Get-Date -Format "yyyyMMdd_HHmm"
$OutDir = Join-Path $Root ("result_" + $Stamp)

Write-Host "  INPUT : $InputFile"
Write-Host "  OUTPUT: $OutDir"
Write-Host ""

& $Py -X utf8 $Engine $InputFile --out $OutDir
$Code = $LASTEXITCODE

if ($Code -ne 0) {
    Write-Host ""
    Write-Host "  [FAILED] Processing did not finish. Common causes:" -ForegroundColor Red
    Write-Host "    1. The file is open in Excel. Close it and retry."
    Write-Host "    2. The file is not .xlsx / .csv. Save as xlsx."
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit $Code
}

Write-Host ""
Write-Host "  ==================================================" -ForegroundColor Green
Write-Host "   DONE. Results are here:" -ForegroundColor Green
Write-Host "   $OutDir" -ForegroundColor Green
Write-Host ""
Write-Host "   Look at these three:"
Write-Host "     report.md    - one page summary, send to the client"
Write-Host "     issues.xlsx  - the problem list"
Write-Host "     clean.xlsx   - the cleaned table"
Write-Host "  ==================================================" -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to open the result folder"
Invoke-Item $OutDir
