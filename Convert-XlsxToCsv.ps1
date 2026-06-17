<#
.SYNOPSIS
    Convert an Excel .xlsx workbook to UTF-8 CSV (via Excel COM).

.DESCRIPTION
    Exports as CSV UTF-8 (xlCSVUTF8). By default exports the first worksheet to
    "<name>.csv" next to the source file as UTF-8 *with* BOM. Use -NoBom for
    BOM-less output, and -AllSheets to export every worksheet to its own file.
    Opens read-only and releases every COM object so Excel quits cleanly (no
    crash-recovery on the next run). Force-kills its own EXCEL.EXE only as a
    fallback if it fails to exit; any Excel you already had open is untouched.

.PARAMETER Path       Path to the .xlsx file. (Required)
.PARAMETER OutDir     Output folder. Defaults to the source file's folder.
.PARAMETER NoBom      Write UTF-8 WITHOUT a byte-order mark.
.PARAMETER AllSheets  Export every worksheet to its own CSV.

.EXAMPLE
    & "C:\temp\Convert-XlsxToCsv.ps1" -Path "C:\temp\Team members.xlsx" -NoBom
#>
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$OutDir,
    [switch]$NoBom,
    [switch]$AllSheets
)

$ErrorActionPreference = 'Stop'
$xlCSVUTF8 = 62   # Excel FileFormat: CSV UTF-8 (comma delimited)

if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
$src = (Resolve-Path -LiteralPath $Path).Path
if (-not $OutDir) { $OutDir = Split-Path -Parent $src }
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
$base = [System.IO.Path]::GetFileNameWithoutExtension($src)

function Remove-Bom([string]$file) {
    $enc  = New-Object System.Text.UTF8Encoding($false)
    $text = [System.IO.File]::ReadAllText($file)
    [System.IO.File]::WriteAllText($file, $text, $enc)
}
function Get-SafeName([string]$n) { ($n -replace '[\\/:*?"<>|]', '_').Trim() }
function Release($o) { if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }

$excel = $null; $books = $null; $excelPid = 0
$written = @()
try {
    $preExisting = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    try { $excel = New-Object -ComObject Excel.Application }
    catch { throw "Microsoft Excel is required (COM) but could not start. $($_.Exception.Message)" }
    $spawn = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $preExisting -notcontains $_.Id })
    if ($spawn.Count -ge 1) { $excelPid = $spawn[0].Id }

    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $books = $excel.Workbooks

    $indexes = @(1)
    if ($AllSheets) {
        $wb0 = $books.Open($src, 0, $true)
        $sheets0 = $wb0.Worksheets; $count = $sheets0.Count
        Release $sheets0; $wb0.Close($false); Release $wb0
        $indexes = 1..$count
    }

    foreach ($i in $indexes) {
        $wb = $books.Open($src, 0, $true)   # UpdateLinks=0, ReadOnly=$true
        $coll = $wb.Worksheets; $ws = $coll.Item($i)
        $name = if ($AllSheets) { "$base - $(Get-SafeName $ws.Name).csv" } else { "$base.csv" }
        $out  = Join-Path $OutDir $name
        $ws.Activate()
        $wb.SaveAs($out, $xlCSVUTF8)
        $wb.Close($false)
        Release $ws; Release $coll; Release $wb
        if ($NoBom) { Remove-Bom $out }
        $written += $out
    }
}
finally {
    Release $books
    if ($excel) { try { $excel.Quit() } catch {}; Release $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # Let it exit cleanly; only force-kill if it's genuinely stuck.
    if ($excelPid -gt 0) {
        $deadline = (Get-Date).AddSeconds(4)
        while ((Get-Date) -lt $deadline -and (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) {
            Start-Sleep -Milliseconds 150
        }
        $p = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
        if ($p) { $p | Stop-Process -Force -ErrorAction SilentlyContinue }
    }
}

$bom = if ($NoBom) { "without BOM" } else { "with BOM" }
Write-Host "Done - UTF-8 CSV ($bom):"
$written | ForEach-Object { Write-Host "  $_" }
