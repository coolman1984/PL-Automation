# Task 3 (V2 plan): deterministic read-only COM acceptance harness.
#
# Runs ONLY the gated read-range integration test and records the evidence
# required by docs/UNIVERSAL_EXCEL_AGENT_CODING_PLAN_V2.md:
#   - source SHA-256 before and after (must match),
#   - exact pytest parameters and output,
#   - pytest exit code (nonzero fails the harness),
#   - Excel PIDs before and after (no new isolated Excel PID may remain).
#
# A handled faulthandler 0x8001010D/0x80010108 diagnostic in the pytest output
# is recorded but is NOT treated as a crash when the pytest process exits zero.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass ^
#     -File tools\run_read_range_acceptance.ps1 ^
#     -Workbook "D:\path\to\workbook.xlsb" -Sheet "VD Total" -Address "A1:C10"
#
# The workbook is opened read-only by the tool under test. The harness itself
# never writes to the workbook. A JSON acceptance report is retained under
# work\acceptance\.

param(
    [Parameter(Mandatory = $true)]
    [string]$Workbook,

    [Parameter(Mandatory = $true)]
    [string]$Sheet,

    [Parameter(Mandatory = $true)]
    [string]$Address,

    [string]$Python = "python",

    [ValidateSet("auto", "attach", "open")]
    [string]$Mode = "auto",

    [string]$ReportDir = ""
)

$ErrorActionPreference = "Stop"

function Fail([string]$Message) {
    Write-Error $Message
    exit 1
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$workbookPath = (Resolve-Path -LiteralPath $Workbook).Path
if ([System.IO.Path]::GetExtension($workbookPath).ToLowerInvariant() -ne ".xlsb") {
    Fail "Acceptance requires an .xlsb workbook; got: $workbookPath"
}
if (-not (Test-Path -LiteralPath $workbookPath -PathType Leaf)) {
    Fail "Workbook not found: $workbookPath"
}

$pytestTarget = Join-Path $repoRoot "tests\integration\test_read_range.py"
if (-not (Test-Path -LiteralPath $pytestTarget)) {
    Fail "Read-range test not found: $pytestTarget"
}

if ([string]::IsNullOrWhiteSpace($ReportDir)) {
    $ReportDir = Join-Path $repoRoot "work\acceptance"
}
New-Item -ItemType Directory -Force -Path $ReportDir | Out-Null

function Get-ExcelPids {
    return @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | ForEach-Object { $_.Id } | Sort-Object)
}

function Get-SourceHash {
    return (Get-FileHash -LiteralPath $workbookPath -Algorithm SHA256).Hash
}

$hashBefore = Get-SourceHash
$pidsBefore = Get-ExcelPids
$startedAt = (Get-Date).ToUniversalTime().ToString("o")

# Run the single gated read-range test with explicit environment gates.
$env:PL_COM_TESTS = "1"
$env:PL_COM_WORKBOOK = $workbookPath
$env:PL_COM_READ_SHEET = $Sheet
$env:PL_COM_READ_ADDRESS = $Address
$env:PL_COM_READ_MODE = $Mode

$pytestLog = Join-Path $ReportDir ("read_range_pytest_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
& $Python -m pytest $pytestTarget -v --no-header > $pytestLog 2>&1
$pytestExit = $LASTEXITCODE

$hashAfter = Get-SourceHash
$pidsAfter = Get-ExcelPids
$finishedAt = (Get-Date).ToUniversalTime().ToString("o")

$orphaned = @($pidsAfter | Where-Object { $pidsBefore -notcontains $_ })

$pytestText = ""
if (Test-Path -LiteralPath $pytestLog) {
    $pytestText = Get-Content -LiteralPath $pytestLog -Raw
}
$comDiagnosticSeen = $pytestText -match "0x8001010"

$checks = [ordered]@{
    pytest_exit_zero        = ($pytestExit -eq 0)
    source_hash_unchanged   = ($hashBefore -eq $hashAfter)
    no_orphaned_excel_pids  = ($orphaned.Count -eq 0)
}
$passed = $true
foreach ($value in $checks.Values) { if (-not $value) { $passed = $false } }

$report = [ordered]@{
    schema_version        = "1.0"
    harness               = "run_read_range_acceptance"
    started_utc           = $startedAt
    finished_utc          = $finishedAt
    workbook              = $workbookPath
    sheet                 = $Sheet
    address               = $Address
    mode                  = $Mode
    pytest_target         = $pytestTarget
    pytest_exit_code      = $pytestExit
    pytest_log            = $pytestLog
    source_sha256_before  = $hashBefore
    source_sha256_after   = $hashAfter
    excel_pids_before     = $pidsBefore
    excel_pids_after      = $pidsAfter
    orphaned_excel_pids   = $orphaned
    com_shutdown_diagnostic_seen = $comDiagnosticSeen
    com_shutdown_diagnostic_treated_as_crash = ($comDiagnosticSeen -and $pytestExit -ne 0)
    checks                = $checks
    passed                = $passed
}

$reportPath = Join-Path $ReportDir ("read_range_acceptance_{0}.json" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
$report | ConvertTo-Json -Depth 4 | Out-File -LiteralPath $reportPath -Encoding utf8

Write-Host "Read-range acceptance report: $reportPath"
Write-Host ("pytest exit: {0} | hash unchanged: {1} | orphaned Excel PIDs: {2} | PASSED: {3}" -f `
    $pytestExit, ($hashBefore -eq $hashAfter), $orphaned.Count, $passed)

if ($pytestExit -ne 0) {
    Write-Host "See pytest log: $pytestLog"
}
if (-not $passed) {
    exit 1
}
exit 0
