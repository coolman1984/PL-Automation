param(
    [string]$PythonCommand = "python"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$Wheelhouse = Join-Path $ProjectRoot "vendor\wheelhouse"
$Runtime = Join-Path $ProjectRoot ".runtime_env"
$LockFile = Join-Path $ProjectRoot "requirements-runtime.lock"
$Verifier = Join-Path $PSScriptRoot "verify_wheelhouse.ps1"

. $Verifier
Assert-WheelhouseIntegrity -Wheelhouse $Wheelhouse

& $PythonCommand -c "import struct,sys; assert sys.version_info[:2] in ((3,12),(3,13),(3,14)), 'Python 3.12-3.14 is required'; assert struct.calcsize('P') == 8, '64-bit Python is required'"
if ($LASTEXITCODE -ne 0) { throw "A working 64-bit Python 3.12, 3.13, or 3.14 is required for source-mode repair." }

if (Test-Path $Runtime) {
    $BackupName = ".runtime_env_failed_{0}" -f [DateTime]::UtcNow.ToString("yyyyMMdd_HHmmss")
    Move-Item -Path $Runtime -Destination (Join-Path $ProjectRoot $BackupName)
}

& $PythonCommand -m venv $Runtime
if ($LASTEXITCODE -ne 0) { throw "Could not create the private runtime." }

$PrivatePython = Join-Path $Runtime "Scripts\python.exe"
& $PrivatePython -m pip install --no-index --only-binary=:all: --find-links $Wheelhouse --requirement $LockFile
if ($LASTEXITCODE -ne 0) { throw "Offline dependency installation failed." }

& $PrivatePython (Join-Path $ProjectRoot "app.py") --self-check
if ($LASTEXITCODE -ne 0) { throw "The repaired private runtime did not pass self-check." }
