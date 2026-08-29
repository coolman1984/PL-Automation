param(
    [string]$PythonCommand = "python"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$Wheelhouse = Join-Path $ProjectRoot "vendor\wheelhouse"
$LockFile = Join-Path $ProjectRoot "requirements-build.lock"

if (Test-Path $Wheelhouse) { Remove-Item -Recurse -Force $Wheelhouse }
New-Item -ItemType Directory -Force -Path $Wheelhouse | Out-Null

& $PythonCommand -c "import struct,sys; assert sys.version_info >= (3,11), 'Python 3.11+ is required'; assert struct.calcsize('P') == 8, '64-bit Python is required'"
if ($LASTEXITCODE -ne 0) { throw "A working 64-bit Python 3.11+ download runtime is required." }

foreach ($Target in @("312", "313", "314")) {
    $Abi = "cp$Target"
    & $PythonCommand -m pip download --only-binary=:all: --platform win_amd64 --implementation cp --python-version $Target --abi $Abi --dest $Wheelhouse --requirement $LockFile
    if ($LASTEXITCODE -ne 0) { throw "Could not download the locked Windows packages for Python $Target." }
}

$Files = Get-ChildItem -Path $Wheelhouse -File | Sort-Object Name | ForEach-Object {
    $Hash = Get-FileHash -Algorithm SHA256 -Path $_.FullName
    [ordered]@{
        file = $_.Name
        size_bytes = $_.Length
        sha256 = $Hash.Hash.ToLowerInvariant()
    }
}

$Manifest = [ordered]@{
    generated_utc = [DateTime]::UtcNow.ToString("o")
    target = "Windows 10/11 x64, Python 3.12/3.13/3.14"
    files = @($Files)
}
$Manifest | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $Wheelhouse "wheelhouse-manifest.json")

Write-Host "Offline package set created at $Wheelhouse"
