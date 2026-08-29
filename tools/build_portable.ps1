param(
    [string]$PythonCommand = "python"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$Wheelhouse = Join-Path $ProjectRoot "vendor\wheelhouse"
$BuildEnvironment = Join-Path $ProjectRoot ".packaging_env"
$DistributionRoot = Join-Path $ProjectRoot "dist"
$ReleaseRoot = Join-Path $ProjectRoot "release"
$LockFile = Join-Path $ProjectRoot "requirements-build.lock"
$SpecFile = Join-Path $ProjectRoot "packaging\PnLExcelAgent.spec"
$Verifier = Join-Path $PSScriptRoot "verify_wheelhouse.ps1"

. $Verifier
Assert-WheelhouseIntegrity -Wheelhouse $Wheelhouse

& $PythonCommand -c "import struct,sys; assert sys.version_info[:2] in ((3,12),(3,13),(3,14)), 'Python 3.12-3.14 is required'; assert struct.calcsize('P') == 8, '64-bit Python is required'"
if ($LASTEXITCODE -ne 0) { throw "A working 64-bit Python 3.12, 3.13, or 3.14 build runtime is required." }

if (Test-Path $BuildEnvironment) { Remove-Item -Recurse -Force $BuildEnvironment }
& $PythonCommand -m venv $BuildEnvironment
$BuildPython = Join-Path $BuildEnvironment "Scripts\python.exe"

& $BuildPython -m pip install --no-index --only-binary=:all: --find-links $Wheelhouse --requirement $LockFile
if ($LASTEXITCODE -ne 0) { throw "Could not install the locked build tools from the offline package folder." }

Push-Location $ProjectRoot
try {
    & $BuildPython -m PyInstaller --noconfirm --clean --distpath $DistributionRoot --workpath (Join-Path $ProjectRoot "build") $SpecFile
    if ($LASTEXITCODE -ne 0) { throw "Portable application build failed." }
}
finally {
    Pop-Location
}

$AppFolder = Join-Path $DistributionRoot "P_L_Excel_Agent"
foreach ($Name in @("RUN_A08.bat", "PREPARE_WORKBOOK.bat", "LIST_AGENT_TOOLS.bat", "SELF_CHECK.bat", "REPAIR_OFFLINE.bat", "config.yaml", "README.md", "PORTABLE_DEPLOYMENT.md", "app.py", "requirements-runtime.lock")) {
    Copy-Item -Force (Join-Path $ProjectRoot $Name) (Join-Path $AppFolder $Name)
}
Copy-Item -Recurse -Force (Join-Path $ProjectRoot "src") (Join-Path $AppFolder "src")
Copy-Item -Recurse -Force (Join-Path $ProjectRoot "tools") (Join-Path $AppFolder "tools")
Get-ChildItem -Path $AppFolder -Recurse -Directory -Filter "__pycache__" | Remove-Item -Recurse -Force

$VendorFolder = Join-Path $AppFolder "vendor"
New-Item -ItemType Directory -Force -Path $VendorFolder | Out-Null
Copy-Item -Recurse -Force $Wheelhouse (Join-Path $VendorFolder "wheelhouse")

$Checksums = Get-ChildItem -Path $AppFolder -File -Recurse | Sort-Object FullName | ForEach-Object {
    $Relative = $_.FullName.Substring($AppFolder.Length + 1)
    $Hash = (Get-FileHash -Algorithm SHA256 -Path $_.FullName).Hash.ToLowerInvariant()
    "$Hash  $Relative"
}
$Checksums | Set-Content -Encoding ASCII (Join-Path $AppFolder "SHA256SUMS.txt")

New-Item -ItemType Directory -Force -Path $ReleaseRoot | Out-Null
$ZipPath = Join-Path $ReleaseRoot "P_L_Excel_Agent_Windows_x64_Portable.zip"
if (Test-Path $ZipPath) { Remove-Item -Force $ZipPath }
Compress-Archive -Path $AppFolder -DestinationPath $ZipPath -CompressionLevel Optimal

Write-Host "Portable release created: $ZipPath"
