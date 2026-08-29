function Assert-WheelhouseIntegrity {
    param([Parameter(Mandatory = $true)][string]$Wheelhouse)

    $ManifestPath = Join-Path $Wheelhouse "wheelhouse-manifest.json"
    if (-not (Test-Path $ManifestPath)) {
        throw "The offline package manifest is missing."
    }
    $Manifest = Get-Content -Raw -Encoding UTF8 $ManifestPath | ConvertFrom-Json
    if (-not $Manifest.files -or $Manifest.files.Count -eq 0) {
        throw "The offline package manifest contains no packages."
    }
    foreach ($Item in $Manifest.files) {
        $PackagePath = Join-Path $Wheelhouse $Item.file
        if (-not (Test-Path $PackagePath -PathType Leaf)) {
            throw "Offline package is missing: $($Item.file)"
        }
        $Actual = (Get-FileHash -Algorithm SHA256 -Path $PackagePath).Hash.ToLowerInvariant()
        if ($Actual -ne $Item.sha256) {
            throw "Offline package hash mismatch: $($Item.file)"
        }
    }
}
