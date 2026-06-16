param(
    [switch]$SkipSidecarBuild
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

if (-not $SkipSidecarBuild) {
    Write-Host "Building Python sidecar ..."
    & powershell -ExecutionPolicy Bypass -File (Join-Path $ProjectRoot "scripts\build-sidecar.ps1")
    if ($LASTEXITCODE -ne 0) {
        throw "Sidecar build failed."
    }
}

Write-Host "Starting Vite dev server ..."
Push-Location $ProjectRoot
try {
    npm run dev
}
finally {
    Pop-Location
}
