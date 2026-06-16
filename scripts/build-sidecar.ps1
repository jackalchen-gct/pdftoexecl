param(
    [switch]$ForceRebuild
)

$ErrorActionPreference = "Stop"

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$ConverterScript = Join-Path $ProjectRoot "converter\convert_pdf.py"
$DistDir = Join-Path $ProjectRoot "src-tauri\bin"
$BuildRoot = Join-Path $ProjectRoot ".sidecar-build"

function Invoke-Python {
    param(
        [Parameter(Mandatory = $true)]$Runner,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    if ($Runner -is [string[]]) {
        & $Runner[0] @($Runner[1..($Runner.Length - 1)] + $Arguments)
    }
    else {
        & $Runner @Arguments
    }

    if ($LASTEXITCODE -ne 0) {
        throw "Python command failed."
    }
}

if (-not (Test-Path $ConverterScript)) {
    throw "Missing converter script: $ConverterScript"
}

if (-not (Test-Path $VenvPython)) {
    throw "Missing virtual environment at $VenvPython. Run scripts/setup-dev.ps1 first."
}

$pyinstallerVersion = & $VenvPython -m PyInstaller --version 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Installing PyInstaller ..."
    Invoke-Python -Runner $VenvPython -Arguments @("-m", "pip", "install", "pyinstaller")
} else {
    Write-Host "PyInstaller detected: $pyinstallerVersion"
}

if ($ForceRebuild -and (Test-Path $BuildRoot)) {
    Remove-Item -Recurse -Force $BuildRoot
}

$targetTriple = if ($env:CARGO_BUILD_TARGET) { $env:CARGO_BUILD_TARGET } else { "x86_64-pc-windows-msvc" }
$sidecarName = "converter-sidecar-$targetTriple"

New-Item -ItemType Directory -Force -Path $DistDir | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $BuildRoot "work") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $BuildRoot "spec") | Out-Null

Write-Host "Building sidecar $sidecarName ..."
Invoke-Python -Runner $VenvPython -Arguments @(
    "-m",
    "PyInstaller",
    "--noconfirm",
    "--clean",
    "--onefile",
    "--console",
    "--name",
    $sidecarName,
    "--distpath",
    $DistDir,
    "--workpath",
    (Join-Path $BuildRoot "work"),
    "--specpath",
    (Join-Path $BuildRoot "spec"),
    $ConverterScript
)

Write-Host "Sidecar written to $(Join-Path $DistDir ($sidecarName + '.exe'))"
