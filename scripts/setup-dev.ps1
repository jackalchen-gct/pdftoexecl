param(
    [string]$PythonVersion = "3.12",
    [switch]$SkipNodeInstall
)

$ErrorActionPreference = "Stop"

function Test-Command {
    param([Parameter(Mandatory = $true)][string]$Name)
    return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

function Ensure-Winget {
    if (-not (Test-Command -Name "winget")) {
        throw "winget is required to install prerequisites automatically."
    }
}

function Install-WingetPackage {
    param(
        [Parameter(Mandatory = $true)][string]$Id,
        [string]$Override
    )

    $args = @(
        "install",
        "--id", $Id,
        "--source", "winget",
        "--accept-package-agreements",
        "--accept-source-agreements"
    )

    if ($Override) {
        $args += @("--override", $Override)
    }

    & winget @args
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to install package: $Id"
    }
}

function Refresh-Path {
    $machine = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $user = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = @($machine, $user) -join ";"
}

function Resolve-PythonRunner {
    $venvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
    if (Test-Path $venvPython) {
        return [pscustomobject]@{ Exe = $venvPython; PrefixArgs = @() }
    }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python -and $python.Source -and $python.Source -notlike "*WindowsApps*") {
        return [pscustomobject]@{ Exe = $python.Source; PrefixArgs = @() }
    }

    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        return [pscustomobject]@{ Exe = $py.Source; PrefixArgs = @("-3") }
    }

    return $null
}

function Invoke-Python {
    param(
        [Parameter(Mandatory = $true)]$Runner,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    if ($Runner.PrefixArgs.Count -gt 0) {
        & $Runner.Exe @($Runner.PrefixArgs + $Arguments)
    }
    else {
        & $Runner.Exe @Arguments
    }

    return $LASTEXITCODE
}

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

Write-Host "Project root: $ProjectRoot"

Ensure-Winget

$pythonRunner = Resolve-PythonRunner
if (-not $pythonRunner) {
    Write-Host "Installing Python $PythonVersion ..."
    Install-WingetPackage -Id "Python.Python.$PythonVersion"
    Refresh-Path
    $pythonRunner = Resolve-PythonRunner
}

if (-not (Test-Command -Name "rustup") -and -not (Test-Command -Name "cargo")) {
    Write-Host "Installing Rust toolchain via rustup ..."
    Install-WingetPackage -Id "Rustlang.Rustup"
    Refresh-Path
}

if (-not $SkipNodeInstall -and -not (Test-Command -Name "node")) {
    Write-Host "Installing Node.js LTS ..."
    Install-WingetPackage -Id "OpenJS.NodeJS.LTS"
    Refresh-Path
}

if (Test-Command -Name "rustup") {
    & rustup default stable-msvc
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to set Rust default toolchain to stable-msvc."
    }
}

if (-not (Get-Command cl.exe -ErrorAction SilentlyContinue)) {
    Write-Host "Installing Visual Studio Build Tools with the C++ workload ..."
    try {
        Install-WingetPackage -Id "Microsoft.VisualStudio.2022.BuildTools" -Override "--passive --wait --norestart --add Microsoft.VisualStudio.Workload.VCTools --includeRecommended"
        Refresh-Path
    }
    catch {
        Write-Warning "Automatic C++ Build Tools installation failed. Install the Desktop development with C++ workload manually if Tauri builds fail."
    }
}

if (-not (Test-Path (Join-Path $ProjectRoot "requirements.txt"))) {
    throw "Missing requirements.txt in project root."
}

if (-not $pythonRunner) {
    throw "Python was not found after installation."
}

if (-not (Test-Path (Join-Path $ProjectRoot ".venv"))) {
    Write-Host "Creating Python virtual environment ..."
    $venvExit = Invoke-Python -Runner $pythonRunner -Arguments @("-m", "venv", (Join-Path $ProjectRoot ".venv"))
    if ($venvExit -ne 0) {
        throw "Failed to create .venv"
    }
}

$venvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $venvPython)) {
    throw "Python virtual environment is incomplete: $venvPython"
}

Write-Host "Upgrading pip ..."
$pipExit = & $venvPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) {
    throw "Failed to upgrade pip."
}

Write-Host "Installing Python dependencies ..."
$depsExit = & $venvPython -m pip install -r (Join-Path $ProjectRoot "requirements.txt")
if ($LASTEXITCODE -ne 0) {
    throw "Failed to install Python dependencies."
}

Write-Host "Installing Node dependencies ..."
Push-Location $ProjectRoot
try {
    npm install
    if ($LASTEXITCODE -ne 0) {
        throw "npm install failed."
    }
}
finally {
    Pop-Location
}

$env:PDFTOEXECL_PYTHON = $venvPython

if (-not (Get-Command cl.exe -ErrorAction SilentlyContinue)) {
    Write-Warning "Visual Studio C++ Build Tools were not detected. Install the Desktop development with C++ workload before running Tauri builds."
}

Write-Host ""
Write-Host "Setup complete."
Write-Host "Python: $venvPython"
Write-Host "Rust: cargo and rustc are managed by rustup in your user profile."
Write-Host "Next: run 'npm run tauri dev' after opening a new PowerShell window, or keep using this session with PDFTOEXECL_PYTHON set."
