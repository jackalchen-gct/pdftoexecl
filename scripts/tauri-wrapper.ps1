# tauri-wrapper.ps1 forwards tauri commands to npx tauri
# and triggers build-portable if the command was "build"

$tauriArgs = $args

# Execute tauri build / tauri dev via npx tauri
Write-Host "Running: npx tauri $tauriArgs" -ForegroundColor Gray
npx tauri $tauriArgs

if ($LASTEXITCODE -eq 0) {
    if ($tauriArgs -contains "build") {
        Write-Host "Tauri build completed successfully. Running portable packager..." -ForegroundColor Green
        powershell -ExecutionPolicy Bypass -File .\scripts\build-portable.ps1
    }
} else {
    Write-Error "Tauri command failed with exit code $LASTEXITCODE"
    exit $LASTEXITCODE
}
