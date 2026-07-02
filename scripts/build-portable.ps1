$workspace = "C:\Users\jackal.chen\Workspace\pdftoexecl"
$binDir = "$workspace\src-tauri\bin"
$releaseDir = "$workspace\src-tauri\target\release"
$portableDir = "$workspace\src-tauri\target\pdftoexecl-portable"
$portableBinDir = "$portableDir\bin"
$zipFile = "$workspace\src-tauri\target\pdftoexecl-portable.zip"

Write-Host "=== Starting Portable Build Packaging ===" -ForegroundColor Green

# Create directories if they don't exist
if (-not (Test-Path $portableBinDir)) {
    New-Item -ItemType Directory -Path $portableBinDir -Force | Out-Null
}

# Copy converter-sidecar-x86_64-pc-windows-msvc.exe to bin
$sidecarSource = "$binDir\converter-sidecar-x86_64-pc-windows-msvc.exe"
if (Test-Path $sidecarSource) {
    Write-Host "Copying sidecar from $sidecarSource to $portableBinDir..." -ForegroundColor Cyan
    Copy-Item -Path $sidecarSource -Destination $portableBinDir -Force
} else {
    Write-Warning "Sidecar source not found at $sidecarSource"
}

# Copy pdftoexecl.exe to portable root
$appSource = "$releaseDir\pdftoexecl.exe"
if (Test-Path $appSource) {
    Write-Host "Copying app from $appSource to $portableDir..." -ForegroundColor Cyan
    Copy-Item -Path $appSource -Destination $portableDir -Force
} else {
    Write-Warning "App source not found at $appSource"
}

# Zip pdftoexecl-portable
if (Test-Path $zipFile) {
    Remove-Item -Path $zipFile -Force | Out-Null
}

Write-Host "Zipping $portableDir into $zipFile..." -ForegroundColor Cyan
Compress-Archive -Path "$portableDir\*" -DestinationPath $zipFile -Force
Write-Host "=== Portable Build Complete! ZIP saved at $zipFile ===" -ForegroundColor Green
