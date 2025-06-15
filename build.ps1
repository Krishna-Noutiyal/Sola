# Enable ANSI escape codes in PowerShell
$esc = [char]27
$version = "1.3.2"
function Write-Info($msg) {
    Write-Host "$esc[1;34m[INFO]$esc[0m $msg"
}
function Write-Success($msg) {
    Write-Host "$esc[1;32m[SUCCESS]$esc[0m $msg"
}
function Write-ErrorMsg($msg) {
    Write-Host "$esc[1;31m[ERROR]$esc[0m $msg"
}
function Write-Section($msg) {
    Write-Host ""
    Write-Host "$esc[1;36m======== $msg ========`n$esc[0m"
}
function Show-Help {
    Write-Host @"
$esc[1;36m
CryptoAIS Build Script
=============================

USAGE:
    build.ps1 [-i|--install-req] [-h|--help]

OPTIONS:
    -i, --install-req   Install Python requirements before building. (requirements.txt needed)
    -h, --help          Show this help message and exit.

DESCRIPTION:
    Builds the CryptoAIS Flet Windows application.
    Optionally installs Python dependencies from requirements.txt.

EXAMPLES:
    .\build.ps1 -i
    .\build.ps1 --help

$esc[0m
"@
}

$installReq = $false
foreach ($arg in $args) {
    if ($arg -eq "-i" -or $arg -eq "--install-req") {
        $installReq = $true
        break
    }
    if ($arg -eq "-h" -or $arg -eq "--help") {
        Show-help
        exit 0
    }


}

# Step 1: Install Requirements (only if flag is set)
if ($installReq) {
    Write-Section "Installing Python Packages"
    try {
        pip install -r .\requirements.txt
        if ($LASTEXITCODE -ne 0) {
            Write-ErrorMsg "Dependency installation failed. Exiting."
            exit 1
        }
        else {
            Write-Success "All Python packages installed successfully."
        }
    }
    catch {
        Write-ErrorMsg "Exception during requirements installation: $_"
        exit 1
    }
}

Clear-Host

# Step 2: Build the Flet App
Write-Section "Building Flet Windows Application"

try {
    flet build windows `
        --project "Sola" `
        --company "Pooja ITR Center" `
        --description "Created by Krish to Automate Form-16 Generation using ITR Format. Created for POOJA ITR CENTER" `
        --product "Sola" `
        --build-version $version `
        --company "Pooja ITR Center" `
        --copyright "Copyright (C) 2025 Pooja ITR Center" `
        --exclude "release, icons, requirements.txt, README.md, certs, test, .venv" `
        --clear-cache --compile-app --compile-packages --cleanup-app --cleanup-packages --cleanup-app --module-name .\main.py

    if ($LASTEXITCODE -eq 0) {
        Write-Success "Flet Windows build completed successfully!"
    }
    else {
        Write-ErrorMsg "Flet build failed. Exit code: $LASTEXITCODE"
        exit 1
    }
}
catch {
    Write-ErrorMsg "Exception during Flet build: $_"
    exit 1
}
