# PowerShell Script to Install CoolProp for LibreOffice Portable
# Run this script if you get "No module named 'CoolProp'" error

param(
    [Parameter(Mandatory=$false)]
    [string]$LibreOfficePath = "C:\Users\$env:USERNAME\AppData\Local\Programs\LibreOfficePortable"
)

Write-Host "=== CoolProp Installer for LibreOffice Portable ===" -ForegroundColor Cyan
Write-Host ""

# Find LibreOffice Python executable
$pythonExe = Get-ChildItem -Path "$LibreOfficePath\App\libreoffice\program" -Filter "python.exe" -ErrorAction SilentlyContinue

if (-not $pythonExe) {
    Write-Host "ERROR: LibreOffice Python not found at $LibreOfficePath" -ForegroundColor Red
    Write-Host "Please specify the correct path using -LibreOfficePath parameter" -ForegroundColor Yellow
    Write-Host "Example: .\install_coolprop_portable.ps1 -LibreOfficePath 'C:\Path\To\LibreOffice'" -ForegroundColor Yellow
    exit 1
}

Write-Host "Found LibreOffice Python: $($pythonExe.FullName)" -ForegroundColor Green

# Check Python version and architecture
Write-Host ""
Write-Host "Checking Python version..." -ForegroundColor Cyan
$pythonVersion = & $pythonExe.FullName -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"
$pythonArch = & $pythonExe.FullName -c "import struct; print('64-bit' if struct.calcsize('P') * 8 == 64 else '32-bit')"
Write-Host "Python Version: $pythonVersion" -ForegroundColor Green
Write-Host "Architecture: $pythonArch" -ForegroundColor Green

# Determine wheel filename pattern
$platform = if ($pythonArch -eq "32-bit") { "win32" } else { "win_amd64" }
$pyVersion = $pythonVersion.Replace(".", "")

Write-Host ""
Write-Host "Downloading CoolProp..." -ForegroundColor Cyan

# Create temp directory
$tempDir = "$env:TEMP\coolprop_install_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

# Download CoolProp wheel
try {
    $downloadCmd = "pip download CoolProp --dest `"$tempDir`" --only-binary :all: --platform $platform --python-version $pyVersion --no-deps"
    Write-Host "Running: $downloadCmd" -ForegroundColor Gray
    Invoke-Expression $downloadCmd
    
    if ($LASTEXITCODE -ne 0) {
        throw "pip download failed"
    }
} catch {
    Write-Host "ERROR: Failed to download CoolProp" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

# Find the downloaded wheel
$wheelFile = Get-ChildItem -Path $tempDir -Filter "coolprop-*.whl" | Select-Object -First 1

if (-not $wheelFile) {
    Write-Host "ERROR: CoolProp wheel file not found in $tempDir" -ForegroundColor Red
    exit 1
}

Write-Host "Downloaded: $($wheelFile.Name)" -ForegroundColor Green

# Find site-packages directory
$pythonCoreDir = Get-ChildItem -Path "$LibreOfficePath\App\libreoffice\program" -Filter "python-core-*" | 
    Sort-Object -Descending | Select-Object -First 1

if (-not $pythonCoreDir) {
    Write-Host "ERROR: python-core directory not found" -ForegroundColor Red
    exit 1
}

$sitePackages = Join-Path $pythonCoreDir.FullName "lib\site-packages"

if (-not (Test-Path $sitePackages)) {
    Write-Host "ERROR: site-packages directory not found at $sitePackages" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Installing CoolProp to: $sitePackages" -ForegroundColor Cyan

# Extract wheel (it's a zip file)
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($wheelFile.FullName, $sitePackages)
    Write-Host "Installation successful!" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to extract wheel" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

# Verify installation
Write-Host ""
Write-Host "Verifying installation..." -ForegroundColor Cyan
try {
    $coolpropVersion = & $pythonExe.FullName -c "import CoolProp; print(CoolProp.__version__)"
    if ($LASTEXITCODE -eq 0) {
        Write-Host "SUCCESS! CoolProp version $coolpropVersion installed" -ForegroundColor Green
        Write-Host ""
        Write-Host "Next steps:" -ForegroundColor Cyan
        Write-Host "1. Install the CoolPropLibre.oxt extension in LibreOffice" -ForegroundColor Yellow
        Write-Host "2. Restart LibreOffice Calc" -ForegroundColor Yellow
        Write-Host "3. Use CPROP() and CPROPHA() functions in your spreadsheet" -ForegroundColor Yellow
    } else {
        throw "Import test failed"
    }
} catch {
    Write-Host "ERROR: CoolProp import test failed" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

# Cleanup
Write-Host ""
Write-Host "Cleaning up temporary files..." -ForegroundColor Cyan
Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Installation complete!" -ForegroundColor Green
