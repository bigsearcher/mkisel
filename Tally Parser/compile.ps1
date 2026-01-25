# PowerShell script to compile tallyconverter.py to executable using PyInstaller

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Tally Converter - Compilation Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if PyInstaller is installed
try {
    python -c "import PyInstaller" 2>$null
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller not found"
    }
} catch {
    Write-Host "PyInstaller is not installed!" -ForegroundColor Yellow
    Write-Host "Installing PyInstaller..." -ForegroundColor Yellow
    pip install pyinstaller
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Failed to install PyInstaller!" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

Write-Host ""
Write-Host "Cleaning previous build..." -ForegroundColor Yellow
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
if (Test-Path "tallyconverter.spec") { Remove-Item -Force "tallyconverter.spec" }

Write-Host ""
Write-Host "Checking for UPX (compression tool)..." -ForegroundColor Yellow
$upxPath = Get-Command upx -ErrorAction SilentlyContinue
if (-not $upxPath) {
    Write-Host "UPX not found. Compression will be skipped." -ForegroundColor Yellow
    Write-Host "To enable UPX compression, download from: https://upx.github.io/" -ForegroundColor Cyan
    Write-Host ""
} else {
    Write-Host "UPX found. Compression will be enabled." -ForegroundColor Green
    Write-Host ""
}

Write-Host "Compiling tallyconverter.py..." -ForegroundColor Yellow
Write-Host ""

# Compile with PyInstaller
$pyinstallerArgs = @(
    "--name=TallyConverter",
    "--onefile",
    "--windowed",
    "--icon=excel_crossed.ico",
    "--add-data", "utils;utils",
    "--hidden-import=openpyxl",
    "--hidden-import=pandas",
    "--hidden-import=xlrd",
    "--hidden-import=tkinter",
    "--hidden-import=win32com.client",
    "--collect-all=openpyxl",
    "--collect-all=pandas",
    "--collect-all=xlrd",
    "tallyconverter.py"
)

pyinstaller @pyinstallerArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "Compilation failed!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "Compilation completed successfully!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Check if executable exists
if (-not (Test-Path "dist\TallyConverter.exe")) {
    Write-Host "Error: Executable not found!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Sign the executable
Write-Host "Signing executable..." -ForegroundColor Yellow
Write-Host ""

# Check if signtool is available
$signtoolPath = Get-Command signtool -ErrorAction SilentlyContinue
if (-not $signtoolPath) {
    Write-Host "Warning: signtool.exe not found in PATH." -ForegroundColor Yellow
    Write-Host "Skipping code signing." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "To sign manually, use:" -ForegroundColor Cyan
    Write-Host '  signtool sign /n "mkisel" /fd sha256 /td sha256 /tr http://timestamp.digicert.com "dist\TallyConverter.exe"' -ForegroundColor Gray
    Write-Host ""
} else {
    # Check if certificate exists, if not create it automatically
    $cert = Get-ChildItem -Path Cert:\CurrentUser\My -ErrorAction SilentlyContinue | Where-Object { 
        $_.Subject -like '*CN=mkisel*' -or $_.FriendlyName -like '*mkisel*' 
    }
    
    if (-not $cert) {
        Write-Host "Certificate not found. Creating self-signed certificate..." -ForegroundColor Yellow
        Write-Host ""
        try {
            $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject 'CN=mkisel' -CertStoreLocation Cert:\CurrentUser\My -KeyUsage DigitalSignature -FriendlyName 'mkisel Code Signing' -ErrorAction Stop
            Write-Host "Certificate created successfully!" -ForegroundColor Green
            Write-Host ""
        } catch {
            Write-Host "Warning: Failed to create certificate automatically." -ForegroundColor Yellow
            Write-Host "Skipping code signing." -ForegroundColor Yellow
            Write-Host ""
        }
    }
    
    if ($cert) {
        # Sign with certificate
        Write-Host "Signing with certificate 'mkisel'..." -ForegroundColor Yellow
        $signResult = & signtool sign /n "mkisel" /fd sha256 /td sha256 /tr http://timestamp.digicert.com "dist\TallyConverter.exe" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Warning: Code signing failed." -ForegroundColor Yellow
            Write-Host ""
        } else {
            Write-Host "Code signing completed successfully!" -ForegroundColor Green
            Write-Host ""
        }
    }
}

Write-Host "========================================" -ForegroundColor Green
Write-Host "Build completed!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Executable location: dist\TallyConverter.exe" -ForegroundColor Cyan
Write-Host ""
Read-Host "Press Enter to exit"
