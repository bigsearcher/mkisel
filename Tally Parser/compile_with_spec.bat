@echo off
REM Batch script to compile using .spec file

echo ========================================
echo Tally Converter - Compilation (using .spec)
echo ========================================
echo.

REM Check if PyInstaller is installed
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo PyInstaller is not installed!
    echo Installing PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo Failed to install PyInstaller!
        pause
        exit /b 1
    )
)

echo.
echo Cleaning previous build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo Checking for UPX (compression tool)...
where upx >nul 2>&1
if errorlevel 1 (
    echo UPX not found. Compression will be skipped.
    echo To enable UPX compression, download from: https://upx.github.io/
    echo.
) else (
    echo UPX found. Compression will be enabled.
    echo.
)

echo Compiling using TallyConverter.spec...
echo.

REM Compile using spec file
pyinstaller TallyConverter.spec

if errorlevel 1 (
    echo.
    echo Compilation failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo Compilation completed successfully!
echo ========================================
echo.

REM Check if executable exists
if not exist "dist\TallyConverter.exe" (
    echo Error: Executable not found!
    pause
    exit /b 1
)

REM Sign the executable
echo Signing executable...
echo.

REM Check if signtool is available
where signtool >nul 2>&1
if errorlevel 1 (
    echo Warning: signtool.exe not found in PATH.
    echo Skipping code signing.
    echo.
    echo To sign manually, use:
    echo   signtool sign /n "mkisel" /fd sha256 /td sha256 /tr http://timestamp.digicert.com "dist\TallyConverter.exe"
    echo.
    goto :skip_signing
)

REM Check if certificate exists, if not create it automatically
powershell -Command "Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Subject -like '*CN=mkisel*' -or $_.FriendlyName -like '*mkisel*' }" >nul 2>&1
if errorlevel 1 (
    echo Certificate not found. Creating self-signed certificate...
    echo.
    powershell -Command "New-SelfSignedCertificate -Type CodeSigningCert -Subject 'CN=mkisel' -CertStoreLocation Cert:\CurrentUser\My -KeyUsage DigitalSignature -FriendlyName 'mkisel Code Signing'" >nul 2>&1
    if errorlevel 1 (
        echo Warning: Failed to create certificate automatically.
        echo Skipping code signing.
        echo.
        goto :skip_signing
    )
    echo Certificate created successfully!
    echo.
)

REM Sign with certificate
echo Signing with certificate "mkisel"...
signtool sign /n "mkisel" /fd sha256 /td sha256 /tr http://timestamp.digicert.com "dist\TallyConverter.exe" 2>nul
if errorlevel 1 (
    echo Warning: Code signing failed.
    echo.
) else (
    echo Code signing completed successfully!
    echo.
)

:skip_signing
echo ========================================
echo Build completed!
echo ========================================
echo.
echo Executable location: dist\TallyConverter.exe
echo.
pause
