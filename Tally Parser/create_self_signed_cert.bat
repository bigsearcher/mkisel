@echo off
REM Create a self-signed code signing certificate for testing
REM Note: Self-signed certificates will still show a warning, but less severe than unsigned
REM For production use, you need a certificate from a trusted CA (paid)

echo ========================================
echo Creating Self-Signed Code Signing Certificate
echo ========================================
echo.
echo This will create a certificate for testing purposes.
echo Windows will still show a warning, but it will be less severe.
echo.
echo For production use, you need a certificate from a trusted CA.
echo.

powershell -Command "New-SelfSignedCertificate -Type CodeSigningCert -Subject 'CN=mkisel' -CertStoreLocation Cert:\CurrentUser\My -KeyUsage DigitalSignature -FriendlyName 'mkisel Code Signing'"

if errorlevel 1 (
    echo.
    echo Failed to create certificate
    pause
    exit /b 1
)

echo.
echo ========================================
echo Certificate created successfully!
echo ========================================
echo.
echo To use this certificate for signing, run:
echo   signtool sign /n "mkisel" /fd sha256 /td sha256 /tr http://timestamp.digicert.com "dist\TallyConverter.exe"
echo.
echo To export the certificate to PFX file:
echo   1. Open certmgr.msc
echo   2. Go to Personal ^> Certificates
echo   3. Find "mkisel Code Signing" certificate
echo   4. Right-click ^> All Tasks ^> Export
echo   5. Export with private key to certificate.pfx
echo.
pause
