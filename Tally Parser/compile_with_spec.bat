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
    python -m pip install pyinstaller
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

REM Compile using spec file (python -m to avoid Device Guard blocking pyinstaller.exe)
python -m PyInstaller TallyConverter.spec

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

call "%~dp0sign_exe.bat" "dist\TallyConverter.exe"
echo ========================================
echo Build completed!
echo ========================================
echo.
echo Executable location: dist\TallyConverter.exe
echo.
pause
