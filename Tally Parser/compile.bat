@echo off
REM Batch script to compile tallyconverter.py to executable using PyInstaller

echo ========================================
echo Tally Converter - Compilation Script
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
if exist tallyconverter.spec del /q tallyconverter.spec

echo.
echo Checking for UPX (compression tool)...
set "UPX_DIR=D:\Archives\Programming\upx-5.1.0-win64"
if exist "%UPX_DIR%\upx.exe" (
    set "PATH=%UPX_DIR%;%PATH%"
    echo UPX found at %UPX_DIR%. Compression will be enabled.
    echo.
    goto :upx_done
)
where upx >nul 2>&1
if errorlevel 1 (
    echo UPX not found. Compression will be skipped.
    echo To enable UPX compression, download from: https://upx.github.io/
    echo.
) else (
    echo UPX found in PATH. Compression will be enabled.
    echo.
)
:upx_done

echo Compiling tallyconverter.py...
echo.

REM Compile with PyInstaller (python -m to avoid Device Guard blocking pyinstaller.exe)
python -m PyInstaller --name="TallyConverter" ^
    --onefile ^
    --windowed ^
    --icon="tallyconverter.ico" ^
    --add-data "utils;utils" ^
    --hidden-import="openpyxl" ^
    --hidden-import="pandas" ^
    --hidden-import="xlrd" ^
    --hidden-import="tkinter" ^
    --hidden-import="win32com.client" ^
    --hidden-import="win32timezone" ^
    --hidden-import="secrets" ^
    --hidden-import="numpy.random" ^
    --exclude-module="pytest" ^
    --exclude-module="unittest" ^
    --exclude-module="IPython" ^
    --exclude-module="jupyter" ^
    --exclude-module="matplotlib" ^
    --exclude-module="scipy" ^
    --exclude-module="PIL" ^
    --exclude-module="Pillow" ^
    --exclude-module="requests" ^
    --exclude-module="sqlite3" ^
    --exclude-module="asyncio" ^
    --exclude-module="multiprocessing" ^
    --collect-all="openpyxl" ^
    --collect-all="xlrd" ^
    tallyconverter.py

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
