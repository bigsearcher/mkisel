@echo off
REM Build both: TkinterTest (onefile) and TallyConverter (onefile).
REM Output: dist\TkinterTest.exe, dist\TallyConverter.exe

setlocal
echo ========================================
echo Building ONEFILE: test + main app
echo ========================================
echo.

echo [1/2] TkinterTest (onefile)...
call compile_test_tkinter_onefile.bat
if errorlevel 1 exit /b 1
echo.

echo [2/2] TallyConverter (onefile)...
call compile_onefile.bat
if errorlevel 1 exit /b 1
echo.

echo ========================================
echo Done. Output:
echo   dist\TkinterTest.exe
echo   dist\TallyConverter.exe
echo ========================================
endlocal
