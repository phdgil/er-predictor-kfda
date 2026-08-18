@echo off
setlocal
cd /d %~dp0

set EXE_PATH=dist\ERTA_Predictor\ERTA_Predictor.exe

if not exist "%EXE_PATH%" (
    echo EXE was not found:
    echo %CD%\%EXE_PATH%
    echo.
    echo Run build_exe.bat first.
    pause
    exit /b 1
)

echo Running:
echo %CD%\%EXE_PATH%
echo.

"%EXE_PATH%"

echo.
echo Program exited with code %errorlevel%.
pause
