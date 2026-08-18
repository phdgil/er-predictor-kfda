@echo off
setlocal EnableExtensions DisableDelayedExpansion

for %%I in ("%~dp0.") do set "ROOT=%%~fI"
set "PYTHONNOUSERSITE=1"
set "PYTHONSAFEPATH=1"

rem Publication is deliberately fixed inside publish_exe.py. No destination argument
rem is accepted, and the approved pre-change ERTA manifest is a mandatory gate.
if not "%~1"=="" (
    echo publish_exe.bat accepts no destination or positional arguments.
    exit /b 1
)
if defined ER_PREDICTOR_PYTHON (
    set "PUBLISH_PY="%ER_PREDICTOR_PYTHON%""
) else (
    where py >nul 2>nul || (
        echo Python launcher 'py' was not found. Set ER_PREDICTOR_PYTHON to Python 3.10 x64.
        exit /b 1
    )
    set "PUBLISH_PY=py -3.10"
)

%PUBLISH_PY% -I "%ROOT%\publish_exe.py" || goto :fail
exit /b 0

:fail
echo Publication failed; the immutable ERTA package was not modified.
exit /b 1
