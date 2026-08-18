@echo off
setlocal EnableExtensions DisableDelayedExpansion
for %%I in ("%~dp0.") do set "ROOT=%%~fI"
set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 7\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 7\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" (
  echo Inno Setup compiler was not found.
  echo Install the signed x64 compiler from https://jrsoftware.org/isdl.php
  exit /b 1
)
if not exist "%ROOT%\dist\ER_Predictor\ER_Predictor.exe" (
  echo Missing packaged application. Run build_exe.bat first.
  exit /b 1
)
if exist "%ROOT%\installer" rmdir /s /q "%ROOT%\installer" || exit /b 1
"%ISCC%" "%ROOT%\ER_Predictor.iss" || exit /b 1
if not exist "%ROOT%\installer\ER_Predictor_Setup_x64.exe" exit /b 1
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$p='%ROOT%\installer\ER_Predictor_Setup_x64.exe'; $h=(Get-FileHash -Algorithm SHA256 $p).Hash.ToLower(); $s=(Get-Item $p).Length; [pscustomobject]@{product='ER Predictor';installer='ER_Predictor_Setup_x64.exe';size_bytes=$s;sha256=$h;architecture='x64';python_runtime='bundled';requires_system_python=$false} | ConvertTo-Json | Set-Content -Encoding UTF8 '%ROOT%\installer\ER_Predictor_Setup_x64.sha256.json'" || exit /b 1
echo Installer created: installer\ER_Predictor_Setup_x64.exe
exit /b 0
