@echo off
setlocal EnableExtensions DisableDelayedExpansion

for %%I in ("%~dp0.") do set "ROOT=%%~fI"
set "OUTPUT=%ROOT%\artifacts\qa-runner"
set "PYTHONNOUSERSITE=1"
set "PYTHONSAFEPATH=1"
set "PIP_NO_INDEX=1"
set "PIP_NO_CACHE_DIR=1"
set "PYTHONHASHSEED=0"
set "SOURCE_DATE_EPOCH=1786320000"
set "TZ=UTC"

if not defined ER_PREDICTOR_QA_PYTHON (
    echo Set ER_PREDICTOR_QA_PYTHON to the Python 3.10 x64 executable from the verified pinned build environment.
    exit /b 1
)
set "QA_PY=%ER_PREDICTOR_QA_PYTHON%"

"%QA_PY%" -I -c "import site,sys; import PIL,PyInstaller,pywinauto; assert sys.version_info[:2] == (3,10), sys.version; assert sys.maxsize > 2**32, 'Python must be x64'; assert not site.ENABLE_USER_SITE, 'user site must be disabled'; assert PyInstaller.__version__ == '6.10.0', PyInstaller.__version__; assert pywinauto.__version__ == '0.6.9', pywinauto.__version__" || goto :fail
if not exist "%ROOT%\qa\external_qa_runner.py" goto :missing_source
if not exist "%ROOT%\qa\ER_Predictor_QA_Runner.spec" goto :missing_source
if not exist "%OUTPUT%" mkdir "%OUTPUT%" || goto :fail

"%QA_PY%" -I -m PyInstaller "%ROOT%\qa\ER_Predictor_QA_Runner.spec" --clean --noconfirm --distpath "%OUTPUT%" --workpath "%OUTPUT%\build" || goto :fail
if not exist "%OUTPUT%\ER_Predictor_QA_Runner.exe" (
    echo Missing QA runner executable: "%OUTPUT%\ER_Predictor_QA_Runner.exe"
    goto :fail
)
"%QA_PY%" -I -c "import hashlib,json,pathlib,sys; output=pathlib.Path(sys.argv[1]); exe=output/'ER_Predictor_QA_Runner.exe'; payload={'product':'ER_Predictor_QA_Runner','executable':str(exe.name),'sha256':hashlib.sha256(exe.read_bytes()).hexdigest(),'size_bytes':exe.stat().st_size}; (output/'ER_Predictor_QA_Runner-build-receipt.json').write_text(json.dumps(payload,indent=2,sort_keys=True)+'\n',encoding='utf-8')" "%OUTPUT%" || goto :fail
echo QA runner finished: artifacts\qa-runner\ER_Predictor_QA_Runner.exe
exit /b 0

:missing_source
echo QA runner source or deterministic spec is missing.
exit /b 1

:fail
echo ER_Predictor QA runner build failed.
exit /b 1
