@echo off
setlocal EnableExtensions DisableDelayedExpansion

for %%I in ("%~dp0.") do set "ROOT=%%~fI"
set "DIST=%ROOT%\dist\ER_Predictor"
set "MANIFEST_DIR=%ROOT%\artifacts"
set "WHEELHOUSE=%MANIFEST_DIR%\ER_Predictor-wheelhouse"
set "WHEELHOUSE_INVENTORY=%MANIFEST_DIR%\ER_Predictor-wheelhouse-inventory.json"
set "FIRST_VENV=%TEMP%\ER_Predictor-build-first.venv"
set "SECOND_VENV=%TEMP%\ER_Predictor-build-second.venv"
set "PYTHONNOUSERSITE=1"
set "PYTHONSAFEPATH=1"
set "PIP_NO_INDEX=1"
set "PIP_NO_CACHE_DIR=1"
set "PIP_DISABLE_PIP_VERSION_CHECK=1"
set "PYTHONHASHSEED=0"
set "SOURCE_DATE_EPOCH=1786320000"
set "TZ=UTC"

if defined ER_PREDICTOR_PYTHON (
    set "BUILD_PY="%ER_PREDICTOR_PYTHON%""
) else (
    where py >nul 2>nul || (
        echo Python launcher 'py' was not found. Set ER_PREDICTOR_PYTHON to Python 3.10 x64.
        exit /b 1
    )
    set "BUILD_PY=py -3.10"
)

rem Phase 1: a separately invoked prepare operation may download wheels; builds only verify
rem the resulting complete local wheelhouse and hash lock, with no online fallback.
%BUILD_PY% -I -c "import site,sys; assert sys.version_info[:2] == (3, 10), sys.version; assert sys.maxsize > 2**32, 'Python must be x64'; assert not site.ENABLE_USER_SITE, 'user site must be disabled'" || goto :fail
%BUILD_PY% -I "%ROOT%\prepare_build_wheelhouse.py" verify || goto :fail
if not exist "%WHEELHOUSE_INVENTORY%" (
    echo Missing verified wheelhouse inventory: "%WHEELHOUSE_INVENTORY%"
    goto :fail
)
set "PIP_REQUIRE_VIRTUALENV=true"
if exist "%WHEELHOUSE%" goto :wheelhouse_ready
echo Missing verified wheelhouse: "%WHEELHOUSE%"
goto :fail

:wheelhouse_ready
if not exist "%MANIFEST_DIR%" mkdir "%MANIFEST_DIR%" || goto :fail

rem Fail before either build if catalog-listed ERBA assets are absent or altered.
%BUILD_PY% -I -c "import hashlib,json,pathlib,sys; root=pathlib.Path(sys.argv[1]); catalog=root/'models'/'erba'/'catalog.v2.json'; payload=json.loads(catalog.read_text(encoding='utf-8')); assert payload.get('schema_version') == 2, 'catalog schema_version must be 2'; routes=[route for route in payload.get('routes',[]) if isinstance(route,dict) and route.get('release_status') == 'released']; assert routes, 'catalog has no released routes'; artifacts=[route.get('artifact',route) for route in routes]; paths=[pathlib.Path(str(item['relative_path'])) for item in artifacts]; assert len(paths) == len(set(paths)), 'duplicate ERBA artifact path'; assert all(not path.is_absolute() and '..' not in path.parts and path != pathlib.Path('.') for path in paths), 'unsafe ERBA artifact path'; files=[catalog.parent/path for path in paths]; assert all(path.is_file() for path in files), 'catalogued ERBA artifact missing'; assert all(path.stat().st_size == int(item['size_bytes']) and hashlib.sha256(path.read_bytes()).hexdigest().lower() == str(item['sha256']).lower() for path,item in zip(files,artifacts)), 'catalogued ERBA artifact integrity failure'; print('ERBA source artifact audit passed:', ', '.join(str(path) for path in paths))" "%ROOT%" || goto :fail

call :create_venv "%FIRST_VENV%" || goto :fail
call :build_one "%FIRST_VENV%\Scripts\python.exe" "first" || goto :fail
call :create_venv "%SECOND_VENV%" || goto :fail
call :build_one "%SECOND_VENV%\Scripts\python.exe" "second" || goto :fail

rem Fail closed with: independent clean-build normalized inventories differ
%BUILD_PY% -I "%ROOT%\build_reproducibility_manifest.py" compare --first "%MANIFEST_DIR%\ER_Predictor-reproducibility-manifest.first.json" --second "%MANIFEST_DIR%\ER_Predictor-reproducibility-manifest.second.json" --output "%MANIFEST_DIR%\ER_Predictor-clean-build-comparison.json" || goto :fail
copy /y "%MANIFEST_DIR%\ER_Predictor-reproducibility-manifest.second.json" "%MANIFEST_DIR%\ER_Predictor-reproducibility-manifest.json" >nul || goto :fail
copy /y "%MANIFEST_DIR%\ER_Predictor-resolved-requirements.second.txt" "%MANIFEST_DIR%\ER_Predictor-resolved-requirements.txt" >nul || goto :fail

"%SECOND_VENV%\Scripts\python.exe" -I -c "import hashlib,json,pathlib,platform,sys; root=pathlib.Path(sys.argv[1]); dist=root/'dist'/'ER_Predictor'; files=sorted((path for path in dist.rglob('*') if path.is_file()),key=lambda path:str(path)); inventory=[{'path':str(path.relative_to(root)).replace('\\','/'),'size_bytes':path.stat().st_size,'sha256':hashlib.sha256(path.read_bytes()).hexdigest()} for path in files]; manifest={'product':'ER_Predictor','executable':'dist/ER_Predictor/ER_Predictor.exe','python':sys.version,'platform':platform.platform(),'catalog':'dist/ER_Predictor/_internal/models/erba/catalog.v2.json','inventory':inventory}; pathlib.Path(sys.argv[2]).write_text(json.dumps(manifest,indent=2,sort_keys=True)+'\n',encoding='utf-8')" "%ROOT%" "%MANIFEST_DIR%\ER_Predictor-build-manifest.json" || goto :fail

echo Build finished: dist\ER_Predictor\ER_Predictor.exe
echo Release receipts written under: artifacts
exit /b 0

:create_venv
if exist "%~1" rmdir /s /q "%~1" || exit /b 1
%BUILD_PY% -I -m venv "%~1" || exit /b 1
"%~1\Scripts\python.exe" -I -c "import site,sys; assert sys.version_info[:2] == (3, 10), sys.version; assert sys.maxsize > 2**32, 'Python must be x64'; assert sys.prefix != sys.base_prefix, 'empty isolated venv required'; assert not site.ENABLE_USER_SITE, 'user site must be disabled'" || exit /b 1
"%~1\Scripts\python.exe" -I -m pip install --require-virtualenv --no-index --find-links "%WHEELHOUSE%" --only-binary=:all: --require-hashes --no-cache-dir -r "%ROOT%\requirements-lock.txt" || exit /b 1
"%~1\Scripts\python.exe" -I -m pip check || exit /b 1
exit /b 0

:build_one
set "VENV_PY=%~1"
set "BUILD_LABEL=%~2"
if exist "%ROOT%\build" rmdir /s /q "%ROOT%\build" || exit /b 1
if exist "%ROOT%\dist" rmdir /s /q "%ROOT%\dist" || exit /b 1
"%VENV_PY%" -I -c "import PyInstaller,h5py,joblib,openpyxl,rdkit,requests,scipy,sklearn,tensorflow,xgboost; print('Python', __import__('sys').version); print('PyInstaller', PyInstaller.__version__); print('TensorFlow', tensorflow.__version__); print('scikit-learn', sklearn.__version__); print('joblib', joblib.__version__); print('XGBoost', xgboost.__version__); print('RDKit', rdkit.__version__); print('scipy', scipy.__version__); print('h5py', h5py.__version__); print('openpyxl', openpyxl.__version__); print('requests', requests.__version__)" || exit /b 1
"%VENV_PY%" -I -m PyInstaller "%ROOT%\app.spec" --clean --noconfirm --distpath "%ROOT%\dist" --workpath "%ROOT%\build" || exit /b 1
"%VENV_PY%" -I -c "import json,pathlib,sys; root=pathlib.Path(sys.argv[1]); dist=root/'dist'/'ER_Predictor'; catalog=dist/'_internal'/'models'/'erba'/'catalog.v2.json'; assert (dist/'ER_Predictor.exe').is_file(), 'missing dist/ER_Predictor/ER_Predictor.exe'; payload=json.loads(catalog.read_text(encoding='utf-8')); assert payload.get('schema_version') == 2, 'catalog schema_version must be 2'; artifacts=[route.get('artifact',route) for route in payload['routes'] if isinstance(route,dict) and route.get('release_status') == 'released']; expected={pathlib.Path(str(item['relative_path'])) for item in artifacts}; assert expected and all((catalog.parent/path).is_file() for path in expected), 'missing catalogued ERBA asset'; found={path.relative_to(catalog.parent) for path in catalog.parent.rglob('*.joblib')}; assert found == {path for path in expected if path.suffix.lower() == '.joblib'}, 'unlisted or missing ERBA joblib'; print('ER_Predictor distribution artifact audit passed')" "%ROOT%" || exit /b 1
"%VENV_PY%" -I -m pip freeze --all > "%MANIFEST_DIR%\ER_Predictor-resolved-requirements.%BUILD_LABEL%.txt" || exit /b 1
"%VENV_PY%" -I "%ROOT%\build_reproducibility_manifest.py" manifest --root "%ROOT%" --output "%MANIFEST_DIR%\ER_Predictor-reproducibility-manifest.%BUILD_LABEL%.json" || exit /b 1
exit /b 0

:fail
echo ER_Predictor build failed; no package outside this repository was targeted.
exit /b 1
