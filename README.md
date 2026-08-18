# ER_Predictor

`ER_Predictor` is a separate Windows one-folder application. It preserves the legacy ERTA workflow and adds the FDA-scoped ERα binding-classification workflow; it does not replace or modify `ERTA_Predictor`.

## Scope and scientific limits

| Tab | Route | Output | Limits |
| --- | --- | --- | --- |
| ERTA (startup default) | Legacy ERTA classification | Active probability/class, ERTA AD, graphs | Existing ERTA behavior is retained. |
| ERalpha | ERα binding classification | Direct-binding probability and label at the approved 0.5 threshold, route-specific AD and nearest reference | Internal historically exposed evidence; not regulatory validation. |

**Classification evidence is internal and based on historically exposed development data; it is not fresh, untouched, independent, external, temporal, population, or regulatory validation. This application is not for regulatory use.**

## Input and output

- Direct SMILES works offline for every supported route.
- CAS lookup requires network access; a lookup failure does not replace an existing valid SMILES.
- Use `templates/ERTA_KRICT_example.xlsx` for legacy ERTA input unchanged.
- Use `templates/ERTA_ERBA_example.xlsx` for the combined template. Its first worksheet, **ERBA_Input**, is the ERBA workbook input and has exactly `Row_ID`, `CAS`, `SMILES`. The **ERTA_Input** worksheet shows the legacy columns `CID`, `CAS`, `SMILES`, `label`; export that worksheet to a separate workbook when submitting it to the legacy ERTA batch flow.
- ERTA and ERalpha use separate models, training references, AD caches, graphs, and state.
- ERalpha batch results begin with `Guide`, retain machine-readable `Predictions`, `Input`, and `Metadata` sheets, and explain every not-predicted row without putting text into numeric probability columns.

Bundled assets are read-only. Normal runtime state is `%LOCALAPPDATA%\ER_Predictor\v1`; exports are `%USERPROFILE%\Documents\ER_Predictor\Exports`. Set `ER_PREDICTOR_PORTABLE=1` before launch only for an explicitly writable portable install; state and exports then use `ER_Predictor_UserData` beside the executable. Do not select an output location inside `ERTA_Predictor`.

## Development setup and run

Build on Windows x64 with Python 3.10. Dependency acquisition and release building are separate, fail-closed phases. The preparation phase resolves the direct requirements, downloads a complete Windows wheelhouse, writes exact transitive pins with wheel SHA-256 values, and verifies the resulting offline set:

```bat
py -3.10 -I prepare_build_wheelhouse.py prepare
```

Set `ER_PREDICTOR_PYTHON` to an approved Python 3.10 x64 executable only when the `py -3.10` launcher is unavailable. For source development, install `requirements.txt` in a separate environment and run `python app.py`; release builds never use that development environment.

`requirements-lock.txt` and `artifacts\ER_Predictor-wheelhouse-inventory.json` are generated evidence, not hand-maintained pins. `build_exe.bat` refuses an empty/stale lock, an incomplete wheelhouse, a hash mismatch, an online dependency fallback, or a non-3.10/non-x64 interpreter. If a released classifier requires CatBoost, the exact modeling-runtime CatBoost pin must first be added to `requirements.txt` and its PyInstaller data/native collection must be enabled; then rerun the preparation phase so the generated transitive lock and inventory prove it.

## Audited build and clean-build comparison

`build_exe.bat` fails closed before packaging when the released ERBA catalog is absent, unsafe, incomplete, duplicated, or its recorded size/SHA-256 differs. It then produces exactly:

```text
dist\ER_Predictor\ER_Predictor.exe
```

The full `dist\ER_Predictor` folder is the release unit. After the preparation phase, run:

```bat
build_exe.bat
```

Every invocation creates **two separate empty virtual environments**, installs exclusively from the verified local wheelhouse with `--require-hashes`, performs two clean PyInstaller builds, and compares normalized component, source, SBOM-style, model/static, and distribution inventories. It writes resolved requirements, wheelhouse inventory, two reproducibility manifests, the comparison receipt, and final build manifest under `artifacts\`. A mismatch fails the build and leaves `artifacts\ER_Predictor-clean-build-comparison.json` for review. Build artifacts contain no runtime dependency on an absolute research path; resources resolve from the installed collection and mutable data resolves to the writable roots above.

## Run and publication

Run `dist\ER_Predictor\ER_Predictor.exe` from the complete one-folder collection. End-user machines need no Python.

After a successful audited build, run `publish_exe.bat` with no destination argument. It refuses the immutable legacy root, stages the **entire** one-folder collection, verifies the staged executable, and atomically promotes only to:

```text
D:\research\FDA_endocrine_disruption\ER_Predictor\ER_Predictor\ER_Predictor.exe
```

Publication captures recursive old-package manifests before and after the operation and fails if they differ. `D:\research\FDA_endocrine_disruption\ERTA_Predictor` remains the rollback package and is never a source or destination for this product.
