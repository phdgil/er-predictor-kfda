# ER Predictor: code intention and ERBA integration contract

## Purpose and baseline

This document records the existing ERTA code intent before implementation of the 2026 ERBA work. It is the first product-source change. The pre-change source and installed package were hashed first in `D:\research\FDA_endocrine_disruption\ER_Predictor_evidence\baseline\prechange_manifest.json`.

The existing package at `D:\research\FDA_endocrine_disruption\ERTA_Predictor` is an immutable behavioral oracle. The integrated application is a separate product and must never overwrite or modify it.

## Existing code intention

The current code is a Tkinter desktop application for **estrogen receptor transcriptional activation (ERTA) classification**, distributed as a PyInstaller one-folder Windows application.

1. `app.py` establishes the source directory as `project_root`, creates `gui.main_window.MainWindow`, and starts Tk.
2. `MainWindow` owns one `KerasPredictor`, one ERTA applicability-domain calculator, and separate single/batch pages.
3. Startup loads `models/best_seed831053.keras` and fits ERTA AD from `data/training_reference_RDKit_FP.xlsx`.
4. Users provide SMILES directly or resolve CAS through PubChem; direct SMILES works offline.
5. `core.fingerprint` computes a legacy 2,048-bit `Chem.RDKFingerprint`. Existing blank/invalid SMILES behavior produces an all-zero vector; stricter ERBA validation must not silently alter this frozen ERTA behavior.
6. `KerasPredictor` loads the model directly or reconstructs its fixed dense architecture for Keras-version compatibility, then assigns Positive/Negative by the larger softmax probability.
7. Single prediction renders structure, probabilities, ERTA AD, nearest training reference, and ERTA plots.
8. Batch prediction reads Excel, preserves identifiers, adds canonical SMILES/validity/probabilities/class, evaluates ERTA AD, writes a workbook, and generates ERTA graphs.
9. `core.paths.resource_path` resolves PyInstaller resources through `_MEIPASS`; current exports default below `project_root/output`.
10. `app.spec` bundles the Keras model, AD data, templates, TensorFlow, RDKit, scikit-learn, SciPy, Pillow, and HDF5.

## ERTA invariants

- ERTA remains startup-selected and keeps its Keras model, legacy fingerprint, probability mapping, labels, decision rule, AD, graphs, examples, batch schema, and filename behavior.
- Existing model/AD browsing remains ERTA-only.
- ERTA must not load ERBA joblibs or display direct-binding claims.
- New ERBA invalid-structure handling must not change historical ERTA behavior.
- The old packaged `ERTA_Predictor` tree remains unchanged.

## New product intention

Build a separate sibling product named **ER Predictor**, with executable exactly `ER_Predictor.exe` and two FDA-facing top-level tabs:

- **ERTA:** the current single/batch experience, preserved rather than generalized.
- **ERalpha:** independent single/batch state for direct ERα binding classification.

| ERBA task | Subtype | Output | Enablement |
|---|---|---|---|
| Classification | ERα | non-binding/binding probabilities and label | Nested scaffold-group CV MCC winner passes preregistered internal gates |
| Classification | ERβ | research artifact only; not exposed or shipped in the FDA application | Excluded by FDA feedback |
| IC50 regression | ERα/ERβ | research artifact only; not exposed or shipped in the FDA application | FDA-facing scope is classification only |

ERBA means **direct receptor binding**, not transcriptional activation, agonism, antagonism, signaling, coactivator recruitment, or general endocrine disruption.

## Classification evidence contract

All current classification rows and prior feature/model results are historically exposed. The scientific owner explicitly accepts this limitation for a rapid-feedback release.

- Partition deterministically by normalized identity and Bemis–Murcko scaffold into development and sealed internal-resplit data.
- Perform nested scaffold-group CV on development data, fitting all learned preprocessing inside folds.
- Select one complete pipeline per subtype solely by highest mean outer-fold MCC with preregistered ties.
- Freeze/hash the development-fit winner before one evaluation on the sealed resplit.
- Never rerank, alter threshold/preprocessing, substitute a runner-up, or refit after evaluation.

The result is an **internal resplit/sensitivity evaluation**—not fresh, untouched, independent, external, temporal, population, or regulatory validation. This caveat must appear in manifests, reports, GUI, workbook metadata, documentation, tests, and release evidence.
### Accepted V7 classification release

The preregistered development-only search completed all 112 subtype/feature/algorithm candidates with five finite outer folds, no recorded warning or failure, exact source-manifest hashes, and no sealed-resplit access. ERα selected RDKit fingerprint (2,048 bits) with SVM (nested MCC `0.8999 ± 0.0343`; internal sensitivity MCC `0.8381`; frozen SHA-256 `db3f2c369cfc9f2a08ecc985424879b70b82f3a4bc8d11a5e2f94ab8c7fe96ea`). ERβ selected Avalon fingerprint (2,048 bits) with random forest (nested MCC `0.9186 ± 0.0224`; internal sensitivity MCC `0.9636`; frozen SHA-256 `5d01aa2729346a0459280dd7a1ea0327731434a24ce4ae99b2e27daeab849d76`). The one-time evaluator preserved the exact frozen bytes and issued external evaluator-bound release manifests. The authoritative report is `D:\research\FDA_endocrine_disruption\data\normalized_erba\er_regression_classification_model_report_final.html`.

## Task-specific ERBA structure policies

### Classification

Reject blank SMILES, wildcard `[*]`, pipe/CXSMILES syntax, invalid molecules, non-carbon structures, and structures retaining metal after largest-fragment/salt handling. Produce canonical isomeric model SMILES. Use the exact selected pipeline feature generator and fail closed on policy, width, dtype, class-order, or artifact mismatch.

### ERα IC50 regression

Treat preprocessing independently. Use Avalon 2,048 bits only after raw-input-to-`rdkit_parent_smiles` equivalence is proved. Use serialized `production.selector` and `production.estimator`; report `pIC50 = -log10(IC50 [mol/L])` and `IC50_nM = 10^(9-pIC50)`. Keep disabled if parity cannot be recovered.

## Model integrity

ERBA loads only fixed bundled artifacts. Before joblib loading: confine the regular non-reparse file to the bundled model directory; verify size and SHA-256 against a frozen allowlist or signed manifest; then validate task, subtype, policy, width, dictionary keys, pipeline structure, and class order. Arbitrary user-selected ERBA joblibs are prohibited.

Without release signing, this detects accidental corruption or model-only replacement while the executable remains trusted; it does not protect against replacement of both executable and assets.

## Input/output contract

A versioned contract shared by adapters, GUI, export, docs, and tests defines exact columns and status codes.

- Nonblank SMILES takes precedence; CAS is fallback only.
- Batch order and passthrough identifiers are preserved.
- Invalid rows remain with stable technical error codes and blank numeric predictions. They also receive user-facing `Result_Status`, `Reason_Category`, `Reason_Description`, and `Recommended_Action` fields; blank numeric cells never stand alone without an explanation.
- Outputs include task/subtype, raw/model SMILES, predictions, model/policy IDs and hashes, row status, and evidence caveat.
- FDA-facing ERalpha exports classification only.
- ERalpha uses the same AD algorithm and UI structure as ERTA, but only with its own route-specific training reference and cache.
- Writes are atomic and never silently overwrite existing files.

## Runtime, GUI, and release boundaries

- Bundled resources and install root are read-only.
- Logs/cache/state use `%LOCALAPPDATA%\ER_Predictor`; exports use an explicit user-owned directory and never silently redirect.
- Optional portable writes require an explicit mode and successful probe.
- ERTA and ERalpha tabs own separate controls/state. Background requests snapshot workflow/task/subtype/model/policy/input and discard stale completions.
- Build from an isolated pinned Windows environment with resolved dependencies, hashes, SBOM/build manifest, audited assets, and target `ER_Predictor.exe`.
- Publish only as a new sibling package. Never write beneath the old `ERTA_Predictor` root.

## Completion gates

1. This file is the only product-source delta against the external pre-change manifest before implementation.
2. Existing source/package ERTA behavior is frozen with exact schemas and justified tolerances.
3. Nested grouped-CV selection uses only mean MCC; the sealed resplit is evaluated once and cannot influence selection.
4. All classification evidence is labeled internal and historically exposed.
5. ERBA fixtures prove raw input → model SMILES → feature hash → prediction.
6. FDA-facing routes are limited to ERTA and ERα binding classification; ERβ and regression routes are absent from the UI, shortcuts, release catalog, packaged assets, and installer.
7. Unit/integration tests cover contracts, preprocessing, integrity, class order, conversion, invalid structures, paths, stale requests, batch order/status, and no ERTA AD/graphs for ERBA.
8. Packaged Windows automation covers ERTA default, ERα classification, invalid/mixed batch, ERTA→ERalpha→ERTA isolation, output capture, progress reporting, and shutdown.
9. Source and packaged predictions match independent golden fixtures.
10. The final artifact is a separate `ER_Predictor.exe`; the old package still matches its pre-change manifest.
11. Cleaner, architecture/product/code, QA/red-team, and terminal critic gates are clean.

## Non-goals

No old-package replacement, ERTA retraining, combined ERTA/ERBA score, general endocrine-disruption claim, FDA-facing ERβ/regression route, external/regulatory claim from current classification data, or arbitrary ERBA model loading.

## FDA feedback amendment: executable implementation steps

This amendment supersedes earlier UI/release scope where it conflicts. The FDA test workbook contained 504 rows: 448 predicted and 56 intentionally not predicted (35 PubChem 404/no recovered SMILES, 18 no-carbon inorganic structures, 2 invalid CAS check digits, and 1 unsupported metal-containing structure).

1. **Scope lock:** retain ERTA and ERalpha only. Remove ERbeta from tabs, shortcuts, native QA, released catalog, packaged model/manifest/AD assets, installer, and user documentation. Keep research artifacts outside the FDA distribution.
2. **Progress:** add a determinate batch progress bar and text for read, CAS/SMILES resolution, preprocessing/prediction, workbook writing, and terminal success/failure. Worker threads communicate progress through Tk `after`; no widget is updated directly from a worker.
3. **Explain non-predictions:** keep probability/label columns blank and numeric-safe when no prediction exists, and add bilingual-friendly result status, category, description, and recommended action columns.
4. **Workbook guidance:** make `Guide` the first worksheet and explain `Predictions`, `Input`, and `Metadata`; include totals and reason counts. Preserve exact machine-readable headers on data sheets. Add filters, frozen headers, widths, wrapped text, and yellow/red outcome highlighting.
5. **Compatibility:** preserve ERTA behavior, ERalpha model/preprocessing/AD calculations, input precedence, row order, passthrough identifiers, atomic publication, read-only install boundaries, and immutable old ERTA package.
6. **Validation:** run unit/integration tests, exercise the FDA-provided 504-row workbook, verify progress reaches 100% or a terminal failure, build twice, publish, install to a clean directory, run packaged ERTA/ERalpha QA, uninstall, and re-audit all 10,280 old ERTA files.
