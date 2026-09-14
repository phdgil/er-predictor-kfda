# ER Predictor: code intention and ERBA integration contract

## Purpose and baseline

This document records the existing ERTA code intent before implementation of the 2026 ERBA work. It is the first product-source change. The pre-change source and installed package were hashed first in `D:\research\FDA_endocrine_disruption\ER_Predictor_evidence\baseline\prechange_manifest.json`.

The active source and working tree are at `D:\research\FDA_endocrine_disruption\ER_Predictor_Code`. The archived legacy package at `D:\research\FDA_endocrine_disruption\_archive\legacy_apps\ERTA_Predictor` is an immutable behavioral oracle. The integrated application is a separate product and must never overwrite or modify the archived package.

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
9. `core.paths.resource_path` resolves PyInstaller resources through `_MEIPASS`; in the pre-integration baseline, exports defaulted below `project_root/output`.
10. `app.spec` bundles the Keras model, AD data, templates, TensorFlow, RDKit, scikit-learn, SciPy, Pillow, and HDF5.

## ERTA invariants

- ERTA remains startup-selected and keeps its Keras model, legacy fingerprint, probability mapping, labels, decision rule, AD, graphs, and batch schema. The authoritative shared-example and publication amendment below controls the distributed example and the `ERTA_` result filename prefix as well as destinations and collision suffixes.
- Existing model/AD browsing remains ERTA-only.
- ERTA must not load ERBA joblibs or display direct-binding claims.
- New ERBA invalid-structure handling must not change historical ERTA behavior.
- The archived legacy `ERTA_Predictor` tree remains unchanged.

## New product intention

Build the separate product named **ER Predictor**, with executable exactly `ER_Predictor.exe` and two FDA-facing top-level tabs:

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
- Invalid rows remain in `Predictions` with blank probabilities, numeric prediction, label, and AD values. The one-to-one row in `Diagnostics` retains the stable technical status, user-facing reason, description, and recommended action, so blank numeric cells never stand alone without an explanation.
- The internal result API continues to expose task/subtype, raw/model SMILES, probabilities, label, model/policy identity, technical status, and evidence caveat. The Excel v3 adapter deliberately separates that API into a plain primary result surface, row-level `Diagnostics`, and execution-level `Metadata` rather than duplicating every internal field on every primary row.
- FDA-facing ERalpha exports classification only.
- ERalpha uses the same AD algorithm and UI structure as ERTA, but only with its own route-specific training reference and cache.
- ERalpha workbooks use contract `erba.binding.classification.excel.v3` and open on `Predictions`, followed by `Guide`, `Diagnostics`, `Input`, and `Metadata`. `Predictions` starts with collision-safe original input columns in their original order. Its trusted suffix is exactly `CAS`, `SMILES`, `Canonical_SMILES`, `Mol_valid`, `Probability_Negative_0`, `Probability_Positive_1`, `Prediction`, `Prediction_label`, the nine route-specific AD columns, `PubChem_CID`, and `PubChem_status`. The untouched source columns remain on `Input`.
- `Prediction` is numeric `0`/`1`; `Prediction_label` is `Non-binding`/`Binding`. In this ERalpha workbook class 0 and class 1 mean direct receptor non-binding and binding. ERTA's similarly named Negative/Positive probability columns mean transcriptional activation outcomes, not ERalpha binding, so the endpoints must not be interpreted interchangeably.
- `Diagnostics` has exact columns `Row_ID`, `CAS`, `row_index`, `Status_Code`, `Status_Message`, `SMILES_Provenance`, `Result_Status`, `Reason_Category`, `Reason_Description`, and `Recommended_Action`. Rows are one-to-one with `Predictions` and preserve order. `Metadata` holds the constant workflow/task/subtype, decision rule, model ID/SHA-256, preprocessing policy, protocol/source/split/CV/report/caveat hashes, evidence scope, and caveat once per run; those values are not repeated on primary rows.
- Workbook consumers must use the v3 `Predictions` names, resolve row status/reasons from the same-position `Diagnostics` row (or its zero-based `row_index`), and read model/run provenance once from `Metadata`. The former lowercase `raw_smiles`, `model_smiles`, `non_binding_probability`, `binding_probability`, and `binding_label` columns and per-row model/hash/caveat fields have no compatibility aliases on the primary worksheet.
- ERalpha workbook sheets use plain pandas/openpyxl cell formatting without colored headers, result highlighting, wrapped-text decoration, or fitted display widths. For the same cell kind, the primary `Predictions` formatting signature is exactly the one generated for the plain ERTA worksheet. This presentation parity does not equate endpoint semantics and does not alter `Guide`, row-level `Diagnostics`, untouched `Input`, or release/provenance `Metadata`.
- The ERalpha AD columns are, in order, `AD`, `AD_MeanDistance`, `AD_DistanceThreshold`, `AD_Distance_InDomain`, `AD_SimilarityMax`, `AD_SimilarityThreshold`, `AD_Similarity_InDomain`, `AD_PC1`, and `AD_PC2`. Their values and graph artifacts use only the ERalpha route's training reference, cache, and batch results.
- `PubChem_CID` and `PubChem_status` contain only values captured from the CAS lookup response used for that row. Direct-SMILES rows and absent response fields stay blank; the exporter neither invents identifiers nor performs a second lookup to populate provenance.
- ERTA and ERalpha batch workbooks are atomically published beside the selected input workbook. They never overwrite the input or an existing result; an available collision-safe filename is allocated instead. ERTA uses `ERTA_<input-stem>_<model-name>_prediction.xlsx`, then `_2`, `_3`, and so on before the extension.
- A protected or unwritable input parent is rejected before prediction with clear guidance to copy the input workbook to a writable folder outside the application files and select that copy. Batch publication never silently redirects to another directory.

## Runtime, GUI, and release boundaries

- Bundled resources and install root are read-only.
- Logs/cache/state and single-prediction plots/exports use writable external user roots under `%LOCALAPPDATA%\ER_Predictor` and `%USERPROFILE%\Documents\ER_Predictor\Exports`. Batch workbooks do not use an independent export root or directory chooser; their sole destination is the selected input workbook's parent.
- Optional portable writes require an explicit mode and successful probe.
- ERTA and ERalpha tabs own separate controls/state. Background requests snapshot workflow/task/subtype/model/policy/input and discard stale completions.
- Build from an isolated pinned Windows environment with resolved dependencies, hashes, SBOM/build manifest, audited assets, and target `ER_Predictor.exe`.
- Publish only as a separate product. Never write beneath the archived legacy `D:\research\FDA_endocrine_disruption\_archive\legacy_apps\ERTA_Predictor` root.

## Completion gates

1. This file is the only product-source delta against the external pre-change manifest before implementation.
2. Existing source/package ERTA behavior is frozen with exact schemas and justified tolerances.
3. Nested grouped-CV selection uses only mean MCC; the sealed resplit is evaluated once and cannot influence selection.
4. All classification evidence is labeled internal and historically exposed.
5. ERBA fixtures prove raw input → model SMILES → feature hash → prediction.
6. FDA-facing routes are limited to ERTA and ERα binding classification; ERβ and regression routes are absent from the UI, shortcuts, release catalog, packaged assets, and installer.
7. Unit/integration tests cover contracts, preprocessing, integrity, class order, conversion, invalid structures, paths, stale requests, batch order/status, and strict route isolation for ERTA and ERalpha AD fields/graphs.
8. Packaged Windows automation covers ERTA default, ERα classification, invalid/mixed batch, ERTA→ERalpha→ERTA isolation, output capture, progress reporting, and shutdown.
9. Source and packaged predictions match independent golden fixtures.
10. The final artifact is a separate `ER_Predictor.exe`; the archived legacy package still matches its pre-change manifest.
11. Cleaner, architecture/product/code, QA/red-team, and terminal critic gates are clean.

## Non-goals

No archived legacy-package replacement, ERTA retraining, combined ERTA/ERBA score, general endocrine-disruption claim, FDA-facing ERβ/regression route, external/regulatory claim from current classification data, or arbitrary ERBA model loading.

## FDA feedback amendment: executable implementation steps

This amendment supersedes earlier UI/release scope where it conflicts. The FDA test workbook contained 504 rows: 448 predicted and 56 intentionally not predicted (35 PubChem 404/no recovered SMILES, 18 no-carbon inorganic structures, 2 invalid CAS check digits, and 1 unsupported metal-containing structure).

1. **Scope lock:** retain ERTA and ERalpha only. Remove ERbeta from tabs, shortcuts, native QA, released catalog, packaged model/manifest/AD assets, installer, and user documentation. Keep research artifacts outside the FDA distribution.
2. **Progress:** add a determinate batch progress bar and text for read, CAS/SMILES resolution, preprocessing/prediction, workbook writing, and terminal success/failure. Worker threads communicate progress through Tk `after`; no widget is updated directly from a worker.
3. **Explain non-predictions:** keep probability/label columns blank and numeric-safe when no prediction exists, and retain bilingual-friendly result status, category, description, and recommended action fields in the row-aligned `Diagnostics` sheet.
4. **Workbook guidance:** make `Predictions` the first and active worksheet, followed by `Guide`, `Diagnostics`, `Input`, and `Metadata`. `Guide` explains all sheets and includes totals and reason counts. Preserve exact machine-readable headers on data sheets. Use plain pandas/openpyxl cell formatting without colored/fancy outcome highlighting; the later v3 result-column amendment is authoritative.
5. **Compatibility:** preserve ERTA behavior, ERalpha model/preprocessing/AD calculations, input precedence, row order, passthrough identifiers, atomic publication, read-only install boundaries, and the immutable archived legacy ERTA package.
6. **Validation:** run unit/integration tests, exercise the FDA-provided 504-row workbook, verify progress reaches 100% or a terminal failure, build twice, publish, install to a clean directory, run packaged ERTA/ERalpha QA, uninstall, and re-audit all 10,280 archived legacy ERTA files.

## FDA usability amendment: batch controls and path consistency

This amendment is applied before the corresponding code changes.

1. ERTA and ERalpha batch pages use the same compact control hierarchy: `Input xlsx`, a read-only `Result folder (same as input)` display, `Download template`, and a normal-sized `Run batch` button.
2. `Download template` appears immediately above `Run batch` with the notice that the `CAS` column is the required batch input; direct SMILES remains supported where the workbook contract permits it.
3. Both pages display determinate progress and a terminal success/failure state.
4. The displayed result directory is derived from the selected input workbook and is not independently selectable. The input path captured when the run starts, its parent shown in the read-only display, and the directory reported in the result/status areas are one canonical snapshot. Input selection and template actions are disabled for the duration of a batch so the displayed path cannot diverge from the running snapshot.
5. The result box reports the exact atomically published file path returned by the exporter. Tests cover changed-input snapshots, control locking, template generation, and path equality.

## FDA usability amendment: batch result simplification

This amendment is applied before the corresponding code changes.

1. Batch applicability-domain graphs are output-file artifacts, not interactive batch-screen content.
2. The ERTA Batch page follows the ERalpha layout exactly: one full-width input/control box, one progress row directly underneath, one full-width `Prediction result` text box, and one status row.
3. Remove the ERTA batch preview table, graph selector, refresh button, and in-app batch AD graph panel. Preserve generation of AD columns and graph files under the input-adjacent `graphs` directory.
4. The result box contains a concise completion summary: row totals, prediction counts, in/out-domain counts when available, exact workbook path, and graph directory/count. On failure it contains a concise terminal error summary.
5. Existing single-prediction visualization, sorting, model behavior, and AD calculations remain unchanged. The authoritative amendments below govern batch workbook layout, graph artifacts, template flow, progress, and canonical paths.

## FDA usability amendment: authoritative input-adjacent batch publication

This amendment supersedes every earlier requirement for an independently selected or configurable batch output directory. It changes publication location and controls only; it does not change the released models, preprocessing, decision rules, applicability-domain calculations, FDA-facing task scope, or evidence limitations.

1. For both ERTA and ERalpha, the selected input workbook's resolved parent directory is the sole destination for the batch result workbook.
2. The independent batch output-directory chooser is removed. Each batch page instead shows the derived parent path in a read-only field labeled `Result folder (same as input)`.
3. Before prediction begins, the application must reject a protected or unwritable input parent and instruct the user to copy the input workbook to a writable folder outside application files and select the copy. It must not fall back or silently redirect to `%LOCALAPPDATA%`, Documents, the install root, the source tree, or any other location.
4. Batch workbook publication remains atomic. It must never overwrite the selected input workbook or any prior result workbook; when the base result name is occupied, it allocates a collision-safe name such as `_2`, `_3`, and so on.
5. ERTA batch graph artifacts remain under an input-adjacent `graphs` directory. ERalpha batch AD graphs are published in a collision-safe sibling directory named from the result workbook stem (`<workbook-stem>_graphs`, then `_2`, `_3`, and so on when occupied). Single-prediction caches, plots, and exports remain in their writable external user roots and are not moved beside the batch input.
6. The result summary and terminal status report the exact published workbook path. The displayed destination, captured input parent, exporter destination, and reported path must agree.

## FDA usability amendment: authoritative ERalpha result and AD parity

This amendment supersedes the earlier `Guide`-first rule and any earlier statement that ERalpha batch produces no AD graph artifacts. It adds reporting parity only: no model change or new model release was identified or authorized, and the accepted V7 classification release remains unchanged.

1. The ERalpha workbook sheet order is `Predictions`, `Guide`, `Diagnostics`, `Input`, `Metadata`; `Predictions` is both first and active when the workbook opens. `Guide`, row-aligned `Diagnostics`, the untouched source `Input`, and release/provenance `Metadata` remain present.
2. Within `Predictions`, collision-safe original input columns appear first in their original order. The exact trusted v3 suffix defined in the Input/output contract follows. Technical status and reasons reside in `Diagnostics`; constant workflow/model/evidence provenance resides in `Metadata`. A normalized input header that collides with a trusted result, diagnostic, metadata, or AD header is omitted from `Predictions`; the generated value remains authoritative on its designated sheet and the untouched original remains available on `Input`.
3. The nine AD fields, in order, are `AD`, `AD_MeanDistance`, `AD_DistanceThreshold`, `AD_Distance_InDomain`, `AD_SimilarityMax`, `AD_SimilarityThreshold`, `AD_Similarity_InDomain`, `AD_PC1`, and `AD_PC2`. They are computed and graphed only from the ERalpha route-specific training reference, cache, and batch results; ERTA AD state is never substituted.
4. ERalpha batch AD graphs use binding-specific labels and the input-adjacent, collision-safe graph directory defined above. In the same full-width `Prediction result` placement used by ERTA, the detailed completion summary reports Total, Binding, Non-binding, Not predicted, AD In-domain, AD Out-of-domain, exact workbook path, graph count/directory, and the evidence caveat.
5. `Binding` and `Non-binding` continue to mean direct ERalpha receptor binding classification only. They do not assert transcriptional activation, agonism, antagonism, signaling, coactivator recruitment, or general endocrine disruption.
6. The accepted V7 evidence remains an internal resplit/sensitivity evaluation using historically exposed development data. It is not fresh independent, external, temporal, population, regulatory, or regulatory-use validation.

## Authoritative shared example, plain workbook, filename, and dialog amendment

This amendment supersedes every earlier reference to
`templates\ERTA_KRICT_example.xlsx` as the UI default, every requirement for
colored/fancy ERalpha workbook styling, and every ERTA result filename pattern
without the `ERTA_` prefix. It changes distribution, presentation, naming, and
batch feedback only. Released models, decision rules, AD calculations, direct
binding meaning, evidence limitations, and provenance sheets remain unchanged.

1. The exact user-supplied
   `D:\research\FDA_endocrine_disruption\ER_Predictor\test.xlsx` (25 rows) is
   bundled as `templates/test.xlsx` without changing its bytes. Its only header
   is the historical `CARSRN` spelling, which both ERTA and ERalpha batch
   adapters recognize as a CAS alias. Neither packaging nor QA may mutate the
   supplied source workbook. Its SHA-256 is
   `5a1f569f8f6a5cd47bff67a189645c3f9461fcf07bd24f4e5b4f83197f3350aa`.
2. Both endpoints receive the same `Path` from
   `core.paths.resolve_shared_example_input(example_path=None)`. A frozen one-folder
   layout `<container>/ER_Predictor/ER_Predictor.exe` uses
   `<container>/test.xlsx`. Other runs atomically initialize
   `%USERPROFILE%/Documents/ER_Predictor/Examples/test.xlsx` from the bundled
   bytes. Existing user or published copies are preserved; deleting that copy
   is the explicit reset operation. Missing/unreadable bundled bytes or a
   non-writable destination fails clearly instead of falling back.
3. Single-input defaults and **Example input** use the first shared CAS and an
   empty SMILES when the source provides no SMILES. CAS lookup is therefore
   required before a CAS-only single prediction. Batch QA copies the workbook
   into QA-owned writable directories and never publishes beside or writes to
   the distributed/source copy. Unattended native QA replaces PubChem responses
   for the listed 25 CAS values with the same valid `C=O` SMILES solely for
   deterministic callback/model/export coverage, records that substitution in
   its automation transcript, and does not present it as chemical-identity or
   online-service evidence. The unmodified 25-CAS online run remains a separate
   packaged-app verification.
4. `gui.main_window.allocate_erta_output_path(input_path, model_name)` reserves
   ERTA batch results as
   `ERTA_<input-stem>_<model-name>_prediction.xlsx` with collision suffixes
   before `.xlsx`. ERalpha keeps its collision-safe ERBA filename and its
   `Predictions`, `Guide`, row-level `Diagnostics`, untouched `Input`, and
   execution-level provenance `Metadata` sheets.
5. ERalpha removes colored/fancy workbook decoration.
   `core.native_qa.compare_plain_primary_workbook_formatting(erta_workbook,
   eralpha_workbook)` verifies that its primary `Predictions` cells have the
   exact same openpyxl formatting signature as ERTA cells of the same kind;
   values and endpoint-specific column semantics are not compared to establish
   formatting parity.
6. An atomically published workbook is a successful batch job on either
   endpoint even when expected unsupported rows could not be predicted or
   optional AD/graph artifacts are unavailable. Every such publication shows
   the blue `messagebox.showinfo("Batch prediction done", ...)` dialog using
   the same template: exact saved path, total/predicted/not-predicted and
   endpoint outcome counts, graph count/directory, and neutral textual
   AD/graph details. It never hides unavailable rows or claims that every row
   was predicted. If none were predicted, both the result and dialog say
   `No rows could be predicted.` A start, read, model execution, or workbook
   save/publication failure that produces no workbook instead shows the red
   `messagebox.showerror("Batch prediction failed", ...)`, restores all batch
   controls, leaves terminal failure text/progress, and emits no success
   dialog. Expected row-level exclusions and missing optional artifacts do not
   use a generic yellow warning dialog.

## FDA usability amendment: exact ERTA/ERalpha interaction parity

This amendment replaces approximate visual similarity with a shared widget and
interaction contract. It also supersedes the earlier statement that model/AD
browsing remains ERTA-only; ERalpha browsing is restricted as specified below.

1. ERTA and ERalpha use the same outer grid, hidden-by-default `Options` box,
   single/batch notebook placement, padding, row/column weights, control order,
   fixed-size probability/result areas, batch progress row, result box, and
   status-row placement. A stable `parity_widgets` map names the corresponding
   controls for native bounding-box comparison.
2. Both single pages initialize and reset from the first CAS value in the shared
   `test.xlsx`. The distributed source has the exact original `CARSRN` header
   and 25 CAS rows but no SMILES column, so the initial/reset SMILES may be
   blank until PubChem lookup. Both batch pages initialize to the exact same
   writable shared workbook; endpoint state remains independent after
   initialization.
3. Both `Options` boxes expose, in order, `Model`, `Browse`, `Reload model`,
   `AD reference`, `Browse`, and `Reload AD`. ERalpha model browsing accepts
   only regular bundled files represented by a `released` catalog entry whose
   model ID, size, and SHA-256 also match an executable release allowlist trust
   anchor. Rejection occurs before `joblib` deserialization. ERalpha AD browsing
   and reload are confined to the approved route-specific bundled reference.
   Both approved defaults are validated and loaded automatically at startup.
4. The ERalpha catalog retains every uniquely identified released model for a
   route while preserving one route default for existing inference APIs.
   Activating a choice creates a route-keyed predictor snapshot from that exact
   approved specification, binds the matching ERalpha AD route, and invalidates
   stale single/batch completions and progress. Reload completions are accepted
   only for the latest operation and unchanged typed selection. AD reload fits
   a fresh route manager before atomically replacing the active manager, so
   in-flight request snapshots keep their original fitted AD state. Additional
   choices appear only when their model and trust-anchor artifacts have
   completed the approved release process; this is not a plugin or
   arbitrary-file mechanism.
5. `Download template` remains an explicit save dialog on both batch pages.
   The destination must be a writable user-selected folder outside bundled or
   installed resources, and neither endpoint silently redirects a protected
   destination.
6. The intentional endpoint wording differences are limited to
   Positive/Negative for ERTA transcriptional activation versus
   Binding/Non-binding and the required binding-evidence caveat for direct
   ERalpha binding. Predictor state, model paths, AD references, request
   snapshots, results, and batch destinations remain endpoint-isolated.
7. Starting either batch immediately replaces any prior completion summary
   with `Batch prediction is running.` The aggregate progress row uses the
   same format and stage map on both endpoints:
   `Reading input workbook` (0%), `Resolving CAS/SMILES` (10–35%),
   `Preprocessing and prediction` (35–88%), `Writing workbook` (90%), and
   terminal `Completed` or `Failed` (100%). The displayed format is
   `<percent>% - <current>/<total> - <stage>`.
8. Aggregate stages remain exclusively in the progress row. During an actual
   CAS lookup, the shared lower status row exclusively shows
   `Fetching SMILES from PubChem: <index> / <total> (<CAS>)`; its common
   lifecycle text is `Batch prediction started.`,
   `Batch prediction completed: <path>`, or
   `Batch prediction failed: <details>`. Worker callbacks reach Tk only through
   `after`, and request/model-generation guards discard queued stale progress,
   lookup detail, and completion updates.
9. ERTA preserves its legacy workbook schema and zero-fingerprint inference
   behavior. Presentation counts a row as Predicted/Positive/Negative only when
   `Mol_valid=True`; legacy labels retained on `Mol_valid=False` workbook/graph
   rows are explicitly described as not usable predictions. ERTA model and AD
   reloads are serialized against batch execution in both directions, and a
   successful reload clears the prior batch summary. AD In-domain,
   Out-of-domain, and Unavailable counts on both endpoints cover predicted rows
   only; not-predicted rows are excluded from that denominator.
