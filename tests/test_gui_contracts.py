import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from core.contracts import (
    ERBAArtifactSpec,
    ERBARoutePreflight,
    ERBAResult,
    ERBAStatusCode,
    ERBASubtype,
    ERBATask,
)
from gui.erba_tab import (
    ErbaTab,
    allocate_output_path,
    canonicalize_batch_input,
    export_erba_batch,
    load_erba_catalog,
    metadata_row,
)
from gui.main_window import MainWindow


class GuiContractTests(unittest.TestCase):
    def test_main_window_keeps_erta_first_and_nests_legacy_notebook(self):
        source = Path("gui/main_window.py").read_text(encoding="utf-8")
        self.assertIn('self.notebook.add(self.erta_tab, text="ERTA")', source)
        self.assertIn('self.notebook.add(self.eralpha_tab, text="ERalpha")', source)
        self.assertNotIn('text="ERbeta"', source)
        self.assertNotIn("self.erbeta_tab", source)
        self.assertIn("self.notebook.select(self.erta_tab)", source)
        self.assertIn("self.erta_notebook = ttk.Notebook(self.erta_tab)", source)
        self.assertIn('self.title("ER Predictor")', source)
        erba_source = Path("gui/erba_tab.py").read_text(encoding="utf-8")
        self.assertIn("self.single_task_var", erba_source)
        self.assertIn("self.batch_task_var", erba_source)
        self.assertNotIn("self.task_var =", erba_source)
        self.assertIn('text="Single prediction"', erba_source)
        self.assertIn('text="Batch prediction"', erba_source)
        self.assertIn('text="Input"', erba_source)
        self.assertIn('text="Prediction result"', erba_source)
        self.assertIn("Applicability domain: Not evaluated", erba_source)
        self.assertNotIn("values = asdict(result)", erba_source)

    def test_catalog_releases_only_approved_routes_and_keeps_beta_regression_disabled(self):
        payload = {
            "schema_version": 2,
            "regression_parity_approved": True,
            "routes": [
                {"task": "classification", "subtype": "er_alpha", "release_status": "released",
                 "relative_path": "alpha.joblib", "size_bytes": 10, "sha256": "a" * 64, "model_id": "alpha"},
                {"task": "ic50_regression", "subtype": "er_beta", "release_status": "released",
                 "relative_path": "beta.joblib", "size_bytes": 10, "sha256": "b" * 64, "model_id": "beta"},
                {"task": "classification", "subtype": "er_beta", "release_status": "pending",
                 "relative_path": "pending.joblib", "size_bytes": 10, "sha256": "c" * 64, "model_id": "pending"},
            ],
        }
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory, "catalog.v2.json")
            path.write_text(json.dumps(payload), encoding="utf-8")
            catalog, parity, _ = load_erba_catalog(path)
        self.assertTrue(parity)
        self.assertIn((ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA), catalog)
        self.assertNotIn((ERBATask.IC50_REGRESSION, ERBASubtype.ER_BETA), catalog)
        self.assertNotIn((ERBATask.CLASSIFICATION, ERBASubtype.ER_BETA), catalog)

    def test_output_names_are_reserved_without_overwriting(self):
        with tempfile.TemporaryDirectory() as directory:
            first = allocate_output_path(directory, ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA)
            second = allocate_output_path(directory, ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA)
            self.assertTrue(first.exists())
            self.assertTrue(second.exists())
            self.assertNotEqual(first, second)



class _Variable:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value


class _StartupPredictor:
    def __init__(self, error=None):
        self.error = error
        self.calls = []

    def load_model(self, path):
        self.calls.append(path)
        if self.error:
            raise self.error


class _StartupADCalculator:
    def __init__(self, error=None):
        self.error = error
        self.calls = []

    def fit_from_excel_cached(self, path):
        self.calls.append(path)
        if self.error:
            raise self.error


class StartupLoadingContractTests(unittest.TestCase):
    def _window(self, model_path, ad_path, predictor=None, ad_calculator=None):
        window = MainWindow.__new__(MainWindow)
        window.model_path_var = _Variable(model_path)
        window.ad_ref_path_var = _Variable(ad_path)
        window.predictor = predictor or _StartupPredictor()
        window.ad_calculator = ad_calculator or _StartupADCalculator()
        window.statuses = []
        window.set_status = window.statuses.append
        window.run_threaded = lambda job: job()
        return window

    def test_startup_reports_blank_or_missing_model_path_and_does_not_claim_ready(self):
        with tempfile.TemporaryDirectory() as directory:
            ad_path = str(Path(directory, "reference.xlsx"))
            missing_model_path = str(Path(directory, "missing.keras"))
            Path(ad_path).touch()
            for model_path, failure in (
                ("", "model path is blank"),
                (missing_model_path, f"model path does not exist: {missing_model_path}"),
            ):
                with self.subTest(model_path=model_path):
                    window = self._window(model_path, ad_path)

                    window.load_defaults_on_startup()

                    self.assertEqual(window.predictor.calls, [])
                    self.assertEqual(window.ad_calculator.calls, [ad_path])
                    self.assertEqual(window.statuses[-1], f"Startup failed: {failure}")
                    self.assertNotIn("Ready", window.statuses)

    def test_startup_reports_blank_or_missing_ad_reference_path_and_does_not_claim_ready(self):
        with tempfile.TemporaryDirectory() as directory:
            model_path = str(Path(directory, "model.keras"))
            missing_ad_path = str(Path(directory, "missing.xlsx"))
            Path(model_path).touch()
            for ad_path, failure in (
                ("", "AD reference path is blank"),
                (missing_ad_path, f"AD reference path does not exist: {missing_ad_path}"),
            ):
                with self.subTest(ad_path=ad_path):
                    window = self._window(model_path, ad_path)

                    window.load_defaults_on_startup()

                    self.assertEqual(window.predictor.calls, [model_path])
                    self.assertEqual(window.ad_calculator.calls, [])
                    self.assertEqual(window.statuses[-1], f"Startup failed: {failure}")
                    self.assertNotIn("Ready", window.statuses)

    def test_startup_reports_model_load_exception_and_still_loads_ad(self):
        with tempfile.TemporaryDirectory() as directory:
            model_path = str(Path(directory, "model.keras"))
            ad_path = str(Path(directory, "reference.xlsx"))
            Path(model_path).touch()
            Path(ad_path).touch()
            window = self._window(
                model_path,
                ad_path,
                predictor=_StartupPredictor(RuntimeError("corrupt model")),
            )

            window.load_defaults_on_startup()

        self.assertEqual(window.predictor.calls, [model_path])
        self.assertEqual(window.ad_calculator.calls, [ad_path])
        self.assertEqual(window.statuses[-1], "Startup failed: model load failed: corrupt model")
        self.assertNotIn("Ready", window.statuses)

    def test_startup_reports_ad_fit_exception_and_does_not_claim_ready(self):
        with tempfile.TemporaryDirectory() as directory:
            model_path = str(Path(directory, "model.keras"))
            ad_path = str(Path(directory, "reference.xlsx"))
            Path(model_path).touch()
            Path(ad_path).touch()
            window = self._window(
                model_path,
                ad_path,
                ad_calculator=_StartupADCalculator(RuntimeError("invalid reference")),
            )

            window.load_defaults_on_startup()

        self.assertEqual(window.predictor.calls, [model_path])
        self.assertEqual(window.ad_calculator.calls, [ad_path])
        self.assertEqual(window.statuses[-1], "Startup failed: AD fit failed: invalid reference")
        self.assertNotIn("Ready", window.statuses)

    def test_startup_loads_both_resources_before_setting_ready(self):
        with tempfile.TemporaryDirectory() as directory:
            model_path = str(Path(directory, "model.keras"))
            ad_path = str(Path(directory, "reference.xlsx"))
            Path(model_path).touch()
            Path(ad_path).touch()
            window = self._window(model_path, ad_path)

            window.load_defaults_on_startup()

        self.assertEqual(window.predictor.calls, [model_path])
        self.assertEqual(window.ad_calculator.calls, [ad_path])
        self.assertEqual(
            window.statuses,
            ["Loading default model...", "Loading default AD reference...", "Ready"],
        )

class _Predictor:
    def __init__(self, status=ERBAStatusCode.OK):
        self.calls = []
        self.status = status

    def preflight(self, task, subtype):
        return ERBARoutePreflight(task, subtype, ERBAStatusCode.OK)

    def predict(self, request):
        self.calls.append(request)
        return ERBAResult(
            row_index=request.row_index, task=request.task, subtype=request.subtype,
            raw_smiles=request.smiles, model_smiles=request.smiles or None,
            status_code=self.status, status_message=self.status.value,
            non_binding_probability=0.2, binding_probability=0.8,
            binding_label="binding", model_id="model-v1", model_sha256="a" * 64,
        )


class ErbaExcelContractTests(unittest.TestCase):
    spec = ERBAArtifactSpec("model.joblib", 1, "a" * 64, "model-v1")
    provenance = {
        "metadata": {
            "protocol_sha256": "a" * 64, "source_manifest_sha256": "b" * 64,
            "historical_exposure_manifest_sha256": "c" * 64,
            "split_manifest_sha256": "d" * 64, "nested_cv_sha256": "e" * 64,
            "internal_resplit_sha256": "f" * 64, "report_sha256": "0" * 64,
            "caveat_sha256": "1" * 64,
        }
    }

    def _input(self, directory, frame):
        path = Path(directory, "input.xlsx")
        frame.to_excel(path, index=False)
        return path

    def test_canonical_aliases_reject_duplicates_and_preserve_source_fields(self):
        frame = pd.DataFrame([["one", "50-00-0", "CCO", "keep"]], columns=["id", "cas", "smiles", "Note"])
        canonical, passthrough = canonicalize_batch_input(frame)
        self.assertEqual(canonical.iloc[0].tolist(), ["one", "50-00-0", "CCO"])
        self.assertEqual(passthrough.columns.tolist(), ["id", "cas", "smiles", "Note"])
        with self.assertRaisesRegex(ValueError, "Duplicate canonical CAS"):
            canonicalize_batch_input(pd.DataFrame([["a", "b"]], columns=["CAS", "cas"]))
        with tempfile.TemporaryDirectory() as directory:
            input_path = self._input(directory, pd.DataFrame([["a", "b"]], columns=["CAS", "cas"]))
            with self.assertRaisesRegex(ValueError, "Duplicate canonical CAS"):
                export_erba_batch(input_path, directory, ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA,
                                  _Predictor(), self.spec, self.provenance)

    def test_mixed_batch_contract_direct_precedence_failures_and_metadata(self):
        frame = pd.DataFrame(
            [["direct", "12-34-5", "CCO", "first"], ["lookup", "50-00-0", "", "second"],
             ["bad", "12-34-5", "", "third"]],
            columns=["identifier", "cas no.", "SMILES", "Analyst note"],
        )
        predictor = _Predictor()
        with tempfile.TemporaryDirectory() as directory, patch(
            "gui.erba_tab.cas_to_smiles", return_value={"CanonicalSMILES": "CCC"}
        ) as lookup:
            destination, count = export_erba_batch(
                self._input(directory, frame), directory, ERBATask.CLASSIFICATION,
                ERBASubtype.ER_ALPHA, predictor, self.spec, self.provenance,
            )
            with pd.ExcelFile(destination) as workbook:
                sheet_names = workbook.sheet_names
            predictions = pd.read_excel(destination, sheet_name="Predictions", dtype=str, keep_default_na=False)
            supplied = pd.read_excel(destination, sheet_name="Input", dtype=str, keep_default_na=False)
            metadata = pd.read_excel(destination, sheet_name="Metadata", dtype=str, keep_default_na=False)
        self.assertEqual(count, 3)
        self.assertEqual(sheet_names, ["Guide", "Predictions", "Input", "Metadata"])
        self.assertEqual([call.smiles for call in predictor.calls], ["CCO", "CCC", ""])
        lookup.assert_called_once_with("50-00-0")
        self.assertEqual(supplied.columns.tolist(), frame.columns.tolist())
        self.assertEqual(supplied["Analyst note"].tolist(), ["first", "second", "third"])
        self.assertEqual(predictions.loc[2, "binding_probability"], "")
        self.assertEqual(predictions.loc[2, "Result_Status"], "Not predicted")
        self.assertEqual(predictions.loc[2, "Reason_Category"], "Invalid CAS")
        self.assertTrue(predictions.loc[2, "Recommended_Action"])
        self.assertEqual(metadata.loc[0, "Performance_Evidence_Scope"], "internal_historically_exposed")
        self.assertEqual(metadata.loc[0, "Protocol_SHA256"], "a" * 64)

    def test_batch_progress_and_guide_explain_all_unavailable_predictions(self):
        class _FdaPredictor(_Predictor):
            def predict(self, request):
                if request.smiles == "O":
                    self.status = ERBAStatusCode.NO_CARBON
                elif request.smiles == "[Na+]":
                    self.status = ERBAStatusCode.METAL_RETAINED
                else:
                    self.status = ERBAStatusCode.OK
                return super().predict(request)

        # Fixture-level equivalent of the FDA category totals; no supplied workbook is altered.
        frame = pd.DataFrame(
            ([["predicted", "CCO"]] * 448)
            + [["inorganic", "O"]] * 18
            + [["metal", "[Na+]"]]
            + [["bad-cas", ""]] * 2
            + [["pubchem", ""]] * 35,
            columns=["Row_ID", "SMILES"],
        )
        frame.loc[467:468, "CAS"] = "12-34-5"
        frame.loc[469:, "CAS"] = "50-00-0"
        progress = []
        with tempfile.TemporaryDirectory() as directory, patch(
            "gui.erba_tab.cas_to_smiles", side_effect=ValueError("404")
        ):
            destination, count = export_erba_batch(
                self._input(directory, frame), directory, ERBATask.CLASSIFICATION,
                ERBASubtype.ER_ALPHA, _FdaPredictor(), self.spec, self.provenance,
                progress_callback=lambda stage, current, total, percent: progress.append(
                    (stage, current, total, percent)
                ),
            )
            predictions = pd.read_excel(
                destination, sheet_name="Predictions", dtype=str, keep_default_na=False
            )
            guide = pd.read_excel(destination, sheet_name="Guide", dtype=str, keep_default_na=False)
        self.assertEqual(count, 504)
        self.assertEqual((predictions["Result_Status"] == "Predicted").sum(), 448)
        self.assertEqual((predictions["Result_Status"] == "Not predicted").sum(), 56)
        self.assertEqual(
            predictions.loc[predictions["Result_Status"] == "Not predicted", "Reason_Category"]
            .value_counts()
            .to_dict(),
            {
                "PubChem lookup unavailable": 35,
                "No carbon / inorganic": 18,
                "Invalid CAS": 2,
                "Unsupported metal-containing structure": 1,
            },
        )
        self.assertEqual(progress[0], ("Reading input workbook", 0, 0, 0))
        self.assertEqual(progress[-1], ("Completed", 504, 504, 100))
        self.assertEqual({entry[0] for entry in progress}, {
            "Reading input workbook", "Resolving CAS/SMILES", "Predicting", "Writing workbook", "Completed",
        })
        self.assertIn("Not predicted rows", guide["Topic"].tolist())
    def test_single_cas_lookup_uses_isomeric_fallback_and_clears_on_missing_smiles(self):
        class _Variable:
            def __init__(self, value=""):
                self.value = value

            def get(self):
                return self.value

            def set(self, value):
                self.value = value

        class _Widget:
            def configure(self, **kwargs):
                pass

        class _Text(_Widget):
            def __init__(self):
                self.value = ""

            def delete(self, start, end):
                self.value = ""

            def insert(self, start, value):
                self.value = value

        class _ImmediateThread:
            def __init__(self, target, daemon):
                self.target = target

            def start(self):
                self.target()

        tab = ErbaTab.__new__(ErbaTab)
        predictor = _Predictor()
        tab.predictor = predictor
        tab.smiles_var, tab.cas_var = _Variable(), _Variable("50-00-0")
        tab.single_button, tab.single_status_var = _Widget(), _Variable()
        tab.single_prediction_summary_var = _Variable()
        tab.single_negative_probability_var = _Variable()
        tab.single_positive_probability_var = _Variable()
        tab.single_pic50_var, tab.single_ic50_var = _Variable(), _Variable()
        tab.single_detail_var = _Variable()
        tab.single_negative_bar, tab.single_positive_bar = _Widget(), _Widget()
        tab._draw_single_structure = lambda smiles: None
        tab.after = lambda delay, callback: callback()
        tab._finish_duration_ms = lambda request_id: 0
        tab._snapshot_is_current = lambda *args: True
        tab._route_changed = lambda workflow: None
        tab._emit = lambda *args, **kwargs: None
        tab._snapshot = lambda workflow: (1, ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA, "model-v1")
        with patch("gui.erba_tab.threading.Thread", _ImmediateThread), patch(
            "gui.erba_tab.cas_to_smiles", return_value={"IsomericSMILES": "C[C@H](O)F"}
        ):
            tab.single_predict_clicked()
        self.assertEqual([request.smiles for request in predictor.calls], ["C[C@H](O)F"])
        self.assertIn("Binding prediction", tab.single_prediction_summary_var.value)
        self.assertIn("Negative probability", tab.single_negative_probability_var.value)
        self.assertIn("Positive probability", tab.single_positive_probability_var.value)
        tab._snapshot = lambda workflow: (2, ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA, "model-v1")
        with patch("gui.erba_tab.threading.Thread", _ImmediateThread), patch(
            "gui.erba_tab.cas_to_smiles", return_value={"PubChem_CID": 712}
        ):
            tab.single_predict_clicked()
        self.assertEqual(tab.single_prediction_summary_var.value, "Prediction failed")
        self.assertIn("neither CanonicalSMILES nor IsomericSMILES", tab.single_status_var.value)

    def test_structured_single_result_renders_classification_details_for_both_receptors(self):
        class _Variable:
            def __init__(self):
                self.value = ""

            def set(self, value):
                self.value = value

        class _Widget:
            def __init__(self):
                self.values = {}

            def configure(self, **kwargs):
                self.values.update(kwargs)

        tab = ErbaTab.__new__(ErbaTab)
        tab.single_prediction_summary_var = _Variable()
        tab.single_negative_probability_var = _Variable()
        tab.single_positive_probability_var = _Variable()
        tab.single_pic50_var, tab.single_ic50_var = _Variable(), _Variable()
        tab.single_detail_var = _Variable()
        tab.single_negative_bar, tab.single_positive_bar = _Widget(), _Widget()
        tab._draw_single_structure = lambda smiles: None
        tab._configure_single_result_task = lambda task: None
        tab._render_single_result(ERBAResult(
            row_index=0, task=ERBATask.CLASSIFICATION, subtype=ERBASubtype.ER_ALPHA,
            raw_smiles="CCO", model_smiles="CCO", status_code=ERBAStatusCode.OK,
            status_message="ok", non_binding_probability=0.2, binding_probability=0.8,
            binding_label="binding", model_id="model-v1",
        ))
        self.assertEqual(tab.single_prediction_summary_var.value, "Binding prediction: Binding")
        self.assertEqual(tab.single_negative_probability_var.value, "Negative probability: 20.0%")
        self.assertEqual(tab.single_positive_probability_var.value, "Positive probability: 80.0%")
        self.assertEqual(tab.single_positive_bar.values["value"], 80.0)
        tab._render_single_result(ERBAResult(
            row_index=0, task=ERBATask.CLASSIFICATION, subtype=ERBASubtype.ER_BETA,
            raw_smiles="CCO", model_smiles="CCO", status_code=ERBAStatusCode.OK,
            status_message="ok", non_binding_probability=0.7, binding_probability=0.3,
            binding_label="non_binding", model_id="model-v2",
        ))
        self.assertEqual(tab.single_prediction_summary_var.value, "Binding prediction: Non Binding")
        self.assertEqual(tab.single_negative_probability_var.value, "Negative probability: 70.0%")
        self.assertEqual(tab.single_positive_probability_var.value, "Positive probability: 30.0%")

    def test_regression_metadata_is_task_specific_and_requires_parity_provenance(self):
        payload = {
            "metadata": {
                "source_manifest_sha256": "a" * 64,
                "preprocessing_parity_sha256": "b" * 64,
                "report_sha256": "c" * 64,
            }
        }
        row = metadata_row(ERBATask.IC50_REGRESSION, self.spec, payload)
        self.assertEqual(row["Performance_Evidence_Scope"], "internal_regression_model")
        self.assertEqual(row["Preprocessing_Parity_SHA256"], "b" * 64)
        self.assertEqual(row["Evidence_Caveat"], "")
        self.assertEqual(row["Protocol_SHA256"], "")
    def test_empty_and_fatal_batches_publish_nothing(self):
        with tempfile.TemporaryDirectory() as directory:
            empty = self._input(directory, pd.DataFrame(columns=["Row_ID", "SMILES"]))
            with self.assertRaisesRegex(ValueError, "no rows"):
                export_erba_batch(empty, directory, ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA,
                                  _Predictor(), self.spec, self.provenance)
            fatal = self._input(directory, pd.DataFrame([["x", "CCO"]], columns=["Row_ID", "SMILES"]))
            with self.assertRaisesRegex(RuntimeError, "no output"):
                export_erba_batch(fatal, directory, ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA,
                                  _Predictor(ERBAStatusCode.MODEL_INTEGRITY_FAILED), self.spec, self.provenance)
            self.assertEqual(sorted(path.name for path in Path(directory).glob("ERBA_*.xlsx")), [])

    def test_batch_completion_ignores_stale_request(self):
        tab = ErbaTab.__new__(ErbaTab)
        tab._active_batch_request_id = 2
        ErbaTab._batch_complete(tab, 1, ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA,
                                "model-v1", Path("ignored.xlsx"), 1, "")
        self.assertEqual(tab._active_batch_request_id, 2)

    def test_worker_progress_is_scheduled_through_after(self):
        class _Variable:
            def __init__(self):
                self.value = None

            def set(self, value):
                self.value = value

        tab = ErbaTab.__new__(ErbaTab)
        tab._active_batch_request_id = 7
        tab.batch_progress_var = _Variable()
        tab.batch_progress_value = _Variable()
        scheduled = []
        tab.after = lambda delay, callback: scheduled.append((delay, callback))
        tab._batch_progress_from_worker(7, "Predicting", 3, 10, 61)
        self.assertEqual(len(scheduled), 1)
        self.assertEqual(tab.batch_progress_var.value, None)
        scheduled[0][1]()
        self.assertEqual(tab.batch_progress_value.value, 61)
        self.assertEqual(tab.batch_progress_var.value, "61% - 3/10 - Predicting")

    def test_export_avoids_existing_destination_collision(self):
        frame = pd.DataFrame([["x", "CCO"]], columns=["Row_ID", "SMILES"])
        with tempfile.TemporaryDirectory() as directory:
            existing = Path(directory, "ERBA_classification_er_alpha_results.xlsx")
            existing.write_bytes(b"prior")
            destination, _ = export_erba_batch(
                self._input(directory, frame), directory, ERBATask.CLASSIFICATION,
                ERBASubtype.ER_ALPHA, _Predictor(), self.spec, self.provenance,
            )
            self.assertEqual(existing.read_bytes(), b"prior")
            self.assertEqual(destination.name, "ERBA_classification_er_alpha_results_2.xlsx")


class EralphaBatchUsabilityContractTests(unittest.TestCase):
    class _Variable:
        def __init__(self, value=""):
            self.value = value

        def get(self):
            return self.value

        def set(self, value):
            self.value = value

    class _Widget:
        def __init__(self):
            self.state = None

        def configure(self, **kwargs):
            self.state = kwargs.get("state", self.state)

    class _Text(_Widget):
        def __init__(self):
            super().__init__()
            self.value = ""

        def delete(self, _start, _end):
            self.value = ""

        def insert(self, _start, value):
            self.value = value

    def _tab_with_batch_controls(self):
        tab = ErbaTab.__new__(ErbaTab)
        tab.batch_input_button = self._Widget()
        tab.batch_output_button = self._Widget()
        tab.batch_template_button = self._Widget()
        tab.batch_button = self._Widget()
        return tab

    def test_downloaded_template_has_only_the_required_cas_header_and_becomes_input(self):
        tab = self._tab_with_batch_controls()
        tab.batch_input_var = self._Variable()
        with tempfile.TemporaryDirectory() as directory:
            destination = Path(directory, "ERalpha_batch_template.xlsx")
            with patch("gui.erba_tab.filedialog.asksaveasfilename", return_value=str(destination)):
                tab._download_batch_template()
            template = pd.read_excel(destination, dtype=str, keep_default_na=False)
        self.assertEqual(template.columns.tolist(), ["CAS"])
        self.assertTrue(template.empty)
        self.assertEqual(tab.batch_input_var.value, str(destination))

    def test_completion_uses_exported_parent_for_displayed_directory_and_exact_destination_everywhere(self):
        tab = self._tab_with_batch_controls()
        tab._active_batch_request_id = 1
        tab.batch_output_var = self._Variable("C:/selected")
        tab.batch_status_var = self._Variable()
        tab.batch_result = self._Text()
        tab._finish_duration_ms = lambda _request_id: 0
        tab._set_batch_progress = lambda *args: None
        tab._snapshot_is_current = lambda *args: True
        tab._route_changed = lambda _workflow: None
        tab._emit = lambda *args, **kwargs: None
        destination = Path("C:/published/output/ERBA_classification_er_alpha_results.xlsx")

        tab._batch_complete(
            1, ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA, "model-v1", destination, 3, ""
        )

        self.assertEqual(tab.batch_output_var.value, str(destination.parent))
        self.assertEqual(tab.batch_status_var.value, f"ERBA batch complete: {destination}")
        self.assertIn(str(destination), tab.batch_result.value)
        self.assertTrue(all(
            widget.state == "normal"
            for widget in (
                tab.batch_input_button, tab.batch_output_button,
                tab.batch_template_button, tab.batch_button,
            )
        ))

    def test_failure_and_stale_completion_restore_or_preserve_control_lock_correctly(self):
        tab = self._tab_with_batch_controls()
        tab._active_batch_request_id = 1
        tab.batch_output_var = self._Variable()
        tab.batch_status_var = self._Variable()
        tab._finish_duration_ms = lambda _request_id: 0
        tab._set_batch_progress = lambda *args: None
        tab._snapshot_is_current = lambda *args: True
        tab._route_changed = lambda _workflow: None
        tab._emit = lambda *args, **kwargs: None

        tab._batch_complete(1, ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA, "model-v1", None, 0, "write failed")

        self.assertIn("write failed", tab.batch_status_var.value)
        self.assertTrue(all(
            widget.state == "normal"
            for widget in (
                tab.batch_input_button, tab.batch_output_button,
                tab.batch_template_button, tab.batch_button,
            )
        ))

        tab._set_batch_controls_running(True)
        tab._active_batch_request_id = None
        tab._route_changed = lambda _workflow: tab._set_batch_controls_running(False)
        tab._batch_complete(2, ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA, "model-v1", None, 0, "stale")
        self.assertTrue(all(
            widget.state == "normal"
            for widget in (
                tab.batch_input_button, tab.batch_output_button,
                tab.batch_template_button, tab.batch_button,
            )
        ))

class _ErtaBatchVariable:
    def __init__(self, value=None):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class _ErtaBatchControl:
    def __init__(self):
        self.state = "normal"

    def configure(self, **kwargs):
        self.state = kwargs.get("state", self.state)


class _ErtaBatchPredictor:
    model_name = "legacy"

    def is_loaded(self):
        return True

    def predict_dataframe(self, frame, smiles_col):
        result = frame.copy()
        result["Prediction_label"] = "Positive"
        return result, pd.DataFrame([[0.0] for _ in range(len(frame))])


class ErtaBatchUsabilityContractTests(unittest.TestCase):
    def _window(self):
        window = MainWindow.__new__(MainWindow)
        window.batch_input_button = _ErtaBatchControl()
        window.output_dir_button = _ErtaBatchControl()
        window.download_template_button = _ErtaBatchControl()
        window.run_batch_button = _ErtaBatchControl()
        window.batch_progress_var = _ErtaBatchVariable()
        window.batch_progress_value = _ErtaBatchVariable()
        window._batch_active = False
        window._active_batch_input_path = None
        window._active_batch_output_dir = None
        window._active_batch_total = 1
        window.statuses = []
        window.set_status = window.statuses.append
        window.after = lambda _delay, callback: callback()
        window.ui = lambda callback, *args: callback(*args)
        window.run_threaded = lambda job: job()
        return window

    def test_compact_erta_batch_controls_and_cas_notice_are_present(self):
        source = Path("gui/main_window.py").read_text(encoding="utf-8")
        self.assertIn('text="Input xlsx"', source)
        self.assertIn('text="Output directory"', source)
        self.assertIn('text="Download template"', source)
        self.assertIn('text="Run batch"', source)
        self.assertIn("required CAS column", source)
        self.assertNotIn('text="Run batch prediction"', source)
        self.assertNotIn('bg="#0969da"', source)

    def test_template_contains_only_the_required_cas_header(self):
        window = self._window()
        window.project_root = str(Path.cwd() / "project-root")
        window.batch_input_var = _ErtaBatchVariable()
        window.batch_input_display_var = _ErtaBatchVariable()
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory, "Template.xlsx")
            with patch("gui.main_window.filedialog.asksaveasfilename", return_value=str(path)), \
                 patch("gui.main_window.messagebox.showinfo"):
                window.download_template_clicked()
            self.assertEqual(pd.read_excel(path).columns.tolist(), ["CAS"])

    def test_batch_controls_lock_and_restore_with_terminal_progress(self):
        window = self._window()
        window._set_batch_controls_active(True)
        self.assertTrue(window._batch_active)
        self.assertEqual(
            [control.state for control in (
                window.batch_input_button, window.output_dir_button,
                window.download_template_button, window.run_batch_button,
            )],
            ["disabled"] * 4,
        )
        window._finish_batch(False, "Batch prediction failed")
        self.assertFalse(window._batch_active)
        self.assertEqual(window.batch_progress_var.value, "100% - 0/1 - Failed")
        self.assertEqual(
            [control.state for control in (
                window.batch_input_button, window.output_dir_button,
                window.download_template_button, window.run_batch_button,
            )],
            ["normal"] * 4,
        )

    def test_progress_updates_are_scheduled_through_after(self):
        window = self._window()
        scheduled = []
        window.after = lambda delay, callback: scheduled.append((delay, callback))
        window._batch_progress_from_worker("Writing workbook", 3, 4, 85)
        self.assertEqual(window.batch_progress_var.value, None)
        self.assertEqual(len(scheduled), 1)
        scheduled[0][1]()
        self.assertEqual(window.batch_progress_var.value, "85% - 3/4 - Writing workbook")

    def test_batch_reports_the_published_file_in_the_captured_output_directory(self):
        window = self._window()
        window.predictor = _ErtaBatchPredictor()
        window.ad_calculator = type("AD", (), {"fitted": False})()
        window.model_path_var = _ErtaBatchVariable("")
        window.ad_ref_path_var = _ErtaBatchVariable("")
        window.output_dir_var = _ErtaBatchVariable()
        window.output_dir_display_var = _ErtaBatchVariable()
        window.update_preview_table = lambda _result: None
        reported = []
        window.update_batch_result_summary = lambda _result, path: reported.append(path)
        window.update_batch_ad_summary = lambda _result: None
        window.generate_batch_graphs = lambda *_args, **_kwargs: []
        window.show_error = lambda _title, error: self.fail(str(error))
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory, "input.xlsx")
            pd.DataFrame({"CAS": ["50-00-0"], "SMILES": ["CCO"]}).to_excel(source, index=False)
            window.batch_input_var = _ErtaBatchVariable(str(source))
            window.validated_output_dir = lambda: directory
            window.prepare_batch_input = lambda frame: frame
            with patch("gui.main_window.messagebox.showinfo"):
                window.batch_predict_clicked()
            self.assertEqual(Path(reported[0]).parent, Path(directory))
            self.assertTrue(Path(reported[0]).exists())
            self.assertEqual(window.output_dir_display_var.value, directory)


if __name__ == "__main__":
    unittest.main()
