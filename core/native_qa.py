"""Opt-in native-package acceptance driver used only by release verification."""
from __future__ import annotations

import json
import os
from pathlib import Path
import time
import traceback

from openpyxl import Workbook, load_workbook

from core.contracts import (
    ERBA_BINDING_CLASSIFICATION_EXCEL_CONTRACT_ID,
    ERBASubtype,
    ERBATask,
)


class NativePackageQa:
    """Drive real Tk callbacks and record fail-closed packaged-app evidence."""

    VALID_SMILES = "C[C@]12CC[C@H]3[C@@H]([C@@H]1CC[C@@H]2O)CCC4=CC(=CC=C34)O"
    LEGACY_INVALID_SMILES = "not a SMILES"
    MOCK_CAS = "50-00-0"
    MOCK_SMILES = "C=O"
    ERBA_BATCH_AD_COLUMNS = (
        "AD",
        "AD_MeanDistance",
        "AD_DistanceThreshold",
        "AD_Distance_InDomain",
        "AD_SimilarityMax",
        "AD_SimilarityThreshold",
        "AD_Similarity_InDomain",
        "AD_PC1",
        "AD_PC2",
    )
    ERBA_BATCH_GRAPH_FILES = {
        "binding_class_count.png",
        "binding_probability_histogram.png",
        "top_binding_chemicals.png",
        "ad_pca_plot.png",
        "ad_decision_plot.png",
    }
    ERBA_PROVENANCE_COLUMNS = (
        "Protocol_SHA256",
        "Source_Manifest_SHA256",
        "Historical_Exposure_Manifest_SHA256",
        "Split_Manifest_SHA256",
        "Nested_CV_SHA256",
        "Internal_Resplit_SHA256",
        "Report_SHA256",
        "Caveat_SHA256",
    )

    def __init__(self, app, destination: str | Path) -> None:
        self.app = app
        self.destination = Path(destination)
        self.destination.parent.mkdir(parents=True, exist_ok=True)
        self.root = Path(app.output_root) / "native-package-qa"
        self.root.mkdir(parents=True, exist_ok=True)
        self.stage = 0
        self.route_index = 0
        self.invalid_index = 0
        self.started_at = time.monotonic()
        self.deadline = self.started_at + 900
        self.erta_summary = ""
        self.erta_probability = ""
        self.erta_batch_input_path: Path | None = None
        self.erta_batch_outputs_before: set[Path] = set()
        self.erba_batch_input_path: Path | None = None
        self.erba_batch_outputs_before: set[Path] = set()
        self.erba_batch_graph_dirs_before: set[Path] = set()
        self.dialogs: list[dict] = []
        self.patches: list[tuple[object, str, object]] = []
        self.mock_cas_enabled = False
        gate_value = os.environ.get("ER_PREDICTOR_EXTERNAL_DRIVER_GATE", "").strip()
        self.external_driver_gate = Path(gate_value) if gate_value else None
        self.run_nonce = os.environ.get("ER_PREDICTOR_AUTOMATION_NONCE", "").strip()
        self.target_pid = os.getpid()
        self.routes = [
            (ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA),
        ]
        self.invalid_cases = [
            ("wildcard", "[*]CC", "", "wildcard_smiles"),
            ("pipe", "CC|CC", "", "cxsmiles_not_allowed"),
            ("metal", "[Na+]", "", "no_carbon"),
            ("blank", "", "", "blank_smiles"),
            ("invalid_cas", "", "123-45-6", "CAS is invalid"),
        ]
        self.transcript = {
            "schema_version": 2,
            "kind": "native-app-automation-transcript",
            "surface": "native-desktop",
            "executable": os.fspath(Path(os.sys.executable).resolve()),
            "run_nonce": self.run_nonce,
            "target_pid": self.target_pid,
            "checks": [],
            "automation_scopes": [],
        }

    def start(self) -> None:
        self._install_dialog_observer()
        callback = self._wait_for_external_driver if self.external_driver_gate else self.tick
        self.app.after(100, callback)

    def _wait_for_external_driver(self) -> None:
        try:
            if time.monotonic() >= self.deadline:
                raise TimeoutError("external pywinauto driver did not release the native QA gate")
            if not self.external_driver_gate.is_file():
                self.app.after(100, self._wait_for_external_driver)
                return
            payload = json.loads(self.external_driver_gate.read_text(encoding="utf-8"))
            self._require(payload.get("passed") is True, "external pywinauto driver gate is not passing")
            self._require(bool(self.run_nonce), "automation run nonce is missing")
            self._require(payload.get("run_nonce") == self.run_nonce, "external driver gate nonce mismatch")
            self._require(payload.get("target_pid") == self.target_pid, "external driver gate target PID mismatch")
            self._record(
                "external_pywinauto_driver_gate",
                driver_kind=payload.get("driver_kind"),
                run_nonce=self.run_nonce,
                target_pid=self.target_pid,
                real_input_actions=payload.get("real_input_actions"),
                evidence_path=str(self.external_driver_gate),
            )
            self.app.after(100, self.tick)
        except Exception:
            self._finish(False, traceback.format_exc())

    def _body(self, tab=None) -> str:
        tab = tab or self.app.erba_tab
        fields = (
            tab.single_prediction_summary_var,
            tab.single_negative_probability_var,
            tab.single_positive_probability_var,
            tab.single_pic50_var,
            tab.single_ic50_var,
            tab.single_ad_domain_var,
            tab.single_detail_var,
            tab.single_status_var,
        )
        return "\n".join(value.get().strip() for value in fields if value.get().strip())

    def _record(self, check: str, **evidence) -> None:
        self.transcript["checks"].append({"check": check, **evidence, "passed": True})

    def _require(self, condition: bool, message: str) -> None:
        if not condition:
            raise AssertionError(message)

    def _patch(self, owner: object, name: str, replacement: object) -> None:
        self.patches.append((owner, name, getattr(owner, name)))
        setattr(owner, name, replacement)

    def _restore_patches(self) -> None:
        while self.patches:
            owner, name, original = self.patches.pop()
            setattr(owner, name, original)

    def _install_dialog_observer(self) -> None:
        # Native message boxes block unattended Tk automation; record their exact
        # visible title/body while preserving the callback path that invokes them.
        from gui import main_window

        def observe(kind):
            def handler(title, message, **_kwargs):
                self.dialogs.append({"kind": kind, "title": str(title), "message": str(message)})
                return "ok"
            return handler

        self._patch(main_window.messagebox, "showerror", observe("error"))
        self._patch(main_window.messagebox, "showinfo", observe("info"))
        self.transcript["automation_scopes"].append({
            "scope": "messagebox_observer",
            "reason": "record dialog-visible semantics without blocking opt-in automation",
        })

    def _install_mock_cas(self) -> None:
        if self.mock_cas_enabled:
            return
        from gui import main_window

        def mocked_cas_to_smiles(cas: str) -> dict:
            if str(cas).strip() != self.MOCK_CAS:
                raise RuntimeError(f"unexpected deterministic QA CAS: {cas}")
            return {"CanonicalSMILES": self.MOCK_SMILES, "PubChem_CID": "712"}

        self._patch(main_window, "cas_to_smiles", mocked_cas_to_smiles)
        self.mock_cas_enabled = True
        self.transcript["automation_scopes"].append({
            "scope": "mocked_pubchem",
            "cas": self.MOCK_CAS,
            "canonical_smiles": self.MOCK_SMILES,
            "cid": "712",
            "reason": "deterministic opt-in UI callback coverage; no network request",
        })

    def _install_picker(self, save_path: Path) -> None:
        from gui import main_window

        def choose_open_file(**kwargs):
            title = str(kwargs.get("title", ""))
            return self.app.ad_ref_path_var.get() if "AD reference" in title else self.app.model_path_var.get()

        self._patch(main_window.filedialog, "askopenfilename", choose_open_file)
        self._patch(main_window.filedialog, "asksaveasfilename", lambda **_kwargs: str(save_path))
        self.transcript["automation_scopes"].append({
            "scope": "filedialog_selection",
            "path": str(save_path),
            "reason": "exercise callback state transitions without coordinate-driven native dialogs",
        })

    def _finish(self, passed: bool, error: str = "") -> None:
        self._restore_patches()
        self.transcript["passed"] = passed
        self.transcript["duration_seconds"] = time.monotonic() - self.started_at
        if error:
            self.transcript["error"] = error
        self.destination.write_text(
            json.dumps(self.transcript, indent=2, ensure_ascii=False, default=str) + "\n",
            encoding="utf-8",
        )
        hold_seconds = max(1, int(os.environ.get("ER_PREDICTOR_AUTOMATION_HOLD_SECONDS", "30")))
        self.app.after(hold_seconds * 1000, self.app.destroy)

    def _reschedule(self, delay: int = 100) -> None:
        if time.monotonic() >= self.deadline:
            raise TimeoutError(f"native package QA timed out at stage {self.stage}")
        self.app.after(delay, self.tick)

    def _wait_erta_result(self, previous: str = "") -> bool:
        return (
            self.app.status_var.get().startswith("Single prediction done:")
            and self.app.prediction_summary_var.get() != "Prediction: -"
            and (not previous or self.app.result_text.get("1.0", "end").strip() != previous)
        )

    def _wait_dialog(self, count: int) -> bool:
        return len(self.dialogs) > count

    def tick(self) -> None:
        try:
            app = self.app
            tab = app.eralpha_tab
            if self.stage == 0:
                self._require(app.notebook.tab(app.notebook.select(), "text") == "ERTA", "ERTA is not startup tab")
                self._require(Path(app.model_path_var.get()).is_file(), "default ERTA model resource missing")
                self._require(Path(app.ad_ref_path_var.get()).is_file(), "default ERTA AD resource missing")
                self._require(app.options_frame.winfo_ismapped() == 0, "options must start hidden")
                self._require(str(app.run_batch_button.cget("state")) == "normal", "ERTA batch control unavailable")
                self._record("erta_startup_resources_and_controls", model=app.model_path_var.get(), ad_reference=app.ad_ref_path_var.get(), options_visible=False, batch_control_state=app.run_batch_button.cget("state"))
                app.toggle_options()
                app.update_idletasks()
                self._require(app.options_visible and bool(app.options_frame.grid_info()), "options did not open")
                self._require(app.model_path_var.get() and app.ad_ref_path_var.get(), "option resources not visible")
                app.toggle_options()
                app.update_idletasks()
                self._record("erta_options_and_resource_controls", model_entry=app.model_path_var.get(), ad_entry=app.ad_ref_path_var.get(), options_hidden_after_toggle=not app.options_visible and not bool(app.options_frame.grid_info()))
                self.stage = 1
            if self.stage == 1:
                if not app.predictor.is_loaded() or not app.ad_calculator.fitted:
                    self._reschedule()
                    return
                app.cas_var.set("")
                app.smiles_var.set(self.VALID_SMILES)
                app.single_predict_clicked()
                self.stage = 2
                self._reschedule()
                return
            if self.stage == 2:
                if not self._wait_erta_result():
                    self._reschedule()
                    return
                probability = float(app.active_probability_var.get().split(":", 1)[1].strip())
                self._require(abs(probability - 0.9801110625267029) < 1e-5, "ERTA exact valid oracle probability drift")
                self._require(app.ad_domain_var.get() in {"Applicability domain: In-domain", "Applicability domain: Out-of-domain"}, "ERTA AD was not rendered")
                self._require(bool(app.structure_label.cget("image")), "input molecule structure was not rendered")
                self._require(bool(app.nearest_reference_structure_label.cget("image")), "nearest reference structure was not rendered")
                self._require(bool(app.single_ad_graph_label.cget("image")), "single AD graph was not rendered")
                self._require(len(app.prob_canvas.find_all()) > 0, "probability graph canvas is empty")
                self.erta_summary = app.prediction_summary_var.get()
                self.erta_probability = app.active_probability_var.get()
                self._record("erta_exact_valid_single", probability_positive=probability, expected_probability_positive=0.9801110625267029, summary=self.erta_summary, ad=app.ad_domain_var.get(), nearest_reference=app.nearest_reference_var.get(), probability_canvas_items=len(app.prob_canvas.find_all()), input_structure=True, reference_structure=True, ad_graph=True)
                self.erta_previous_text = app.result_text.get("1.0", "end").strip()
                app.smiles_var.set(self.LEGACY_INVALID_SMILES)
                app.single_predict_clicked()
                self.stage = 3
                self._reschedule()
                return
            if self.stage == 3:
                if not self._wait_erta_result(self.erta_previous_text):
                    self._reschedule()
                    return
                detail = app.result_text.get("1.0", "end").strip()
                self._require("Mol valid: False" in detail, "legacy invalid SMILES did not reach ERTA zero-vector path")
                self._require(app.prediction_summary_var.get() in {"Prediction: Positive", "Prediction: Negative"}, "legacy invalid prediction label missing")
                self._record("erta_legacy_invalid_direct_smiles", smiles=self.LEGACY_INVALID_SMILES, rendered_summary=app.prediction_summary_var.get(), positive_probability=app.active_probability_var.get(), details=detail)
                self._install_mock_cas()
                app.cas_var.set(self.MOCK_CAS)
                app.smiles_var.set("")
                app.pubchem_clicked()
                self.stage = 4
                self._reschedule()
                return
            if self.stage == 4:
                if app.smiles_var.get() != self.MOCK_SMILES:
                    self._reschedule()
                    return
                self._require("PubChem found CID 712" in app.status_var.get(), "mocked PubChem callback did not update status")
                self._record("erta_mocked_pubchem_callback", cas=self.MOCK_CAS, smiles=app.smiles_var.get(), status=app.status_var.get())
                self.erta_previous_text = app.result_text.get("1.0", "end").strip()
                app.single_predict_clicked()
                self.stage = 5
                self._reschedule()
                return
            if self.stage == 5:
                if not self._wait_erta_result(self.erta_previous_text):
                    self._reschedule()
                    return
                self._require(f"CAS: {self.MOCK_CAS}" in app.result_text.get("1.0", "end"), "CAS-resolved prediction details missing CAS")
                self._record("erta_mocked_pubchem_prediction", summary=app.prediction_summary_var.get(), probability=app.active_probability_var.get(), canonical_details=app.result_text.get("1.0", "end").strip())
                self.stage = 6
            if self.stage == 6:
                template_path = self.root / "downloaded-template.xlsx"
                self._install_picker(template_path)
                model_before = app.model_path_var.get()
                ad_before = app.ad_ref_path_var.get()
                app.browse_model()
                app.browse_ad_reference()
                self._require(app.model_path_var.get() == model_before and app.ad_ref_path_var.get() == ad_before, "model or AD browse callback changed bundled resource selection")
                self._record("erta_model_ad_browse_callbacks", model=app.model_path_var.get(), ad_reference=app.ad_ref_path_var.get())
                dialogs_before = len(self.dialogs)
                app.download_template_clicked()
                self._require(template_path.is_file(), "template callback did not create workbook")
                self._require(self._wait_dialog(dialogs_before), "template callback did not expose completion dialog")
                template = load_workbook(template_path, read_only=True, data_only=True)
                self._require(template.active.max_row == 1 and [cell.value for cell in template.active[1]] == ["CAS"], "template schema drift")
                self._record("erta_template_callback_and_dialog", output=str(template_path), headers=[cell.value for cell in template.active[1]], dialog=self.dialogs[-1])
                input_path = self.root / "erta-mixed-input.xlsx"
                workbook = Workbook()
                sheet = workbook.active
                sheet.append(["No.", "CAS", "Chemical Name", "SMILES"])
                sheet.append([2, "", "valid", self.VALID_SMILES])
                sheet.append([1, self.MOCK_CAS, "mocked-cas", ""])
                sheet.append([3, "", "legacy-invalid", self.LEGACY_INVALID_SMILES])
                workbook.save(input_path)
                self.erta_batch_input_path = input_path.resolve()
                self.erta_batch_outputs_before = {
                    path.resolve()
                    for path in self.erta_batch_input_path.parent.glob(
                        "erta-mixed-input_*_prediction*.xlsx"
                    )
                }
                app.batch_input_var.set(str(input_path))
                app.batch_input_display_var.set(input_path.name)
                app.batch_destination_var.set(str(self.erta_batch_input_path.parent))
                self.batch_dialog_count = len(self.dialogs)
                app.batch_predict_clicked()
                self.stage = 7
                self._reschedule()
                return
            if self.stage == 7:
                if not self._wait_dialog(self.batch_dialog_count):
                    self._reschedule()
                    return
                self._require(
                    self.erta_batch_input_path is not None,
                    "ERTA batch input path was not recorded",
                )
                input_parent = self.erta_batch_input_path.parent
                outputs = {
                    path.resolve()
                    for path in input_parent.glob(
                        "erta-mixed-input_*_prediction*.xlsx"
                    )
                } - self.erta_batch_outputs_before
                self._require(
                    len(outputs) == 1,
                    f"expected one fresh ERTA workbook, found {len(outputs)}",
                )
                output = outputs.pop()
                self._require(
                    output.parent == input_parent,
                    "ERTA workbook was not saved beside the input workbook",
                )
                workbook = load_workbook(output, read_only=True, data_only=True)
                sheet = workbook.active
                headers = [cell.value for cell in sheet[1]]
                rows = list(sheet.iter_rows(min_row=2, values_only=True))
                self._require(sheet.title == "Sheet1", "legacy ERTA workbook sheet name drift")
                self._require(headers == ["No.", "CAS", "Chemical Name", "SMILES", "Canonical_SMILES", "Mol_valid", "Probability_Negative_0", "Probability_Positive_1", "Prediction", "Prediction_label", "Decision_rule", "AD", "AD_MeanDistance", "AD_DistanceThreshold", "AD_Distance_InDomain", "AD_SimilarityMax", "AD_SimilarityThreshold", "AD_Similarity_InDomain", "AD_PC1", "AD_PC2", "PubChem_CID", "PubChem_status"], "legacy ERTA workbook schema/order drift")
                self._require([row[0] for row in rows] == [2, 1, 3], "ERTA batch input order drift")
                self._require(rows[1][3] == self.MOCK_SMILES and rows[1][-1] == "Found", "mocked CAS batch row was not resolved")
                self._require(rows[2][5] is False, "legacy invalid batch row did not preserve invalid marker")
                graph_paths = sorted(
                    path.name for path in (self.root / "graphs").glob("*.png")
                )
                self._require(graph_paths, "ERTA batch graph files were not generated")
                summary = app.batch_result.get("1.0", "end").strip()
                self._require(
                    f"Output workbook: {output}" in summary
                    and "AD In-domain:" in summary
                    and "Graph directory:" in summary,
                    "ERTA batch completion summary is incomplete",
                )
                self._require(
                    not hasattr(app, "batch_tree")
                    and not hasattr(app, "graph_combo")
                    and not hasattr(app, "graph_label"),
                    "obsolete ERTA batch table/graph preview remains in the UI",
                )
                self._record(
                    "erta_batch_workbook_graphs_and_order",
                    output=str(output),
                    sheet=sheet.title,
                    headers=headers,
                    input_order=[row[0] for row in rows],
                    mocked_cas_status=rows[1][-1],
                    invalid_row_mol_valid=rows[2][5],
                    graph_files=graph_paths,
                    batch_summary=summary,
                    completion_dialog=self.dialogs[-1],
                )
                self._restore_patches()
                self.mock_cas_enabled = False
                self._install_dialog_observer()
                app.notebook.select(tab)
                self.stage = 8
            if self.stage == 8:
                task, subtype = self.routes[self.route_index]
                app.notebook.select(tab)
                if self.route_index == 0:
                    dialogs_before = len(self.dialogs)
                    tab._show_evidence_details()
                    self._require(self._wait_dialog(dialogs_before), "ERBA evidence-details dialog was not exposed")
                    self._require(bool(tab.single_evidence_caveat_var.get()), "ERBA compact evidence caveat missing")
                    self._record("erba_evidence_caveat_and_dialog", compact_caveat=tab.single_evidence_caveat_var.get(), dialog=self.dialogs[-1])
                tab._route_changed("single")
                tab.smiles_var.set("Oc1ccccc1")
                tab.cas_var.set("")
                tab.single_predict_clicked()
                self.stage = 9
                self._reschedule()
                return
            if self.stage == 9:
                if (
                    tab._active_single_request_id is not None
                    or tab._active_ad_request_id is not None
                    or "Evaluating" in tab.single_ad_domain_var.get()
                ):
                    self._reschedule()
                    return
                body = self._body(tab)
                task, subtype = self.routes[self.route_index]
                self._require(tab.single_status_var.get() == "ok" and body, "ERBA route did not render success")
                self._require(
                    tab.single_ad_domain_var.get()
                    in (
                        "Applicability domain: In-domain",
                        "Applicability domain: Out-of-domain",
                    ),
                    "ERBA route did not complete applicability-domain analysis",
                )
                self._require(
                    bool(tab._single_ad_graph_image)
                    and "Morgan similarity=" in tab.single_nearest_reference_var.get(),
                    "ERBA route did not render AD graph and nearest reference",
                )
                self._record("erba_route", task=task.value, subtype=subtype.value, status=tab.single_status_var.get(), rendered_result=body)
                self.route_index += 1
                if self.route_index < len(self.routes):
                    self.stage = 8
                    self._reschedule()
                    return
                self.stage = 10
            if self.stage == 10:
                tab = app.eralpha_tab
                app.notebook.select(tab)
                name, smiles, cas, _expected = self.invalid_cases[self.invalid_index]
                tab.single_task_var.set(ERBATask.CLASSIFICATION.value)
                tab.single_subtype_var.set(ERBASubtype.ER_ALPHA.value)
                tab.smiles_var.set(smiles)
                tab.cas_var.set(cas)
                tab.single_predict_clicked()
                self.stage = 11
                self._reschedule()
                return
            if self.stage == 11:
                if tab._active_single_request_id is not None:
                    self._reschedule()
                    return
                body = self._body(tab)
                self._require(tab.single_status_var.get() != "ok", f"ERBA invalid {self.invalid_cases[self.invalid_index][0]} succeeded")
                self._record(f"erba_invalid_{self.invalid_cases[self.invalid_index][0]}", rendered_result=body, rendered_status=tab.single_status_var.get())
                self.invalid_index += 1
                if self.invalid_index < len(self.invalid_cases):
                    self.stage = 10
                    self._reschedule()
                    return
                self.stage = 12
            if self.stage == 12:
                app.notebook.select(tab)
                tab.notebook.select(tab.batch_tab)
                tab.batch_task_var.set(ERBATask.CLASSIFICATION.value)
                tab.batch_subtype_var.set(ERBASubtype.ER_ALPHA.value)
                input_path = self.root / "erba-alpha-batch-input.xlsx"
                workbook = Workbook()
                sheet = workbook.active
                sheet.title = "Original_Input"
                sheet.append(["Row_ID", "SMILES", "Analyst_Note"])
                sheet.append(["native-valid", "Oc1ccccc1", "preserve-valid"])
                sheet.append(["native-invalid", "[*]CC", "preserve-invalid"])
                workbook.save(input_path)
                self.erba_batch_input_path = input_path.resolve()
                input_parent = self.erba_batch_input_path.parent
                self.erba_batch_outputs_before = {
                    path.resolve()
                    for path in input_parent.glob(
                        "ERBA_classification_er_alpha_results*.xlsx"
                    )
                }
                self.erba_batch_graph_dirs_before = {
                    path.resolve()
                    for path in input_parent.glob(
                        "ERBA_classification_er_alpha_results*_graphs*"
                    )
                    if path.is_dir()
                }
                tab.batch_input_var.set(str(self.erba_batch_input_path))
                tab.batch_destination_var.set(str(input_parent))
                tab.batch_predict_clicked()
                self._require(
                    tab._active_batch_request_id is not None,
                    f"ERalpha batch callback did not start: {tab.batch_status_var.get()}",
                )
                self.stage = 13
                self._reschedule()
                return
            if self.stage == 13:
                if tab._active_batch_request_id is not None:
                    self._reschedule()
                    return
                self._require(
                    self.erba_batch_input_path is not None,
                    "ERalpha batch input path was not recorded",
                )
                input_parent = self.erba_batch_input_path.parent
                outputs = {
                    path.resolve()
                    for path in input_parent.glob(
                        "ERBA_classification_er_alpha_results*.xlsx"
                    )
                } - self.erba_batch_outputs_before
                self._require(
                    len(outputs) == 1,
                    f"expected one new ERalpha workbook, found {len(outputs)}: "
                    f"{tab.batch_status_var.get()}",
                )
                output = outputs.pop()
                self._require(
                    output.parent == input_parent
                    and tab.batch_destination_var.get() == str(input_parent),
                    "ERalpha workbook destination is not the input workbook parent",
                )
                graph_dirs = {
                    path.resolve()
                    for path in input_parent.glob(
                        "ERBA_classification_er_alpha_results*_graphs*"
                    )
                    if path.is_dir()
                } - self.erba_batch_graph_dirs_before
                self._require(
                    len(graph_dirs) == 1,
                    f"expected one new input-adjacent ERalpha graph directory, found {len(graph_dirs)}",
                )
                graph_directory = graph_dirs.pop()
                graph_files = {
                    path.name for path in graph_directory.iterdir() if path.is_file()
                }
                self._require(
                    graph_directory.parent == input_parent
                    and graph_files == self.ERBA_BATCH_GRAPH_FILES,
                    "ERalpha batch graphs are incomplete or outside the input parent",
                )

                workbook = load_workbook(output, read_only=True, data_only=True)
                sheet_names = workbook.sheetnames
                self._require(
                    sheet_names == ["Predictions", "Guide", "Input", "Metadata"]
                    and workbook.active.title == "Predictions",
                    "ERalpha batch workbook is not Predictions-first and active",
                )
                prediction_sheet = workbook["Predictions"]
                prediction_headers = [cell.value for cell in prediction_sheet[1]]
                self._require(
                    all(column in prediction_headers for column in self.ERBA_BATCH_AD_COLUMNS),
                    "ERalpha batch workbook is missing applicability-domain columns",
                )
                prediction_rows = [
                    dict(zip(prediction_headers, row))
                    for row in prediction_sheet.iter_rows(min_row=2, values_only=True)
                ]
                self._require(
                    len(prediction_rows) == 2
                    and prediction_rows[0]["Analyst_Note"] == "preserve-valid"
                    and prediction_rows[1]["Analyst_Note"] == "preserve-invalid",
                    "ERalpha primary predictions did not retain safe input fields and row order",
                )
                self._require(
                    prediction_rows[0]["Result_Status"] == "Predicted"
                    and prediction_rows[0]["AD"] in {"In-domain", "Out-of-domain"}
                    and prediction_rows[1]["Result_Status"] == "Not predicted"
                    and all(
                        prediction_rows[1][column] is None
                        for column in self.ERBA_BATCH_AD_COLUMNS
                    ),
                    "ERalpha prediction/AD row mapping is incorrect",
                )

                input_sheet = workbook["Input"]
                input_headers = [cell.value for cell in input_sheet[1]]
                input_rows = list(input_sheet.iter_rows(min_row=2, values_only=True))
                self._require(
                    input_headers == ["Row_ID", "SMILES", "Analyst_Note"]
                    and input_rows
                    == [
                        ("native-valid", "Oc1ccccc1", "preserve-valid"),
                        ("native-invalid", "[*]CC", "preserve-invalid"),
                    ],
                    "ERalpha Input sheet did not preserve the supplied input",
                )

                metadata_sheet = workbook["Metadata"]
                metadata_headers = [cell.value for cell in metadata_sheet[1]]
                metadata_values = [
                    cell.value for cell in next(metadata_sheet.iter_rows(min_row=2, max_row=2))
                ]
                metadata = dict(zip(metadata_headers, metadata_values))
                expected_spec = tab.catalog[
                    (ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA)
                ]
                hex_digits = set("0123456789abcdef")
                self._require(
                    metadata.get("Excel_Contract_ID")
                    == ERBA_BINDING_CLASSIFICATION_EXCEL_CONTRACT_ID
                    and metadata.get("Model_ID") == expected_spec.model_id
                    and metadata.get("Model_SHA256") == expected_spec.sha256
                    and metadata.get("Performance_Evidence_Scope")
                    == "internal_historically_exposed"
                    and bool(metadata.get("Evidence_Caveat"))
                    and all(
                        isinstance(metadata.get(column), str)
                        and len(metadata[column]) == 64
                        and set(metadata[column]) <= hex_digits
                        for column in self.ERBA_PROVENANCE_COLUMNS
                    ),
                    "ERalpha batch metadata did not preserve catalog model provenance",
                )
                workbook.close()

                binding_count = sum(
                    row["binding_label"] == "binding" for row in prediction_rows
                )
                non_binding_count = sum(
                    row["binding_label"] == "non_binding" for row in prediction_rows
                )
                not_predicted_count = sum(
                    row["Result_Status"] != "Predicted" for row in prediction_rows
                )
                summary = tab.batch_result.get("1.0", "end").strip()
                self._require(
                    tab.batch_status_var.get() == f"ERBA batch complete: {output}"
                    and f"Total rows: {len(prediction_rows)}" in summary
                    and f"Binding: {binding_count}" in summary
                    and f"Non-binding: {non_binding_count}" in summary
                    and f"Not predicted: {not_predicted_count}" in summary
                    and f"Output workbook: {output}" in summary
                    and f"Graph files: {len(graph_files)}" in summary
                    and f"Graph directory: {graph_directory}" in summary,
                    "ERalpha batch completion did not expose binding and artifact summaries",
                )
                self._record(
                    "erba_eralpha_batch_workbook_ad_graphs_and_summary",
                    input=str(self.erba_batch_input_path),
                    output=str(output),
                    sheets=sheet_names,
                    prediction_headers=prediction_headers,
                    preserved_input_rows=[list(row) for row in input_rows],
                    metadata=metadata,
                    graph_directory=str(graph_directory),
                    graph_files=sorted(graph_files),
                    binding_count=binding_count,
                    non_binding_count=non_binding_count,
                    not_predicted_count=not_predicted_count,
                    batch_summary=summary,
                )
                self.stage = 14
            if self.stage == 14:
                self._require(
                    app.eralpha_tab._selected_route("single")
                    == (ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA),
                    "ERalpha page is not fixed to classification",
                )
                alpha_model = app.eralpha_tab.catalog[
                    (ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA)
                ].model_id
                self._require(
                    not hasattr(app, "erbeta_tab")
                    and app.notebook.tabs()
                    and "ERbeta"
                    not in [
                        app.notebook.tab(tab_id, "text")
                        for tab_id in app.notebook.tabs()
                    ],
                    "ERbeta remains exposed in the FDA application",
                )
                self._record(
                    "fixed_eralpha_classification_tab",
                    eralpha_model=alpha_model,
                    erbeta_present=False,
                )
                app.notebook.select(app.erta_tab)
                self._require(app.prediction_summary_var.get() != self.erta_summary or app.active_probability_var.get() != self.erta_probability, "ERTA state was not independently updated after CAS journey")
                self._require(app.last_batch_result is not None and app.notebook.tab(app.notebook.select(), "text") == "ERTA", "ERTA batch state lost after ERBA journey")
                self._record("erta_erba_erta_state_isolation", erta_summary_after_return=app.prediction_summary_var.get(), erta_probability_after_return=app.active_probability_var.get(), erba_result_retained=bool(self._body(app.eralpha_tab)), erta_batch_rows=len(app.last_batch_result))
                self._record(
                    "read_only_install_and_writable_state",
                    resource_root=str(app.project_root),
                    output_root=str(app.output_root),
                    state_root=str(app.state_root),
                )
                app.deiconify()
                app.state("normal")
                app.update()
                self._finish(True)
        except Exception:
            self._finish(False, traceback.format_exc())
