"""ERBA-only Tkinter journey; it never exposes legacy ERTA models or AD tools."""
from __future__ import annotations

from collections.abc import Mapping
import json
import os
import tempfile
import threading
import re
import time
from dataclasses import asdict
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from typing import Callable

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

from core.contracts import (
    ERBA_BINDING_CLASSIFICATION_EXCEL_CONTRACT_ID,
    ERBA_CLASSIFICATION_PREDICTION_COLUMNS,
    ERBA_REGRESSION_PREDICTION_COLUMNS,
    ERBA_CANONICAL_INPUT_COLUMNS,
    ERBA_FATAL_ARTIFACT_STATUS_CODES,
    ERBA_INPUT_HEADER_ALIASES,
    ERBA_INPUT_SHEET_NAME,
    ERBA_GUIDE_SHEET_NAME,
    ERBA_METADATA_COLUMNS,
    ERBA_METADATA_SHEET_NAME,
    ERBA_PERFORMANCE_EVIDENCE_SCOPE,
    ERBA_REGRESSION_EVIDENCE_SCOPE,
    ERBA_PREDICTIONS_SHEET_NAME,
    ERBA_ERALPHA_IC50_EXCEL_CONTRACT_ID,
    HISTORICAL_EXPOSURE_CAVEAT_COMPACT,
    HISTORICAL_EXPOSURE_CAVEAT_FULL,
    REGRESSION_EVIDENCE_CAVEAT_COMPACT,
    REGRESSION_EVIDENCE_CAVEAT_FULL,
    ERBAArtifactSpec,
    ERBARequest,
    ERBAStatusCode,
    ERBASubtype,
    ERBATask,
    output_columns,
)
from core.erba_predictor import ERBAPredictor
from core.erba_ad import ERBAApplicabilityDomain
from core.graph import save_single_ad_plot
from core.molecule_image import save_molecule_image
from core.pubchem import cas_to_smiles
from core.paths import validate_mutable_directory

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False

_RELEASED = "released"
_CAS_PATTERN = re.compile(r"^\d{2,7}-\d{2}-\d$")


def load_erba_catalog(path: str | Path):
    """Parse only released, fixed-catalog routes; malformed entries fail closed."""
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(payload, dict) or payload.get("schema_version") != 2:
        raise ValueError("ERBA catalog schema version must be 2.")
    routes = payload.get("routes")
    if not isinstance(routes, list):
        raise ValueError("ERBA catalog routes are missing.")
    catalog = {}
    for route in routes:
        if not isinstance(route, dict) or route.get("release_status") != _RELEASED:
            continue
        try:
            task = ERBATask(route["task"])
            subtype = ERBASubtype(route["subtype"])
            artifact = route.get("artifact", route)
            spec = ERBAArtifactSpec(
                relative_path=str(artifact["relative_path"]), size_bytes=int(artifact["size_bytes"]),
                sha256=str(artifact["sha256"]), model_id=str(artifact["model_id"]),
            )
        except (KeyError, TypeError, ValueError) as error:
            raise ValueError("ERBA catalog contains an invalid released route.") from error
        if task is ERBATask.IC50_REGRESSION and subtype is ERBASubtype.ER_BETA:
            continue
        if (task, subtype) in catalog:
            raise ValueError("ERBA catalog contains a duplicate released route.")
        catalog[(task, subtype)] = spec
    parity = payload.get("regression_parity_approved") is True
    return catalog, parity, payload


def allocate_output_path(output_dir: str | Path, task: ERBATask, subtype: ERBASubtype) -> Path:
    """Reserve a unique result name so an existing output is never overwritten."""
    directory = Path(output_dir)
    directory.mkdir(parents=True, exist_ok=True)
    stem = f"ERBA_{task.value}_{subtype.value}_results"
    for index in range(1, 10000):
        suffix = "" if index == 1 else f"_{index}"
        candidate = directory / f"{stem}{suffix}.xlsx"
        try:
            fd = os.open(candidate, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            continue
        os.close(fd)
        return candidate
    raise RuntimeError("Could not allocate a non-overwriting ERBA output filename.")
def validate_cas(cas: str) -> bool:
    """Accept only a syntactically valid CAS Registry Number with a valid check digit."""
    value = cas.strip()
    if not _CAS_PATTERN.fullmatch(value):
        return False
    digits = value.replace("-", "")
    return sum(int(digit) * factor for factor, digit in enumerate(reversed(digits[:-1]), 1)) % 10 == int(digits[-1])
def cas_lookup_smiles(cas: str) -> str:
    """Extract a usable SMILES string from the PubChem mapping contract."""
    lookup = cas_to_smiles(cas)
    if not isinstance(lookup, Mapping):
        raise ValueError("CAS lookup returned an invalid response.")
    for field in ("CanonicalSMILES", "IsomericSMILES"):
        value = lookup.get(field)
        if isinstance(value, str) and value.strip():
            return value.strip()
    raise ValueError("CAS lookup returned neither CanonicalSMILES nor IsomericSMILES.")




def read_batch_input(path: str | Path) -> pd.DataFrame:
    """Read the first worksheet as strings without pandas mangling duplicate headers."""
    raw = pd.read_excel(path, sheet_name=0, header=None, dtype=str, keep_default_na=False)
    if raw.empty:
        return pd.DataFrame()
    frame = raw.iloc[1:].copy()
    frame.columns = [str(value) for value in raw.iloc[0].tolist()]
    return frame.reset_index(drop=True)


def canonicalize_batch_input(frame: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return canonical request fields and untouched source rows, rejecting ambiguity."""
    normalized = {}
    for header in frame.columns:
        normalized.setdefault(str(header).strip().lower(), []).append(header)
    canonical = {}
    for target, aliases in ERBA_INPUT_HEADER_ALIASES.items():
        matches = [header for alias in aliases for header in normalized.get(alias, [])]
        if len(matches) > 1:
            raise ValueError(f"Duplicate canonical {target} headers are not allowed.")
        if matches:
            canonical[target] = matches[0]

    canonical_rows = pd.DataFrame(index=frame.index)
    canonical_rows["Row_ID"] = (
        frame[canonical["Row_ID"]].fillna("").astype(str)
        if "Row_ID" in canonical else [str(index) for index in frame.index]
    )
    for target in ("CAS", "SMILES"):
        canonical_rows[target] = (
            frame[canonical[target]].fillna("").astype(str)
            if target in canonical else ""
        )
    return canonical_rows.loc[:, ERBA_CANONICAL_INPUT_COLUMNS].copy(), frame.copy()


def excel_contract_id(task: ERBATask) -> str:
    return ERBA_BINDING_CLASSIFICATION_EXCEL_CONTRACT_ID if task is ERBATask.CLASSIFICATION else ERBA_ERALPHA_IC50_EXCEL_CONTRACT_ID


def metadata_row(task: ERBATask, spec: ERBAArtifactSpec, catalog_payload: dict) -> dict:
    shared = dict(catalog_payload.get("metadata") or {})
    shared.update(catalog_payload.get("provenance") or {})
    route_metadata = (catalog_payload.get("route_metadata") or {}).get(spec.model_id, {})
    provenance = {**shared, **route_metadata}
    if task is ERBATask.CLASSIFICATION:
        required = (
            "protocol_sha256", "source_manifest_sha256", "historical_exposure_manifest_sha256",
            "split_manifest_sha256", "nested_cv_sha256", "internal_resplit_sha256",
            "report_sha256", "caveat_sha256",
        )
    else:
        required = ("source_manifest_sha256", "preprocessing_parity_sha256", "report_sha256")
    missing = [name for name in required if not re.fullmatch(r"[0-9a-f]{64}", str(provenance.get(name, "")))]
    if missing:
        raise ValueError(f"ERBA catalog provenance is incomplete: {', '.join(missing)}")
    classification = task is ERBATask.CLASSIFICATION
    return {
        "Excel_Contract_ID": excel_contract_id(task), "Model_ID": spec.model_id, "Model_SHA256": spec.sha256,
        "Protocol_SHA256": provenance.get("protocol_sha256", ""),
        "Source_Manifest_SHA256": provenance["source_manifest_sha256"],
        "Historical_Exposure_Manifest_SHA256": provenance.get("historical_exposure_manifest_sha256", ""),
        "Split_Manifest_SHA256": provenance.get("split_manifest_sha256", ""),
        "Nested_CV_SHA256": provenance.get("nested_cv_sha256", ""),
        "Internal_Resplit_SHA256": provenance.get("internal_resplit_sha256", ""),
        "Preprocessing_Parity_SHA256": provenance.get("preprocessing_parity_sha256", ""),
        "Report_SHA256": provenance["report_sha256"], "Caveat_SHA256": provenance.get("caveat_sha256", ""),
        "Performance_Evidence_Scope": (
            ERBA_PERFORMANCE_EVIDENCE_SCOPE if classification else ERBA_REGRESSION_EVIDENCE_SCOPE
        ),
        "Evidence_Caveat": HISTORICAL_EXPOSURE_CAVEAT_FULL if classification else "",
    }


def result_row(result, task: ERBATask, row_id: str, cas: str) -> dict:
    row = {key: getattr(result, key, "") for key in output_columns(task)}
    row["Row_ID"], row["CAS"] = row_id, cas
    if result.status_code is not ERBAStatusCode.OK:
        for column in ("non_binding_probability", "binding_probability", "binding_label", "pic50", "ic50_nm"):
            if column in row:
                row[column] = ""
    row.update(result_explanation(result))
    return row


def fatal_artifact_failure(results) -> bool:
    return any(result.status_code in ERBA_FATAL_ARTIFACT_STATUS_CODES for result in results)


def result_explanation(result) -> dict[str, str]:
    """Translate stable technical statuses into concise workbook-facing guidance."""
    if result.status_code is ERBAStatusCode.OK:
        return {
            "Result_Status": "Predicted",
            "Reason_Category": "Prediction available",
            "Reason_Description": "A model prediction was generated.",
            "Recommended_Action": "Review the prediction with the workbook caveat.",
        }
    message = (result.status_message or "").lower()
    if "cas is invalid" in message:
        category, description, action = (
            "Invalid CAS",
            "The CAS Registry Number format or check digit is invalid.",
            "Correct the CAS number or provide a valid SMILES.",
        )
    elif "cas lookup" in message or "pubchem" in message:
        category, description, action = (
            "PubChem lookup unavailable",
            "PubChem did not return a usable SMILES for this CAS number.",
            "Provide a verified SMILES and run the batch again.",
        )
    elif result.status_code is ERBAStatusCode.NO_CARBON:
        category, description, action = (
            "No carbon / inorganic",
            "The structure contains no carbon and is outside the organic model domain.",
            "Provide an organic parent structure where applicable.",
        )
    elif result.status_code is ERBAStatusCode.METAL_RETAINED:
        category, description, action = (
            "Unsupported metal-containing structure",
            "A metal remains after salt and largest-fragment handling.",
            "Provide an organic parent structure without the metal component.",
        )
    elif result.status_code in {
        ERBAStatusCode.INVALID_SMILES, ERBAStatusCode.BLANK_SMILES,
        ERBAStatusCode.WILDCARD_SMILES, ERBAStatusCode.CXSMILES_NOT_ALLOWED,
    }:
        category, description, action = (
            "Invalid or ambiguous SMILES",
            "The supplied SMILES is blank, invalid, or not supported by the model input policy.",
            "Provide one valid, unambiguous organic SMILES.",
        )
    elif result.status_code in ERBA_FATAL_ARTIFACT_STATUS_CODES or result.status_code in {
        ERBAStatusCode.PREDICTION_FAILED, ERBAStatusCode.UNSUPPORTED_COMBINATION,
        ERBAStatusCode.REGRESSION_PARITY_NOT_APPROVED,
    }:
        category, description, action = (
            "System or model failure",
            "The prediction could not be completed because of a system or model error.",
            "Record the technical status and contact the application support team.",
        )
    else:
        category, description, action = (
            "Prediction unavailable",
            "The model could not produce a prediction for this row.",
            "Review the input and technical status, then correct and rerun.",
        )
    return {
        "Result_Status": "Not predicted",
        "Reason_Category": category,
        "Reason_Description": description,
        "Recommended_Action": action,
    }


def _progress(callback, stage: str, current: int, total: int, percent: int) -> None:
    if callback is not None:
        callback(stage, current, total, percent)


def _format_workbook(writer, predictions: pd.DataFrame, input_rows: pd.DataFrame, metadata: pd.DataFrame) -> None:
    """Apply readable presentation without changing data-sheet headers or values."""
    guide_rows = [
        ("ERalpha batch prediction guide", "ERalpha direct receptor-binding classification output."),
        ("Predictions", "One row per input row. Numeric probabilities and labels are blank when Result_Status is Not predicted."),
        ("Input", "Original first-sheet input, retained in its original row order."),
        ("Metadata", "Model identity, integrity hashes, and evidence caveat for this export."),
        ("Caveat", HISTORICAL_EXPOSURE_CAVEAT_FULL),
        ("Total rows", str(len(predictions))),
        ("Predicted rows", str((predictions["Result_Status"] == "Predicted").sum())),
        ("Not predicted rows", str((predictions["Result_Status"] != "Predicted").sum())),
        ("Reason counts", "Counts below include not-predicted categories only."),
    ]
    for category, count in (
        predictions.loc[predictions["Result_Status"] != "Predicted", "Reason_Category"]
        .value_counts()
        .sort_index()
        .items()
    ):
        guide_rows.append((f"Reason: {category}", str(count)))
    pd.DataFrame(guide_rows, columns=("Topic", "Details")).to_excel(
        writer, sheet_name=ERBA_GUIDE_SHEET_NAME, index=False
    )

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    warning_fill = PatternFill("solid", fgColor="FFF2CC")
    fatal_fill = PatternFill("solid", fgColor="F4CCCC")
    for sheet_name, frame in (
        (ERBA_GUIDE_SHEET_NAME, pd.DataFrame(guide_rows)),
        (ERBA_PREDICTIONS_SHEET_NAME, predictions),
        (ERBA_INPUT_SHEET_NAME, input_rows),
        (ERBA_METADATA_SHEET_NAME, metadata),
    ):
        sheet = writer.book[sheet_name]
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        for column_cells in sheet.columns:
            letter = column_cells[0].column_letter
            width = max(len(str(cell.value or "")) for cell in column_cells)
            sheet.column_dimensions[letter].width = min(max(width + 2, 12), 48)
            for cell in column_cells[1:]:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        if sheet_name == ERBA_PREDICTIONS_SHEET_NAME:
            status_column = predictions.columns.get_loc("Result_Status") + 1
            technical_column = predictions.columns.get_loc("status_code") + 1
            reason_column = predictions.columns.get_loc("Reason_Category") + 1
            for row in range(2, sheet.max_row + 1):
                status = sheet.cell(row, status_column).value
                technical = sheet.cell(row, technical_column).value
                reason = sheet.cell(row, reason_column).value
                fill = (
                    fatal_fill
                    if technical in {code.value for code in ERBA_FATAL_ARTIFACT_STATUS_CODES}
                    or reason == "System or model failure"
                    else warning_fill
                )
                if status != "Predicted":
                    for cell in sheet[row]:
                        cell.fill = fill
    writer.book._sheets = [
        writer.book[ERBA_GUIDE_SHEET_NAME],
        writer.book[ERBA_PREDICTIONS_SHEET_NAME],
        writer.book[ERBA_INPUT_SHEET_NAME],
        writer.book[ERBA_METADATA_SHEET_NAME],
    ]
def export_erba_batch(
    input_path: str | Path,
    output_dir: str | Path,
    task: ERBATask,
    subtype: ERBASubtype,
    predictor: ERBAPredictor,
    spec: ERBAArtifactSpec,
    catalog_payload: dict,
    progress_callback: Callable[[str, int, int, int], None] | None = None,
) -> tuple[Path, int]:
    """Predict a first-sheet ERBA workbook and atomically publish a new result workbook."""
    _progress(progress_callback, "Reading input workbook", 0, 0, 0)
    output_dir = validate_mutable_directory(output_dir)
    preflight = predictor.preflight(task, subtype)
    if not preflight.ready:
        raise RuntimeError(
            f"ERBA route preflight failed ({preflight.status_code.value}): {preflight.status_message}"
        )
    frame = read_batch_input(input_path)
    if frame.empty:
        raise ValueError("ERBA batch input has no rows; no output was published.")
    canonical_rows, input_rows = canonicalize_batch_input(frame)
    total = len(canonical_rows)
    resolved_rows = []
    _progress(progress_callback, "Resolving CAS/SMILES", 0, total, 10)
    for position, (_, values) in enumerate(canonical_rows.iterrows()):
        row_id, cas, direct_smiles = (
            str(values[column]).strip() for column in ERBA_CANONICAL_INPUT_COLUMNS
        )
        lookup_error, smiles = "", direct_smiles
        if not smiles and cas:
            if not validate_cas(cas):
                lookup_error = "CAS is invalid: expected a valid CAS Registry Number check digit."
            else:
                try:
                    smiles = cas_lookup_smiles(cas)
                except Exception as error:
                    lookup_error = f"CAS lookup failed: {error}"
        resolved_rows.append((position, row_id, cas, smiles, lookup_error))
        _progress(progress_callback, "Resolving CAS/SMILES", position + 1, total, 10 + int(30 * (position + 1) / total))
    results = []
    _progress(progress_callback, "Predicting", 0, total, 40)
    for position, row_id, cas, smiles, lookup_error in resolved_rows:
        if lookup_error:
            result = predictor.predict(ERBARequest("", task, subtype, position))
            result = result.__class__(**{
                **asdict(result),
                "status_code": ERBAStatusCode.PREDICTION_FAILED,
                "status_message": lookup_error,
            })
        else:
            result = predictor.predict(ERBARequest(smiles, task, subtype, position))
        results.append((result, row_id, cas))
        _progress(progress_callback, "Predicting", position + 1, total, 40 + int(45 * (position + 1) / total))

    if fatal_artifact_failure([result for result, _, _ in results]):
        raise RuntimeError("ERBA artifact validation failed; no output was published.")

    prediction_columns = (
        ERBA_CLASSIFICATION_PREDICTION_COLUMNS
        if task is ERBATask.CLASSIFICATION
        else ERBA_REGRESSION_PREDICTION_COLUMNS
    )
    predictions = pd.DataFrame(
        [result_row(result, task, row_id, cas) for result, row_id, cas in results],
        columns=prediction_columns,
    )
    destination = allocate_output_path(output_dir, task, subtype)
    _progress(progress_callback, "Writing workbook", total, total, 90)
    temp_fd, temp_name = tempfile.mkstemp(suffix=".xlsx", dir=str(destination.parent))
    os.close(temp_fd)
    try:
        with pd.ExcelWriter(temp_name, engine="openpyxl") as writer:
            predictions.to_excel(writer, sheet_name=ERBA_PREDICTIONS_SHEET_NAME, index=False)
            input_rows.to_excel(writer, sheet_name=ERBA_INPUT_SHEET_NAME, index=False)
            metadata = pd.DataFrame(
                [metadata_row(task, spec, catalog_payload)],
                columns=ERBA_METADATA_COLUMNS,
            )
            metadata.to_excel(writer, sheet_name=ERBA_METADATA_SHEET_NAME, index=False)
            _format_workbook(writer, predictions, input_rows, metadata)
        os.replace(temp_name, destination)
    except Exception:
        Path(temp_name).unlink(missing_ok=True)
        Path(destination).unlink(missing_ok=True)
        raise
    _progress(progress_callback, "Completed", total, total, 100)
    return destination, len(results)


class ErbaTab(ttk.Frame):
    def __init__(
        self,
        parent,
        project_root: str,
        output_root: str | None = None,
        event_log=None,
        structure_editor=None,
        fixed_subtype: ERBASubtype = ERBASubtype.ER_ALPHA,
    ):
        super().__init__(parent)
        self.project_root = Path(project_root)
        self.output_root = Path(output_root) if output_root else self.project_root / "output"
        self.event_log = event_log
        self.structure_editor = structure_editor
        self.fixed_subtype = fixed_subtype
        self.catalog_path = self.project_root / "models" / "erba" / "catalog.v2.json"
        self.catalog = {}
        self.catalog_payload = {}
        self.predictor = None
        self.catalog_error = ""
        self.regression_parity_approved = False
        self._request_id = 0
        self._started_at = {}
        self._load_catalog()
        self.single_task_var = tk.StringVar(value=ERBATask.CLASSIFICATION.value)
        self.single_subtype_var = tk.StringVar(value=fixed_subtype.value)
        self.batch_task_var = tk.StringVar(value=ERBATask.CLASSIFICATION.value)
        self.batch_subtype_var = tk.StringVar(value=fixed_subtype.value)
        self.smiles_var = tk.StringVar()
        self.cas_var = tk.StringVar()
        self.single_status_var = tk.StringVar(value=self.catalog_error or "Ready")
        self.single_prediction_summary_var = tk.StringVar(value="Prediction: -")
        self.single_negative_probability_var = tk.StringVar(value="Negative probability: -")
        self.single_positive_probability_var = tk.StringVar(value="Positive probability: -")
        self.single_pic50_var = tk.StringVar(value="pIC50: -")
        self.single_ic50_var = tk.StringVar(value="IC50: - nM")
        self.single_ad_domain_var = tk.StringVar(
            value="Applicability domain: Not evaluated"
        )
        self.single_detail_var = tk.StringVar(value="Awaiting prediction.")
        self.batch_input_var = tk.StringVar()
        self.batch_output_var = tk.StringVar(value=str(self.output_root))
        self.single_evidence_caveat_var = tk.StringVar(value=HISTORICAL_EXPOSURE_CAVEAT_COMPACT)
        self.batch_status_var = tk.StringVar(value=self.catalog_error or "Ready")
        self.batch_progress_var = tk.StringVar(value="0% - Ready")
        self.batch_progress_value = tk.DoubleVar(value=0)
        self._active_single_request_id = None
        self._active_ad_request_id = None
        self._active_batch_request_id = None
        self._single_structure_image = None
        self._nearest_reference_structure_image = None
        self._single_ad_graph_image = None
        self._single_detail_lines = []
        self.erba_ad = ERBAApplicabilityDomain(self.project_root)
        self.options_visible = False
        self._build_ui()
        self.single_task_var.trace_add("write", lambda *_: self._route_changed("single"))
        self.single_subtype_var.trace_add("write", lambda *_: self._route_changed("single"))
        self.batch_task_var.trace_add("write", lambda *_: self._route_changed("batch"))
        self.batch_subtype_var.trace_add("write", lambda *_: self._route_changed("batch"))
        self._route_changed()

    def _emit(self, event: str, **fields):
        event_log = getattr(self, "event_log", None)
        if event_log is not None:
            event_log.emit(event, **fields)

    def _workflow_tab_changed(self, _event=None):
        selected = self.notebook.select()
        tab = self.notebook.tab(selected, "text") if selected else ""
        self._emit("workflow.tab_changed", workflow="erba", tab=tab)

    def _load_catalog(self):
        try:
            self.catalog, self.regression_parity_approved, self.catalog_payload = load_erba_catalog(self.catalog_path)
            self.predictor = ERBAPredictor(self.catalog_path.parent, self.catalog,
                                           regression_parity_approved=self.regression_parity_approved)
            provenance = {
                key: value
                for section in ("metadata", "provenance")
                for key, value in (self.catalog_payload.get(section) or {}).items()
                if str(key).endswith("_sha256")
            }
            self._emit(
                "model.catalog_loaded",
                workflow="erba",
                route_count=len(self.catalog),
                status="ok",
                **provenance,
            )
            if not self.catalog:
                self.catalog_error = "ERBA catalog has no released routes. Predictions are unavailable."
        except Exception as error:
            self.catalog_error = f"ERBA catalog unavailable: {error}"
            self._emit(
                "model.catalog_loaded",
                workflow="erba",
                route_count=0,
                status="failed",
                exception_type=type(error).__name__,
            )

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        summary = ttk.Frame(self)
        summary.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 2))
        summary.columnconfigure(0, weight=1)
        ttk.Button(summary, text="Options", command=self.toggle_options).grid(row=0, column=1, sticky="e")
        self.options_frame = ttk.LabelFrame(self, text="Options")
        self.options_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(2, 8))
        self.options_frame.columnconfigure(1, weight=1)
        receptor_name = (
            "Estrogen receptor alpha (ERα)"
            if self.fixed_subtype is ERBASubtype.ER_ALPHA
            else "Estrogen receptor beta (ERβ)"
        )
        ttk.Label(self.options_frame, text="Task").grid(
            row=0, column=0, sticky="w", padx=6, pady=5
        )
        ttk.Label(self.options_frame, text="Binding classification").grid(
            row=0, column=1, sticky="w", padx=6, pady=5
        )
        ttk.Label(self.options_frame, text="Receptor").grid(
            row=0, column=2, sticky="w", padx=6, pady=5
        )
        ttk.Label(self.options_frame, text=receptor_name).grid(
            row=0, column=3, sticky="w", padx=6, pady=5
        )
        self.options_frame.grid_remove()
        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        self.single_tab = ttk.Frame(self.notebook)
        self.batch_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.single_tab, text="Single prediction")
        self.notebook.add(self.batch_tab, text="Batch prediction")
        self._build_single()
        self._build_batch()
        self.notebook.bind("<<NotebookTabChanged>>", self._workflow_tab_changed)
        ttk.Label(self, textvariable=self.single_status_var, anchor="w").grid(
            row=3, column=0, sticky="ew", padx=10, pady=(2, 8)
        )

    def toggle_options(self):
        self.options_visible = not self.options_visible
        if self.options_visible:
            self.options_frame.grid()
        else:
            self.options_frame.grid_remove()

    def _route_controls(self, parent, task_var, subtype_var, row=0):
        ttk.Label(parent, text="Task").grid(row=row, column=0, sticky="w", padx=6, pady=5)
        task_combo = ttk.Combobox(
            parent,
            textvariable=task_var,
            state="readonly",
            width=20,
            values=[ERBATask.CLASSIFICATION.value, ERBATask.IC50_REGRESSION.value],
        )
        task_combo.grid(row=row, column=1, sticky="w", padx=6, pady=5)
        ttk.Label(parent, text="Subtype").grid(row=row, column=2, sticky="w", padx=6, pady=5)
        subtype_combo = ttk.Combobox(
            parent,
            textvariable=subtype_var,
            state="readonly",
            width=14,
            values=[ERBASubtype.ER_ALPHA.value, ERBASubtype.ER_BETA.value],
        )
        subtype_combo.grid(row=row, column=3, sticky="w", padx=6, pady=5)

    def pubchem_clicked(self):
        def work():
            try:
                self.smiles_var.set(cas_lookup_smiles(self.cas_var.get().strip()))
                self.single_status_var.set("PubChem search completed.")
            except Exception as error:
                self.single_status_var.set(f"PubChem search failed: {error}")
        threading.Thread(target=work, daemon=True).start()

    def open_jsme_popup(self):
        if self.structure_editor is None:
            messagebox.showerror("Draw structure", "Structure editor is unavailable.")
            return
        self.structure_editor(self.smiles_var, self.single_status_var)

    def _build_single(self):
        self.single_tab.columnconfigure(0, weight=1)
        self.single_tab.columnconfigure(1, weight=1)
        self.single_tab.rowconfigure(1, weight=1)
        controls = ttk.LabelFrame(self.single_tab, text="Input")
        controls.grid(row=0, column=0, columnspan=2, sticky="ew", padx=8, pady=8)
        controls.columnconfigure(1, weight=1)
        ttk.Label(controls, text="CAS").grid(row=0, column=0, sticky="w", padx=6, pady=5)
        ttk.Entry(controls, textvariable=self.cas_var).grid(row=0, column=1, sticky="ew", padx=6, pady=5)
        ttk.Button(controls, text="PubChem search", command=self.pubchem_clicked).grid(row=0, column=2, padx=6, pady=5)
        ttk.Label(controls, text="SMILES").grid(row=1, column=0, sticky="w", padx=6, pady=5)
        ttk.Entry(controls, textvariable=self.smiles_var).grid(row=1, column=1, sticky="ew", padx=6, pady=5)
        self.single_button = ttk.Button(controls, text="Predict", command=self.single_predict_clicked)
        self.single_button.grid(row=1, column=2, padx=6, pady=5)
        ttk.Button(controls, text="Draw structure", command=self.open_jsme_popup).grid(row=1, column=3, padx=6, pady=5)
        self.single_route_reason = ttk.Label(self.options_frame, foreground="#b42318", wraplength=700)
        self.single_route_reason.grid(row=1, column=0, columnspan=4, sticky="w", padx=6, pady=(0, 5))

        result_frame = ttk.LabelFrame(self.single_tab, text="Prediction result")
        result_frame.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(3, weight=1)
        summary = ttk.Frame(result_frame)
        summary.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))
        summary.columnconfigure(0, weight=1)
        self.single_prediction_label = tk.Label(
            summary, textvariable=self.single_prediction_summary_var,
            font=("Segoe UI", 14, "bold"), anchor="w",
        )
        self.single_prediction_label.grid(row=0, column=0, sticky="w")
        self.single_ad_domain_label = tk.Label(
            summary, textvariable=self.single_ad_domain_var,
            font=("Segoe UI", 13, "bold"), fg="#57606a", anchor="w",
        )
        self.single_ad_domain_label.grid(row=1, column=0, sticky="w", pady=(6, 0))

        self.single_probability_frame = ttk.Frame(result_frame)
        self.single_probability_frame.grid(row=1, column=0, sticky="ew", padx=8, pady=(4, 8))
        self.single_probability_frame.columnconfigure(0, weight=1)
        self.single_prob_canvas = tk.Canvas(self.single_probability_frame, height=92, bg="white", highlightthickness=1, highlightbackground="#d0d7de")
        self.single_prob_canvas.grid(row=0, column=0, sticky="ew")
        self.single_prob_canvas.bind("<Configure>", lambda _event: self.draw_probability_graph())
        self.single_negative_bar = self.single_positive_bar = None

        self.single_regression_frame = ttk.Frame(result_frame)
        self.single_regression_frame.grid(row=1, column=0, sticky="ew", padx=8, pady=4)
        ttk.Label(
            self.single_regression_frame, textvariable=self.single_pic50_var,
            font=("Segoe UI", 13, "bold"),
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            self.single_regression_frame, textvariable=self.single_ic50_var,
            font=("Segoe UI", 13, "bold"),
        ).grid(row=0, column=1, sticky="w", padx=(24, 0))
        ttk.Label(result_frame, text="Details").grid(row=2, column=0, sticky="w", padx=8)
        self.single_result_text = tk.Text(result_frame, height=7, wrap="word", font=("Consolas", 9))
        self.single_result_text.grid(row=3, column=0, sticky="nsew", padx=8, pady=(2, 8))

        right_frame = ttk.Frame(self.single_tab)
        right_frame.grid(row=1, column=1, sticky="nsew", padx=8, pady=8)
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(1, weight=1)
        molecule_pair = ttk.Frame(right_frame)
        molecule_pair.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        molecule_pair.columnconfigure(0, weight=1)
        molecule_pair.columnconfigure(1, weight=1)
        structure_frame = ttk.LabelFrame(molecule_pair, text="Input molecule")
        structure_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self.single_structure_label = ttk.Label(structure_frame, text="No structure")
        self.single_structure_label.pack(expand=True, fill="both", padx=6, pady=6)
        reference_frame = ttk.LabelFrame(molecule_pair, text="Nearest training reference")
        reference_frame.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        self.single_nearest_reference_structure_label = ttk.Label(reference_frame, text="No reference")
        self.single_nearest_reference_structure_label.pack(expand=True, fill="both", padx=6, pady=6)
        ad_frame = ttk.LabelFrame(right_frame, text="Applicability domain")
        ad_frame.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        ad_frame.columnconfigure(0, weight=1)
        ad_frame.rowconfigure(0, weight=1)
        self.single_ad_graph_label = ttk.Label(ad_frame, text="AD graph will appear after prediction.")
        self.single_ad_graph_label.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.single_nearest_reference_var = tk.StringVar(value="Nearest training reference: -")
        ttk.Label(ad_frame, textvariable=self.single_nearest_reference_var, wraplength=420, justify="left").grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))

        caveat = ttk.Frame(self.single_tab)
        caveat.grid(row=2, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 5))
        ttk.Label(
            caveat,
            textvariable=self.single_evidence_caveat_var,
            foreground="#8a4b00",
            wraplength=760,
        ).pack(side="left", fill="x", expand=True)
        ttk.Button(caveat, text="Evidence details", command=self._show_evidence_details).pack(side="right")
        ttk.Label(self.single_tab, text="CAS to SMILES requires internet access. Direct SMILES prediction works offline.").grid(row=3, column=0, columnspan=2, sticky="sw", padx=8, pady=8)

    def _build_batch(self):
        self.batch_tab.columnconfigure(0, weight=1)
        self.batch_tab.rowconfigure(2, weight=1)
        controls = ttk.LabelFrame(self.batch_tab, text="Input")
        controls.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        controls.columnconfigure(1, weight=1)
        ttk.Button(controls, text="Input xlsx", command=self._choose_batch_input).grid(
            row=0, column=0, padx=6, pady=5
        )
        ttk.Label(controls, textvariable=self.batch_input_var, anchor="w").grid(
            row=0, column=1, columnspan=3, sticky="ew", padx=6, pady=5
        )
        ttk.Button(controls, text="Output directory", command=self._choose_output_dir).grid(
            row=1, column=0, padx=6, pady=5
        )
        ttk.Label(controls, textvariable=self.batch_output_var, anchor="w").grid(
            row=1, column=1, columnspan=2, sticky="ew", padx=6, pady=5
        )
        self.batch_button = ttk.Button(controls, text="Run batch", command=self.batch_predict_clicked)
        self.batch_button.grid(row=1, column=3, padx=6, pady=5)
        self.batch_route_reason = ttk.Label(controls, foreground="#b42318", wraplength=700)
        self.batch_route_reason.grid(row=2, column=0, columnspan=4, sticky="w", padx=6, pady=(0, 5))
        progress = ttk.Frame(self.batch_tab)
        progress.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 4))
        progress.columnconfigure(0, weight=1)
        self.batch_progress_bar = ttk.Progressbar(
            progress, maximum=100, mode="determinate", variable=self.batch_progress_value,
        )
        self.batch_progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Label(progress, textvariable=self.batch_progress_var, anchor="w").grid(
            row=0, column=1, sticky="w"
        )
        result_frame = ttk.LabelFrame(self.batch_tab, text="Prediction result")
        result_frame.grid(row=2, column=0, sticky="nsew", padx=8, pady=8)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        self.batch_result = tk.Text(result_frame, height=12, wrap="word", state="disabled")
        self.batch_result.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        ttk.Label(self.batch_tab, textvariable=self.batch_status_var, anchor="w").grid(
            row=3, column=0, sticky="ew", padx=8, pady=(0, 8)
        )

    def _selected_route(self, workflow: str):
        return ERBATask.CLASSIFICATION, self.fixed_subtype

    def _route_available(self, task, subtype):
        if task is ERBATask.IC50_REGRESSION and subtype is ERBASubtype.ER_BETA:
            return False, "ERbeta IC50 regression is not supported."
        if task is ERBATask.IC50_REGRESSION and not self.regression_parity_approved:
            return False, "IC50 regression is disabled until catalog parity approval is recorded."
        if (task, subtype) not in self.catalog:
            return False, "This ERBA route is not released in the bundled catalog."
        return True, ""

    def _show_evidence_details(self):
        route = self._selected_route("single")
        classification = bool(route and route[0] is ERBATask.CLASSIFICATION)
        details = HISTORICAL_EXPOSURE_CAVEAT_FULL if classification else REGRESSION_EVIDENCE_CAVEAT_FULL
        messagebox.showinfo("ERBA evidence caveat", details)

    def _route_changed(self, workflow: str | None = None):
        workflows = (workflow,) if workflow else ("single", "batch")
        for current in workflows:
            route = self._selected_route(current)
            ok, reason = self._route_available(*route) if route else (False, "Invalid ERBA route.")
            if current == "single":
                self._configure_single_result_task(route[0] if route else None)
                classification = bool(route and route[0] is ERBATask.CLASSIFICATION)
                self.single_evidence_caveat_var.set(
                    HISTORICAL_EXPOSURE_CAVEAT_COMPACT
                    if classification
                    else REGRESSION_EVIDENCE_CAVEAT_COMPACT
                )
                self.single_route_reason.configure(text=reason)
                state = "disabled" if self._active_single_request_id is not None else ("normal" if ok else "disabled")
                self.single_button.configure(state=state)
            else:
                self.batch_route_reason.configure(text=reason)
                state = "disabled" if self._active_batch_request_id is not None else ("normal" if ok else "disabled")
                self.batch_button.configure(state=state)

    def _configure_single_result_task(self, task):
        probability_frame = getattr(self, "single_probability_frame", None)
        regression_frame = getattr(self, "single_regression_frame", None)
        if probability_frame is None or regression_frame is None:
            return
        if task is ERBATask.CLASSIFICATION:
            regression_frame.grid_remove()
            probability_frame.grid()
        else:
            probability_frame.grid_remove()
            regression_frame.grid()

    def _snapshot(self, workflow: str):
        route = self._selected_route(workflow)
        if not route:
            return None
        task, subtype = route
        ok, reason = self._route_available(task, subtype)
        if not ok or not self.predictor:
            status = self.single_status_var if workflow == "single" else self.batch_status_var
            status.set(reason or self.catalog_error)
            return None
        self._request_id += 1
        self._started_at[self._request_id] = time.perf_counter()
        return self._request_id, task, subtype, self.catalog[(task, subtype)].model_id

    def _finish_duration_ms(self, request_id: int) -> int:
        started_at = getattr(self, "_started_at", {})
        started = started_at.pop(request_id, time.perf_counter())
        return max(0, int((time.perf_counter() - started) * 1000))

    def single_predict_clicked(self):
        snapshot = self._snapshot("single")
        if not snapshot:
            return
        request_id, task, subtype, model_id = snapshot
        direct_smiles, cas = self.smiles_var.get().strip(), self.cas_var.get().strip()
        self._active_single_request_id = request_id
        self.single_button.configure(state="disabled")
        self.single_status_var.set("Predicting ERBA route...")
        def work():
            smiles = direct_smiles
            if not smiles and cas:
                if not validate_cas(cas):
                    self.after(0, lambda: self._single_complete(
                        request_id, task, subtype, model_id, None,
                        "CAS is invalid: expected a valid CAS Registry Number check digit.",
                    ))
                    return
                try:
                    smiles = cas_lookup_smiles(cas)
                except Exception as error:
                    self.after(0, lambda: self._single_complete(request_id, task, subtype, model_id, None, f"CAS lookup failed: {error}"))
                    return
            result = self.predictor.predict(ERBARequest(smiles=smiles, task=task, subtype=subtype))
            self.after(0, lambda: self._single_complete(request_id, task, subtype, model_id, result, "", smiles))
        threading.Thread(target=work, daemon=True).start()

    def _single_complete(self, request_id, task, subtype, model_id, result, error, smiles=""):
        duration_ms = self._finish_duration_ms(request_id)
        if request_id != self._active_single_request_id:
            self._emit("model.inference_stale", workflow="erba", task=task.value, subtype=subtype.value,
                       model_id=model_id, correlation_id=request_id, duration_ms=duration_ms)
            return
        self._active_single_request_id = None
        if not self._snapshot_is_current("single", request_id, task, subtype, model_id):
            self._route_changed("single")
            self._emit("model.inference_stale", workflow="erba", task=task.value, subtype=subtype.value,
                       model_id=model_id, correlation_id=request_id, duration_ms=duration_ms)
            return
        self._route_changed("single")
        if error:
            self._clear_single_result("Prediction failed")
            self.single_status_var.set(error)
            self._emit("model.inference_complete", workflow="erba", task=task.value, subtype=subtype.value,
                       model_id=model_id, correlation_id=request_id, duration_ms=duration_ms,
                       status="failed", exception_type="InputOrLookupError")
            return
        self.single_status_var.set(result.status_message or result.status_code.value)
        self._emit("model.inference_complete", workflow="erba", task=task.value, subtype=subtype.value,
                   model_id=model_id, correlation_id=request_id, duration_ms=duration_ms,
                   status=result.status_code.value)
        self._render_single_result(result)
        if result.status_code is ERBAStatusCode.OK and hasattr(self, "single_ad_domain_var"):
            self._active_ad_request_id = request_id
            self.single_ad_domain_var.set("Applicability domain: Evaluating")
            threading.Thread(
                target=self._evaluate_single_ad,
                args=(request_id, task, subtype, result.model_smiles or smiles),
                daemon=True,
            ).start()

    def _evaluate_single_ad(self, request_id, task, subtype, smiles):
        try:
            calculator, ad_result, fingerprint = self.erba_ad.evaluate(task, subtype, smiles)
            graph_path = save_single_ad_plot(calculator, fingerprint, str(validate_mutable_directory(self.output_root) / "erba_ad"))
            self.after(
                0,
                lambda: self._render_single_ad(
                    request_id, task, subtype, ad_result, graph_path, ""
                ),
            )
        except Exception as error:
            self.after(
                0,
                lambda: self._render_single_ad(
                    request_id, task, subtype, None, "", str(error)
                ),
            )

    def _render_single_ad(
        self, request_id, task, subtype, ad_result, graph_path, error
    ):
        if (
            request_id != self._active_ad_request_id
            or self._selected_route("single") != (task, subtype)
        ):
            return
        self._active_ad_request_id = None
        if error:
            self.single_ad_domain_var.set(f"Applicability domain: AD error — {error}")
            self._single_detail_lines.append(f"Applicability domain: Error: {error}")
            self._render_detail_lines()
            return
        domain = "In-domain" if ad_result.in_domain else "Out-of-domain"
        self.single_ad_domain_var.set(f"Applicability domain: {domain}")
        self.single_ad_domain_label.configure(fg="#1a7f37" if ad_result.in_domain else "#cf222e")
        ref = ad_result.display_nearest_reference or ad_result.nearest_reference or {}
        self.single_nearest_reference_var.set(f"Nearest training reference: Morgan similarity={ad_result.display_similarity:.4f} / {' / '.join(f'{key}: {value}' for key, value in list(ref.items())[:4])}")
        self._single_detail_lines.extend(
            [
                f"Applicability domain: {domain}",
                f"kNN mean distance: {ad_result.distance:.4f} / 95% threshold {ad_result.threshold:.4f}",
                f"Distance AD: {'In-domain' if ad_result.distance_in_domain else 'Out-of-domain'}",
                f"AD RDKit similarity: {ad_result.max_similarity:.4f} / cutoff {ad_result.similarity_threshold:.4f}",
                f"Similarity AD: {'In-domain' if ad_result.similarity_in_domain else 'Out-of-domain'}",
                self.single_nearest_reference_var.get(),
                f"PCA position: PC1={ad_result.pc1:.4f}, PC2={ad_result.pc2:.4f}",
            ]
        )
        self._render_detail_lines()
        self._draw_nearest_reference_structure(ref)
        if PIL_AVAILABLE and graph_path:
            with Image.open(graph_path) as image:
                image.thumbnail((520, 300))
                self._single_ad_graph_image = ImageTk.PhotoImage(image.copy())
            self.single_ad_graph_label.configure(image=self._single_ad_graph_image, text="")

    def _clear_single_result(self, summary):
        if hasattr(self, "_active_ad_request_id"):
            self._active_ad_request_id = None
        self.single_prediction_summary_var.set(summary)
        self.single_negative_probability_var.set("Negative probability: -")
        self.single_positive_probability_var.set("Positive probability: -")
        self.single_pic50_var.set("pIC50: -")
        self.single_ic50_var.set("IC50: - nM")
        self.single_detail_var.set("No prediction result is available.")
        self._single_detail_lines = [f"Status: {summary}"]
        for name in ("single_negative_bar", "single_positive_bar"):
            widget = getattr(self, name, None)
            if widget is not None:
                widget.configure(value=0)
        structure = getattr(self, "single_structure_label", None)
        if structure is not None:
            structure.configure(text="No structure", image="")
        self._single_structure_image = None
        if hasattr(self, "single_ad_domain_var"):
            self.single_ad_domain_var.set("Applicability domain: Not evaluated")
        if hasattr(self, "single_nearest_reference_var"):
            self.single_nearest_reference_var.set("Nearest training reference: -")
        if hasattr(self, "single_ad_graph_label"):
            self.single_ad_graph_label.configure(
                text="AD graph will appear after prediction.", image=""
            )
        self._render_detail_lines()

    def _render_single_result(self, result):
        self._configure_single_result_task(result.task)
        status = result.status_message or result.status_code.value
        if result.status_code is not ERBAStatusCode.OK:
            self._clear_single_result("Prediction unavailable")
            self.single_detail_var.set(f"Status: {status}")
            return
        if result.task is ERBATask.CLASSIFICATION:
            label = (result.binding_label or "Unknown").replace("_", " ").title()
            self.single_prediction_summary_var.set(f"Binding prediction: {label}")
            negative = float(result.non_binding_probability or 0)
            positive = float(result.binding_probability or 0)
            self.single_negative_probability_var.set(f"Negative probability: {negative:.1%}")
            self.single_positive_probability_var.set(f"Positive probability: {positive:.1%}")
            if self.single_negative_bar is not None:
                self.single_negative_bar.configure(value=negative * 100)
            if self.single_positive_bar is not None:
                self.single_positive_bar.configure(value=positive * 100)
            self._single_positive_probability = positive
            self.draw_probability_graph()
            self.single_detail_var.set(f"Status: {status}  •  Model: {result.model_id or '-'}")
        else:
            self.single_prediction_summary_var.set("ERα IC50 regression prediction")
            self.single_pic50_var.set(
                f"pIC50: {result.pic50:.4f}" if result.pic50 is not None else "pIC50: -"
            )
            self.single_ic50_var.set(
                f"IC50: {result.ic50_nm:.2f} nM" if result.ic50_nm is not None else "IC50: - nM"
            )
            self.single_detail_var.set(f"Status: {status}  •  Model: {result.model_id or '-'}")
        label = (
            (result.binding_label or "Unknown").replace("_", " ").title()
            if result.task is ERBATask.CLASSIFICATION
            else "-"
        )
        cas_var = getattr(self, "cas_var", None)
        cas = cas_var.get().strip() if cas_var is not None else ""
        self._single_detail_lines = [
            f"CAS: {cas}",
            f"Input SMILES: {result.raw_smiles}",
            f"Canonical SMILES: {result.model_smiles or ''}",
            f"Task: Binding classification",
            f"Receptor subtype: {'ERalpha' if result.subtype is ERBASubtype.ER_ALPHA else 'ERbeta'}",
            f"Negative probability: {float(result.non_binding_probability or 0):.6f}",
            f"Positive probability: {float(result.binding_probability or 0):.6f}",
            f"Binding prediction: {label}",
            f"Status: {result.status_code.value}",
            f"Status message: {status}",
            f"Model ID: {result.model_id or ''}",
            f"Model SHA256: {result.model_sha256 or ''}",
            f"Preprocessing policy ID: {result.preprocessing_policy_id or ''}",
            f"Evidence scope: {result.evidence_scope or ''}",
        ]
        self._draw_single_structure(result.model_smiles)
        self._render_detail_lines()

    def _render_detail_lines(self):
        details = getattr(self, "single_result_text", None)
        if details is None:
            return
        details.delete("1.0", "end")
        details.insert("1.0", "\n".join(self._single_detail_lines))

    def _draw_nearest_reference_structure(self, reference):
        widget = getattr(self, "single_nearest_reference_structure_label", None)
        if widget is None:
            return
        smiles = next(
            (
                str(reference[key]).strip()
                for key in (
                    "SMILES",
                    "Canonical_SMILES",
                    "Canonical SMILES",
                    "model_smiles",
                )
                if reference.get(key)
            ),
            "",
        )
        if not smiles or not PIL_AVAILABLE:
            widget.configure(text="No reference", image="")
            self._nearest_reference_structure_image = None
            return
        try:
            directory = validate_mutable_directory(self.output_root) / "structures"
            directory.mkdir(parents=True, exist_ok=True)
            path = directory / f"{self.fixed_subtype.value}_nearest_reference.png"
            save_molecule_image(smiles, str(path), size=(260, 200))
            with Image.open(path) as image:
                image.thumbnail((260, 200))
                self._nearest_reference_structure_image = ImageTk.PhotoImage(image.copy())
            widget.configure(
                image=self._nearest_reference_structure_image,
                text="",
            )
        except Exception as error:
            widget.configure(text=f"Reference preview unavailable: {error}", image="")
            self._nearest_reference_structure_image = None

    def draw_probability_graph(self):
        if not hasattr(self, "single_prob_canvas"):
            return
        canvas = self.single_prob_canvas
        canvas.delete("all")
        positive = getattr(self, "_single_positive_probability", None)
        width, margin, label_width, bar_h = max(canvas.winfo_width(), 320), 18, 72, 18
        bar_x, bar_w = margin + label_width, max(width - margin - label_width - 48, 120)
        def bar(y, label, value, color):
            canvas.create_text(margin, y + bar_h / 2, text=label, anchor="w")
            canvas.create_rectangle(bar_x, y, bar_x + bar_w, y + bar_h, fill="#eef2f7", outline="#d0d7de")
            canvas.create_rectangle(bar_x, y, bar_x + bar_w * value, y + bar_h, fill=color, outline=color)
            canvas.create_text(bar_x + bar_w + 8, y + bar_h / 2, text=f"{value:.3f}", anchor="w")
        if positive is None:
            bar(22, "Negative", 0, "#2da44e")
            bar(58, "Positive", 0, "#cf222e")
            canvas.create_text(bar_x, 8, text="Run prediction to show probabilities", anchor="w", fill="#57606a")
            return
        bar(22, "Negative", 1 - positive, "#2da44e")
        bar(58, "Positive", positive, "#cf222e")

    def _draw_single_structure(self, smiles):
        structure = getattr(self, "single_structure_label", None)
        if structure is None:
            return
        if not smiles:
            structure.configure(text="No structure", image="")
            self._single_structure_image = None
            return
        if not PIL_AVAILABLE:
            structure.configure(text="Structure preview unavailable (Pillow is not installed).", image="")
            return
        try:
            directory = validate_mutable_directory(self.output_root) / "structures"
            directory.mkdir(parents=True, exist_ok=True)
            path = directory / "erba_single_structure.png"
            save_molecule_image(smiles, str(path), size=(260, 200))
            with Image.open(path) as image:
                image.thumbnail((260, 200))
                self._single_structure_image = ImageTk.PhotoImage(image.copy())
            structure.configure(image=self._single_structure_image, text="")
        except Exception as error:
            structure.configure(text=f"Structure preview unavailable: {error}", image="")
            self._single_structure_image = None

    def _choose_batch_input(self):
        path = filedialog.askopenfilename(title="Select ERBA input", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.batch_input_var.set(path)

    def _choose_output_dir(self):
        path = filedialog.askdirectory(title="Select ERBA output directory")
        if not path:
            return
        try:
            validated = validate_mutable_directory(path, forbidden_roots=(self.project_root,))
        except Exception as error:
            messagebox.showerror("Invalid output directory", str(error))
            return
        self.batch_output_var.set(str(validated))

    @staticmethod
    def _column_lookup(frame, aliases):
        normalized = {str(column).strip().lower(): column for column in frame.columns}
        return next((normalized[name] for name in aliases if name in normalized), None)

    def batch_predict_clicked(self):
        snapshot = self._snapshot("batch")
        if not snapshot:
            return
        request_id, task, subtype, model_id = snapshot
        input_path, output_dir = self.batch_input_var.get().strip(), self.batch_output_var.get().strip()
        if not input_path or not output_dir:
            self.batch_status_var.set("Choose an input xlsx and output directory.")
            self._started_at.pop(request_id, None)
            return
        try:
            output_dir = str(validate_mutable_directory(output_dir, forbidden_roots=(self.project_root,)))
            preflight = self.predictor.preflight(task, subtype)
            if not preflight.ready:
                raise RuntimeError(
                    f"ERBA route preflight failed ({preflight.status_code.value}): {preflight.status_message}"
                )
        except Exception as error:
            self.batch_status_var.set(str(error))
            self._started_at.pop(request_id, None)
            return
        self.batch_button.configure(state="disabled")
        self._active_batch_request_id = request_id
        self._set_batch_progress(request_id, "Ready to read input", 0, 0, 0)
        self.batch_status_var.set("Running ERBA batch...")
        threading.Thread(target=self._batch_work, args=(request_id, task, subtype, model_id, input_path, output_dir), daemon=True).start()

    def _batch_progress_from_worker(self, request_id, stage, current, total, percent):
        """Schedule all Tk mutation on the event loop, never on the batch worker."""
        self.after(
            0,
            lambda: self._set_batch_progress(request_id, stage, current, total, percent),
        )

    def _set_batch_progress(self, request_id, stage, current, total, percent):
        if request_id != getattr(self, "_active_batch_request_id", None):
            return
        percent = max(0, min(100, int(percent)))
        detail = f"{current}/{total}" if total else "0/0"
        progress_var = getattr(self, "batch_progress_var", None)
        progress_value = getattr(self, "batch_progress_value", None)
        if progress_var is not None:
            progress_var.set(f"{percent}% - {detail} - {stage}")
        if progress_value is not None:
            progress_value.set(percent)

    def _batch_work(self, request_id, task, subtype, model_id, input_path, output_dir):
        try:
            destination, count = export_erba_batch(
                input_path,
                output_dir,
                task,
                subtype,
                self.predictor,
                self.catalog[(task, subtype)],
                self.catalog_payload,
                progress_callback=lambda stage, current, total, percent: self._batch_progress_from_worker(
                    request_id, stage, current, total, percent
                ),
            )
            self.after(0, lambda: self._batch_complete(
                request_id, task, subtype, model_id, destination, count, ""
            ))
        except Exception as error:
            self.after(0, lambda: self._batch_complete(request_id, task, subtype, model_id, None, 0, str(error)))

    def _batch_complete(self, request_id, task, subtype, model_id, destination, count, error):
        duration_ms = self._finish_duration_ms(request_id)
        if request_id != self._active_batch_request_id:
            self._emit("batch.inference_stale", workflow="erba", task=task.value, subtype=subtype.value,
                       model_id=model_id, correlation_id=request_id, duration_ms=duration_ms)
            return
        self._set_batch_progress(
            request_id,
            f"Failed: {error}" if error else "Completed",
            count if not error else 0,
            count if not error else 0,
            100 if not error else 0,
        )
        self._active_batch_request_id = None
        if not self._snapshot_is_current("batch", request_id, task, subtype, model_id):
            self._route_changed("batch")
            self._emit("batch.inference_stale", workflow="erba", task=task.value, subtype=subtype.value,
                       model_id=model_id, correlation_id=request_id, duration_ms=duration_ms)
            return
        self._route_changed("batch")
        if error:
            self.batch_status_var.set(f"ERBA batch failed: {error}")
            self._emit("batch.inference_complete", workflow="erba", task=task.value, subtype=subtype.value,
                       model_id=model_id, correlation_id=request_id, duration_ms=duration_ms,
                       row_count=0, status="failed", exception_type="BatchError")
            return
        self.batch_status_var.set(f"ERBA batch complete: {destination}")
        self._emit("batch.inference_complete", workflow="erba", task=task.value, subtype=subtype.value,
                   model_id=model_id, correlation_id=request_id, duration_ms=duration_ms,
                   row_count=count, status="ok")
        caveat = (
            HISTORICAL_EXPOSURE_CAVEAT_COMPACT
            if task is ERBATask.CLASSIFICATION
            else REGRESSION_EVIDENCE_CAVEAT_COMPACT
        )
        self._set_text(self.batch_result, f"Exported {count} rows to\n{destination}\n\n{caveat}")

    def _snapshot_is_current(self, workflow, request_id, task, subtype, model_id):
        active_id = (
            self._active_single_request_id if workflow == "single" else self._active_batch_request_id
        )
        return (
            active_id in (None, request_id)
            and self._selected_route(workflow) == (task, subtype)
            and self.catalog.get((task, subtype), None) is not None
            and self.catalog[(task, subtype)].model_id == model_id
        )
    @staticmethod
    def _set_text(widget, text):
        widget.configure(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert("1.0", text)
        widget.configure(state="disabled")
