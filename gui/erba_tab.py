"""ERalpha Tkinter journey with released-model and route-isolated AD controls."""
from __future__ import annotations

from collections.abc import Mapping
from dataclasses import asdict, dataclass
import hashlib
import json
import os
import shutil
import tempfile
import threading
import re
import time
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from typing import Callable

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
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
from core.erba_ad import (
    ERBAApplicabilityDomain,
    erba_ad_cache_path,
    erba_ad_reference_path,
)
from core.erba_release_allowlist import TRUSTED_ERBA_ARTIFACTS
from core.graph import save_ad_decision_plot, save_ad_plot, save_single_ad_plot
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
_RESOURCE_ROOT = Path(__file__).resolve().parents[1]
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


@dataclass(frozen=True)
class SharedExampleInput:
    workbook: Path
    cas: str
    smiles: str


@dataclass(frozen=True)
class ERBABatchExportResult:
    destination: Path
    count: int
    binding_count: int
    non_binding_count: int
    not_predicted_count: int
    ad_in_domain_count: int
    ad_out_of_domain_count: int
    ad_unavailable_count: int
    graph_paths: tuple[Path, ...]
    graph_directory: Path | None
    graph_error: str = ""
    ad_error: str = ""


class ERBAReleasedCatalog(dict):
    """Expose a route default while retaining every approved model for that route."""

    def __init__(self):
        super().__init__()
        self._choices: dict[
            tuple[ERBATask, ERBASubtype], list[ERBAArtifactSpec]
        ] = {}

    def add(
        self,
        task: ERBATask,
        subtype: ERBASubtype,
        spec: ERBAArtifactSpec,
    ) -> None:
        route = (task, subtype)
        choices = self._choices.setdefault(route, [])
        if any(choice.model_id == spec.model_id for choice in choices):
            raise ValueError(
                "ERBA catalog contains a duplicate released model ID for one route."
            )
        choices.append(spec)
        self.setdefault(route, spec)

    def choices_for(
        self,
        task: ERBATask,
        subtype: ERBASubtype,
    ) -> tuple[ERBAArtifactSpec, ...]:
        return tuple(self._choices.get((task, subtype), ()))

    def model_for(
        self,
        task: ERBATask,
        subtype: ERBASubtype,
        model_id: str,
    ) -> ERBAArtifactSpec:
        for spec in self.choices_for(task, subtype):
            if spec.model_id == model_id:
                return spec
        raise KeyError(model_id)


def load_shared_example_input(project_root: str | Path) -> SharedExampleInput:
    """Read the existing ERTA example used to initialize both endpoint tabs."""
    workbook = Path(project_root) / "templates" / "ERTA_KRICT_example.xlsx"
    if not workbook.is_file():
        return SharedExampleInput(workbook, "", "")
    try:
        frame = pd.read_excel(
            workbook,
            sheet_name=0,
            nrows=1,
            dtype=str,
            keep_default_na=False,
        )
    except Exception:
        return SharedExampleInput(workbook, "", "")
    if frame.empty:
        return SharedExampleInput(workbook, "", "")
    columns = {
        re.sub(r"\s+", " ", str(column).strip()).casefold(): column
        for column in frame.columns
    }
    cas_column = columns.get("cas")
    smiles_column = columns.get("smiles")
    cas = str(frame.iloc[0][cas_column]).strip() if cas_column is not None else ""
    smiles = (
        str(frame.iloc[0][smiles_column]).strip()
        if smiles_column is not None
        else ""
    )
    return SharedExampleInput(workbook, cas, smiles)


def batch_destination_display(input_path: str | Path) -> str:
    """Describe the fixed batch destination without probing or creating directories."""
    if not str(input_path).strip():
        return "Select an input workbook; the result workbook is saved in the same folder."
    return str(Path(input_path).expanduser().resolve(strict=False).parent)


def load_erba_catalog(path: str | Path):
    """Parse released choices while preserving the route-keyed default API."""
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(payload, dict) or payload.get("schema_version") != 2:
        raise ValueError("ERBA catalog schema version must be 2.")
    routes = payload.get("routes")
    if not isinstance(routes, list):
        raise ValueError("ERBA catalog routes are missing.")
    catalog = ERBAReleasedCatalog()
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
        catalog.add(task, subtype, spec)
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


def _normalized_header(value) -> str:
    return re.sub(r"\s+", " ", str(value).strip()).casefold()


def build_primary_predictions(
    input_rows: pd.DataFrame,
    trusted_predictions: pd.DataFrame,
    ad_rows: pd.DataFrame,
) -> pd.DataFrame:
    """Prefix safe input fields while preventing them from spoofing trusted output fields."""
    if not (len(input_rows) == len(trusted_predictions) == len(ad_rows)):
        raise ValueError("ERBA batch result rows are not aligned.")
    reserved = {
        _normalized_header(column)
        for column in (*trusted_predictions.columns, *ad_rows.columns)
    }
    used = set(reserved)
    passthrough = []
    for position, header in enumerate(input_rows.columns):
        normalized = _normalized_header(header)
        if normalized in reserved:
            continue
        base = str(header).strip() or f"Input_{position + 1}"
        candidate = base
        suffix = 2
        while _normalized_header(candidate) in used:
            candidate = f"{base}_input_{suffix}"
            suffix += 1
        used.add(_normalized_header(candidate))
        passthrough.append(
            input_rows.iloc[:, position].reset_index(drop=True).rename(candidate)
        )
    passthrough_frame = (
        pd.concat(passthrough, axis=1)
        if passthrough
        else pd.DataFrame(index=range(len(input_rows)))
    )
    return pd.concat(
        (
            passthrough_frame,
            trusted_predictions.reset_index(drop=True),
            ad_rows.reset_index(drop=True),
        ),
        axis=1,
    )


def _blank_ad_row() -> dict[str, object]:
    return {column: "" for column in ERBA_BATCH_AD_COLUMNS}


def _ad_result_row(ad_result) -> dict[str, object]:
    if not getattr(ad_result, "fitted", False) or ad_result.in_domain is None:
        raise RuntimeError(getattr(ad_result, "message", "") or "AD evaluation is unavailable.")
    return {
        "AD": "In-domain" if ad_result.in_domain else "Out-of-domain",
        "AD_MeanDistance": ad_result.distance,
        "AD_DistanceThreshold": ad_result.threshold,
        "AD_Distance_InDomain": ad_result.distance_in_domain,
        "AD_SimilarityMax": ad_result.max_similarity,
        "AD_SimilarityThreshold": ad_result.similarity_threshold,
        "AD_Similarity_InDomain": ad_result.similarity_in_domain,
        "AD_PC1": ad_result.pc1,
        "AD_PC2": ad_result.pc2,
    }


def allocate_graph_directory(destination: str | Path) -> Path:
    """Reserve a batch-specific graph directory without reusing prior artifacts."""
    workbook = Path(destination)
    base = workbook.parent / f"{workbook.stem}_graphs"
    for index in range(1, 10000):
        suffix = "" if index == 1 else f"_{index}"
        candidate = base.with_name(f"{base.name}{suffix}")
        try:
            candidate.mkdir()
        except FileExistsError:
            continue
        return candidate
    raise RuntimeError("Could not allocate a non-overwriting ERBA graph directory.")


def _binding_display_label(value) -> str:
    normalized = str(value or "").strip().replace("-", "_").replace(" ", "_").casefold()
    if normalized in {"binding", "1", "1.0"}:
        return "Binding"
    if normalized in {"non_binding", "nonbinding", "0", "0.0"}:
        return "Non-binding"
    return ""


def _graph_row_label(row: pd.Series) -> str:
    for column in ("CAS", "Row_ID", "model_smiles", "raw_smiles"):
        value = str(row.get(column, "") or "").strip()
        if value and value.casefold() != "nan":
            return value[:24]
    return "-"


def save_erba_batch_graphs(
    predictions: pd.DataFrame,
    destination: str | Path,
    calculator,
    fingerprints: np.ndarray,
) -> tuple[tuple[Path, ...], Path]:
    """Write binding-labelled graphs using only aligned ERBA AD-success rows."""
    if calculator is None or not getattr(calculator, "fitted", False):
        raise RuntimeError("The route-specific ERBA AD calculator is unavailable.")
    matrix = np.asarray(fingerprints, dtype=float)
    if matrix.ndim != 2 or len(matrix) != len(predictions) or not len(predictions):
        raise ValueError("ERBA graph rows and AD fingerprints are not aligned.")

    directory = allocate_graph_directory(destination)
    paths: list[Path] = []
    try:
        labels = predictions["binding_label"].map(_binding_display_label)
        counts = labels.value_counts().reindex(("Non-binding", "Binding"), fill_value=0)
        figure, axis = plt.subplots(figsize=(5, 3.5))
        axis.bar(counts.index, counts.values)
        axis.set_ylabel("Count")
        axis.set_title("ERalpha direct-binding class count")
        for index, value in enumerate(counts.values):
            axis.text(index, value, str(int(value)), ha="center", va="bottom")
        figure.tight_layout()
        path = directory / "binding_class_count.png"
        figure.savefig(path, dpi=150)
        plt.close(figure)
        paths.append(path)

        probabilities = pd.to_numeric(
            predictions["binding_probability"], errors="coerce"
        ).dropna()
        figure, axis = plt.subplots(figsize=(5, 3.5))
        axis.hist(probabilities, bins=20)
        axis.set_xlabel("Binding probability")
        axis.set_ylabel("Count")
        axis.set_title("ERalpha binding-probability distribution")
        figure.tight_layout()
        path = directory / "binding_probability_histogram.png"
        figure.savefig(path, dpi=150)
        plt.close(figure)
        paths.append(path)

        top = predictions.assign(
            _binding_probability=pd.to_numeric(
                predictions["binding_probability"], errors="coerce"
            )
        ).dropna(subset=["_binding_probability"])
        top = top.sort_values("_binding_probability", ascending=False).head(10)
        top_labels = [_graph_row_label(row) for _, row in top.iterrows()]
        figure, axis = plt.subplots(figsize=(6, 4))
        axis.barh(
            top_labels[::-1],
            top["_binding_probability"].to_numpy()[::-1],
        )
        axis.set_xlabel("Binding probability")
        axis.set_title(f"Top {len(top)} predicted ERalpha binders")
        axis.set_xlim(0, 1)
        figure.tight_layout()
        path = directory / "top_binding_chemicals.png"
        figure.savefig(path, dpi=150)
        plt.close(figure)
        paths.append(path)

        paths.append(Path(save_ad_plot(calculator, matrix, predictions, str(directory))))
        paths.append(
            Path(save_ad_decision_plot(calculator, matrix, predictions, str(directory)))
        )
        return tuple(paths), directory
    except Exception:
        plt.close("all")
        shutil.rmtree(directory, ignore_errors=True)
        raise


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


def _format_workbook(
    writer,
    predictions: pd.DataFrame,
    input_rows: pd.DataFrame,
    metadata: pd.DataFrame,
    task: ERBATask,
    graph_paths: tuple[Path, ...],
    graph_directory: Path | None,
    graph_error: str,
    ad_error: str,
) -> None:
    """Apply readable presentation without changing data-sheet headers or values."""
    classification = task is ERBATask.CLASSIFICATION
    guide_rows = [
        (
            "ERalpha batch prediction guide",
            "ERalpha direct receptor-binding classification output."
            if classification
            else "ERalpha direct receptor-binding potency output.",
        ),
        (
            "Predictions",
            "Primary input-first result table. Original fields that do not collide with trusted "
            "result fields appear first, followed by validated prediction and applicability-domain fields.",
        ),
        (
            "Unavailable predictions",
            (
                "Binding probabilities, binding label, and AD values are blank when Result_Status is Not predicted; "
                "the reason and recommended action remain on the same row."
                if classification
                else "pIC50 and IC50 values are blank when Result_Status is Not predicted; "
                "the reason and recommended action remain on the same row."
            ),
        ),
        (
            "Applicability domain",
            (
                "AD fields use only the bundled reference for this ERBA route. AD can be Unavailable without "
                "changing a successful binding prediction."
                if classification
                else "Batch AD fields are not part of this regression workbook contract."
            ),
        ),
        ("Input", "Original first-sheet input, retained in its original row order."),
        ("Metadata", "Model identity, integrity hashes, and evidence caveat for this export."),
        ("Graph files", str(len(graph_paths))),
        (
            "Graph directory",
            str(graph_directory) if graph_directory is not None else "Not generated",
        ),
        (
            "Graph warning",
            ""
            if graph_paths
            else (
                graph_error or "No route-specific AD rows were available."
                if classification
                else "Graphs are not generated for this regression task."
            ),
        ),
        (
            "Caveat",
            HISTORICAL_EXPOSURE_CAVEAT_FULL
            if classification
            else REGRESSION_EVIDENCE_CAVEAT_FULL,
        ),
        ("Total rows", str(len(predictions))),
        ("Predicted rows", str((predictions["Result_Status"] == "Predicted").sum())),
        ("Not predicted rows", str((predictions["Result_Status"] != "Predicted").sum())),
        ("Reason counts", "Counts below include not-predicted categories only."),
    ]
    if "binding_label" in predictions:
        labels = predictions["binding_label"].map(_binding_display_label)
        guide_rows.extend(
            (
                ("Binding rows", str((labels == "Binding").sum())),
                ("Non-binding rows", str((labels == "Non-binding").sum())),
            )
        )
    if "AD" in predictions:
        guide_rows.extend(
            (
                ("AD In-domain rows", str((predictions["AD"] == "In-domain").sum())),
                ("AD Out-of-domain rows", str((predictions["AD"] == "Out-of-domain").sum())),
                ("AD Unavailable rows", str((predictions["AD"] == "Unavailable").sum())),
            )
        )
    if ad_error:
        guide_rows.append(("AD evaluation warning", ad_error))
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
        (ERBA_PREDICTIONS_SHEET_NAME, predictions),
        (ERBA_GUIDE_SHEET_NAME, pd.DataFrame(guide_rows)),
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
        writer.book[ERBA_PREDICTIONS_SHEET_NAME],
        writer.book[ERBA_GUIDE_SHEET_NAME],
        writer.book[ERBA_INPUT_SHEET_NAME],
        writer.book[ERBA_METADATA_SHEET_NAME],
    ]
    writer.book.active = writer.book[ERBA_PREDICTIONS_SHEET_NAME]


def export_erba_batch(
    input_path: str | Path,
    task: ERBATask,
    subtype: ERBASubtype,
    predictor: ERBAPredictor,
    erba_ad: ERBAApplicabilityDomain,
    spec: ERBAArtifactSpec,
    catalog_payload: dict,
    progress_callback: Callable[[str, int, int, int], None] | None = None,
    *,
    forbidden_roots: tuple[str | Path, ...] = (),
) -> ERBABatchExportResult:
    """Predict an ERBA workbook and publish a new result beside that input workbook."""
    _progress(progress_callback, "Reading input workbook", 0, 0, 0)
    source = Path(input_path).expanduser().resolve(strict=False)
    if not source.is_file():
        raise FileNotFoundError(f"ERBA batch input does not exist: {source}")
    output_dir = validate_mutable_directory(
        source.parent,
        forbidden_roots=(_RESOURCE_ROOT, *forbidden_roots),
    )
    preflight = predictor.preflight(task, subtype)
    if not preflight.ready:
        raise RuntimeError(
            f"ERBA route preflight failed ({preflight.status_code.value}): {preflight.status_message}"
        )
    frame = read_batch_input(source)
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
        _progress(
            progress_callback,
            "Resolving CAS/SMILES",
            position + 1,
            total,
            10 + int(25 * (position + 1) / total),
        )
    results = []
    _progress(progress_callback, "Predicting", 0, total, 35)
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
        _progress(
            progress_callback,
            "Predicting",
            position + 1,
            total,
            35 + int(35 * (position + 1) / total),
        )

    if fatal_artifact_failure([result for result, _, _ in results]):
        raise RuntimeError("ERBA artifact validation failed; no output was published.")

    prediction_columns = (
        ERBA_CLASSIFICATION_PREDICTION_COLUMNS
        if task is ERBATask.CLASSIFICATION
        else ERBA_REGRESSION_PREDICTION_COLUMNS
    )
    trusted_predictions = pd.DataFrame(
        [result_row(result, task, row_id, cas) for result, row_id, cas in results],
        columns=prediction_columns,
    )

    ad_rows = [_blank_ad_row() for _ in range(total)]
    ad_errors = []
    ad_positions = []
    ad_fingerprints = []
    graph_calculator = None
    _progress(progress_callback, "Evaluating applicability domain", 0, total, 70)
    for current, (result, row_id, _) in enumerate(results, 1):
        if (
            task is ERBATask.CLASSIFICATION
            and subtype is ERBASubtype.ER_ALPHA
            and result.status_code is ERBAStatusCode.OK
        ):
            try:
                calculator, ad_result, fingerprint = erba_ad.evaluate(
                    task,
                    subtype,
                    result.model_smiles or "",
                )
                ad_rows[current - 1] = _ad_result_row(ad_result)
                graph_calculator = graph_calculator or calculator
                ad_positions.append(current - 1)
                ad_fingerprints.append(
                    np.asarray(fingerprint, dtype=float).reshape(1, -1)
                )
            except Exception as error:
                ad_rows[current - 1]["AD"] = "Unavailable"
                ad_errors.append(f"Row {row_id or current}: {error}")
        _progress(
            progress_callback,
            "Evaluating applicability domain",
            current,
            total,
            70 + int(18 * current / total),
        )

    ad_frame = pd.DataFrame(ad_rows, columns=ERBA_BATCH_AD_COLUMNS)
    predictions = (
        build_primary_predictions(input_rows, trusted_predictions, ad_frame)
        if task is ERBATask.CLASSIFICATION
        and subtype is ERBASubtype.ER_ALPHA
        else trusted_predictions
    )
    unique_ad_errors = list(dict.fromkeys(ad_errors))
    ad_error = "; ".join(unique_ad_errors[:3])
    if len(unique_ad_errors) > 3:
        ad_error += f"; and {len(unique_ad_errors) - 3} more row(s)"

    destination = allocate_output_path(output_dir, task, subtype)
    graph_paths: tuple[Path, ...] = ()
    graph_directory = None
    graph_error = ""
    if ad_positions:
        try:
            graph_predictions = predictions.iloc[ad_positions].reset_index(drop=True)
            graph_paths, graph_directory = save_erba_batch_graphs(
                graph_predictions,
                destination,
                graph_calculator,
                np.vstack(ad_fingerprints),
            )
        except Exception as error:
            graph_error = str(error)
    elif task is ERBATask.CLASSIFICATION and subtype is ERBASubtype.ER_ALPHA:
        graph_error = (
            "No rows had an available route-specific AD evaluation; "
            "binding graphs were not generated."
        )

    temp_name = None
    try:
        _progress(progress_callback, "Writing workbook", total, total, 90)
        temp_fd, temp_name = tempfile.mkstemp(suffix=".xlsx", dir=str(destination.parent))
        os.close(temp_fd)
        with pd.ExcelWriter(temp_name, engine="openpyxl") as writer:
            predictions.to_excel(writer, sheet_name=ERBA_PREDICTIONS_SHEET_NAME, index=False)
            input_rows.to_excel(writer, sheet_name=ERBA_INPUT_SHEET_NAME, index=False)
            metadata = pd.DataFrame(
                [metadata_row(task, spec, catalog_payload)],
                columns=ERBA_METADATA_COLUMNS,
            )
            metadata.to_excel(writer, sheet_name=ERBA_METADATA_SHEET_NAME, index=False)
            _format_workbook(
                writer,
                predictions,
                input_rows,
                metadata,
                task,
                graph_paths,
                graph_directory,
                graph_error,
                ad_error,
            )
        os.replace(temp_name, destination)
    except Exception:
        if temp_name is not None:
            Path(temp_name).unlink(missing_ok=True)
        Path(destination).unlink(missing_ok=True)
        if graph_directory is not None:
            shutil.rmtree(graph_directory, ignore_errors=True)
        raise
    _progress(progress_callback, "Completed", total, total, 100)
    binding_labels = (
        trusted_predictions["binding_label"].map(_binding_display_label)
        if "binding_label" in trusted_predictions
        else pd.Series("", index=trusted_predictions.index)
    )
    return ERBABatchExportResult(
        destination=destination,
        count=len(results),
        binding_count=int((binding_labels == "Binding").sum()),
        non_binding_count=int((binding_labels == "Non-binding").sum()),
        not_predicted_count=int(
            (trusted_predictions["Result_Status"] != "Predicted").sum()
        ),
        ad_in_domain_count=int((ad_frame["AD"] == "In-domain").sum()),
        ad_out_of_domain_count=int((ad_frame["AD"] == "Out-of-domain").sum()),
        ad_unavailable_count=int((ad_frame["AD"] == "Unavailable").sum()),
        graph_paths=graph_paths,
        graph_directory=graph_directory,
        graph_error=graph_error,
        ad_error=ad_error,
    )


class ErbaTab(ttk.Frame):
    def __init__(
        self,
        parent,
        project_root: str,
        install_root: str,
        output_root: str | None = None,
        event_log=None,
        structure_editor=None,
        fixed_subtype: ERBASubtype = ERBASubtype.ER_ALPHA,
    ):
        super().__init__(parent)
        self.project_root = Path(project_root)
        self.install_root = Path(install_root)
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
        self._request_generations = {}
        self._request_predictors = {}
        self._request_specs = {}
        self._request_ad_calculators = {}
        self._model_generation = 0
        self._model_reload_id = 0
        self._ad_reload_id = 0
        self.selected_model_spec = None
        self._load_catalog()
        example = load_shared_example_input(self.project_root)
        self.single_task_var = tk.StringVar(value=ERBATask.CLASSIFICATION.value)
        self.single_subtype_var = tk.StringVar(value=fixed_subtype.value)
        self.batch_task_var = tk.StringVar(value=ERBATask.CLASSIFICATION.value)
        self.batch_subtype_var = tk.StringVar(value=fixed_subtype.value)
        self.example_cas = example.cas
        self.example_smiles = example.smiles
        self.example_workbook = example.workbook
        self.smiles_var = tk.StringVar(value=self.example_smiles)
        self.cas_var = tk.StringVar(value=self.example_cas)
        model_path = (
            self.catalog_path.parent / self.selected_model_spec.relative_path
            if self.selected_model_spec is not None
            else Path()
        )
        ad_path = erba_ad_reference_path(
            self.project_root,
            ERBATask.CLASSIFICATION,
            self.fixed_subtype,
        )
        self.model_path_var = tk.StringVar(
            value=str(model_path) if self.selected_model_spec is not None else ""
        )
        self.ad_ref_path_var = tk.StringVar(
            value=str(ad_path) if ad_path.is_file() else ""
        )
        self.loaded_model_path_var = tk.StringVar()
        self.loaded_ad_ref_path_var = tk.StringVar()
        self.single_status_var = tk.StringVar(value=self.catalog_error or "Ready")
        self.single_prediction_summary_var = tk.StringVar(value="Prediction: -")
        self.single_negative_probability_var = tk.StringVar(value="Non-binding probability: -")
        self.single_positive_probability_var = tk.StringVar(value="Binding probability: -")
        self.single_pic50_var = tk.StringVar(value="pIC50: -")
        self.single_ic50_var = tk.StringVar(value="IC50: - nM")
        self.single_ad_domain_var = tk.StringVar(
            value="Applicability domain: Not evaluated"
        )
        self.single_detail_var = tk.StringVar(value="Awaiting prediction.")
        self.batch_input_var = tk.StringVar(value=str(self.example_workbook))
        default_batch_name = (
            self.example_workbook.name
            if self.example_workbook.is_file()
            else "No input template selected"
        )
        self.batch_input_display_var = tk.StringVar(value=default_batch_name)
        self.batch_destination_var = tk.StringVar(
            value=batch_destination_display(self.batch_input_var.get())
        )
        self.batch_status_var = tk.StringVar(value=self.catalog_error or "Ready")
        self.batch_progress_var = tk.StringVar(value="0% - 0/0 - Ready")
        self.batch_progress_value = tk.DoubleVar(value=0)
        self._active_single_request_id = None
        self._active_ad_request_id = None
        self._active_batch_request_id = None
        self._single_structure_image = None
        self._nearest_reference_structure_image = None
        self._single_ad_graph_image = None
        self._single_detail_lines = [
            "Awaiting prediction.",
            "",
            f"Evidence caveat: {HISTORICAL_EXPOSURE_CAVEAT_COMPACT}",
        ]
        self.single_preview_size = (250, 180)
        self.nearest_reference_preview_size = (250, 180)
        self.single_ad_preview_size = (520, 300)
        self.erba_ad = ERBAApplicabilityDomain(self.project_root)
        self.options_visible = False
        self._build_ui()
        self.single_task_var.trace_add("write", lambda *_: self._route_changed("single"))
        self.single_subtype_var.trace_add("write", lambda *_: self._route_changed("single"))
        self.batch_task_var.trace_add("write", lambda *_: self._route_changed("batch"))
        self.batch_subtype_var.trace_add("write", lambda *_: self._route_changed("batch"))
        self._route_changed()
        self.after(300, self.load_defaults_on_startup)

    def _emit(self, event: str, **fields):
        event_log = getattr(self, "event_log", None)
        if event_log is not None:
            event_log.emit(event, **fields)

    def _workflow_tab_changed(self, _event=None):
        selected = self.notebook.select()
        tab = self.notebook.tab(selected, "text") if selected else ""
        status_var = (
            self.batch_status_var
            if selected == str(self.batch_tab)
            else self.single_status_var
        )
        self.workflow_status_label.configure(textvariable=status_var)
        self._emit("workflow.tab_changed", workflow="erba", tab=tab)

    def _load_catalog(self):
        try:
            self.catalog, self.regression_parity_approved, self.catalog_payload = load_erba_catalog(self.catalog_path)
            choices = self.catalog.choices_for(
                ERBATask.CLASSIFICATION,
                self.fixed_subtype,
            )
            self.selected_model_spec = choices[0] if choices else None
            self.predictor = (
                self._predictor_for_spec(
                    ERBATask.CLASSIFICATION,
                    self.fixed_subtype,
                    self.selected_model_spec,
                )
                if self.selected_model_spec is not None
                else None
            )
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
                model_count=sum(
                    len(self.catalog.choices_for(*route))
                    for route in self.catalog
                ),
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

    @staticmethod
    def _sha256_path(path: Path) -> str:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()

    @staticmethod
    def _trusted_artifacts_for_spec(
        task: ERBATask,
        subtype: ERBASubtype,
        spec: ERBAArtifactSpec,
    ) -> dict[str, Mapping]:
        route_id = f"{task.value}:{subtype.value}"
        for anchor_id, anchor in TRUSTED_ERBA_ARTIFACTS.items():
            if not (
                anchor_id == route_id or anchor_id.startswith(f"{route_id}:")
            ):
                continue
            if (
                isinstance(anchor, Mapping)
                and anchor.get("release_status") == _RELEASED
                and anchor.get("model_id") == spec.model_id
                and anchor.get("sha256") == spec.sha256.lower()
                and anchor.get("size_bytes") == spec.size_bytes
            ):
                return {route_id: anchor}
        raise ValueError(
            f"Model {spec.model_id!r} is not approved by the executable release allowlist."
        )

    def _predictor_for_spec(
        self,
        task: ERBATask,
        subtype: ERBASubtype,
        spec: ERBAArtifactSpec,
    ) -> ERBAPredictor:
        return ERBAPredictor(
            self.catalog_path.parent,
            {(task, subtype): spec},
            regression_parity_approved=self.regression_parity_approved,
            trusted_artifacts=self._trusted_artifacts_for_spec(task, subtype, spec),
        )

    def _released_model_for_path(
        self,
        path: str | Path,
    ) -> tuple[ERBATask, ERBASubtype, ERBAArtifactSpec, Path]:
        candidate = Path(path).expanduser()
        try:
            resolved = candidate.resolve(strict=True)
        except (OSError, RuntimeError) as error:
            raise ValueError(f"Selected model file does not exist: {candidate}") from error
        route = self._selected_route("single")
        choices = self.catalog.choices_for(*route)
        for spec in choices:
            expected = (self.catalog_path.parent / spec.relative_path).resolve(
                strict=True
            )
            if resolved != expected:
                continue
            self._trusted_artifacts_for_spec(*route, spec)
            if (
                resolved.stat().st_size != spec.size_bytes
                or self._sha256_path(resolved) != spec.sha256.lower()
            ):
                raise ValueError(
                    "Selected released model failed its size or SHA-256 integrity check."
                )
            return route[0], route[1], spec, resolved
        raise ValueError(
            "Select a released ERalpha model from the bundled model catalog. "
            "Arbitrary joblib files are not allowed."
        )

    def _approved_ad_reference_for_path(self, path: str | Path) -> Path:
        task, subtype = self._selected_route("single")
        expected = erba_ad_reference_path(self.project_root, task, subtype)
        try:
            expected_resolved = expected.resolve(strict=True)
            selected_resolved = Path(path).expanduser().resolve(strict=True)
        except (OSError, RuntimeError) as error:
            raise ValueError(f"Selected AD reference does not exist: {path}") from error
        if selected_resolved != expected_resolved:
            raise ValueError(
                "Select the approved bundled AD reference for the active ERalpha model."
            )
        if not selected_resolved.is_file():
            raise ValueError("Selected AD reference is not a regular file.")
        return selected_resolved

    def set_status(self, text: str) -> None:
        self.single_status_var.set(text)
        self.batch_status_var.set(text)

    def show_error(self, title: str, error: Exception | str) -> None:
        messagebox.showerror(title, str(error))
        self.set_status(f"Error: {error}")

    def browse_model(self):
        path = filedialog.askopenfilename(
            title="Select model",
            initialdir=str(self.catalog_path.parent),
            filetypes=[("Model file", "*.joblib"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            _, _, _, resolved = self._released_model_for_path(path)
        except Exception as error:
            self.show_error("Model selection failed", error)
            return
        self.model_path_var.set(str(resolved))
        self.set_status("Released model selected. Click Reload model to activate it.")

    def browse_ad_reference(self):
        path = filedialog.askopenfilename(
            title="Select AD reference Excel",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            resolved = self._approved_ad_reference_for_path(path)
        except Exception as error:
            self.show_error("AD reference selection failed", error)
            return
        self.ad_ref_path_var.set(str(resolved))
        self.set_status("Approved AD reference selected. Click Reload AD to activate it.")

    def _activate_model(
        self,
        task: ERBATask,
        subtype: ERBASubtype,
        spec: ERBAArtifactSpec,
        path: Path,
        predictor: ERBAPredictor,
    ) -> None:
        self.model_path_var.set(str(path))
        self.predictor = predictor
        self.selected_model_spec = spec
        self._model_generation += 1
        self.loaded_model_path_var.set(str(path))
        approved_ad = erba_ad_reference_path(self.project_root, task, subtype)
        self.ad_ref_path_var.set(str(approved_ad))
        self.loaded_ad_ref_path_var.set("")
        self._active_ad_request_id = None
        if hasattr(self, "single_prediction_summary_var"):
            self._clear_single_result("Prediction: -")
        if hasattr(self, "batch_result"):
            self._set_text(
                self.batch_result,
                "Run a batch to show the completion summary.",
            )
            self.batch_progress_value.set(0)
            self.batch_progress_var.set("0% - 0/0 - Ready")
        self._route_changed()
        self.set_status(f"Model loaded: {spec.model_id}")

    def load_model_clicked(self):
        selected_path = self.model_path_var.get().strip()
        self._model_reload_id += 1
        reload_id = self._model_reload_id

        def work():
            try:
                task, subtype, spec, path = self._released_model_for_path(
                    selected_path
                )
                predictor = self._predictor_for_spec(task, subtype, spec)
                preflight = predictor.preflight(task, subtype)
                if not preflight.ready:
                    raise RuntimeError(
                        "ERBA route preflight failed "
                        f"({preflight.status_code.value}): {preflight.status_message}"
                    )
                self.after(
                    0,
                    lambda: self._finish_model_reload(
                        reload_id,
                        selected_path,
                        task,
                        subtype,
                        spec,
                        path,
                        predictor,
                    ),
                )
            except Exception as error:
                error_text = str(error)
                self.after(
                    0,
                    lambda: self._model_reload_failed(
                        reload_id,
                        selected_path,
                        error_text,
                    ),
                )

        self.set_status("Loading model...")
        threading.Thread(target=work, daemon=True).start()

    def _finish_model_reload(
        self,
        reload_id,
        selected_path,
        task,
        subtype,
        spec,
        path,
        predictor,
    ):
        if (
            reload_id != self._model_reload_id
            or self.model_path_var.get().strip() != selected_path
        ):
            return
        self._activate_model(task, subtype, spec, path, predictor)
        messagebox.showinfo(
            "Model loaded",
            f"Loaded released model:\n{path}\n\nModel ID: {spec.model_id}",
        )

    def _model_reload_failed(self, reload_id, selected_path, error):
        if (
            reload_id == self._model_reload_id
            and self.model_path_var.get().strip() == selected_path
        ):
            self.show_error("Model load failed", error)

    def fit_ad_clicked(self):
        selected_path = self.ad_ref_path_var.get().strip()
        task, subtype = self._selected_route("single")
        generation = self._model_generation
        self._ad_reload_id += 1
        reload_id = self._ad_reload_id

        def work():
            try:
                path = self._approved_ad_reference_for_path(selected_path)
                fresh_ad = ERBAApplicabilityDomain(self.project_root)
                calculator = fresh_ad.calculator_for(task, subtype)
                calculator.ensure_fitted_from_excel_cached(
                    str(path),
                    cache_path=str(
                        erba_ad_cache_path(self.project_root, task, subtype)
                    ),
                    force_refit=True,
                )
                self.after(
                    0,
                    lambda: self._finish_ad_reload(
                        reload_id,
                        selected_path,
                        path,
                        fresh_ad,
                        calculator,
                        generation,
                    ),
                )
            except Exception as error:
                error_text = str(error)
                self.after(
                    0,
                    lambda: self._ad_reload_failed(
                        reload_id,
                        selected_path,
                        generation,
                        error_text,
                    ),
                )

        self.set_status("Rebuilding AD reference cache...")
        threading.Thread(target=work, daemon=True).start()

    def _finish_ad_reload(
        self,
        reload_id: int,
        selected_path: str,
        path: Path,
        fresh_ad: ERBAApplicabilityDomain,
        calculator,
        generation: int,
    ) -> None:
        if (
            reload_id != self._ad_reload_id
            or generation != self._model_generation
            or self.ad_ref_path_var.get().strip() != selected_path
        ):
            return
        self.ad_ref_path_var.set(str(path))
        self.erba_ad = fresh_ad
        self.loaded_ad_ref_path_var.set(str(path))
        cache_status = calculator.last_cache_status or "AD cache status unavailable."
        self.set_status(f"AD fitted. {cache_status}")
        messagebox.showinfo(
            "AD fitted",
            f"AD reference fitted:\n{path}\n\n{cache_status}",
        )

    def _ad_reload_failed(
        self,
        reload_id,
        selected_path,
        generation,
        error,
    ):
        if (
            reload_id == self._ad_reload_id
            and generation == self._model_generation
            and self.ad_ref_path_var.get().strip() == selected_path
        ):
            self.show_error("AD fitting failed", error)

    def load_defaults_on_startup(self):
        if self.catalog_error or self.selected_model_spec is None:
            return
        model_selection = self.model_path_var.get().strip()
        ad_selection = self.ad_ref_path_var.get().strip()
        generation = self._model_generation
        self._model_reload_id += 1
        self._ad_reload_id += 1
        model_reload_id = self._model_reload_id
        ad_reload_id = self._ad_reload_id

        def work():
            try:
                task, subtype, spec, model_path = self._released_model_for_path(
                    model_selection
                )
                predictor = self._predictor_for_spec(task, subtype, spec)
                preflight = predictor.preflight(task, subtype)
                if not preflight.ready:
                    raise RuntimeError(
                        "ERBA route preflight failed "
                        f"({preflight.status_code.value}): {preflight.status_message}"
                    )
                ad_path = self._approved_ad_reference_for_path(ad_selection)
                fresh_ad = ERBAApplicabilityDomain(self.project_root)
                calculator = fresh_ad.calculator_for(task, subtype)
                calculator.ensure_fitted_from_excel_cached(
                    str(ad_path),
                    cache_path=str(
                        erba_ad_cache_path(self.project_root, task, subtype)
                    ),
                )
                self.after(
                    0,
                    lambda: self._finish_default_loading(
                        model_reload_id,
                        ad_reload_id,
                        model_selection,
                        ad_selection,
                        generation,
                        task,
                        subtype,
                        spec,
                        model_path,
                        predictor,
                        ad_path,
                        fresh_ad,
                        calculator,
                    ),
                )
            except Exception as error:
                error_text = str(error)
                self.after(
                    0,
                    lambda: self._default_loading_failed(
                        model_reload_id,
                        ad_reload_id,
                        model_selection,
                        ad_selection,
                        generation,
                        error_text,
                    ),
                )

        self.set_status("Loading default model and AD reference...")
        threading.Thread(target=work, daemon=True).start()

    def _default_loading_is_current(
        self,
        model_reload_id,
        ad_reload_id,
        model_selection,
        ad_selection,
        generation,
    ):
        return (
            model_reload_id == self._model_reload_id
            and ad_reload_id == self._ad_reload_id
            and generation == self._model_generation
            and self.model_path_var.get().strip() == model_selection
            and self.ad_ref_path_var.get().strip() == ad_selection
        )

    def _finish_default_loading(
        self,
        model_reload_id,
        ad_reload_id,
        model_selection,
        ad_selection,
        generation,
        task,
        subtype,
        spec,
        model_path,
        predictor,
        ad_path,
        fresh_ad,
        calculator,
    ):
        if not self._default_loading_is_current(
            model_reload_id,
            ad_reload_id,
            model_selection,
            ad_selection,
            generation,
        ):
            return
        self._activate_model(task, subtype, spec, model_path, predictor)
        self.ad_ref_path_var.set(str(ad_path))
        self.erba_ad = fresh_ad
        self.loaded_ad_ref_path_var.set(str(ad_path))
        self.set_status(
            f"Ready: {spec.model_id}; "
            f"{calculator.last_cache_status or 'AD reference fitted.'}"
        )

    def _default_loading_failed(
        self,
        model_reload_id,
        ad_reload_id,
        model_selection,
        ad_selection,
        generation,
        error,
    ):
        if self._default_loading_is_current(
            model_reload_id,
            ad_reload_id,
            model_selection,
            ad_selection,
            generation,
        ):
            self.set_status(f"Startup failed: {error}")

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        summary = ttk.Frame(self)
        summary.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 2))
        summary.columnconfigure(0, weight=1)

        self.options_button = ttk.Button(
            summary,
            text="Options",
            command=self.toggle_options,
        )
        self.options_button.grid(row=0, column=1, sticky="e")

        self.options_frame = ttk.LabelFrame(self, text="Options")
        self.options_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(2, 8))
        self.options_frame.columnconfigure(1, weight=1)

        ttk.Label(self.options_frame, text="Model").grid(
            row=0, column=0, sticky="w", padx=6, pady=4
        )
        self.model_entry = ttk.Entry(
            self.options_frame,
            textvariable=self.model_path_var,
        )
        self.model_entry.grid(
            row=0, column=1, sticky="ew", padx=6, pady=4
        )
        self.model_browse_button = ttk.Button(
            self.options_frame,
            text="Browse",
            command=self.browse_model,
        )
        self.model_browse_button.grid(row=0, column=2, padx=6, pady=4)
        self.model_reload_button = ttk.Button(
            self.options_frame,
            text="Reload model",
            command=self.load_model_clicked,
        )
        self.model_reload_button.grid(row=0, column=3, padx=6, pady=4)

        ttk.Label(self.options_frame, text="AD reference").grid(
            row=1, column=0, sticky="w", padx=6, pady=4
        )
        self.ad_entry = ttk.Entry(
            self.options_frame,
            textvariable=self.ad_ref_path_var,
        )
        self.ad_entry.grid(row=1, column=1, sticky="ew", padx=6, pady=4)
        self.ad_browse_button = ttk.Button(
            self.options_frame,
            text="Browse",
            command=self.browse_ad_reference,
        )
        self.ad_browse_button.grid(row=1, column=2, padx=6, pady=4)
        self.ad_reload_button = ttk.Button(
            self.options_frame,
            text="Reload AD",
            command=self.fit_ad_clicked,
        )
        self.ad_reload_button.grid(row=1, column=3, padx=6, pady=4)

        ttk.Label(
            self.options_frame,
            text=(
                "AD is fitted automatically at startup. Reload AD only when "
                "changing the reference file."
            ),
        ).grid(
            row=2,
            column=0,
            columnspan=4,
            sticky="w",
            padx=6,
            pady=(0, 4),
        )
        self.options_frame.grid_remove()

        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        self.mode_notebook = self.notebook
        self.single_tab = ttk.Frame(self.notebook)
        self.batch_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.single_tab, text="Single prediction")
        self.notebook.add(self.batch_tab, text="Batch prediction")
        self._build_single()
        self._build_batch()
        self.workflow_status_label = ttk.Label(
            self,
            textvariable=self.single_status_var,
            anchor="w",
        )
        self.workflow_status_label.grid(
            row=3, column=0, sticky="ew", padx=10, pady=(2, 8)
        )
        self.notebook.bind("<<NotebookTabChanged>>", self._workflow_tab_changed)
        self._register_parity_widgets()

    def toggle_options(self):
        self.options_visible = not self.options_visible
        if self.options_visible:
            self.options_frame.grid()
        else:
            self.options_frame.grid_remove()

    def _register_parity_widgets(self):
        self.parity_widgets = {
            "endpoint_tab": self,
            "options_button": self.options_button,
            "options_frame": self.options_frame,
            "model_entry": self.model_entry,
            "model_browse_button": self.model_browse_button,
            "model_reload_button": self.model_reload_button,
            "ad_entry": self.ad_entry,
            "ad_browse_button": self.ad_browse_button,
            "ad_reload_button": self.ad_reload_button,
            "mode_notebook": self.mode_notebook,
            "single_tab": self.single_tab,
            "batch_tab": self.batch_tab,
            "cas_label": self.cas_label,
            "cas_entry": self.cas_entry,
            "pubchem_button": self.pubchem_button,
            "smiles_label": self.smiles_label,
            "smiles_entry": self.smiles_entry,
            "example_button": self.example_button,
            "single_predict_button": self.single_predict_button,
            "draw_structure_button": self.draw_structure_button,
            "batch_input_entry": self.batch_input_entry,
            "batch_browse_button": self.batch_input_button,
            "batch_destination_entry": self.batch_destination_entry,
            "batch_example_button": self.download_template_button,
            "batch_predict_button": self.run_batch_button,
            "batch_progress": self.batch_progress,
            "batch_status": self.workflow_status_label,
            "batch_result": self.batch_result,
        }

    def pubchem_clicked(self):
        cas = self.cas_var.get()

        def work():
            try:
                self.after(0, lambda: self.set_status("Searching PubChem..."))
                result = cas_to_smiles(cas)
                smiles = result.get("CanonicalSMILES") or result.get(
                    "IsomericSMILES"
                )
                if not smiles:
                    raise RuntimeError(
                        "PubChem did not return a SMILES string."
                    )
                self.after(
                    0,
                    lambda: self._pubchem_complete(
                        smiles,
                        result.get("PubChem_CID"),
                    ),
                )
            except Exception as error:
                error_text = str(error)
                self.after(
                    0,
                    lambda: self.show_error("PubChem search failed", error_text),
                )
        threading.Thread(target=work, daemon=True).start()

    def _pubchem_complete(self, smiles, cid):
        self.smiles_var.set(smiles)
        self.set_status(f"PubChem found CID {cid}")

    def load_example_input(self):
        self.cas_var.set(self.example_cas)
        self.smiles_var.set(self.example_smiles)
        self.single_status_var.set(
            f"Example input restored from {self.example_workbook.name}."
        )

    def open_jsme_popup(self):
        if self.structure_editor is None:
            messagebox.showerror("Draw structure", "Structure editor is unavailable.")
            return
        self.structure_editor(self.smiles_var, self.single_status_var)

    def _build_single(self):
        self.single_tab.columnconfigure(0, weight=1, uniform="single")
        self.single_tab.columnconfigure(1, weight=1, uniform="single")
        self.single_tab.rowconfigure(1, weight=1)

        input_frame = ttk.LabelFrame(self.single_tab, text="Input")
        input_frame.grid(
            row=0,
            column=0,
            columnspan=2,
            sticky="ew",
            padx=8,
            pady=8,
        )
        input_frame.columnconfigure(1, weight=1)

        self.cas_label = ttk.Label(input_frame, text="CAS")
        self.cas_label.grid(row=0, column=0, sticky="w", padx=6, pady=5)
        self.cas_entry = ttk.Entry(input_frame, textvariable=self.cas_var)
        self.cas_entry.grid(row=0, column=1, sticky="ew", padx=6, pady=5)
        self.pubchem_button = ttk.Button(
            input_frame,
            text="PubChem search",
            command=self.pubchem_clicked,
        )
        self.pubchem_button.grid(row=0, column=2, padx=6, pady=5)
        self.example_button = ttk.Button(
            input_frame,
            text="Example input",
            command=self.load_example_input,
        )
        self.example_button.grid(row=0, column=3, padx=6, pady=5)

        self.smiles_label = ttk.Label(input_frame, text="SMILES")
        self.smiles_label.grid(row=1, column=0, sticky="w", padx=6, pady=5)
        self.smiles_entry = ttk.Entry(input_frame, textvariable=self.smiles_var)
        self.smiles_entry.grid(row=1, column=1, sticky="ew", padx=6, pady=5)
        self.single_predict_button = ttk.Button(
            input_frame,
            text="Predict",
            command=self.single_predict_clicked,
        )
        self.single_predict_button.grid(row=1, column=2, padx=6, pady=5)
        self.draw_structure_button = ttk.Button(
            input_frame,
            text="Draw structure",
            command=self.open_jsme_popup,
        )
        self.draw_structure_button.grid(row=1, column=3, padx=6, pady=5)

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
        self.single_ad_domain_label.grid(
            row=1,
            column=0,
            sticky="w",
            pady=(8, 0),
        )

        self.single_prob_canvas = tk.Canvas(
            result_frame,
            height=92,
            bg="white",
            highlightthickness=1,
            highlightbackground="#d0d7de",
        )
        self.single_prob_canvas.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=8,
            pady=(4, 8),
        )
        self.single_prob_canvas.bind("<Configure>", lambda _event: self.draw_probability_graph())
        self.single_probability_frame = self.single_prob_canvas
        self.single_regression_frame = None
        self.single_negative_bar = self.single_positive_bar = None

        ttk.Label(result_frame, text="Details").grid(row=2, column=0, sticky="w", padx=8)
        self.single_result_text = tk.Text(result_frame, height=7, wrap="word", font=("Consolas", 9))
        self.single_result_text.tag_configure(
            "detail_key",
            font=("Consolas", 9, "bold"),
        )
        self.single_result_text.tag_configure(
            "detail_value",
            font=("Consolas", 9),
        )
        self.single_result_text.grid(row=3, column=0, sticky="nsew", padx=8, pady=(2, 8))

        right_frame = ttk.Frame(self.single_tab)
        right_frame.grid(row=1, column=1, sticky="nsew", padx=8, pady=8)
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=0)
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
        self.single_nearest_reference_label = ttk.Label(
            ad_frame,
            textvariable=self.single_nearest_reference_var,
            wraplength=420,
            justify="left",
        )
        self.single_nearest_reference_label.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=8,
            pady=(0, 8),
        )

        ttk.Label(
            self.single_tab,
            text=(
                "CAS to SMILES requires internet access. "
                "Direct SMILES prediction works offline."
            ),
        ).grid(
            row=2,
            column=0,
            columnspan=2,
            sticky="sw",
            padx=8,
            pady=8,
        )
        self._render_detail_lines()

    def _build_batch(self):
        self.batch_tab.columnconfigure(0, weight=1)
        self.batch_tab.rowconfigure(2, weight=1)

        frame = ttk.LabelFrame(self.batch_tab, text="Batch")
        frame.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        frame.columnconfigure(1, weight=1)

        self.batch_input_button = ttk.Button(
            frame,
            text="Input xlsx",
            command=self.browse_batch_input,
        )
        self.batch_input_button.grid(
            row=0,
            column=0,
            sticky="w",
            padx=6,
            pady=4,
        )
        self.batch_input_entry = ttk.Entry(
            frame,
            textvariable=self.batch_input_display_var,
            state="readonly",
        )
        self.batch_input_entry.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=6,
            pady=4,
        )

        ttk.Label(frame, text="Result folder (same as input)").grid(
            row=1,
            column=0,
            sticky="w",
            padx=6,
            pady=4,
        )
        self.batch_destination_entry = ttk.Entry(
            frame,
            textvariable=self.batch_destination_var,
            state="readonly",
        )
        self.batch_destination_entry.grid(
            row=1,
            column=1,
            sticky="ew",
            padx=6,
            pady=4,
        )

        ttk.Label(
            frame,
            text="Enter CAS numbers in the required CAS column.",
        ).grid(
            row=2,
            column=0,
            columnspan=2,
            sticky="w",
            padx=6,
            pady=(3, 1),
        )
        self.download_template_button = ttk.Button(
            frame,
            text="Download template",
            command=self.download_template_clicked,
        )
        self.download_template_button.grid(
            row=3,
            column=0,
            sticky="w",
            padx=6,
            pady=(1, 4),
        )
        self.run_batch_button = ttk.Button(
            frame,
            text="Run batch",
            command=self.batch_predict_clicked,
        )
        self.run_batch_button.grid(
            row=4,
            column=0,
            sticky="w",
            padx=6,
            pady=(0, 5),
        )

        progress = ttk.Frame(self.batch_tab)
        progress.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 4))
        progress.columnconfigure(0, weight=1)
        self.batch_progress = ttk.Progressbar(
            progress,
            maximum=100,
            variable=self.batch_progress_value,
            mode="determinate",
        )
        self.batch_progress.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 8),
        )
        ttk.Label(progress, textvariable=self.batch_progress_var, anchor="w").grid(
            row=0, column=1, sticky="w"
        )

        result_frame = ttk.LabelFrame(self.batch_tab, text="Prediction result")
        result_frame.grid(row=2, column=0, sticky="nsew", padx=8, pady=8)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        self.batch_result = tk.Text(
            result_frame,
            height=12,
            wrap="word",
            state="disabled",
        )
        self.batch_result.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self._set_text(
            self.batch_result,
            "Run a batch to show the completion summary.",
        )

    def _selected_route(self, workflow: str):
        return ERBATask.CLASSIFICATION, self.fixed_subtype

    def _route_available(self, task, subtype):
        if self.catalog_error:
            return False, self.catalog_error
        if task is ERBATask.IC50_REGRESSION and subtype is ERBASubtype.ER_BETA:
            return False, "ERbeta IC50 regression is not supported."
        if task is ERBATask.IC50_REGRESSION and not self.regression_parity_approved:
            return False, "IC50 regression is disabled until catalog parity approval is recorded."
        if (task, subtype) not in self.catalog:
            return False, "This ERBA route is not released in the bundled catalog."
        selected = getattr(self, "selected_model_spec", None)
        if selected is None or selected not in self.catalog.choices_for(task, subtype):
            return False, "Select and reload a released ERalpha model."
        if self.predictor is None:
            return False, "The selected ERalpha model is unavailable."
        return True, ""

    def _route_changed(self, workflow: str | None = None):
        workflows = (workflow,) if workflow else ("single", "batch")
        for current in workflows:
            route = self._selected_route(current)
            ok, reason = self._route_available(*route) if route else (False, "Invalid ERBA route.")
            if current == "single":
                self._configure_single_result_task(route[0] if route else None)
                state = "disabled" if self._active_single_request_id is not None else ("normal" if ok else "disabled")
                self.single_predict_button.configure(state=state)
                if reason and self._active_single_request_id is None:
                    self.single_status_var.set(reason)
            else:
                state = "disabled" if self._active_batch_request_id is not None else ("normal" if ok else "disabled")
                self.run_batch_button.configure(state=state)
                if self._active_batch_request_id is None:
                    self.batch_input_button.configure(state="normal")
                    self.download_template_button.configure(state="normal")
                    if reason:
                        self.batch_status_var.set(reason)

    def _configure_single_result_task(self, task):
        probability_frame = getattr(self, "single_probability_frame", None)
        if probability_frame is None:
            return
        if task is ERBATask.CLASSIFICATION:
            probability_frame.grid()
        else:
            probability_frame.grid_remove()

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
        request_id = self._request_id
        self._started_at[request_id] = time.perf_counter()
        self._request_generations[request_id] = self._model_generation
        self._request_predictors[request_id] = self.predictor
        self._request_specs[request_id] = self.selected_model_spec
        self._request_ad_calculators[request_id] = self.erba_ad
        return request_id, task, subtype, self.selected_model_spec.model_id

    def _finish_duration_ms(self, request_id: int) -> int:
        started_at = getattr(self, "_started_at", {})
        started = started_at.pop(request_id, time.perf_counter())
        return max(0, int((time.perf_counter() - started) * 1000))

    def _forget_request(self, request_id: int) -> None:
        for name in (
            "_request_generations",
            "_request_predictors",
            "_request_specs",
            "_request_ad_calculators",
        ):
            getattr(self, name, {}).pop(request_id, None)

    def single_predict_clicked(self):
        snapshot = self._snapshot("single")
        if not snapshot:
            return
        request_id, task, subtype, model_id = snapshot
        direct_smiles, cas = self.smiles_var.get().strip(), self.cas_var.get().strip()
        self._active_single_request_id = request_id
        self.single_predict_button.configure(state="disabled")
        self.single_status_var.set("Predicting ERBA route...")
        predictor = getattr(self, "_request_predictors", {}).get(
            request_id,
            self.predictor,
        )

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
                    error_message = f"CAS lookup failed: {error}"
                    self.after(
                        0,
                        lambda: self._single_complete(
                            request_id,
                            task,
                            subtype,
                            model_id,
                            None,
                            error_message,
                        ),
                    )
                    return
            result = predictor.predict(
                ERBARequest(smiles=smiles, task=task, subtype=subtype)
            )
            self.after(0, lambda: self._single_complete(request_id, task, subtype, model_id, result, "", smiles))
        threading.Thread(target=work, daemon=True).start()

    def _single_complete(self, request_id, task, subtype, model_id, result, error, smiles=""):
        duration_ms = self._finish_duration_ms(request_id)
        if request_id != self._active_single_request_id:
            self._forget_request(request_id)
            self._emit("model.inference_stale", workflow="erba", task=task.value, subtype=subtype.value,
                       model_id=model_id, correlation_id=request_id, duration_ms=duration_ms)
            return
        self._active_single_request_id = None
        if not self._snapshot_is_current("single", request_id, task, subtype, model_id):
            self._forget_request(request_id)
            self._route_changed("single")
            self._emit("model.inference_stale", workflow="erba", task=task.value, subtype=subtype.value,
                       model_id=model_id, correlation_id=request_id, duration_ms=duration_ms)
            return
        self._route_changed("single")
        if error:
            self._forget_request(request_id)
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
        else:
            self._forget_request(request_id)

    def _evaluate_single_ad(self, request_id, task, subtype, smiles):
        try:
            route_ad = getattr(self, "_request_ad_calculators", {}).get(
                request_id,
                self.erba_ad,
            )
            calculator, ad_result, fingerprint = route_ad.evaluate(
                task,
                subtype,
                smiles,
            )
            graph_path = save_single_ad_plot(calculator, fingerprint, str(validate_mutable_directory(self.output_root) / "erba_ad"))
            self.after(
                0,
                lambda: self._render_single_ad(
                    request_id, task, subtype, ad_result, graph_path, ""
                ),
            )
        except Exception as error:
            error_message = str(error)
            self.after(
                0,
                lambda: self._render_single_ad(
                    request_id, task, subtype, None, "", error_message
                ),
            )

    def _render_single_ad(
        self, request_id, task, subtype, ad_result, graph_path, error
    ):
        if (
            request_id != self._active_ad_request_id
            or not self._snapshot_is_current(
                "single",
                request_id,
                task,
                subtype,
                getattr(
                    getattr(self, "_request_specs", {}).get(request_id),
                    "model_id",
                    "",
                ),
            )
        ):
            self._forget_request(request_id)
            return
        self._active_ad_request_id = None
        if error:
            self._forget_request(request_id)
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
            self._single_ad_graph_image = self._make_preview_photo(
                graph_path,
                self.single_ad_preview_size,
            )
            self.single_ad_graph_label.configure(image=self._single_ad_graph_image, text="")
        self._forget_request(request_id)

    def _clear_single_result(self, summary):
        if hasattr(self, "_active_ad_request_id"):
            self._active_ad_request_id = None
        self.single_prediction_summary_var.set(summary)
        if hasattr(self, "single_prediction_label"):
            self.single_prediction_label.configure(fg="#57606a")
        self.single_negative_probability_var.set("Non-binding probability: -")
        self.single_positive_probability_var.set("Binding probability: -")
        self.single_pic50_var.set("pIC50: -")
        self.single_ic50_var.set("IC50: - nM")
        self.single_detail_var.set("No prediction result is available.")
        self._single_detail_lines = [
            f"Status: {summary}",
            "",
            f"Evidence caveat: {HISTORICAL_EXPOSURE_CAVEAT_COMPACT}",
        ]
        self._single_positive_probability = None
        self.draw_probability_graph()
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
        reference_structure = getattr(
            self,
            "single_nearest_reference_structure_label",
            None,
        )
        if reference_structure is not None:
            reference_structure.configure(text="No reference", image="")
        self._nearest_reference_structure_image = None
        if hasattr(self, "single_ad_graph_label"):
            self.single_ad_graph_label.configure(
                text="AD graph will appear after prediction.", image=""
            )
        self._single_ad_graph_image = None
        self._render_detail_lines()

    def _render_single_result(self, result):
        self._configure_single_result_task(result.task)
        status = result.status_message or result.status_code.value
        if result.status_code is not ERBAStatusCode.OK:
            self._clear_single_result("Prediction unavailable")
            self.single_detail_var.set(f"Status: {status}")
            self._single_detail_lines = [
                f"Status: {result.status_code.value}",
                f"Status message: {status}",
                f"Model ID: {result.model_id or ''}",
                f"Preprocessing policy ID: {result.preprocessing_policy_id or ''}",
                f"Evidence scope: {result.evidence_scope or ''}",
                f"Evidence caveat: {result.evidence_caveat or HISTORICAL_EXPOSURE_CAVEAT_COMPACT}",
            ]
            self._render_detail_lines()
            return
        if result.task is ERBATask.CLASSIFICATION:
            label = _binding_display_label(result.binding_label) or "Unknown"
            self.single_prediction_summary_var.set(f"Binding prediction: {label}")
            prediction_label = getattr(self, "single_prediction_label", None)
            if prediction_label is not None:
                prediction_label.configure(
                    fg=(
                        "#cf222e"
                        if label == "Binding"
                        else "#1a7f37"
                        if label == "Non-binding"
                        else "#57606a"
                    )
                )
            negative = float(result.non_binding_probability or 0)
            positive = float(result.binding_probability or 0)
            self.single_negative_probability_var.set(f"Non-binding probability: {negative:.1%}")
            self.single_positive_probability_var.set(f"Binding probability: {positive:.1%}")
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
            _binding_display_label(result.binding_label) or "Unknown"
            if result.task is ERBATask.CLASSIFICATION
            else "-"
        )
        cas_var = getattr(self, "cas_var", None)
        cas = cas_var.get().strip() if cas_var is not None else ""
        self._single_detail_lines = [
            f"CAS: {cas}",
            f"Input SMILES: {result.raw_smiles}",
            f"Canonical SMILES: {result.model_smiles or ''}",
            f"Task: {'Binding classification' if result.task is ERBATask.CLASSIFICATION else 'IC50 regression'}",
            f"Receptor subtype: {'ERalpha' if result.subtype is ERBASubtype.ER_ALPHA else 'ERbeta'}",
            f"Non-binding probability: {float(result.non_binding_probability or 0):.6f}",
            f"Binding probability: {float(result.binding_probability or 0):.6f}",
            f"Binding prediction: {label}",
            f"Status: {result.status_code.value}",
            f"Status message: {status}",
            f"Model ID: {result.model_id or ''}",
            f"Model SHA256: {result.model_sha256 or ''}",
            f"Preprocessing policy ID: {result.preprocessing_policy_id or ''}",
            f"Evidence scope: {result.evidence_scope or ''}",
            f"Evidence caveat: {result.evidence_caveat or HISTORICAL_EXPOSURE_CAVEAT_COMPACT}",
        ]
        self._draw_single_structure(result.model_smiles)
        self._render_detail_lines()

    def _render_detail_lines(self):
        details = getattr(self, "single_result_text", None)
        if details is None:
            return
        details.delete("1.0", "end")
        for line in self._single_detail_lines:
            text = str(line)
            if ":" in text:
                key, value = text.split(":", 1)
                details.insert("end", f"{key}:", "detail_key")
                details.insert("end", f"{value}\n", "detail_value")
            else:
                details.insert("end", f"{text}\n", "detail_value")

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
            self._nearest_reference_structure_image = self._make_preview_photo(
                path,
                self.nearest_reference_preview_size,
            )
            label_parts = []
            for key in ("label", "Label", "Activity", "active", "inactive"):
                if reference.get(key):
                    label = _binding_display_label(reference.get(key))
                    label_parts.append(
                        f"label: {label or reference.get(key)}"
                    )
                    break
            for key in ("CID", "PubChem_CID"):
                if reference.get(key):
                    label_parts.append(f"CID: {reference.get(key)}")
                    break
            widget.configure(
                image=self._nearest_reference_structure_image,
                text=" / ".join(label_parts) or "Nearest reference",
                compound="top",
                wraplength=180,
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
        width = max(canvas.winfo_width(), 360)
        margin = 18
        label_width = 94
        bar_x = margin + label_width
        bar_w = max(width - bar_x - margin - 48, 120)
        inactive_y = 22
        active_y = 58
        bar_h = 18

        def draw_bar(y, label, value, color):
            value = max(0.0, min(1.0, float(value)))
            canvas.create_text(
                margin,
                y + bar_h / 2,
                text=label,
                anchor="w",
                fill="#24292f",
            )
            canvas.create_rectangle(bar_x, y, bar_x + bar_w, y + bar_h, fill="#eef2f7", outline="#d0d7de")
            canvas.create_rectangle(bar_x, y, bar_x + bar_w * value, y + bar_h, fill=color, outline=color)
            canvas.create_text(
                bar_x + bar_w + 8,
                y + bar_h / 2,
                text=f"{value:.3f}",
                anchor="w",
                fill="#24292f",
            )

        if positive is None:
            draw_bar(inactive_y, "Non-binding", 0.0, "#2da44e")
            draw_bar(active_y, "Binding", 0.0, "#cf222e")
            canvas.create_text(bar_x, 8, text="Run prediction to show probabilities", anchor="w", fill="#57606a")
            return
        positive = max(0.0, min(1.0, float(positive)))
        draw_bar(inactive_y, "Non-binding", 1.0 - positive, "#2da44e")
        draw_bar(active_y, "Binding", positive, "#cf222e")

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
            self._single_structure_image = self._make_preview_photo(
                path,
                self.single_preview_size,
            )
            structure.configure(image=self._single_structure_image, text="")
        except Exception as error:
            structure.configure(text=f"Structure preview unavailable: {error}", image="")
            self._single_structure_image = None

    @staticmethod
    def _make_preview_photo(path: str | Path, size: tuple[int, int]):
        image = Image.open(path).convert("RGB")
        image.thumbnail(size)
        canvas = Image.new("RGB", size, "white")
        x = (size[0] - image.width) // 2
        y = (size[1] - image.height) // 2
        canvas.paste(image, (x, y))
        return ImageTk.PhotoImage(canvas)

    def browse_batch_input(self):
        path = filedialog.askopenfilename(
            title="Select input Excel",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            source = Path(path).expanduser().resolve(strict=False)
            self.batch_input_var.set(str(source))
            self.batch_input_display_var.set(source.name)
            try:
                destination = validate_mutable_directory(
                    source.parent,
                    forbidden_roots=(self.project_root, self.install_root),
                )
            except Exception as error:
                self.batch_destination_var.set(
                    f"Unavailable — input folder cannot receive results: {source.parent}"
                )
                self.batch_status_var.set(
                    "Copy or download the input workbook to a writable folder outside "
                    f"the application files. Details: {error}"
                )
                return
            self.batch_destination_var.set(str(destination))
            self.batch_status_var.set(
                f"Result workbook will be saved beside the input: {destination}"
            )

    def download_template_clicked(self):
        try:
            path = filedialog.asksaveasfilename(
                title="Download batch template",
                defaultextension=".xlsx",
                initialfile="Template.xlsx",
                filetypes=[("Excel", "*.xlsx")],
            )
            if not path:
                return
            destination = Path(path).expanduser().resolve(strict=False)
            validate_mutable_directory(
                destination.parent,
                forbidden_roots=(self.project_root, self.install_root),
            )
            pd.DataFrame(columns=["CAS"]).to_excel(destination, index=False)
        except Exception as error:
            message = (
                "Save the template in a writable folder outside the application files. "
                f"Details: {error}"
            )
            self.batch_status_var.set(message)
            messagebox.showerror("Template download failed", message)
            return
        self.batch_input_var.set(str(destination))
        self.batch_input_display_var.set(destination.name)
        self.batch_destination_var.set(batch_destination_display(destination))
        self.batch_status_var.set(f"Template downloaded: {destination}")
        messagebox.showinfo(
            "Template downloaded",
            f"Saved:\n{destination}\n\nEnter CAS values, save the file, then run batch.",
        )

    def _set_batch_controls_active(self, active: bool):
        state = "disabled" if active else "normal"
        for widget in (
            self.batch_input_button,
            self.download_template_button,
            self.run_batch_button,
        ):
            widget.configure(state=state)

    @staticmethod
    def _column_lookup(frame, aliases):
        normalized = {str(column).strip().lower(): column for column in frame.columns}
        return next((normalized[name] for name in aliases if name in normalized), None)

    def _reject_batch_start(self, request_id: int, message: str) -> None:
        self.batch_status_var.set(message)
        self.batch_progress_value.set(100)
        self.batch_progress_var.set("100% - 0/0 - Failed")
        result = getattr(self, "batch_result", None)
        if result is not None:
            self._set_text(
                result,
                f"Batch prediction failed.\n\nTechnical details: {message}",
            )
        self._started_at.pop(request_id, None)
        self._forget_request(request_id)

    def batch_predict_clicked(self):
        snapshot = self._snapshot("batch")
        if not snapshot:
            return
        request_id, task, subtype, model_id = snapshot
        input_path = self.batch_input_var.get().strip()
        self.batch_destination_var.set(batch_destination_display(input_path))
        if not input_path:
            self._reject_batch_start(request_id, "Choose an input xlsx.")
            return
        try:
            source = Path(input_path).expanduser().resolve(strict=False)
            if not source.is_file():
                raise FileNotFoundError(f"Input workbook does not exist: {source}")
            output_dir = validate_mutable_directory(
                source.parent,
                forbidden_roots=(self.project_root, self.install_root),
            )
        except Exception as error:
            self._reject_batch_start(
                request_id,
                "The ERBA result workbook must be saved beside the input workbook. "
                "Choose an input workbook in a writable folder outside the application files. "
                f"Details: {error}",
            )
            return
        try:
            predictor = getattr(self, "_request_predictors", {}).get(
                request_id,
                self.predictor,
            )
            preflight = predictor.preflight(task, subtype)
            if not preflight.ready:
                raise RuntimeError(
                    f"ERBA route preflight failed ({preflight.status_code.value}): {preflight.status_message}"
                )
        except Exception as error:
            self._reject_batch_start(request_id, str(error))
            return
        input_path = str(source)
        self.batch_destination_var.set(str(output_dir))
        self._set_batch_controls_active(True)
        self._active_batch_request_id = request_id
        self._set_batch_progress(request_id, "Ready to read input", 0, 0, 0)
        self.batch_status_var.set("Running ERBA batch...")
        threading.Thread(
            target=self._batch_work,
            args=(request_id, task, subtype, model_id, input_path),
            daemon=True,
        ).start()

    def _batch_progress_from_worker(self, request_id, stage, current, total, percent):
        """Schedule all Tk mutation on the event loop, never on the batch worker."""
        self.after(
            0,
            lambda: self._set_batch_progress(request_id, stage, current, total, percent),
        )

    def _set_batch_progress(self, request_id, stage, current, total, percent):
        if request_id != getattr(self, "_active_batch_request_id", None):
            return
        current_generation = getattr(self, "_model_generation", 0)
        request_generation = getattr(self, "_request_generations", {}).get(
            request_id,
            current_generation,
        )
        if request_generation != current_generation:
            return
        percent = max(0, min(100, int(percent)))
        detail = f"{current}/{total}" if total else "0/0"
        progress_var = getattr(self, "batch_progress_var", None)
        progress_value = getattr(self, "batch_progress_value", None)
        if progress_var is not None:
            progress_var.set(f"{percent}% - {detail} - {stage}")
        if progress_value is not None:
            progress_value.set(percent)

    def _batch_work(self, request_id, task, subtype, model_id, input_path):
        try:
            predictor = getattr(self, "_request_predictors", {}).get(
                request_id,
                self.predictor,
            )
            route_ad = getattr(self, "_request_ad_calculators", {}).get(
                request_id,
                self.erba_ad,
            )
            spec = getattr(self, "_request_specs", {}).get(request_id)
            if spec is None:
                selected = getattr(self, "selected_model_spec", None)
                spec = selected or self.catalog[(task, subtype)]
            export_result = export_erba_batch(
                input_path,
                task,
                subtype,
                predictor,
                route_ad,
                spec,
                self.catalog_payload,
                progress_callback=lambda stage, current, total, percent: self._batch_progress_from_worker(
                    request_id, stage, current, total, percent
                ),
                forbidden_roots=(self.project_root, self.install_root),
            )
            self.after(0, lambda: self._batch_complete(
                request_id, task, subtype, model_id, export_result, ""
            ))
        except Exception as error:
            error_message = str(error)
            self.after(0, lambda: self._batch_complete(
                request_id, task, subtype, model_id, None, error_message
            ))

    def _batch_complete(
        self,
        request_id,
        task,
        subtype,
        model_id,
        export_result: ERBABatchExportResult | None,
        error,
    ):
        duration_ms = self._finish_duration_ms(request_id)
        if request_id != self._active_batch_request_id:
            self._forget_request(request_id)
            if self._active_batch_request_id is None:
                self._route_changed("batch")
            self._emit("batch.inference_stale", workflow="erba", task=task.value, subtype=subtype.value,
                       model_id=model_id, correlation_id=request_id, duration_ms=duration_ms)
            return
        if not self._snapshot_is_current(
            "batch",
            request_id,
            task,
            subtype,
            model_id,
        ):
            self._active_batch_request_id = None
            self._set_batch_controls_active(False)
            self._forget_request(request_id)
            self._route_changed("batch")
            self._emit("batch.inference_stale", workflow="erba", task=task.value, subtype=subtype.value,
                       model_id=model_id, correlation_id=request_id, duration_ms=duration_ms)
            return
        self._set_batch_progress(
            request_id,
            "Failed" if error else "Completed",
            export_result.count if export_result is not None else 0,
            export_result.count if export_result is not None else 0,
            100,
        )
        self._active_batch_request_id = None
        self._set_batch_controls_active(False)
        self._forget_request(request_id)
        self._route_changed("batch")
        if error:
            self.batch_status_var.set(f"ERBA batch failed: {error}")
            self._set_text(
                self.batch_result,
                f"Batch prediction failed.\n\nTechnical details: {error}",
            )
            self._emit("batch.inference_complete", workflow="erba", task=task.value, subtype=subtype.value,
                       model_id=model_id, correlation_id=request_id, duration_ms=duration_ms,
                       row_count=0, status="failed", exception_type="BatchError")
            return
        if export_result is None:
            raise RuntimeError("ERBA batch completed without an export result.")
        destination = export_result.destination
        self.batch_destination_var.set(str(destination.parent))
        self.batch_status_var.set(f"ERBA batch complete: {destination}")
        self._emit("batch.inference_complete", workflow="erba", task=task.value, subtype=subtype.value,
                   model_id=model_id, correlation_id=request_id, duration_ms=duration_ms,
                   row_count=export_result.count, status="ok")
        caveat = (
            HISTORICAL_EXPOSURE_CAVEAT_COMPACT
            if task is ERBATask.CLASSIFICATION
            else REGRESSION_EVIDENCE_CAVEAT_COMPACT
        )
        lines = [
            "ERalpha batch prediction completed.",
            "",
            f"Total rows: {export_result.count}",
            f"Binding: {export_result.binding_count}",
            f"Non-binding: {export_result.non_binding_count}",
            f"Not predicted: {export_result.not_predicted_count}",
            f"AD In-domain: {export_result.ad_in_domain_count}",
            f"AD Out-of-domain: {export_result.ad_out_of_domain_count}",
            f"AD Unavailable: {export_result.ad_unavailable_count}",
            "",
            f"Output workbook: {destination}",
        ]
        if export_result.graph_paths:
            lines.extend(
                (
                    f"Graph files: {len(export_result.graph_paths)}",
                    f"Graph directory: {export_result.graph_directory}",
                )
            )
        else:
            lines.extend(
                (
                    "Graph files: 0",
                    "Graph directory: Not generated",
                    f"Graph warning: {export_result.graph_error or 'no graphs were generated.'}",
                )
            )
        if export_result.ad_error:
            lines.append(f"AD warning: {export_result.ad_error}")
        lines.extend(("", caveat))
        self._set_text(self.batch_result, "\n".join(lines))

    def _snapshot_is_current(self, workflow, request_id, task, subtype, model_id):
        active_id = (
            self._active_single_request_id if workflow == "single" else self._active_batch_request_id
        )
        selected = getattr(self, "selected_model_spec", None)
        if selected is None:
            selected = self.catalog.get((task, subtype))
        generation = getattr(self, "_model_generation", 0)
        request_generation = getattr(self, "_request_generations", {}).get(
            request_id,
            generation,
        )
        return (
            active_id in (None, request_id)
            and self._selected_route(workflow) == (task, subtype)
            and selected is not None
            and selected.model_id == model_id
            and request_generation == generation
        )

    @staticmethod
    def _set_text(widget, text):
        widget.configure(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert("1.0", text)
        widget.configure(state="disabled")
