"""Stable, ERBA-only public contract shared by adapters and the GUI."""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Any, Mapping, Sequence


class ERBAWorkflow(str, Enum):
    ERBA = "erba"


class ERBATask(str, Enum):
    CLASSIFICATION = "classification"
    IC50_REGRESSION = "ic50_regression"


class ERBASubtype(str, Enum):
    ER_ALPHA = "er_alpha"
    ER_BETA = "er_beta"


class ERBAStatusCode(str, Enum):
    OK = "ok"
    UNSUPPORTED_COMBINATION = "unsupported_combination"
    INVALID_SMILES = "invalid_smiles"
    BLANK_SMILES = "blank_smiles"
    WILDCARD_SMILES = "wildcard_smiles"
    CXSMILES_NOT_ALLOWED = "cxsmiles_not_allowed"
    NO_CARBON = "no_carbon"
    METAL_RETAINED = "metal_retained"
    MODEL_NOT_FOUND = "model_not_found"
    MODEL_PATH_UNSAFE = "model_path_unsafe"
    MODEL_INTEGRITY_FAILED = "model_integrity_failed"
    MODEL_SCHEMA_INVALID = "model_schema_invalid"
    MODEL_POLICY_MISMATCH = "model_policy_mismatch"
    PREDICTION_FAILED = "prediction_failed"
    INFERENCE_INTERNAL_ERROR = "inference_internal_error"
    REGRESSION_PARITY_NOT_APPROVED = "regression_parity_not_approved"


CLASSIFICATION_PARENT_POLICY_ID = "erba_binding_classification_parent_v2"
REGRESSION_PARENT_POLICY_ID = "erba_ic50_regression_parent_parity_v1"
RAW_SMILES_PIPELINE_SCHEMA_ID = "raw_smiles_pipeline_schema_v1"
MODEL_SCHEMA_ID = RAW_SMILES_PIPELINE_SCHEMA_ID
ARTIFACT_SCHEMA_VERSION = 1
CLASSIFICATION_FEATURE_SCHEMA_ID = "rdkit_fp_2048"
REGRESSION_FEATURE_SCHEMA_ID = "avalon_fp_2048"
ERBA_BINDING_CLASSIFICATION_EXCEL_CONTRACT_ID = "erba.binding.classification.excel.v2"
ERBA_ERALPHA_IC50_EXCEL_CONTRACT_ID = "erba.eralpha.ic50.excel.v1"
ERTA_LEGACY_EXCEL_CONTRACT_ID = "erta.legacy.excel.v1"
ERBA_INPUT_SHEET_NAME = "Input"
ERBA_PREDICTIONS_SHEET_NAME = "Predictions"
ERBA_METADATA_SHEET_NAME = "Metadata"
ERBA_GUIDE_SHEET_NAME = "Guide"
ERBA_CANONICAL_INPUT_COLUMNS = ("Row_ID", "CAS", "SMILES")
ERBA_INPUT_HEADER_ALIASES = {
    "Row_ID": ("row_id", "row id", "id", "identifier"),
    "CAS": ("cas", "cas no", "cas no.", "cas rn", "cas_number", "cas number"),
    "SMILES": ("smiles", "canonical_smiles", "canonical smiles", "isomeric_smiles", "isomeric smiles"),
}
ERBA_METADATA_COLUMNS = (
    "Excel_Contract_ID", "Model_ID", "Model_SHA256", "Protocol_SHA256",
    "Source_Manifest_SHA256", "Historical_Exposure_Manifest_SHA256",
    "Split_Manifest_SHA256", "Nested_CV_SHA256", "Internal_Resplit_SHA256",
    "Preprocessing_Parity_SHA256", "Report_SHA256", "Caveat_SHA256",
    "Performance_Evidence_Scope", "Evidence_Caveat",
)
ERBA_PERFORMANCE_EVIDENCE_SCOPE = "internal_historically_exposed"
ERBA_REGRESSION_EVIDENCE_SCOPE = "internal_regression_model"
ERBA_FATAL_ARTIFACT_STATUS_CODES = frozenset({
    ERBAStatusCode.MODEL_NOT_FOUND, ERBAStatusCode.MODEL_PATH_UNSAFE,
    ERBAStatusCode.MODEL_INTEGRITY_FAILED, ERBAStatusCode.MODEL_SCHEMA_INVALID,
    ERBAStatusCode.MODEL_POLICY_MISMATCH, ERBAStatusCode.INFERENCE_INTERNAL_ERROR,
})

HISTORICAL_EXPOSURE_CAVEAT_COMPACT = "Internal model; historically exposed development data — not for regulatory use"
HISTORICAL_EXPOSURE_CAVEAT_FULL = (
    "This classification evidence uses an internal resplit of model-ready rows with known historical "
    "exposure to prior feature generation, training, scoring, and model-family analysis. Nested grouped "
    "CV controls the new selection run, and the sealed resplit is a one-time sensitivity/release gate. "
    "Reported performance is internal and must not be represented as regulatory evidence or as an estimate "
    "from a historically unexposed population."
)
REGRESSION_EVIDENCE_CAVEAT_COMPACT = "Internal ERalpha pIC50 regression model — not for regulatory use"
REGRESSION_EVIDENCE_CAVEAT_FULL = (
    "ERalpha pIC50 regression performance is internal model evidence. The route predicts "
    "pIC50 = -log10(IC50 [mol/L]) and does not provide a classification label or confidence interval."
)

CLASSIFICATION_OUTPUT_COLUMNS = (
    "row_index", "workflow", "task", "subtype", "raw_smiles", "model_smiles", "status_code",
    "status_message", "non_binding_probability", "binding_probability", "binding_label",
    "model_id", "model_sha256", "preprocessing_policy_id", "evidence_caveat",
    "Result_Status", "Reason_Category", "Reason_Description", "Recommended_Action",
)
REGRESSION_OUTPUT_COLUMNS = (
    "row_index", "workflow", "task", "subtype", "raw_smiles", "model_smiles", "status_code",
    "status_message", "pic50", "ic50_nm", "model_id", "model_sha256",
    "preprocessing_policy_id", "evidence_caveat",
    "Result_Status", "Reason_Category", "Reason_Description", "Recommended_Action",
)
ERBA_CLASSIFICATION_PREDICTION_COLUMNS = ("Row_ID", "CAS", *CLASSIFICATION_OUTPUT_COLUMNS)
ERBA_REGRESSION_PREDICTION_COLUMNS = ("Row_ID", "CAS", *REGRESSION_OUTPUT_COLUMNS)

SUPPORTED_COMBINATIONS = frozenset({
    (ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA),
    (ERBATask.CLASSIFICATION, ERBASubtype.ER_BETA),
    (ERBATask.IC50_REGRESSION, ERBASubtype.ER_ALPHA),
})
SUPPORT_MATRIX = {
    (ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA): True,
    (ERBATask.CLASSIFICATION, ERBASubtype.ER_BETA): True,
    (ERBATask.IC50_REGRESSION, ERBASubtype.ER_ALPHA): True,
    (ERBATask.IC50_REGRESSION, ERBASubtype.ER_BETA): False,
}


@dataclass(frozen=True)
class ERBARequest:
    smiles: str
    task: ERBATask
    subtype: ERBASubtype
    row_index: int = 0
    passthrough: Mapping[str, Any] | None = None
    workflow: ERBAWorkflow = ERBAWorkflow.ERBA


@dataclass(frozen=True)
class ERBAArtifactSpec:
    relative_path: str
    size_bytes: int
    sha256: str
    model_id: str

@dataclass(frozen=True)
class ERBARoutePreflight:
    task: ERBATask
    subtype: ERBASubtype
    status_code: ERBAStatusCode
    status_message: str = ""

    @property
    def ready(self) -> bool:
        return self.status_code is ERBAStatusCode.OK


@dataclass(frozen=True)
class ERBAResult:
    row_index: int
    task: ERBATask
    subtype: ERBASubtype
    raw_smiles: str
    model_smiles: str | None
    status_code: ERBAStatusCode
    status_message: str
    non_binding_probability: float | None = None
    binding_probability: float | None = None
    binding_label: str | None = None
    pic50: float | None = None
    ic50_nm: float | None = None
    model_id: str | None = None
    model_sha256: str | None = None
    preprocessing_policy_id: str | None = None
    evidence_caveat: str = HISTORICAL_EXPOSURE_CAVEAT_COMPACT
    passthrough: Mapping[str, Any] | None = None
    workflow: ERBAWorkflow = ERBAWorkflow.ERBA
    protocol_sha256: str | None = None
    nested_cv_sha256: str | None = None
    internal_resplit_sha256: str | None = None
    report_sha256: str | None = None
    source_manifest_sha256: str | None = None
    preprocessing_parity_sha256: str | None = None
    evidence_scope: str | None = None


def combination_is_supported(task: ERBATask, subtype: ERBASubtype) -> bool:
    return SUPPORT_MATRIX.get((task, subtype), False)


def output_columns(task: ERBATask) -> Sequence[str]:
    return CLASSIFICATION_OUTPUT_COLUMNS if task is ERBATask.CLASSIFICATION else REGRESSION_OUTPUT_COLUMNS
