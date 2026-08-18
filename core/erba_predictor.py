"""Safe, task-isolated ERBA inference.  This module never imports legacy ERTA code."""
from __future__ import annotations

import hashlib
import io
import re
import stat
from pathlib import Path
from typing import Any, Iterable, Mapping

import numpy as np

from .contracts import (
    ARTIFACT_SCHEMA_VERSION,
    CLASSIFICATION_PARENT_POLICY_ID,
    HISTORICAL_EXPOSURE_CAVEAT_COMPACT,
    HISTORICAL_EXPOSURE_CAVEAT_FULL,
    RAW_SMILES_PIPELINE_SCHEMA_ID,
    REGRESSION_PARENT_POLICY_ID,
    ERBAArtifactSpec,
    ERBARoutePreflight,
    ERBARequest,
    ERBAResult,
    ERBAStatusCode,
    ERBASubtype,
    ERBATask,
    ERBAWorkflow,
    combination_is_supported,
)
from .erba_features import ERBARawSmilesFeatures
from .erba_preprocessing import avalon_fp_2048, classification_parent_smiles, regression_parent_smiles
from .erba_release_allowlist import TRUSTED_ERBA_ARTIFACTS


_CLASSIFICATION_EVIDENCE_SCOPE = "internal_historically_exposed"
_CLASSIFICATION_CAVEAT = HISTORICAL_EXPOSURE_CAVEAT_FULL
_REGRESSION_TARGET = "pIC50 = -log10(IC50 [mol/L])"
_SHA256 = re.compile(r"^[0-9a-f]{64}$")


class _ArtifactError(Exception):
    def __init__(self, code: ERBAStatusCode, message: str):
        super().__init__(message)
        self.code = code


def _as_enum(value: Any, enum_type: type[Any]) -> Any:
    return value if isinstance(value, enum_type) else enum_type(value)




def _has_sha256(metadata: Mapping[str, Any], key: str) -> bool:
    value = metadata.get(key)
    return isinstance(value, str) and _SHA256.fullmatch(value) is not None


class ERBAPredictor:
    """Loads only catalogued bundled artifacts, after path and integrity checks."""

    def __init__(
        self,
        bundled_model_root: str | Path,
        catalog: Mapping[tuple[ERBATask, ERBASubtype], ERBAArtifactSpec],
        *,
        regression_parity_approved: bool = False,
        trusted_artifacts: Mapping[str, Mapping[str, Any]] | None = None,
    ) -> None:
        self._root = Path(bundled_model_root).resolve(strict=False)
        self._catalog = dict(catalog)
        self._regression_parity_approved = regression_parity_approved
        self._trusted_artifacts = dict(
            TRUSTED_ERBA_ARTIFACTS if trusted_artifacts is None else trusted_artifacts
        )
        self._loaded: dict[tuple[ERBATask, ERBASubtype], tuple[Any, ERBAArtifactSpec, Mapping[str, Any]]] = {}

    def predict(self, request: ERBARequest) -> ERBAResult:
        try:
            workflow = _as_enum(request.workflow, ERBAWorkflow)
            task = _as_enum(request.task, ERBATask)
            subtype = _as_enum(request.subtype, ERBASubtype)
        except (TypeError, ValueError):
            return self._result(request, ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA,
                                ERBAStatusCode.UNSUPPORTED_COMBINATION, "Unknown ERBA workflow, task, or subtype.")
        if workflow is not ERBAWorkflow.ERBA:
            return self._result(request, task, subtype, ERBAStatusCode.UNSUPPORTED_COMBINATION,
                                "This predictor accepts ERBA workflow requests only.")
        if not combination_is_supported(task, subtype):
            return self._result(request, task, subtype, ERBAStatusCode.UNSUPPORTED_COMBINATION,
                                "This ERBA task/subtype combination is unavailable.")
        prepared = (classification_parent_smiles(request.smiles) if task is ERBATask.CLASSIFICATION
                    else regression_parent_smiles(request.smiles, parity_approved=self._regression_parity_approved))
        if not prepared.valid:
            return self._result(request, task, subtype, prepared.status_code, prepared.status_message,
                                evidence_caveat=(HISTORICAL_EXPOSURE_CAVEAT_COMPACT if task is ERBATask.CLASSIFICATION else ""))
        try:
            artifact, spec, metadata = self._load(task, subtype)
            if task is ERBATask.CLASSIFICATION:
                return self._classify(request, task, subtype, prepared.model_smiles, artifact, spec, metadata)
            return self._regress(request, task, subtype, prepared.model_smiles, artifact, spec, metadata)
        except _ArtifactError as error:
            return self._result(request, task, subtype, error.code, str(error), model_smiles=prepared.model_smiles,
                                evidence_caveat=(HISTORICAL_EXPOSURE_CAVEAT_COMPACT if task is ERBATask.CLASSIFICATION else ""))
        except Exception:
            return self._result(request, task, subtype, ERBAStatusCode.INFERENCE_INTERNAL_ERROR,
                                "ERBA internal inference error.", model_smiles=prepared.model_smiles,
                                evidence_caveat=(HISTORICAL_EXPOSURE_CAVEAT_COMPACT if task is ERBATask.CLASSIFICATION else ""))

    def predict_batch(self, requests: Iterable[ERBARequest]) -> list[ERBAResult]:
        return [self.predict(request) for request in requests]
    def preflight(self, task: ERBATask, subtype: ERBASubtype) -> ERBARoutePreflight:
        """Load and validate a route before a batch writer creates an output workbook."""
        try:
            parsed_task = _as_enum(task, ERBATask)
            parsed_subtype = _as_enum(subtype, ERBASubtype)
            if not combination_is_supported(parsed_task, parsed_subtype):
                raise _ArtifactError(ERBAStatusCode.UNSUPPORTED_COMBINATION, "This ERBA task/subtype combination is unavailable.")
            if parsed_task is ERBATask.IC50_REGRESSION and not self._regression_parity_approved:
                raise _ArtifactError(ERBAStatusCode.REGRESSION_PARITY_NOT_APPROVED,
                                     "Regression preprocessing parity is not approved.")
            self._load(parsed_task, parsed_subtype)
            return ERBARoutePreflight(parsed_task, parsed_subtype, ERBAStatusCode.OK)
        except _ArtifactError as error:
            return ERBARoutePreflight(
                task if isinstance(task, ERBATask) else ERBATask.CLASSIFICATION,
                subtype if isinstance(subtype, ERBASubtype) else ERBASubtype.ER_ALPHA,
                error.code,
                str(error),
            )
        except (TypeError, ValueError):
            return ERBARoutePreflight(
                ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA, ERBAStatusCode.UNSUPPORTED_COMBINATION,
                "Unknown ERBA task or subtype.",
            )

    def _result(self, request: ERBARequest, task: ERBATask, subtype: ERBASubtype,
                code: ERBAStatusCode, message: str, *, model_smiles: str | None = None,
                spec: ERBAArtifactSpec | None = None, policy: str | None = None,
                evidence_caveat: str = "", **values: Any) -> ERBAResult:
        return ERBAResult(row_index=request.row_index, task=task, subtype=subtype,
                          raw_smiles=request.smiles, model_smiles=model_smiles, status_code=code,
                          status_message=message, model_id=spec.model_id if spec else None,
                          model_sha256=spec.sha256 if spec else None, preprocessing_policy_id=policy,
                          evidence_caveat=evidence_caveat, passthrough=request.passthrough, **values)

    def _validate_spec(self, spec: Any) -> ERBAArtifactSpec:
        if (not isinstance(spec, ERBAArtifactSpec) or not isinstance(spec.relative_path, str)
                or not isinstance(spec.size_bytes, int) or isinstance(spec.size_bytes, bool)
                or spec.size_bytes < 0 or not isinstance(spec.model_id, str) or not spec.model_id
                or not isinstance(spec.sha256, str) or _SHA256.fullmatch(spec.sha256.lower()) is None):
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Catalogued artifact specification is invalid.")
        return spec

    def _validate_trust_anchor(
        self,
        task: ERBATask,
        subtype: ERBASubtype,
        spec: ERBAArtifactSpec,
    ) -> Mapping[str, Any]:
        route_id = f"{task.value}:{subtype.value}"
        trusted = self._trusted_artifacts.get(route_id)
        if not isinstance(trusted, Mapping):
            raise _ArtifactError(
                ERBAStatusCode.MODEL_INTEGRITY_FAILED,
                "Artifact is not present in the executable release allowlist.",
            )
        if (
            trusted.get("model_id") != spec.model_id
            or trusted.get("sha256") != spec.sha256.lower()
            or trusted.get("size_bytes") != spec.size_bytes
            or trusted.get("release_status") != "released"
            or not isinstance(trusted.get("release_manifest_sha256"), str)
            or _SHA256.fullmatch(str(trusted["release_manifest_sha256"])) is None
        ):
            raise _ArtifactError(
                ERBAStatusCode.MODEL_INTEGRITY_FAILED,
                "Catalog entry does not match the executable release allowlist.",
            )
        return trusted

    def _safe_path(self, spec: ERBAArtifactSpec) -> Path:

        relative = Path(spec.relative_path)
        if relative.is_absolute() or ".." in relative.parts or relative == Path("."):
            raise _ArtifactError(ERBAStatusCode.MODEL_PATH_UNSAFE, "Artifact path is not a safe relative path.")
        candidate = self._root / relative
        try:
            resolved, root = candidate.resolve(strict=True), self._root.resolve(strict=True)
        except FileNotFoundError as error:
            raise _ArtifactError(ERBAStatusCode.MODEL_NOT_FOUND, "Catalogued artifact is missing.") from error
        try:
            resolved.relative_to(root)
        except ValueError as error:
            raise _ArtifactError(ERBAStatusCode.MODEL_PATH_UNSAFE, "Artifact resolves outside the bundled root.") from error
        current = self._root
        for part in relative.parts:
            current = current / part
            try:
                file_stat = current.lstat()
            except FileNotFoundError as error:
                raise _ArtifactError(ERBAStatusCode.MODEL_NOT_FOUND, "Catalogued artifact is missing.") from error
            attrs = getattr(file_stat, "st_file_attributes", 0)
            if current.is_symlink() or attrs & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0):
                raise _ArtifactError(ERBAStatusCode.MODEL_PATH_UNSAFE, "Artifact path contains a link or reparse point.")
        if not stat.S_ISREG(resolved.stat().st_mode):
            raise _ArtifactError(ERBAStatusCode.MODEL_PATH_UNSAFE, "Artifact is not a regular file.")
        return resolved

    def _load(self, task: ERBATask, subtype: ERBASubtype) -> tuple[Any, ERBAArtifactSpec, Mapping[str, Any]]:
        key = (task, subtype)
        if key in self._loaded:
            return self._loaded[key]
        try:
            spec = self._validate_spec(self._catalog[key])
        except KeyError as error:
            raise _ArtifactError(ERBAStatusCode.MODEL_NOT_FOUND, "No bundled artifact is catalogued for this ERBA route.") from error
        trusted = self._validate_trust_anchor(task, subtype, spec)
        path = self._safe_path(spec)
        try:
            payload = path.read_bytes()
        except OSError as error:
            raise _ArtifactError(ERBAStatusCode.MODEL_NOT_FOUND, "Catalogued artifact cannot be read.") from error
        if len(payload) != spec.size_bytes or hashlib.sha256(payload).hexdigest() != spec.sha256.lower():
            raise _ArtifactError(ERBAStatusCode.MODEL_INTEGRITY_FAILED, "Artifact size or SHA-256 does not match its catalog entry.")
        try:
            import joblib
            artifact = joblib.load(io.BytesIO(payload))
        except Exception as error:
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Artifact cannot be loaded as a valid ERBA joblib.") from error
        metadata, pipeline = self._validate_artifact(artifact, task, subtype, spec, trusted)
        runtime_metadata = dict(metadata)
        runtime_metadata.update({
            name: value for name, value in trusted.items()
            if name.endswith("_sha256") or name == "evidence_scope"
        })
        self._loaded[key] = (pipeline, spec, runtime_metadata)
        return self._loaded[key]

    def _validate_metadata(
        self,
        metadata: Mapping[str, Any],
        task: ERBATask,
        subtype: ERBASubtype,
        spec: ERBAArtifactSpec,
        trusted: Mapping[str, Any],
    ) -> None:
        if metadata.get("task") != task.value or metadata.get("subtype") != subtype.value or metadata.get("model_id") != spec.model_id:
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Artifact metadata does not match its catalog route.")
        policy = CLASSIFICATION_PARENT_POLICY_ID if task is ERBATask.CLASSIFICATION else REGRESSION_PARENT_POLICY_ID
        if metadata.get("preprocessing_policy_id") != policy:
            raise _ArtifactError(ERBAStatusCode.MODEL_POLICY_MISMATCH, "Artifact preprocessing policy is not approved for this route.")
        artifact_state = metadata.get("release_status")
        if artifact_state not in ("released", "frozen_pending_evaluation"):
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Artifact has an invalid frozen release state.")
        if artifact_state == "frozen_pending_evaluation" and trusted.get("release_status") != "released":
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Frozen artifact lacks external release authority.")

    def _validate_artifact(
        self,
        artifact: Any,
        task: ERBATask,
        subtype: ERBASubtype,
        spec: ERBAArtifactSpec,
        trusted: Mapping[str, Any],
    ) -> tuple[Mapping[str, Any], Any]:
        if not isinstance(artifact, Mapping) or artifact.get("schema_version") != ARTIFACT_SCHEMA_VERSION:
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Artifact schema version is invalid.")
        metadata = artifact.get("metadata")
        if not isinstance(metadata, Mapping):
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Artifact metadata is missing.")
        self._validate_metadata(metadata, task, subtype, spec, trusted)
        if task is ERBATask.CLASSIFICATION:
            return metadata, self._validate_classification_artifact(artifact, metadata)
        return metadata, self._validate_regression_artifact(artifact, metadata)

    def _validate_classification_artifact(self, artifact: Mapping[str, Any], metadata: Mapping[str, Any]) -> Any:
        required_hashes = ("source_manifest_sha256", "split_manifest_sha256", "protocol_sha256",
                           "candidate_registry_sha256", "transformer_source_sha256")
        if (metadata.get("pipeline_schema_id") != RAW_SMILES_PIPELINE_SCHEMA_ID
                or metadata.get("evidence_scope") != _CLASSIFICATION_EVIDENCE_SCOPE
                or metadata.get("caveat") != _CLASSIFICATION_CAVEAT
                or any(not _has_sha256(metadata, key) for key in required_hashes)
                or not isinstance(metadata.get("nested_metrics"), Mapping)
                or metadata.get("cached_feature_parity") is not True
                or metadata.get("round_trip_parity") is not True):
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Classification provenance or evidence metadata is invalid.")
        pipeline = artifact.get("pipeline")
        steps = getattr(pipeline, "named_steps", None)
        transformer = steps.get("raw_smiles") if isinstance(steps, Mapping) else None
        if (not isinstance(transformer, ERBARawSmilesFeatures)
                or not callable(getattr(transformer, "transform", None))
                or not isinstance(getattr(transformer, "n_features_out_", None), (int, np.integer))
                or transformer.n_features_out_ <= 0):
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Classification raw-SMILES transformer is incompatible.")
        feature_set = metadata.get("feature_set")
        if not isinstance(feature_set, str) or transformer.feature_set != feature_set:
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Classification feature transformer does not match metadata.")
        if not callable(getattr(pipeline, "predict_proba", None)) or not hasattr(pipeline, "classes_"):
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Classification pipeline is incomplete.")
        if list(pipeline.classes_) != [0, 1] or list(metadata.get("classes", ())) != [0, 1]:
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Classification class order must be [0, 1].")
        threshold = metadata.get("binding_threshold")
        if not isinstance(threshold, (float, int)) or isinstance(threshold, bool) or float(threshold) != 0.5:
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Classification threshold must be the approved 0.5 value.")
        return pipeline

    def _validate_regression_artifact(self, artifact: Mapping[str, Any], metadata: Mapping[str, Any]) -> Mapping[str, Any]:
        if (metadata.get("feature_set") != "avalon_fp_2048" or metadata.get("target") != _REGRESSION_TARGET
                or metadata.get("regression_parity_approved") is not True
                or not isinstance(metadata.get("evidence_scope"), str) or not metadata["evidence_scope"].strip()
                or metadata.get("evidence_scope") == _CLASSIFICATION_EVIDENCE_SCOPE
                or not _has_sha256(metadata, "source_manifest_sha256")
                or not _has_sha256(metadata, "parity_fixture_sha256")):
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Regression parity or provenance metadata is invalid.")
        production = artifact.get("production")
        if not isinstance(production, Mapping):
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Regression production payload is missing.")
        selector, estimator = production.get("selector"), production.get("estimator")
        if not callable(getattr(selector, "transform", None)) or not callable(getattr(estimator, "predict", None)):
            raise _ArtifactError(ERBAStatusCode.MODEL_SCHEMA_INVALID, "Regression selector or estimator is incomplete.")
        return production

    def _classify(self, request: ERBARequest, task: ERBATask, subtype: ERBASubtype, model_smiles: str,
                  pipeline: Any, spec: ERBAArtifactSpec, metadata: Mapping[str, Any]) -> ERBAResult:
        transformer = pipeline.named_steps["raw_smiles"]
        features = np.asarray(transformer.transform([model_smiles]))
        if (features.ndim != 2 or features.shape != (1, transformer.n_features_out_)
                or not np.issubdtype(features.dtype, np.number) or not np.all(np.isfinite(features))
                or not np.any(features)):
            raise _ArtifactError(ERBAStatusCode.PREDICTION_FAILED,
                                 "Classification feature generation produced an invalid or all-zero vector.")
        probabilities = np.asarray(pipeline.predict_proba([model_smiles]), dtype=float)
        if (probabilities.shape != (1, 2) or not np.all(np.isfinite(probabilities))
                or np.any(probabilities < 0.0) or np.any(probabilities > 1.0)
                or not np.isclose(probabilities[0].sum(), 1.0, rtol=0.0, atol=1e-12)):
            raise _ArtifactError(ERBAStatusCode.PREDICTION_FAILED, "Classification pipeline returned invalid probabilities.")
        non_binding, binding = map(float, probabilities[0])
        return self._result(
            request, task, subtype, ERBAStatusCode.OK, "", model_smiles=model_smiles,
            spec=spec, policy=CLASSIFICATION_PARENT_POLICY_ID, evidence_caveat=metadata["caveat"],
            non_binding_probability=non_binding, binding_probability=binding,
            binding_label="binding" if binding >= 0.5 else "non_binding",
            protocol_sha256=metadata.get("protocol_sha256"),
            nested_cv_sha256=metadata.get("nested_cv_sha256"),
            internal_resplit_sha256=metadata.get("internal_resplit_sha256"),
            report_sha256=metadata.get("report_sha256"),
            source_manifest_sha256=metadata.get("source_manifest_sha256"),
            evidence_scope=metadata.get("evidence_scope"),
        )

    def _regress(self, request: ERBARequest, task: ERBATask, subtype: ERBASubtype, model_smiles: str,
                 production: Mapping[str, Any], spec: ERBAArtifactSpec, metadata: Mapping[str, Any]) -> ERBAResult:
        features = avalon_fp_2048(model_smiles).reshape(1, 2048)
        if features.dtype != np.uint8 or not np.any(features):
            raise _ArtifactError(ERBAStatusCode.PREDICTION_FAILED, "Regression feature generation failed.")
        selected = np.asarray(production["selector"].transform(features))
        if selected.ndim != 2 or selected.shape[0] != 1 or selected.shape[1] == 0 or not np.all(np.isfinite(selected)):
            raise _ArtifactError(ERBAStatusCode.PREDICTION_FAILED, "Regression feature selection failed.")
        prediction = np.asarray(production["estimator"].predict(selected), dtype=float).reshape(-1)
        if prediction.size != 1 or not np.isfinite(prediction[0]):
            raise _ArtifactError(ERBAStatusCode.PREDICTION_FAILED, "Regression estimator returned an invalid pIC50.")
        pic50 = float(prediction[0])
        return self._result(
            request, task, subtype, ERBAStatusCode.OK, "", model_smiles=model_smiles,
            spec=spec, policy=REGRESSION_PARENT_POLICY_ID,
            evidence_caveat=f"Regression evidence scope: {metadata['evidence_scope']}",
            pic50=pic50, ic50_nm=float(10 ** (9 - pic50)),
            report_sha256=metadata.get("report_sha256"),
            source_manifest_sha256=metadata.get("source_manifest_sha256"),
            preprocessing_parity_sha256=metadata.get(
                "preprocessing_parity_sha256", metadata.get("parity_fixture_sha256")
            ),
            evidence_scope=metadata.get("evidence_scope"),
        )
