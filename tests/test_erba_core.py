import hashlib
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import joblib
import numpy as np
import pandas as pd

from core.contracts import (
    ARTIFACT_SCHEMA_VERSION, CLASSIFICATION_PARENT_POLICY_ID, ERBA_FATAL_ARTIFACT_STATUS_CODES,
    HISTORICAL_EXPOSURE_CAVEAT_FULL, RAW_SMILES_PIPELINE_SCHEMA_ID, REGRESSION_PARENT_POLICY_ID,
    ERBAArtifactSpec, ERBARoutePreflight, ERBARequest, ERBAResult, ERBAStatusCode, ERBASubtype, ERBATask,
)
from core.erba_features import ERBARawSmilesFeatures
from core.erba_predictor import ERBAPredictor
from core.erba_preprocessing import classification_parent_smiles
from gui.erba_tab import cas_lookup_smiles, export_erba_batch


class RawSmilesPipeline:
    classes_ = np.array([0, 1])

    def __init__(self, probabilities=(0.8, 0.2), transformer=None):
        self.named_steps = {"raw_smiles": transformer or ERBARawSmilesFeatures("rdkit_fp_2048").fit(["CC"])}
        self.probabilities = np.asarray([probabilities], dtype=float)

    def predict_proba(self, smiles):
        assert smiles == ["CCO"]
        return self.probabilities


class ReversedClassesPipeline(RawSmilesPipeline):
    classes_ = np.array([1, 0])


class ZeroTransformer(ERBARawSmilesFeatures):
    def fit(self, X, y=None):
        self.n_features_out_ = 4
        return self

    def transform(self, X):
        return np.zeros((len(X), 4), dtype=np.float32)


class Selector:
    def transform(self, values):
        assert values.shape == (1, 2048)
        return values[:, :2]


class Estimator:
    def predict(self, values):
        return np.array([6.0])
class ExplodingTransformer(ERBARawSmilesFeatures):
    def transform(self, X):
        raise RuntimeError("transformer defect")


class _BatchPredictor:
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
            non_binding_probability=0.2, binding_probability=0.8, binding_label="binding",
        )


class _BatchAD:
    def __init__(self):
        self.calls = []
        self.calculator = SimpleNamespace(fitted=True)

    def evaluate(self, task, subtype, model_smiles):
        self.calls.append((task, subtype, model_smiles))
        return (
            self.calculator,
            SimpleNamespace(
                fitted=True,
                in_domain=True,
                distance=1.25,
                threshold=2.5,
                distance_in_domain=True,
                max_similarity=0.75,
                similarity_threshold=0.6,
                similarity_in_domain=True,
                pc1=0.5,
                pc2=-0.25,
            ),
            np.asarray([[1.0, 0.0, 1.0]], dtype=float),
        )



def _hashes():
    return {key: "a" * 64 for key in (
        "source_manifest_sha256", "split_manifest_sha256", "protocol_sha256",
        "candidate_registry_sha256", "transformer_source_sha256",
    )}


def artifact_metadata(task, subtype, model_id, policy):
    metadata = {
        "task": task.value, "subtype": subtype.value, "model_id": model_id,
        "preprocessing_policy_id": policy, "release_status": "released",
    }
    if task is ERBATask.CLASSIFICATION:
        metadata.update({
            "pipeline_schema_id": RAW_SMILES_PIPELINE_SCHEMA_ID, "classes": [0, 1],
            "binding_threshold": 0.5, "feature_set": "rdkit_fp_2048",
            "evidence_scope": "internal_historically_exposed", "caveat": HISTORICAL_EXPOSURE_CAVEAT_FULL,
            "nested_metrics": {}, "cached_feature_parity": True, "round_trip_parity": True, **_hashes(),
        })
    else:
        metadata.update({
            "feature_set": "avalon_fp_2048", "target": "pIC50 = -log10(IC50 [mol/L])",
            "regression_parity_approved": True, "evidence_scope": "validated_internal_holdout",
            "source_manifest_sha256": "a" * 64, "parity_fixture_sha256": "b" * 64,
        })
    return metadata


class ERBACoreTests(unittest.TestCase):
    def _write(self, root, name, payload):
        path = Path(root, name)
        joblib.dump(payload, path)
        return ERBAArtifactSpec(name, path.stat().st_size, hashlib.sha256(path.read_bytes()).hexdigest(), "fixture")

    def _predictor(self, root, task, subtype, spec, **kwargs):
        trusted = {
            f"{task.value}:{subtype.value}": {
                "model_id": spec.model_id,
                "sha256": spec.sha256,
                "size_bytes": spec.size_bytes,
                "release_status": "released",
                "release_manifest_sha256": "0" * 64,
            }
        }
        return ERBAPredictor(
            root,
            {(task, subtype): spec},
            trusted_artifacts=trusted,
            **kwargs,
        )

    def _classification_payload(self, pipeline=None, metadata=None):
        task, subtype = ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA
        return {
            "schema_version": ARTIFACT_SCHEMA_VERSION,
            "metadata": metadata or artifact_metadata(task, subtype, "fixture", CLASSIFICATION_PARENT_POLICY_ID),
            "pipeline": pipeline or RawSmilesPipeline(),
        }

    def test_preprocessing_exact_policy_edges(self):
        self.assertEqual(classification_parent_smiles(" CC(=O)O.[Na+] ").model_smiles, "CC(=O)O")
        for smiles, code in [
            ("", ERBAStatusCode.BLANK_SMILES), ("C[*]", ERBAStatusCode.WILDCARD_SMILES),
            ("C[1*]", ERBAStatusCode.WILDCARD_SMILES), ("CC |foo|", ERBAStatusCode.CXSMILES_NOT_ALLOWED),
            ("not smiles", ERBAStatusCode.INVALID_SMILES), ("[Na+]", ERBAStatusCode.NO_CARBON),
            ("CC[Fe]", ERBAStatusCode.METAL_RETAINED),
        ]:
            self.assertEqual(classification_parent_smiles(smiles).status_code, code)

    def test_classification_preflight_and_output_contract(self):
        with tempfile.TemporaryDirectory() as root:
            task, subtype = ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA
            spec = self._write(root, "classification.joblib", self._classification_payload())
            predictor = self._predictor(root, task, subtype, spec)
            self.assertTrue(predictor.preflight(task, subtype).ready)
            result = predictor.predict(ERBARequest("CCO", task, subtype, 7))
            self.assertEqual(result.status_code, ERBAStatusCode.OK)
            self.assertEqual((result.non_binding_probability, result.binding_probability, result.binding_label), (0.8, 0.2, "non_binding"))
            self.assertEqual(result.evidence_caveat, HISTORICAL_EXPOSURE_CAVEAT_FULL)

    def test_adjacent_catalog_cannot_authorize_untrusted_joblib(self):
        with tempfile.TemporaryDirectory() as root:
            task, subtype = ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA
            spec = self._write(root, "untrusted.joblib", self._classification_payload())
            result = ERBAPredictor(
                root,
                {(task, subtype): spec},
                trusted_artifacts={},
            ).preflight(task, subtype)
            self.assertEqual(result.status_code, ERBAStatusCode.MODEL_INTEGRITY_FAILED)

    def test_integrity_path_metadata_and_pipeline_fail_closed(self):
        with tempfile.TemporaryDirectory() as root:
            task, subtype = ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA
            bad = self._classification_payload(pipeline=ReversedClassesPipeline())
            spec = self._write(root, "bad.joblib", bad)
            self.assertEqual(self._predictor(root, task, subtype, spec).preflight(task, subtype).status_code,
                             ERBAStatusCode.MODEL_SCHEMA_INVALID)
            traversal = ERBAArtifactSpec("../bad.joblib", spec.size_bytes, spec.sha256, "fixture")
            self.assertEqual(self._predictor(root, task, subtype, traversal).preflight(task, subtype).status_code,
                             ERBAStatusCode.MODEL_PATH_UNSAFE)
            corrupt = ERBAArtifactSpec("bad.joblib", spec.size_bytes + 1, spec.sha256, "fixture")
            self.assertEqual(self._predictor(root, task, subtype, corrupt).preflight(task, subtype).status_code,
                             ERBAStatusCode.MODEL_INTEGRITY_FAILED)
            metadata = artifact_metadata(task, subtype, "fixture", CLASSIFICATION_PARENT_POLICY_ID)
            metadata["caveat"] = "short caveat"
            malformed = self._write(root, "malformed.joblib", self._classification_payload(metadata=metadata))
            self.assertEqual(self._predictor(root, task, subtype, malformed).preflight(task, subtype).status_code,
                             ERBAStatusCode.MODEL_SCHEMA_INVALID)

    def test_all_zero_features_and_invalid_output_shapes_fail_closed(self):
        with tempfile.TemporaryDirectory() as root:
            task, subtype = ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA
            zero = RawSmilesPipeline(transformer=ZeroTransformer("rdkit_fp_2048").fit(["CC"]))
            zero_spec = self._write(root, "zero.joblib", self._classification_payload(pipeline=zero))
            self.assertEqual(self._predictor(root, task, subtype, zero_spec).predict(ERBARequest("CCO", task, subtype)).status_code,
                             ERBAStatusCode.PREDICTION_FAILED)
            wrong_shape = self._write(root, "shape.joblib", self._classification_payload(pipeline=RawSmilesPipeline((1.0, 0.0, 0.0))))
            self.assertEqual(self._predictor(root, task, subtype, wrong_shape).predict(ERBARequest("CCO", task, subtype)).status_code,
                             ERBAStatusCode.PREDICTION_FAILED)
    def test_unexpected_transformer_failure_is_fatal_internal_error(self):
        with tempfile.TemporaryDirectory() as root:
            task, subtype = ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA
            exploding = RawSmilesPipeline(
                transformer=ExplodingTransformer("rdkit_fp_2048").fit(["CC"])
            )
            spec = self._write(root, "exploding.joblib", self._classification_payload(pipeline=exploding))
            predictor = self._predictor(root, task, subtype, spec)
            results = predictor.predict_batch([
                ERBARequest("", task, subtype),
                ERBARequest("CCO", task, subtype),
            ])
            self.assertEqual(results[0].status_code, ERBAStatusCode.BLANK_SMILES)
            self.assertEqual(results[1].status_code, ERBAStatusCode.INFERENCE_INTERNAL_ERROR)
            self.assertIn(results[1].status_code, ERBA_FATAL_ARTIFACT_STATUS_CODES)

    def test_cas_only_batch_uses_pubchem_mapping_and_internal_error_blocks_export(self):
        task, subtype = ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA
        spec = ERBAArtifactSpec("model.joblib", 1, "a" * 64, "fixture")
        provenance = {"metadata": {key: "a" * 64 for key in (
            "protocol_sha256", "source_manifest_sha256", "historical_exposure_manifest_sha256",
            "split_manifest_sha256", "nested_cv_sha256", "internal_resplit_sha256",
            "report_sha256", "caveat_sha256",
        )}}
        with tempfile.TemporaryDirectory() as directory:
            input_path = Path(directory, "input.xlsx")
            pd.DataFrame([["lookup", "50-00-0", ""]], columns=["Row_ID", "CAS", "SMILES"]).to_excel(
                input_path, index=False
            )
            predictor = _BatchPredictor()
            erba_ad = _BatchAD()
            with patch("gui.erba_tab.cas_to_smiles", return_value={
                "CAS": "50-00-0", "PubChem_CID": 712, "CanonicalSMILES": "C=O", "IsomericSMILES": "C=O",
            }) as lookup, patch(
                "gui.erba_tab.save_erba_batch_graphs", return_value=((), None)
            ):
                export_result = export_erba_batch(
                    input_path, task, subtype, predictor, erba_ad, spec, provenance
                )
            self.assertTrue(export_result.destination.exists())
            self.assertEqual(export_result.destination.parent, input_path.resolve().parent)
            self.assertEqual(export_result.count, 1)
            self.assertEqual(export_result.binding_count, 1)
            self.assertEqual(export_result.ad_in_domain_count, 1)
            self.assertEqual([request.smiles for request in predictor.calls], ["C=O"])
            self.assertEqual(erba_ad.calls, [(task, subtype, "C=O")])
            lookup.assert_called_once_with("50-00-0")
        with patch("gui.erba_tab.cas_to_smiles", return_value={"IsomericSMILES": "C[C@H](O)F"}):
            self.assertEqual(cas_lookup_smiles("75-05-8"), "C[C@H](O)F")
        with patch("gui.erba_tab.cas_to_smiles", return_value={"PubChem_CID": 1}):
            with self.assertRaisesRegex(ValueError, "neither CanonicalSMILES nor IsomericSMILES"):
                cas_lookup_smiles("50-00-0")
        with tempfile.TemporaryDirectory() as directory:
            input_path = Path(directory, "fatal.xlsx")
            pd.DataFrame([["x", "CCO"]], columns=["Row_ID", "SMILES"]).to_excel(input_path, index=False)
            with self.assertRaisesRegex(RuntimeError, "no output"):
                export_erba_batch(
                    input_path, task, subtype,
                    _BatchPredictor(ERBAStatusCode.INFERENCE_INTERNAL_ERROR),
                    _BatchAD(), spec, provenance,
                )
            self.assertEqual(list(Path(directory).glob("ERBA_*.xlsx")), [])

    def test_carsrn_alias_resolves_all_25_rows_from_the_supplied_example(self):
        cas_values = [
            "6422-86-2", "64-17-5", "50-00-0", "67-66-3", "107-13-1",
            "5129-00-0", "24038-68-4", "1571-75-1", "94-18-8", "5397-34-2",
            "41481-66-7", "97042-18-7", "63134-33-8", "95235-30-6",
            "191680-83-8", "93589-69-6", "232938-43-1", "151882-81-4",
            "321860-75-7", "1763-23-1", "29420-49-3", "3871-99-6",
            "2923-26-4", "2043-47-2", "647-42-7",
        ]
        task, subtype = ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA
        spec = ERBAArtifactSpec("model.joblib", 1, "a" * 64, "fixture")
        provenance = {"metadata": {key: "a" * 64 for key in (
            "protocol_sha256", "source_manifest_sha256",
            "historical_exposure_manifest_sha256", "split_manifest_sha256",
            "nested_cv_sha256", "internal_resplit_sha256", "report_sha256",
            "caveat_sha256",
        )}}
        predictor = _BatchPredictor()
        erba_ad = _BatchAD()
        with tempfile.TemporaryDirectory() as directory:
            input_path = Path(directory, "test.xlsx")
            pd.DataFrame({"CARSRN": cas_values}).to_excel(input_path, index=False)
            with patch(
                "gui.erba_tab.cas_to_smiles",
                return_value={"CanonicalSMILES": "CCO", "PubChem_CID": 702},
            ) as lookup, patch(
                "gui.erba_tab.save_erba_batch_graphs",
                return_value=((), None),
            ):
                export_result = export_erba_batch(
                    input_path,
                    task,
                    subtype,
                    predictor,
                    erba_ad,
                    spec,
                    provenance,
                )
            predictions = pd.read_excel(
                export_result.destination,
                sheet_name="Predictions",
                dtype=str,
                keep_default_na=False,
            )

        self.assertEqual(export_result.count, 25)
        self.assertEqual(
            [entry.args[0] for entry in lookup.call_args_list],
            cas_values,
        )
        self.assertEqual([request.smiles for request in predictor.calls], ["CCO"] * 25)
        self.assertEqual(predictions["CARSRN"].tolist(), cas_values)
        self.assertEqual(predictions["CAS"].tolist(), cas_values)
        self.assertEqual(predictions["Result_Status"].tolist(), ["Predicted"] * 25)

    def test_regression_remains_separate_and_parity_gated(self):
        with tempfile.TemporaryDirectory() as root:
            task, subtype = ERBATask.IC50_REGRESSION, ERBASubtype.ER_ALPHA
            payload = {
                "schema_version": ARTIFACT_SCHEMA_VERSION,
                "metadata": artifact_metadata(task, subtype, "fixture", REGRESSION_PARENT_POLICY_ID),
                "production": {"selector": Selector(), "estimator": Estimator()},
            }
            spec = self._write(root, "regression.joblib", payload)
            request = ERBARequest("CCO", task, subtype)
            self.assertEqual(self._predictor(root, task, subtype, spec).preflight(task, subtype).status_code,
                             ERBAStatusCode.REGRESSION_PARITY_NOT_APPROVED)
            result = self._predictor(root, task, subtype, spec, regression_parity_approved=True).predict(request)
            self.assertEqual((result.status_code, result.pic50, result.ic50_nm), (ERBAStatusCode.OK, 6.0, 1000.0))
            self.assertNotIn("historically exposed", result.evidence_caveat)
            beta = self._predictor(root, task, subtype, spec, regression_parity_approved=True).preflight(
                task, ERBASubtype.ER_BETA,
            )
            self.assertEqual(beta.status_code, ERBAStatusCode.UNSUPPORTED_COMBINATION)


if __name__ == "__main__":
    unittest.main()
