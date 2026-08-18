"""One-time V7 evaluator for the historically exposed internal sensitivity split.

This module imports neither trainer nor refit code. It evaluates exact frozen model
bytes, verifies that those bytes remain unchanged, and issues external release
manifests without rewriting the joblib artifacts.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import subprocess
import sys
import time
sys.path.insert(0, str(Path(__file__).resolve().parent))

import joblib
import numpy as np

from erba_classification_pipeline import CAVEAT, OUT, dump, metric, rows, sha, write_rows

SUBTYPES = ("ERalpha", "ERbeta")


def _read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def _clean_process_check(project_root: Path, artifact_path: Path) -> dict:
    started = time.perf_counter()
    code = (
        "import sys; "
        f"sys.path.insert(0, {str(project_root)!r}); "
        "import joblib; "
        f"a=joblib.load({str(artifact_path)!r}); "
        "p=a['pipeline'].predict_proba(['CCO']); "
        "assert p.shape == (1, 2); "
        "print('clean-load-predict-ok')"
    )
    completed = subprocess.run(
        [sys.executable, "-I", "-c", code],
        cwd=str(project_root),
        capture_output=True,
        text=True,
        timeout=180,
        check=False,
    )
    return {
        "passed": completed.returncode == 0 and "clean-load-predict-ok" in completed.stdout,
        "returncode": completed.returncode,
        "stdout": completed.stdout.strip(),
        "stderr": completed.stderr.strip(),
        "command": [sys.executable, "-I", "-c", "<isolated load/predict probe>"],
        "runtime_seconds": time.perf_counter() - started,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=OUT)
    args = parser.parse_args()
    out = args.output_dir.resolve()
    total_started = time.perf_counter()
    project_root = Path(__file__).resolve().parent.parent
    receipt_path = out / "internal_resplit_sensitivity_evaluation.json"

    prohibited_existing = [
        receipt_path,
        *(out / f"internal_resplit_predictions_{subtype}.csv" for subtype in SUBTYPES),
        *(out / f"release_manifest_{subtype}.json" for subtype in SUBTYPES),
    ]
    existing = [str(path) for path in prohibited_existing if path.exists()]
    if existing:
        raise RuntimeError(
            "one-time evaluator outputs already exist; refusing repeat evaluation: "
            + ", ".join(existing)
        )

    split_path = out / "split_manifest.json"
    split = _read_json(split_path)
    training = _read_json(out / "classification_training_summary.json")
    expected_hashes = {
        "source_manifest_sha256": sha(out / "source_manifest.json"),
        "historical_exposure_manifest_sha256": sha(out / "historical_exposure_manifest.json"),
        "split_manifest_sha256": sha(split_path),
        "protocol_sha256": sha(out / "protocol.json"),
        "candidate_registry_sha256": sha(out / "candidate_registry.json"),
        "nested_cv_sha256": sha(out / "nested_cv_summary.json"),
    }
    report = {
        "schema_version": 1,
        "protocol_version": "classification_modeling_nested_mcc_v7",
        "caveat": CAVEAT,
        "evidence_scope": "internal_historically_exposed",
        "one_time_evaluation": True,
        "subtypes": {},
    }

    release_inputs: dict[str, dict] = {}
    for subtype in SUBTYPES:
        subtype_started = time.perf_counter()
        app_subtype = "er_alpha" if subtype == "ERalpha" else "er_beta"
        holdout_path = out / f"internal_resplit_{subtype}.csv"
        development_path = out / f"development_{subtype}.csv"
        partition = split["partitions"][subtype]
        if sha(holdout_path) != partition["internal_resplit"]["sha256"]:
            raise RuntimeError(f"{subtype} exact sealed internal-resplit hash mismatch")
        if sha(development_path) != partition["development"]["sha256"]:
            raise RuntimeError(f"{subtype} exact sealed development hash mismatch")

        holdout = rows(holdout_path)
        development = rows(development_path)
        overlap_counts = {}
        for key in ("row_id", "inchikey", "group"):
            overlap_counts[key] = len(
                {item[key] for item in holdout} & {item[key] for item in development}
            )
        if any(overlap_counts.values()):
            raise RuntimeError(f"{subtype} development/internal-resplit overlap: {overlap_counts}")
        if {int(item["label"]) for item in holdout} != {0, 1}:
            raise RuntimeError(f"{subtype} internal-resplit partition lacks a class")
        if {int(item["label"]) for item in development} != {0, 1}:
            raise RuntimeError(f"{subtype} development partition lacks a class")

        artifact_path = out / f"production_classification_model_{subtype}.joblib"
        frozen_manifest_path = out / f"production_classification_model_{subtype}_manifest.json"
        frozen = _read_json(frozen_manifest_path)
        model_sha_before = sha(artifact_path)
        model_size_before = artifact_path.stat().st_size
        if (
            frozen.get("release_status") != "frozen_pending_evaluation"
            or frozen.get("model_sha256") != model_sha_before
            or frozen.get("model_size_bytes") != model_size_before
        ):
            raise RuntimeError(f"{subtype} frozen artifact binding mismatch")

        artifact = joblib.load(artifact_path)
        metadata = artifact.get("metadata", {}) if isinstance(artifact, dict) else {}
        required = {
            "task": "classification",
            "subtype": app_subtype,
            "model_id": f"erba_classification_v7_{app_subtype}",
            "preprocessing_policy_id": "erba_binding_classification_parent_v2",
            "pipeline_schema_id": "raw_smiles_pipeline_schema_v1",
            "classes": [0, 1],
            "binding_threshold": 0.5,
            "release_status": "frozen_pending_evaluation",
            "evidence_scope": "internal_historically_exposed",
            "caveat": CAVEAT,
        }
        if artifact.get("schema_version") != 1 or any(
            metadata.get(key) != value for key, value in required.items()
        ):
            raise RuntimeError(f"{subtype} artifact metadata contract mismatch")
        for key, expected in expected_hashes.items():
            if metadata.get(key) != expected:
                raise RuntimeError(f"{subtype} artifact provenance mismatch for {key}")
        if metadata.get("internal_resplit_sha256") != partition["internal_resplit"]["sha256"]:
            raise RuntimeError(f"{subtype} artifact is not bound to the sealed sensitivity split")
        if not metadata.get("cached_feature_parity") or not metadata.get("round_trip_parity"):
            raise RuntimeError(f"{subtype} artifact parity gates are not clean")

        y_true = np.asarray([int(item["label"]) for item in holdout])
        probability = artifact["pipeline"].predict_proba(
            np.asarray([item["model_smiles"] for item in holdout])
        )[:, 1]
        metrics = metric(y_true, probability)
        prediction_rows = [
            {
                "subtype": subtype,
                "row_id": item["row_id"],
                "inchikey": item["inchikey"],
                "group": item["group"],
                "label": int(item["label"]),
                "probability_positive": float(score),
                "prediction": int(score >= 0.5),
            }
            for item, score in zip(holdout, probability)
        ]
        prediction_path = out / f"internal_resplit_predictions_{subtype}.csv"
        write_rows(prediction_path, prediction_rows)

        model_sha_after = sha(artifact_path)
        model_size_after = artifact_path.stat().st_size
        if model_sha_before != model_sha_after or model_size_before != model_size_after:
            raise RuntimeError(f"{subtype} artifact bytes changed across evaluation")
        clean_process = _clean_process_check(project_root, artifact_path)
        if not clean_process["passed"]:
            raise RuntimeError(
                f"{subtype} isolated load/predict failed: {clean_process['stderr']}"
            )

        subtype_report = {
            "metrics": metrics,
            "calibration": {
                "mean_predicted_probability": float(np.mean(probability)),
                "observed_prevalence": float(np.mean(y_true)),
                "brier_score": metrics["brier_score"],
            },
            "partition_overlap_counts": overlap_counts,
            "prediction_sha256": sha(prediction_path),
            "model_sha256_pre_evaluator": model_sha_before,
            "model_sha256_post_evaluator": model_sha_after,
            "model_size_bytes_pre_evaluator": model_size_before,
            "model_size_bytes_post_evaluator": model_size_after,
            "clean_pinned_process": clean_process,
            "runtime_seconds": time.perf_counter() - subtype_started,
            "integrity": (
                "passed: evaluator loaded frozen bytes only; no refit, rerank, "
                "preprocessing change, threshold change, or reserialization"
            ),
        }
        report["subtypes"][subtype] = subtype_report
        release_inputs[subtype] = {
            "app_subtype": app_subtype,
            "metadata": metadata,
            "artifact_path": artifact_path,
            "frozen_manifest_path": frozen_manifest_path,
            "subtype_report": subtype_report,
        }
        report["runtime_seconds"] = time.perf_counter() - total_started

    dump(receipt_path, report)
    receipt_sha256 = sha(receipt_path)

    for subtype, release_input in release_inputs.items():
        metadata = release_input["metadata"]
        subtype_report = release_input["subtype_report"]
        gates = dict(training["release_gates"])
        gates.update(
            {
                "sealed_artifact_hash_equality": (
                    subtype_report["model_sha256_pre_evaluator"]
                    == subtype_report["model_sha256_post_evaluator"]
                ),
                "one_time_evaluator_integrity": True,
                "clean_pinned_process_import": subtype_report["clean_pinned_process"]["passed"],
            }
        )
        release_eligible = bool(gates) and all(value is True for value in gates.values())
        if not release_eligible:
            raise RuntimeError(f"{subtype} release gates are not all clean: {gates}")
        release_manifest = {
            "schema_version": 1,
            "task": "classification",
            "subtype": release_input["app_subtype"],
            "model_id": metadata["model_id"],
            "artifact_filename": release_input["artifact_path"].name,
            "model_sha256": subtype_report["model_sha256_pre_evaluator"],
            "model_size_bytes": subtype_report["model_size_bytes_pre_evaluator"],
            "release_status": "released",
            "release_authority": "evaluator_bound_release_manifest",
            "frozen_manifest_sha256": sha(release_input["frozen_manifest_path"]),
            "source_manifest_sha256": metadata["source_manifest_sha256"],
            "historical_exposure_manifest_sha256": metadata["historical_exposure_manifest_sha256"],
            "split_manifest_sha256": metadata["split_manifest_sha256"],
            "protocol_sha256": metadata["protocol_sha256"],
            "candidate_registry_sha256": metadata["candidate_registry_sha256"],
            "transformer_source_sha256": metadata["transformer_source_sha256"],
            "nested_cv_sha256": metadata["nested_cv_sha256"],
            "internal_resplit_sha256": receipt_sha256,
            "internal_resplit_partition_sha256": metadata["internal_resplit_sha256"],
            "gates": gates,
            "release_eligible": True,
            "evidence_scope": "internal_historically_exposed",
            "caveat": CAVEAT,
        }
        dump(out / f"release_manifest_{subtype}.json", release_manifest)


if __name__ == "__main__":
    main()
