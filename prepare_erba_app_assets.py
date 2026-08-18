"""Assemble exactly released ERBA assets, catalog, and executable trust allowlist."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

import joblib
import numpy as np

from core.contracts import (
    ERBAArtifactSpec,
    ERBARequest,
    ERBASubtype,
    ERBATask,
    HISTORICAL_EXPOSURE_CAVEAT_FULL,
)
from core.erba_predictor import ERBAPredictor

DATA_ROOT = Path(r"D:/research/FDA_endocrine_disruption/data/normalized_erba")
DEFAULT_CLASSIFICATION_DIR = DATA_ROOT / "classification_modeling_nested_mcc_v7"
MODEL_DIR = ROOT / "models" / "erba"
ALLOWLIST_PATH = ROOT / "core" / "erba_release_allowlist.py"
REPORT = DATA_ROOT / "er_regression_classification_model_report_final.html"
REGRESSION_NAME = "production_regression_model_ERalpha.joblib"
REGRESSION_MANIFEST_NAME = "production_regression_model_ERalpha_manifest.json"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def copy_exact(source: Path, target: Path) -> None:
    shutil.copyfile(source, target)
    if source.stat().st_size != target.stat().st_size or sha256(source) != sha256(target):
        raise RuntimeError(f"byte-exact copy failed: {source} -> {target}")


def classification_route(source_dir: Path, subtype: str) -> tuple[dict, dict, dict]:
    source_model = source_dir / f"production_classification_model_{subtype}.joblib"
    frozen_manifest = source_dir / f"production_classification_model_{subtype}_manifest.json"
    release_manifest = source_dir / f"release_manifest_{subtype}.json"
    frozen = read_json(frozen_manifest)
    release = read_json(release_manifest)
    app_subtype = "er_alpha" if subtype == "ERalpha" else "er_beta"
    if (
        release.get("schema_version") != 1
        or release.get("task") != "classification"
        or release.get("subtype") != app_subtype
        or release.get("release_status") != "released"
        or release.get("release_authority") != "evaluator_bound_release_manifest"
        or release.get("release_eligible") is not True
    ):
        raise RuntimeError(f"{subtype} external release authority is invalid")
    if (
        frozen.get("release_status") != "frozen_pending_evaluation"
        or release.get("model_sha256") != sha256(source_model)
        or release.get("model_size_bytes") != source_model.stat().st_size
        or release.get("frozen_manifest_sha256") != sha256(frozen_manifest)
    ):
        raise RuntimeError(f"{subtype} frozen model/release binding failed")

    artifact = joblib.load(source_model)
    metadata = artifact.get("metadata", {}) if isinstance(artifact, dict) else {}
    required = {
        "task": "classification",
        "subtype": app_subtype,
        "model_id": release["model_id"],
        "preprocessing_policy_id": "erba_binding_classification_parent_v2",
        "pipeline_schema_id": "raw_smiles_pipeline_schema_v1",
        "classes": [0, 1],
        "binding_threshold": 0.5,
        "release_status": "frozen_pending_evaluation",
        "evidence_scope": "internal_historically_exposed",
        "caveat": HISTORICAL_EXPOSURE_CAVEAT_FULL,
    }
    if artifact.get("schema_version") != 1 or any(
        metadata.get(key) != value for key, value in required.items()
    ):
        raise RuntimeError(f"{subtype} artifact metadata is not app-compatible")

    model_name = source_model.name
    release_name = release_manifest.name
    target_model = MODEL_DIR / model_name
    target_release = MODEL_DIR / release_name
    copy_exact(source_model, target_model)
    copy_exact(release_manifest, target_release)
    report_hash = sha256(REPORT)
    caveat_hash = hashlib.sha256(HISTORICAL_EXPOSURE_CAVEAT_FULL.encode("utf-8")).hexdigest()
    route_metadata = {
        "protocol_sha256": release["protocol_sha256"],
        "source_manifest_sha256": release["source_manifest_sha256"],
        "historical_exposure_manifest_sha256": release["historical_exposure_manifest_sha256"],
        "split_manifest_sha256": release["split_manifest_sha256"],
        "nested_cv_sha256": release["nested_cv_sha256"],
        "internal_resplit_sha256": release["internal_resplit_sha256"],
        "report_sha256": report_hash,
        "caveat_sha256": caveat_hash,
        "candidate_registry_sha256": release["candidate_registry_sha256"],
        "transformer_source_sha256": release["transformer_source_sha256"],
        "release_manifest_sha256": sha256(target_release),
    }
    route = {
        "task": "classification",
        "subtype": app_subtype,
        "release_status": "released",
        "artifact": {
            "relative_path": model_name,
            "size_bytes": target_model.stat().st_size,
            "sha256": sha256(target_model),
            "model_id": metadata["model_id"],
        },
        "evidence_scope": metadata["evidence_scope"],
        "feature_set": metadata["feature_set"],
        "selected_feature_count": metadata["selected_feature_count"],
        "nested_metrics": metadata["nested_metrics"],
    }
    trusted = {
        "model_id": metadata["model_id"],
        "sha256": sha256(target_model),
        "size_bytes": target_model.stat().st_size,
        "release_status": "released",
        "release_manifest_relative_path": release_name,
        "release_manifest_sha256": sha256(target_release),
        "evidence_scope": metadata["evidence_scope"],
        **route_metadata,
    }
    return route, trusted, route_metadata


def regression_route() -> tuple[dict, dict, dict]:
    target = MODEL_DIR / REGRESSION_NAME
    manifest_path = MODEL_DIR / REGRESSION_MANIFEST_NAME
    if not target.is_file() or not manifest_path.is_file():
        raise RuntimeError("Pinned-runtime ERalpha regression artifact and manifest are required")
    manifest = read_json(manifest_path)
    if (
        manifest.get("release_status") != "released"
        or manifest.get("sha256") != sha256(target)
        or manifest.get("size_bytes") != target.stat().st_size
        or manifest.get("regression_parity_approved") is not True
    ):
        raise RuntimeError("ERalpha regression artifact integrity/parity gate failed")
    artifact = joblib.load(target)
    metadata = artifact.get("metadata", {}) if isinstance(artifact, dict) else {}
    required = {
        "task": "ic50_regression",
        "subtype": "er_alpha",
        "model_id": manifest["model_id"],
        "preprocessing_policy_id": "erba_ic50_regression_parent_parity_v1",
    }
    if artifact.get("schema_version") != 1 or any(
        metadata.get(key) != value for key, value in required.items()
    ):
        raise RuntimeError("ERalpha regression metadata is invalid")
    route_metadata = {
        "source_manifest_sha256": manifest["source_manifest_sha256"],
        "preprocessing_parity_sha256": manifest["parity_fixture_sha256"],
        "report_sha256": sha256(REPORT),
        "release_manifest_sha256": sha256(manifest_path),
    }
    route = {
        "task": "ic50_regression",
        "subtype": "er_alpha",
        "release_status": "released",
        "artifact": {
            "relative_path": REGRESSION_NAME,
            "size_bytes": target.stat().st_size,
            "sha256": sha256(target),
            "model_id": metadata["model_id"],
        },
        "evidence_scope": metadata.get("evidence_scope"),
    }
    trusted = {
        "model_id": metadata["model_id"],
        "sha256": sha256(target),
        "size_bytes": target.stat().st_size,
        "release_status": "released",
        "release_manifest_relative_path": REGRESSION_MANIFEST_NAME,
        "release_manifest_sha256": sha256(manifest_path),
        "evidence_scope": metadata.get("evidence_scope"),
        **route_metadata,
    }
    return route, trusted, route_metadata


def write_allowlist(trusted: dict[str, dict]) -> None:
    payload = json.dumps(trusted, indent=4, sort_keys=True, ensure_ascii=False)
    ALLOWLIST_PATH.write_text(
        '"""Build-time ERBA artifact trust anchors; generated by prepare_erba_app_assets.py."""\n'
        "from __future__ import annotations\n\n"
        f"TRUSTED_ERBA_ARTIFACTS: dict[str, dict[str, object]] = {payload}\n",
        encoding="utf-8",
    )


def smoke_test(catalog_payload: dict, trusted: dict[str, dict]) -> None:
    specs = {}
    for route in catalog_payload["routes"]:
        artifact = route["artifact"]
        specs[(ERBATask(route["task"]), ERBASubtype(route["subtype"]))] = ERBAArtifactSpec(
            relative_path=artifact["relative_path"],
            size_bytes=artifact["size_bytes"],
            sha256=artifact["sha256"],
            model_id=artifact["model_id"],
        )
    predictor = ERBAPredictor(
        MODEL_DIR,
        specs,
        regression_parity_approved=False,
        trusted_artifacts=trusted,
    )
    for task, subtype in specs:
        result = predictor.predict(
            ERBARequest(smiles="Oc1ccccc1", task=task, subtype=subtype)
        )
        if result.status_code.value != "ok":
            raise RuntimeError(
                f"App adapter smoke failed for {task.value}/{subtype.value}: "
                f"{result.status_message}"
            )

    code = (
        "import sys; "
        f"sys.path.insert(0, {str(ROOT)!r}); "
        "from gui.erba_tab import load_erba_catalog; "
        f"c,p,_=load_erba_catalog(__import__('pathlib').Path({str(MODEL_DIR / 'catalog.v2.json')!r})); "
        "from core.contracts import ERBATask,ERBASubtype; "
        "assert set(c)=={(ERBATask.CLASSIFICATION,ERBASubtype.ER_ALPHA)} and p is False; "
        "print('isolated-catalog-ok')"
    )
    completed = subprocess.run(
        [sys.executable, "-I", "-c", code],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        timeout=180,
        check=False,
    )
    if completed.returncode or "isolated-catalog-ok" not in completed.stdout:
        raise RuntimeError(f"isolated catalog smoke failed: {completed.stderr}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--classification-dir", type=Path, default=DEFAULT_CLASSIFICATION_DIR)
    args = parser.parse_args()
    if not REPORT.is_file():
        raise RuntimeError("Authoritative V7 report is required before app asset release")
    MODEL_DIR.mkdir(parents=True, exist_ok=True)

    routes: list[dict] = []
    trusted: dict[str, dict] = {}
    route_metadata: dict[str, dict] = {}
    route, trust, metadata = classification_route(args.classification_dir, "ERalpha")
    route_id = f"{route['task']}:{route['subtype']}"
    routes.append(route)
    trusted[route_id] = trust
    route_metadata[route["artifact"]["model_id"]] = metadata

    catalog = {
        "schema_version": 2,
        "product": "ER_Predictor",
        "regression_parity_approved": False,
        "classification_evidence_caveat": HISTORICAL_EXPOSURE_CAVEAT_FULL,
        "routes": routes,
        "route_metadata": route_metadata,
    }
    catalog_path = MODEL_DIR / "catalog.v2.json"
    catalog_path.write_text(
        json.dumps(catalog, indent=2, sort_keys=True, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    write_allowlist(trusted)
    smoke_test(catalog, trusted)
    print(
        json.dumps(
            {
                "catalog": str(catalog_path),
                "catalog_sha256": sha256(catalog_path),
                "allowlist": str(ALLOWLIST_PATH),
                "allowlist_sha256": sha256(ALLOWLIST_PATH),
                "routes": len(routes),
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
