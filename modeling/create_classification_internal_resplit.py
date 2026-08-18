"""Create and receipt the V7 grouped development/internal-sensitivity partition."""
from __future__ import annotations

import argparse
import importlib.metadata
import json
from pathlib import Path
import platform
import sys

import numpy as np
from rdkit import rdBase
import sklearn
from sklearn.model_selection import StratifiedGroupKFold

sys.path.insert(0, str(Path(__file__).resolve().parent))
from erba_classification_pipeline import (
    CAVEAT,
    FEATURES,
    OUT,
    PIPELINE_SCHEMA_ID,
    POLICY_ID,
    SOURCE,
    dump,
    group_for,
    model_smiles,
    rows,
    sha,
    text_sha,
    write_rows,
)

ALGORITHMS = (
    "logistic_regression",
    "svm",
    "random_forest",
    "mlp",
    "knn",
    "xgboost",
    "catboost",
)
GRIDS = {
    "logistic_regression": {
        "model__C": [0.1, 1.0],
        "model__class_weight": ["balanced", None],
    },
    "svm": {"model__C": [0.5, 2.0], "model__gamma": ["scale", "auto"]},
    "random_forest": {
        "model__n_estimators": [40, 80],
        "model__max_depth": [None, 12],
    },
    "mlp": {
        "model__hidden_layer_sizes": [(16,), (32,)],
        "model__alpha": [0.0001, 0.001],
    },
    "knn": {"model__n_neighbors": [3, 7], "model__p": [1, 2]},
    "xgboost": {
        "model__n_estimators": [30, 60],
        "model__max_depth": [2, 4],
    },
    "catboost": {"model__iterations": [30, 60], "model__depth": [3, 5]},
}
SELECT_K = [8, 16, "all"]
SUBTYPES = ("ERalpha", "ERbeta")
PROTOCOL_VERSION = "classification_modeling_nested_mcc_v7"
PROJECT = Path(__file__).resolve().parent.parent
NORMALIZED_ROOT = OUT.parent


def dependency_versions() -> dict[str, str]:
    names = ("numpy", "scikit-learn", "rdkit", "joblib", "xgboost", "catboost")
    versions = {}
    for name in names:
        try:
            versions[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            versions[name] = "not-installed"
    versions.update(
        {
            "python": platform.python_version(),
            "python_executable": str(Path(sys.executable).resolve()),
            "sklearn_runtime": sklearn.__version__,
            "rdkit_runtime": rdBase.rdkitVersion,
        }
    )
    return versions


def script_inventory() -> list[dict]:
    paths = (
        Path(__file__).resolve(),
        Path(__file__).with_name("erba_classification_pipeline.py"),
        Path(__file__).with_name("train_classification_nested_mcc.py"),
        Path(__file__).with_name("evaluate_classification_internal_resplit.py"),
        PROJECT / "core" / "erba_features.py",
        PROJECT / "core" / "erba_preprocessing.py",
        PROJECT / "core" / "contracts.py",
    )
    return [
        {
            "path": str(path),
            "size_bytes": path.stat().st_size,
            "sha256": sha(path),
        }
        for path in paths
    ]


def historical_inventory() -> list[dict]:
    candidates = []
    for directory in sorted(NORMALIZED_ROOT.glob("classification_modeling*")):
        if not directory.is_dir() or directory.resolve() == OUT.resolve():
            continue
        for path in sorted(item for item in directory.rglob("*") if item.is_file()):
            candidates.append(path)
    old_report = NORMALIZED_ROOT / "er_regression_classification_model_report_final.html"
    if old_report.is_file():
        candidates.append(old_report)
    return [
        {
            "path": str(path.relative_to(NORMALIZED_ROOT)).replace("\\", "/"),
            "size_bytes": path.stat().st_size,
            "sha256": sha(path),
        }
        for path in candidates
    ]


def assign_components(subtype_rows: list[dict]) -> list[dict]:
    parent = list(range(len(subtype_rows)))

    def root(index: int) -> int:
        while parent[index] != index:
            parent[index] = parent[parent[index]]
            index = parent[index]
        return index

    def join(left: int, right: int) -> None:
        left_root, right_root = root(left), root(right)
        if left_root != right_root:
            parent[right_root] = left_root

    seen: dict[str, int] = {}
    for index, row in enumerate(subtype_rows):
        for token in (f"inchikey:{row['inchikey']}", f"scaffold:{row['group']}"):
            if token in seen:
                join(index, seen[token])
            else:
                seen[token] = index
    for index, row in enumerate(subtype_rows):
        members = sorted(
            subtype_rows[other]["inchikey"]
            for other in range(len(subtype_rows))
            if root(other) == root(index)
        )
        row["group"] = "identity_scaffold_component:" + text_sha("|".join(members))[:24]
    return subtype_rows


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=OUT)
    args = parser.parse_args()
    out = args.output_dir.resolve()
    if out.exists() and any(out.iterdir()):
        raise RuntimeError("V7 output already exists; immutable evidence is never overwritten")
    out.mkdir(parents=True, exist_ok=True)

    source_rows = rows(SOURCE)
    valid: list[dict] = []
    excluded: list[dict] = []
    for source_row in source_rows:
        if source_row.get("receptor_subtype") not in set(SUBTYPES):
            continue
        if source_row.get("resolved_label") not in {"positive", "negative"}:
            raise RuntimeError("nonbinary label encountered in model-ready source")
        row = dict(source_row)
        try:
            row["model_smiles"] = model_smiles(row["standardized_smiles"])
            row["group"] = group_for(row["standardized_smiles"])
        except ValueError as error:
            excluded.append({"row_id": row.get("row_id"), "reason": str(error)})
            continue
        row["label"] = int(row["resolved_label"] == "positive")
        if not row.get("row_id") or not row.get("inchikey"):
            raise RuntimeError("model-ready row lacks stable row identity")
        valid.append(row)
    if len({row["row_id"] for row in valid}) != len(valid):
        raise RuntimeError("duplicate row_id in model-ready source")

    versions = dependency_versions()
    scripts = script_inventory()
    historical = historical_inventory()
    protocol = {
        "schema_version": 1,
        "protocol_version": PROTOCOL_VERSION,
        "caveat": CAVEAT,
        "evidence_scope": "internal_historically_exposed",
        "split": {
            "method": "StratifiedGroupKFold",
            "n_splits": 5,
            "seed": 20260810,
            "internal_resplit_fold": 0,
            "grouping": "connected components over exact InChIKey and chiral Bemis-Murcko scaffold",
        },
        "nested_cv": {
            "outer_folds": 5,
            "outer_seed": 20260811,
            "inner_folds": 4,
            "inner_seed": "20260812 + outer_fold",
            "final_folds": 5,
            "final_seed": 20260813,
            "metric": "MCC at fixed threshold 0.5",
            "search": "GridSearchCV",
            "candidate_families_per_subtype": len(FEATURES) * len(ALGORITHMS),
            "parameter_configurations_per_family": 12,
            "failure_policy": "any candidate, fold, warning, non-finite MCC, or parity failure disables release; no runner-up substitution",
        },
        "dependencies": versions,
        "scripts": scripts,
        "release_gates": {
            "all_outer_folds_successful": True,
            "finite_mcc_each_fold": True,
            "no_unresolved_warnings_or_failures": True,
            "zero_partition_overlap": True,
            "serialization_round_trip": True,
            "cached_feature_parity": True,
            "sealed_artifact_hash_equality": True,
            "one_time_evaluator_integrity": True,
            "clean_pinned_process_import": True,
            "scientific_owner_approval": "approved_plan_conditional_when_all_preregistered_gates_pass",
        },
    }
    registry = {
        "schema_version": 1,
        "protocol_version": PROTOCOL_VERSION,
        "frozen_before_scoring": True,
        "historical_provenance": (
            "Candidate families were historically exposed exploratory knowledge only; "
            "no old artifact, split, matrix, metric, model, or prediction is a training input."
        ),
        "search_method": "exhaustive GridSearchCV with MCC refit",
        "parameter_configurations_per_family": 12,
        "feature_families": [
            {
                "family": feature,
                "status": "enabled",
                "transformer": "core.erba_features.ERBARawSmilesFeatures",
                "select__k": SELECT_K,
            }
            for feature in FEATURES
        ],
        "algorithm_families": [
            {"family": algorithm, "status": "enabled", "grid": GRIDS[algorithm]}
            for algorithm in ALGORITHMS
        ],
        "candidate_count": len(FEATURES) * len(ALGORITHMS),
        "selection": (
            "highest arithmetic mean development outer MCC; ties within 1e-12 use "
            "lower outer MCC standard deviation, then higher mean outer PR-AUC, then "
            "lexicographically smaller config_id"
        ),
        "seeds": protocol["nested_cv"],
        "failure_policy": protocol["nested_cv"]["failure_policy"],
        "dependencies": versions,
    }
    dump(out / "protocol.json", protocol)
    dump(out / "candidate_registry.json", registry)
    dump(
        out / "source_manifest.json",
        {
            "schema_version": 1,
            "source_path": str(SOURCE),
            "source_sha256": sha(SOURCE),
            "source_rows": len(source_rows),
            "accepted_rows": len(valid),
            "excluded_rows": len(excluded),
            "excluded_row_records": excluded,
            "subtype_label_counts": {
                subtype: {
                    "rows": sum(row["receptor_subtype"] == subtype for row in valid),
                    "positive": sum(
                        row["receptor_subtype"] == subtype and row["label"] == 1
                        for row in valid
                    ),
                    "negative": sum(
                        row["receptor_subtype"] == subtype and row["label"] == 0
                        for row in valid
                    ),
                }
                for subtype in SUBTYPES
            },
            "preprocessing_policy_id": POLICY_ID,
            "pipeline_schema_id": PIPELINE_SCHEMA_ID,
            "dependencies": versions,
            "scripts": scripts,
            "caveat": CAVEAT,
        },
    )
    dump(
        out / "historical_exposure_manifest.json",
        {
            "schema_version": 1,
            "evidence_scope": "internal_historically_exposed",
            "caveat": CAVEAT,
            "prohibition": (
                "Trainer receives only sealed development CSVs and manifest metadata; "
                "it cannot access internal-resplit rows or historical artifacts."
            ),
            "historical_artifact_inventory": historical,
            "historical_artifact_count": len(historical),
        },
    )

    partitioned_rows: list[dict] = []
    for subtype in SUBTYPES:
        subtype_rows = assign_components(
            [row for row in valid if row["receptor_subtype"] == subtype]
        )
        y = np.asarray([row["label"] for row in subtype_rows])
        groups = np.asarray([row["group"] for row in subtype_rows])
        splitter = StratifiedGroupKFold(
            n_splits=5, shuffle=True, random_state=20260810
        )
        for fold, (_, test_indices) in enumerate(
            splitter.split(np.zeros(len(y)), y, groups)
        ):
            for index in test_indices:
                subtype_rows[index]["partition"] = (
                    "internal_resplit" if fold == 0 else "development"
                )
                subtype_rows[index]["split_fold"] = fold
                partitioned_rows.append(subtype_rows[index])

    split_manifest = {
        "schema_version": 1,
        "source_manifest_sha256": sha(out / "source_manifest.json"),
        "historical_exposure_manifest_sha256": sha(out / "historical_exposure_manifest.json"),
        "protocol_sha256": sha(out / "protocol.json"),
        "candidate_registry_sha256": sha(out / "candidate_registry.json"),
        "caveat": CAVEAT,
        "partitions": {},
    }
    for subtype in SUBTYPES:
        subtype_rows = [
            row for row in partitioned_rows if row["receptor_subtype"] == subtype
        ]
        development = [
            row for row in subtype_rows if row["partition"] == "development"
        ]
        internal_resplit = [
            row for row in subtype_rows if row["partition"] == "internal_resplit"
        ]
        for key in ("row_id", "inchikey", "group"):
            if {row[key] for row in development} & {
                row[key] for row in internal_resplit
            }:
                raise RuntimeError(f"{subtype} {key} overlap")
        if {row["label"] for row in development} != {0, 1}:
            raise RuntimeError(f"{subtype} development partition lacks a class")
        if {row["label"] for row in internal_resplit} != {0, 1}:
            raise RuntimeError(f"{subtype} internal-resplit partition lacks a class")
        write_rows(out / f"development_{subtype}.csv", development)
        write_rows(out / f"internal_resplit_{subtype}.csv", internal_resplit)
        split_manifest["partitions"][subtype] = {
            name: {
                "rows": len(partition),
                "positive": sum(int(row["label"]) for row in partition),
                "negative": sum(1 - int(row["label"]) for row in partition),
                "sha256": sha(out / f"{name}_{subtype}.csv"),
            }
            for name, partition in (
                ("development", development),
                ("internal_resplit", internal_resplit),
            )
        }
    dump(out / "split_manifest.json", split_manifest)


if __name__ == "__main__":
    main()
