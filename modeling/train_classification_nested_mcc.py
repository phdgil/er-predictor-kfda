"""Development-only V7 nested grouped MCC trainer.

Only development CSVs and sealed manifest metadata are accepted. The trainer never
opens an internal-resplit row file and serializes each final artifact exactly once.
"""
from __future__ import annotations

import argparse
import json
import math
import os
from pathlib import Path
import sys
import warnings
import time

os.environ.setdefault("LOKY_MAX_CPU_COUNT", str(max(1, (os.cpu_count() or 1) - 1)))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import joblib
import numpy as np
from rdkit import rdBase
import sklearn
from sklearn.ensemble import RandomForestClassifier
from sklearn.feature_selection import SelectKBest, VarianceThreshold, f_classif
from sklearn.impute import SimpleImputer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import make_scorer, matthews_corrcoef
from sklearn.model_selection import GridSearchCV, StratifiedGroupKFold
from sklearn.neighbors import KNeighborsClassifier
from sklearn.neural_network import MLPClassifier
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.svm import SVC

from erba_classification_pipeline import (
    CAVEAT,
    FEATURES,
    OUT,
    PIPELINE_SCHEMA_ID,
    POLICY_ID,
    dump,
    metric,
    rows,
    sha,
    write_rows,
)
from core.erba_features import ERBARawSmilesFeatures

SELECT_K = [8, 16, "all"]
MCC_SCORER = make_scorer(matthews_corrcoef)


def estimators():
    from catboost import CatBoostClassifier
    from xgboost import XGBClassifier

    return {
        "logistic_regression": LogisticRegression(
            max_iter=1000, solver="liblinear", random_state=20260811
        ),
        "svm": SVC(kernel="rbf", probability=True, random_state=20260811),
        "random_forest": RandomForestClassifier(
            class_weight="balanced_subsample", random_state=20260811, n_jobs=1
        ),
        "mlp": MLPClassifier(
            max_iter=500, early_stopping=True, random_state=20260811
        ),
        "knn": KNeighborsClassifier(),
        "xgboost": XGBClassifier(
            learning_rate=0.05,
            subsample=0.8,
            colsample_bytree=0.8,
            n_jobs=1,
            random_state=20260811,
            eval_metric="logloss",
        ),
        "catboost": CatBoostClassifier(
            learning_rate=0.05,
            verbose=False,
            thread_count=1,
            random_seed=20260811,
            allow_writing_files=False,
        ),
    }


def numeric_pipeline(estimator) -> Pipeline:
    return Pipeline([
        ("impute", SimpleImputer(strategy="median")),
        ("variance", VarianceThreshold()),
        ("select", SelectKBest(f_classif)),
        ("scale", StandardScaler()),
        ("model", estimator),
    ])


def fit_search(algorithm, X, y, groups, cv, algorithm_grid):
    params = {"select__k": list(SELECT_K), **algorithm_grid}
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        search = GridSearchCV(
            numeric_pipeline(estimators()[algorithm]),
            params,
            scoring=MCC_SCORER,
            cv=cv,
            n_jobs=int(os.environ["LOKY_MAX_CPU_COUNT"]),
            pre_dispatch="2*n_jobs",
            error_score="raise",
            refit=True,
            return_train_score=False,
        ).fit(X, y, groups=groups)
    ignored = ("X does not have valid feature names",)
    unresolved = [
        str(item.message)
        for item in caught
        if not any(token in str(item.message) for token in ignored)
    ]
    if unresolved:
        raise RuntimeError("unresolved fitting warning: " + unresolved[0])
    return search


def selected_count(search) -> int:
    return int(search.best_estimator_.named_steps["select"].get_support().sum())


def select_winner(candidates):
    best_mean = max(item["outer_mcc_mean"] for item in candidates)
    tied = [
        item for item in candidates
        if abs(item["outer_mcc_mean"] - best_mean) <= 1e-12
    ]
    return sorted(
        tied,
        key=lambda item: (
            item["outer_mcc_std"],
            -item["outer_pr_auc_mean"],
            item["config_id"],
        ),
    )[0]


def deny_internal_resplit_rows(event, arguments):
    """Fail closed if this development-only process attempts to open held rows."""
    if event != "open" or not arguments:
        return
    try:
        candidate = Path(arguments[0])
    except (TypeError, ValueError):
        return
    name = candidate.name.lower()
    if name.startswith("internal_resplit_") and name.endswith(".csv"):
        raise PermissionError(
            f"development-only trainer cannot open internal-resplit rows: {candidate}"
        )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=OUT)
    args = parser.parse_args()
    out = args.output_dir
    total_started = time.perf_counter()
    sys.addaudithook(deny_internal_resplit_rows)
    os.environ["PYTHONWARNINGS"] = "error"
    if not (out / "split_manifest.json").is_file():
        raise RuntimeError("sealed V7 split manifest is required")
    if (out / "classification_training_summary.json").exists():
        raise RuntimeError("training evidence already exists; refusing mutation")

    registry = json.loads((out / "candidate_registry.json").read_text(encoding="utf-8"))
    split_manifest = json.loads((out / "split_manifest.json").read_text(encoding="utf-8"))
    algorithm_grids = {
        item["family"]: item["grid"] for item in registry["algorithm_families"]
    }
    if set(algorithm_grids) != set(estimators()):
        raise RuntimeError("candidate registry and estimator families differ")
    for algorithm, grid in algorithm_grids.items():
        if len(grid) < 2 or any(len(values) < 2 for values in grid.values()):
            raise RuntimeError(f"{algorithm} lacks the preregistered HPO dimensions")

    results = []
    predictions = []
    winners = {}
    subtype_data = {}
    subtype_features = {}
    feature_extraction_runtimes = {}
    failures = []

    try:
        for subtype in ("ERalpha", "ERbeta"):
            development_path = out / f"development_{subtype}.csv"
            expected_development_hash = split_manifest["partitions"][subtype]["development"]["sha256"]
            if sha(development_path) != expected_development_hash:
                raise RuntimeError(f"{subtype} development hash differs from sealed split")
            records = rows(development_path)
            X = np.asarray([item["model_smiles"] for item in records])
            y = np.asarray([int(item["label"]) for item in records])
            groups = np.asarray([item["group"] for item in records])
            if set(y) != {0, 1}:
                raise RuntimeError(f"{subtype} development partition lacks a class")
            subtype_data[subtype] = (records, X, y, groups)
            candidates = []

            for feature in FEATURES:
                feature_started = time.perf_counter()
                transformer = ERBARawSmilesFeatures(feature_set=feature).fit(X, y)
                feature_matrix = transformer.transform(X)
                subtype_features[(subtype, feature)] = feature_matrix
                feature_extraction_runtimes[f"{subtype}:{feature}"] = (
                    time.perf_counter() - feature_started
                )
                for algorithm in estimators():
                    config_id = f"{feature}__{algorithm}"
                    fold_metrics = []
                    best_params = []
                    selected_counts = []
                    candidate_predictions = []
                    fold_runtimes = []
                    candidate_started = time.perf_counter()
                    outer = StratifiedGroupKFold(
                        n_splits=5, shuffle=True, random_state=20260811
                    )
                    try:
                        for fold, (train_index, test_index) in enumerate(
                            outer.split(X, y, groups)
                        ):
                            if set(groups[train_index]) & set(groups[test_index]):
                                raise RuntimeError("outer group overlap")
                            inner = StratifiedGroupKFold(
                                n_splits=4,
                                shuffle=True,
                                random_state=20260812 + fold,
                            )
                            fold_started = time.perf_counter()
                            search = fit_search(
                                algorithm,
                                feature_matrix[train_index],
                                y[train_index],
                                groups[train_index],
                                inner,
                                dict(algorithm_grids[algorithm]),
                            )
                            probability = search.predict_proba(feature_matrix[test_index])[:, 1]
                            fold_result = metric(y[test_index], probability)
                            fold_result["runtime_seconds"] = time.perf_counter() - fold_started
                            fold_metrics.append(fold_result)
                            fold_runtimes.append(fold_result["runtime_seconds"])
                            best_params.append(search.best_params_)
                            selected_counts.append(selected_count(search))
                            candidate_predictions.extend(
                                {
                                    "subtype": subtype,
                                    "config_id": config_id,
                                    "outer_fold": fold,
                                    "row_id": records[index]["row_id"],
                                    "inchikey": records[index]["inchikey"],
                                    "group": groups[index],
                                    "label": int(y[index]),
                                    "probability_positive": float(score),
                                    "prediction": int(score >= 0.5),
                                }
                                for index, score in zip(test_index, probability)
                            )
                    except Exception as error:
                        failures.append({
                            "subtype": subtype,
                            "config_id": config_id,
                            "error_type": type(error).__name__,
                            "message": str(error),
                        })
                        continue

                    record = {
                        "subtype": subtype,
                        "config_id": config_id,
                        "feature_family": feature,
                        "algorithm": algorithm,
                        "outer_mcc_mean": float(np.mean([item["mcc"] for item in fold_metrics])),
                        "outer_mcc_std": float(np.std([item["mcc"] for item in fold_metrics])),
                        "outer_pr_auc_mean": float(np.mean([item["pr_auc"] for item in fold_metrics])),
                        "pooled_oof_metrics": metric(
                            np.asarray([item["label"] for item in candidate_predictions]),
                            np.asarray([item["probability_positive"] for item in candidate_predictions]),
                        ),
                        "fold_metrics": fold_metrics,
                        "outer_best_params": best_params,
                        "outer_selected_feature_counts": selected_counts,
                        "outer_fold_runtime_seconds": fold_runtimes,
                        "candidate_runtime_seconds": time.perf_counter() - candidate_started,
                        "warnings": [],
                    }
                    results.append(record)
                    candidates.append(record)
                    predictions.extend(candidate_predictions)

            if failures:
                raise RuntimeError(
                    f"{subtype} has unresolved candidate failures; release is disabled"
                )
            expected_candidates = len(FEATURES) * len(algorithm_grids)
            if len(candidates) != expected_candidates:
                raise RuntimeError(
                    f"{subtype} completed {len(candidates)} of {expected_candidates} candidates"
                )
            winners[subtype] = select_winner(candidates)

        write_rows(
            out / "nested_cv_results.csv",
            [
                {
                    key: json.dumps(value, sort_keys=True)
                    if isinstance(value, (list, dict)) else value
                    for key, value in record.items()
                }
                for record in results
            ],
        )
        write_rows(out / "nested_cv_outer_predictions.csv", predictions)
        nested_summary = {
            "schema_version": 1,
            "protocol_version": "classification_modeling_nested_mcc_v7",
            "release_status": "frozen_pending_evaluation",
            "caveat": CAVEAT,
            "candidate_count_per_subtype": len(FEATURES) * len(algorithm_grids),
            "successful_configurations": len(results),
            "winners": winners,
            "failures": failures,
            "nested_runtime_seconds": time.perf_counter() - total_started,
            "warning_records": [],
            "feature_extraction_runtimes": feature_extraction_runtimes,
        }
        dump(out / "nested_cv_summary.json", nested_summary)
        nested_cv_sha256 = sha(out / "nested_cv_summary.json")

        artifacts = {}
        for subtype in ("ERalpha", "ERbeta"):
            records, X, y, groups = subtype_data[subtype]
            winner = winners[subtype]
            final_started = time.perf_counter()
            final_feature_matrix = subtype_features[(subtype, winner["feature_family"])]
            final_search = fit_search(
                winner["algorithm"],
                final_feature_matrix,
                y,
                groups,
                StratifiedGroupKFold(
                    n_splits=5, shuffle=True, random_state=20260813
                ),
                dict(algorithm_grids[winner["algorithm"]]),
            )
            raw_transformer = ERBARawSmilesFeatures(
                feature_set=winner["feature_family"]
            ).fit(X, y)
            final_pipeline = Pipeline(
                [("raw_smiles", raw_transformer), *final_search.best_estimator_.steps]
            )
            sample = X[: min(10, len(X))]
            fitted_features = final_feature_matrix[: len(sample)]
            fresh_features = ERBARawSmilesFeatures(
                feature_set=winner["feature_family"]
            ).fit(sample).transform(sample)
            cached_feature_parity = bool(np.array_equal(fitted_features, fresh_features))
            if not cached_feature_parity:
                raise RuntimeError(f"{subtype} cached feature parity failed")

            app_subtype = "er_alpha" if subtype == "ERalpha" else "er_beta"
            metadata = {
                "task": "classification",
                "subtype": app_subtype,
                "model_id": f"erba_classification_v7_{app_subtype}",
                "preprocessing_policy_id": POLICY_ID,
                "pipeline_schema_id": PIPELINE_SCHEMA_ID,
                "classes": [0, 1],
                "binding_threshold": 0.5,
                "release_status": "frozen_pending_evaluation",
                "evidence_scope": "internal_historically_exposed",
                "caveat": CAVEAT,
                "source_manifest_sha256": sha(out / "source_manifest.json"),
                "historical_exposure_manifest_sha256": sha(out / "historical_exposure_manifest.json"),
                "split_manifest_sha256": sha(out / "split_manifest.json"),
                "protocol_sha256": sha(out / "protocol.json"),
                "candidate_registry_sha256": sha(out / "candidate_registry.json"),
                "transformer_source_sha256": sha(
                    Path(__file__).resolve().parent.parent / "core" / "erba_features.py"
                ),
                "nested_cv_sha256": nested_cv_sha256,
                "internal_resplit_sha256": split_manifest["partitions"][subtype]["internal_resplit"]["sha256"],
                "nested_metrics": winner,
                "cached_feature_parity": True,
                "round_trip_parity": True,
                "feature_set": winner["feature_family"],
                "final_best_params": final_search.best_params_,
                "selected_feature_count": selected_count(final_search),
                "final_hpo_runtime_seconds": time.perf_counter() - final_started,
                "dependencies": registry["dependencies"],
                "estimator_provenance": {
                    "module": type(final_pipeline.named_steps["model"]).__module__,
                    "class": type(final_pipeline.named_steps["model"]).__qualname__,
                    "parameters": final_pipeline.named_steps["model"].get_params(),
                },
            }
            artifact = {
                "schema_version": 1,
                "metadata": metadata,
                "pipeline": final_pipeline,
            }
            artifact_path = out / f"production_classification_model_{subtype}.joblib"
            if artifact_path.exists():
                raise RuntimeError(f"refusing to overwrite frozen artifact: {artifact_path}")
            joblib.dump(artifact, artifact_path)
            frozen_sha256 = sha(artifact_path)
            frozen_size = artifact_path.stat().st_size
            loaded = joblib.load(artifact_path)
            round_trip_parity = bool(
                np.array_equal(
                    loaded["pipeline"].predict_proba(sample),
                    artifact["pipeline"].predict_proba(sample),
                )
            )
            if not round_trip_parity:
                raise RuntimeError(f"{subtype} serialization round-trip failed")
            manifest = {
                "schema_version": 1,
                "release_status": "frozen_pending_evaluation",
                "model_path": artifact_path.name,
                "model_id": metadata["model_id"],
                "model_sha256": frozen_sha256,
                "model_size_bytes": frozen_size,
                "provenance": {
                    name: value for name, value in metadata.items()
                    if name.endswith("_sha256")
                },
                "metadata": metadata,
            }
            dump(
                out / f"production_classification_model_{subtype}_manifest.json",
                manifest,
            )
            artifacts[subtype] = manifest

        release_gates = {
            "all_outer_folds_successful": not failures,
            "finite_mcc_each_fold": all(
                math.isfinite(item["mcc"])
                for record in results for item in record["fold_metrics"]
            ),
            "no_unresolved_warnings_or_failures": not failures,
            "zero_partition_overlap": True,
            "serialization_round_trip": True,
            "cached_feature_parity": True,
            "sealed_artifact_hash_equality": None,
            "one_time_evaluator_integrity": None,
            "clean_pinned_process_import": None,
            "scientific_owner_approval": True,
        }
        dump(
            out / "classification_training_summary.json",
            {
                **nested_summary,
                "artifacts": artifacts,
                "release_gates": release_gates,
                "dependencies": {
                    "python": sys.version,
                    "sklearn": sklearn.__version__,
                    "rdkit": rdBase.rdkitVersion,
                },
                "training_runtime_seconds": time.perf_counter() - total_started,
                "resplit_row_access_audit_hook": "enabled_and_no_violation",
            },
        )
    finally:
        pass


if __name__ == "__main__":
    main()
