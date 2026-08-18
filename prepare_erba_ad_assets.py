"""Build route-specific ERBA applicability-domain reference assets.

These references are a reliability aid only; they are not regulatory validation.
"""
from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from pathlib import Path
import sys

import numpy as np
import pandas as pd
from rdkit import Chem

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from core.ad import ADCalculator
from core.fingerprint import RDKIT_COLUMNS, detect_rdkit_columns, rdkit_fp_from_smiles

DATA_ROOT = Path(r"D:/research/FDA_endocrine_disruption/data/normalized_erba")
CLASSIFICATION_ROOT = DATA_ROOT / "classification_modeling_nested_mcc_v7"
REGRESSION_ROOT = DATA_ROOT / "ic50_resolution_step3"
OUTPUT_ROOT = ROOT / "models" / "erba" / "ad"


@dataclass(frozen=True)
class Route:
    key: str
    filename: str
    expected_rows: int


ROUTES = (
    Route("classification/er_alpha", "classification_er_alpha_reference.xlsx", 1481),
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _required_columns(frame: pd.DataFrame, columns: set[str], source: Path) -> None:
    missing = columns.difference(frame.columns)
    if missing:
        raise ValueError(f"{source} is missing required columns: {sorted(missing)}")


def _validate_unique_ids(frame: pd.DataFrame, source: Path) -> None:
    if frame["row_id"].isna().any() or frame["row_id"].astype(str).str.strip().eq("").any():
        raise ValueError(f"{source} contains blank row_id values")
    if frame["row_id"].duplicated().any():
        raise ValueError(f"{source} contains duplicate row_id values")


def classification_rows(subtype: str) -> tuple[pd.DataFrame, list[Path]]:
    source = CLASSIFICATION_ROOT / f"development_{subtype}.csv"
    frame = pd.read_csv(source)
    _required_columns(frame, {"row_id", "receptor_subtype", "model_smiles", "label", "partition"}, source)
    _validate_unique_ids(frame, source)
    expected_subtype = subtype
    if not frame["receptor_subtype"].eq(expected_subtype).all():
        raise ValueError(f"{source} contains rows outside {expected_subtype}")
    if not frame["partition"].eq("development").all():
        raise ValueError(f"{source} contains non-development rows")
    if frame["label"].isna().any() or not frame["label"].isin((0, 1)).all():
        raise ValueError(f"{source} contains invalid classification labels")
    columns = [
        column for column in (
            "row_id", "source_name", "source_record_id", "compound_name", "cas",
            "inchikey", "receptor_subtype", "partition", "model_smiles", "label",
        ) if column in frame.columns
    ]
    return frame[columns].copy(), [source]


def regression_rows() -> tuple[pd.DataFrame, list[Path]]:
    source = REGRESSION_ROOT / "ic50_fast_regression_exact_pic50.csv"
    split_source = REGRESSION_ROOT / "model_build_subtype_v1" / "ESR1_ERalpha" / "train_test_split.csv"
    source_frame = pd.read_csv(source)
    split_frame = pd.read_csv(split_source)
    _required_columns(source_frame, {"row_id", "receptor_subtype", "rdkit_parent_smiles", "pic50"}, source)
    _required_columns(split_frame, {"row_id", "split"}, split_source)
    _validate_unique_ids(source_frame, source)
    _validate_unique_ids(split_frame, split_source)
    if not split_frame["split"].isin(("train", "test")).all():
        raise ValueError(f"{split_source} contains an unexpected split label")
    expected_rows = 1069
    if (split_frame["split"] == "train").sum() != expected_rows:
        raise ValueError(f"{split_source} has an unexpected training-row count")
    joined = source_frame.merge(split_frame, on="row_id", how="inner", validate="one_to_one")
    if len(joined) != len(split_frame):
        raise ValueError("Regression source/split join dropped or duplicated rows")
    train = joined.loc[joined["split"].eq("train")].copy()
    if len(train) != expected_rows:
        raise ValueError("Regression training reference has an unexpected row count")
    if not train["receptor_subtype"].eq("ESR1_ERalpha").all():
        raise ValueError("Regression training reference contains rows outside ESR1_ERalpha")
    if train["pic50"].isna().any() or not np.isfinite(train["pic50"].astype(float)).all():
        raise ValueError("Regression training reference contains invalid pIC50 labels")
    columns = [
        column for column in (
            "row_id", "source_name", "source_record_id", "original_compound_identifier",
            "standardized_inchikey", "receptor_subtype", "rdkit_parent_smiles", "pic50",
        ) if column in train.columns
    ]
    return train[columns].rename(
        columns={"rdkit_parent_smiles": "model_smiles", "pic50": "pIC50"}
    ), [source, split_source]


def fingerprint_reference(frame: pd.DataFrame, route: Route) -> pd.DataFrame:
    rows: list[np.ndarray] = []
    canonical: list[str] = []
    for row_id, smiles in zip(frame["row_id"], frame["model_smiles"]):
        fp, valid, normalized = rdkit_fp_from_smiles(smiles)
        if not valid or not normalized or Chem.MolFromSmiles(normalized) is None:
            raise ValueError(f"{route.key} has malformed model structure for row_id={row_id}")
        rows.append(fp)
        canonical.append(normalized)
    if len(rows) != route.expected_rows:
        raise ValueError(f"{route.key} has an unexpected row count")
    metadata = frame.copy()
    metadata["model_smiles"] = canonical
    metadata["SMILES"] = canonical
    fingerprints = pd.DataFrame(np.vstack(rows), columns=RDKIT_COLUMNS, dtype=np.int8)
    output = pd.concat([metadata.reset_index(drop=True), fingerprints], axis=1)
    if detect_rdkit_columns(output) != RDKIT_COLUMNS:
        raise RuntimeError(f"{route.key} did not produce exactly the expected RDKit columns")
    return output


def build_route(route: Route) -> dict[str, object]:
    if route.key == "classification/er_alpha":
        frame, sources = classification_rows("ERalpha")
    elif route.key == "classification/er_beta":
        frame, sources = classification_rows("ERbeta")
    elif route.key == "ic50_regression/er_alpha":
        frame, sources = regression_rows()
    else:
        raise ValueError(f"Unsupported ERBA route: {route.key}")
    reference = OUTPUT_ROOT / route.filename
    cache = Path(ADCalculator.adjacent_cache_path(str(reference)))
    reference.parent.mkdir(parents=True, exist_ok=True)
    fingerprint_reference(frame, route).to_excel(reference, index=False)
    calculator = ADCalculator()
    calculator.fit_from_excel_cached(str(reference), str(cache), force_refit=True)
    if not cache.is_file():
        raise RuntimeError(f"{route.key} AD cache was not written")
    return {
        "route": route.key,
        "rows": len(frame),
        "sources": {str(path): sha256(path) for path in sources},
        "reference": str(reference),
        "reference_sha256": sha256(reference),
        "cache": str(cache),
        "cache_sha256": sha256(cache),
    }


def main() -> None:
    print(json.dumps({"routes": [build_route(route) for route in ROUTES]}, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
