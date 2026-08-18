"""Route-specific applicability-domain support for ERBA predictions."""
from __future__ import annotations

import os
import tempfile
from pathlib import Path

import numpy as np

from core.ad import ADCalculator
from core.contracts import ERBASubtype, ERBATask
from core.fingerprint import rdkit_fp_from_smiles


ERBA_AD_REFERENCE_FILENAMES = {
    (ERBATask.CLASSIFICATION, ERBASubtype.ER_ALPHA): "classification_er_alpha_reference.xlsx",
    (ERBATask.CLASSIFICATION, ERBASubtype.ER_BETA): "classification_er_beta_reference.xlsx",
    (ERBATask.IC50_REGRESSION, ERBASubtype.ER_ALPHA): "ic50_regression_er_alpha_reference.xlsx",
}


def erba_ad_reference_path(project_root: str | Path, task: ERBATask, subtype: ERBASubtype) -> Path:
    """Return the bundled training reference for one released ERBA route."""
    try:
        filename = ERBA_AD_REFERENCE_FILENAMES[(task, subtype)]
    except KeyError as error:
        raise ValueError(f"ERBA AD is unavailable for route {task.value}/{subtype.value}.") from error
    return Path(project_root) / "models" / "erba" / "ad" / filename


def erba_ad_cache_path(project_root: str | Path, task: ERBATask, subtype: ERBASubtype) -> Path:
    """Return a per-route writable state-cache path outside the immutable project."""
    root = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()) / "ERTA_Predictor" / "erba_ad"
    filename = erba_ad_reference_path(project_root, task, subtype).stem + "_ad_cache.npz"
    return root / filename


class ERBAApplicabilityDomain:
    """Own one unchanged ADCalculator per ERBA model route."""

    def __init__(self, project_root: str | Path):
        self.project_root = Path(project_root)
        self.calculators: dict[tuple[ERBATask, ERBASubtype], ADCalculator] = {}

    def calculator_for(self, task: ERBATask, subtype: ERBASubtype) -> ADCalculator:
        route = (task, subtype)
        if route not in ERBA_AD_REFERENCE_FILENAMES:
            raise ValueError(f"ERBA AD is unavailable for route {task.value}/{subtype.value}.")
        calculator = self.calculators.get(route)
        if calculator is None:
            calculator = ADCalculator(cache_root=str(erba_ad_cache_path(self.project_root, task, subtype).parent))
            self.calculators[route] = calculator
        return calculator

    def evaluate(self, task: ERBATask, subtype: ERBASubtype, smiles: str):
        """Fit/load this route's reference and evaluate a valid input molecule."""
        fingerprint, valid, canonical_smiles = rdkit_fp_from_smiles(smiles)
        if not valid:
            raise ValueError("Applicability domain requires a valid SMILES string.")
        reference_path = erba_ad_reference_path(self.project_root, task, subtype)
        cache_path = erba_ad_cache_path(self.project_root, task, subtype)
        calculator = self.calculator_for(task, subtype)
        calculator.ensure_fitted_from_excel_cached(str(reference_path), cache_path=str(cache_path))
        result = calculator.predict(np.asarray([fingerprint], dtype=float), input_smiles=canonical_smiles)
        return calculator, result, np.asarray([fingerprint], dtype=float)
