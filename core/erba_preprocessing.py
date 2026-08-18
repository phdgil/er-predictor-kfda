"""ERBA structure policies.  These are intentionally independent from frozen ERTA code."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

import numpy as np
from rdkit import Chem
from rdkit.Avalon import pyAvalonTools
from rdkit.Chem.MolStandardize import rdMolStandardize

from .contracts import (
    CLASSIFICATION_PARENT_POLICY_ID,
    REGRESSION_PARENT_POLICY_ID,
    ERBAStatusCode,
)


@dataclass(frozen=True)
class PreprocessingResult:
    raw_smiles: str
    model_smiles: str | None
    status_code: ERBAStatusCode
    status_message: str

    @property
    def valid(self) -> bool:
        return self.status_code is ERBAStatusCode.OK


def _failure(raw_smiles: str, code: ERBAStatusCode, message: str) -> PreprocessingResult:
    return PreprocessingResult(raw_smiles, None, code, message)


def classification_parent_smiles(raw_smiles: str) -> PreprocessingResult:
    """Apply the fixed classification parent policy and fail closed on unsafe structures."""
    if not isinstance(raw_smiles, str) or not raw_smiles.strip():
        return _failure("" if raw_smiles is None else str(raw_smiles), ERBAStatusCode.BLANK_SMILES, "SMILES is blank.")
    text = raw_smiles.strip()
    if "[*]" in text:
        return _failure(raw_smiles, ERBAStatusCode.WILDCARD_SMILES, "Wildcard atoms are not supported.")
    if "|" in text:
        return _failure(raw_smiles, ERBAStatusCode.CXSMILES_NOT_ALLOWED, "CXSMILES syntax is not supported.")
    try:
        molecule = Chem.MolFromSmiles(text, sanitize=True)
    except Exception:
        molecule = None
    if molecule is None:
        return _failure(raw_smiles, ERBAStatusCode.INVALID_SMILES, "SMILES cannot be parsed and sanitized.")
    if any(atom.GetAtomicNum() == 0 for atom in molecule.GetAtoms()):
        return _failure(raw_smiles, ERBAStatusCode.WILDCARD_SMILES, "Wildcard atoms are not supported.")
    try:
        parent = rdMolStandardize.LargestFragmentChooser().choose(molecule)
    except Exception:
        return _failure(raw_smiles, ERBAStatusCode.INVALID_SMILES, "SMILES has no molecular fragments.")
    if parent is None or parent.GetNumAtoms() == 0:
        return _failure(raw_smiles, ERBAStatusCode.INVALID_SMILES, "SMILES has no molecular fragments.")
    atomic_numbers = [atom.GetAtomicNum() for atom in parent.GetAtoms()]
    if 6 not in atomic_numbers:
        return _failure(raw_smiles, ERBAStatusCode.NO_CARBON, "Largest fragment does not contain carbon.")
    allowed = {1, 5, 6, 7, 8, 9, 14, 15, 16, 17, 34, 35, 53}
    if any(number not in allowed for number in atomic_numbers):
        return _failure(raw_smiles, ERBAStatusCode.METAL_RETAINED, "Largest fragment retains a metal or unsupported element.")
    try:
        model_smiles = Chem.MolToSmiles(parent, isomericSmiles=True, canonical=True)
    except Exception:
        return _failure(raw_smiles, ERBAStatusCode.INVALID_SMILES, "SMILES cannot be canonicalized.")
    if not model_smiles:
        return _failure(raw_smiles, ERBAStatusCode.INVALID_SMILES, "SMILES canonicalization produced no parent.")
    return PreprocessingResult(raw_smiles, model_smiles, ERBAStatusCode.OK, "")


def regression_parent_smiles(raw_smiles: str, *, parity_approved: bool) -> PreprocessingResult:
    """Canonicalize/desalt only after the independent regression parity gate is approved."""
    if not parity_approved:
        return _failure(raw_smiles, ERBAStatusCode.REGRESSION_PARITY_NOT_APPROVED, "Regression preprocessing parity is not approved.")
    if not isinstance(raw_smiles, str) or not raw_smiles.strip():
        return _failure("" if raw_smiles is None else str(raw_smiles), ERBAStatusCode.BLANK_SMILES, "SMILES is blank.")
    text = raw_smiles.strip()
    if "|" in text:
        return _failure(raw_smiles, ERBAStatusCode.CXSMILES_NOT_ALLOWED, "CXSMILES syntax is not supported.")
    molecule = Chem.MolFromSmiles(text, sanitize=True)
    if molecule is None:
        return _failure(raw_smiles, ERBAStatusCode.INVALID_SMILES, "SMILES cannot be sanitized.")
    fragments = Chem.GetMolFrags(molecule, asMols=True, sanitizeFrags=True)
    if not fragments:
        return _failure(raw_smiles, ERBAStatusCode.INVALID_SMILES, "SMILES has no molecular fragments.")
    parent = max(fragments, key=lambda fragment: (fragment.GetNumHeavyAtoms(), Chem.MolToSmiles(fragment, isomericSmiles=True)))
    return PreprocessingResult(raw_smiles, Chem.MolToSmiles(parent, isomericSmiles=True), ERBAStatusCode.OK, "")


def _fingerprint(smiles: str, generator: Callable[[Chem.Mol], object]) -> np.ndarray:
    molecule = Chem.MolFromSmiles(smiles, sanitize=True)
    if molecule is None:
        raise ValueError("Cannot fingerprint an invalid model SMILES.")
    bit_vector = generator(molecule)
    values = np.zeros((2048,), dtype=np.uint8)
    # RDKit's explicit bit-vector supports copying directly into a numpy array.
    from rdkit import DataStructs
    DataStructs.ConvertToNumpyArray(bit_vector, values)
    return values


def rdkit_fp_2048(model_smiles: str) -> np.ndarray:
    return _fingerprint(model_smiles, lambda molecule: Chem.RDKFingerprint(molecule, fpSize=2048))


def avalon_fp_2048(model_smiles: str) -> np.ndarray:
    return _fingerprint(model_smiles, lambda molecule: pyAvalonTools.GetAvalonFP(molecule, nBits=2048))


__all__ = [
    "CLASSIFICATION_PARENT_POLICY_ID", "REGRESSION_PARENT_POLICY_ID", "PreprocessingResult",
    "classification_parent_smiles", "regression_parent_smiles", "rdkit_fp_2048", "avalon_fp_2048",
]
