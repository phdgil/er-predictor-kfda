"""Serialized raw-SMILES feature transformer for ERBA classification artifacts."""
from __future__ import annotations

import math
import numpy as np
from rdkit import Chem, DataStructs, rdBase
from rdkit.Avalon import pyAvalonTools
from rdkit.Chem import Descriptors, MACCSkeys, rdFingerprintGenerator
from sklearn.base import BaseEstimator, TransformerMixin
from .erba_preprocessing import classification_parent_smiles


class ERBARawSmilesFeatures(BaseEstimator, TransformerMixin):
    """Fixed ERBA parent policy followed by deterministic RDKit feature dispatch."""
    def __init__(self, feature_set: str = "morgan_r2_2048") -> None:
        self.feature_set = feature_set

    @staticmethod
    def _bits(fingerprint: object, size: int) -> np.ndarray:
        values = np.zeros(size, dtype=np.float32)
        DataStructs.ConvertToNumpyArray(fingerprint, values)
        return values

    def _vector(self, model_smiles: str) -> np.ndarray:
        molecule = Chem.MolFromSmiles(model_smiles, sanitize=True)
        if molecule is None:
            raise ValueError("Classification parent SMILES cannot be reparsed.")
        if self.feature_set == "rdkit_2d":
            # RDKit's descriptor registry still invokes deprecated Morgan helpers
            # internally. Preserve the vector while containing library-owned logs.
            with rdBase.BlockLogs():
                values = np.asarray(
                    [float(fn(molecule)) for _, fn in Descriptors._descList],
                    dtype=np.float64,
                )
            values[~np.isfinite(values)] = np.nan
            values[np.abs(values) > 1e15] = np.nan
            return values.astype(np.float32)
        if self.feature_set == "morgan_r2_2048":
            return self._bits(rdFingerprintGenerator.GetMorganGenerator(radius=2, fpSize=2048).GetFingerprint(molecule), 2048)
        if self.feature_set == "morgan_r3_2048":
            return self._bits(rdFingerprintGenerator.GetMorganGenerator(radius=3, fpSize=2048).GetFingerprint(molecule), 2048)
        if self.feature_set == "rdkit_fp_2048":
            return self._bits(rdFingerprintGenerator.GetRDKitFPGenerator(fpSize=2048).GetFingerprint(molecule), 2048)
        if self.feature_set == "maccs_167":
            return self._bits(MACCSkeys.GenMACCSKeys(molecule), 167)
        if self.feature_set == "atom_pair_2048":
            return self._bits(rdFingerprintGenerator.GetAtomPairGenerator(fpSize=2048).GetFingerprint(molecule), 2048)
        if self.feature_set == "torsion_2048":
            return self._bits(rdFingerprintGenerator.GetTopologicalTorsionGenerator(fpSize=2048).GetFingerprint(molecule), 2048)
        if self.feature_set == "avalon_2048":
            return self._bits(pyAvalonTools.GetAvalonFP(molecule, nBits=2048), 2048)
        raise ValueError(f"Unsupported ERBA feature set: {self.feature_set}")

    def fit(self, X: object, y: object = None) -> "ERBARawSmilesFeatures":
        self.n_features_out_ = int(self._vector("CC").shape[0])
        return self

    def transform(self, X: object) -> np.ndarray:
        vectors = []
        for raw_smiles in X:
            prepared = classification_parent_smiles(raw_smiles)
            if not prepared.valid or not prepared.model_smiles:
                raise ValueError(f"Invalid ERBA raw SMILES: {prepared.status_code.value}")
            vectors.append(self._vector(prepared.model_smiles))
        return np.vstack(vectors).astype(np.float32)
