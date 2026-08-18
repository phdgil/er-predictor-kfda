from __future__ import annotations

import numpy as np
import pandas as pd

N_BITS = 2048
RDKIT_COLUMNS = [f"RDKit_{i + 1}" for i in range(N_BITS)]


def normalize_smiles(smiles) -> str:
    if smiles is None or pd.isna(smiles):
        return ""
    return str(smiles).strip()


def rdkit_fp_from_smiles(smiles: str, n_bits: int = N_BITS):
    """Return (fp_array, mol_valid, canonical_smiles). Invalid SMILES -> all-zero vector."""
    from rdkit import Chem

    smiles = normalize_smiles(smiles)
    if not smiles:
        return np.zeros(n_bits, dtype=np.int8), False, ""

    mol = Chem.MolFromSmiles(smiles)
    if mol is None:
        return np.zeros(n_bits, dtype=np.int8), False, smiles

    fp = Chem.RDKFingerprint(mol, fpSize=n_bits)
    fp_array = np.fromiter(fp.ToBitString(), dtype=np.int8).astype(np.int8)
    canonical = Chem.MolToSmiles(mol, canonical=True)
    return fp_array, True, canonical


def make_rdkit_fp_dataframe(df: pd.DataFrame, smiles_col: str = "SMILES"):
    if smiles_col not in df.columns:
        raise ValueError(f"Input file must contain a '{smiles_col}' column.")

    fp_list = []
    valid_list = []
    canonical_list = []

    for smiles in df[smiles_col].tolist():
        fp_array, is_valid, canonical = rdkit_fp_from_smiles(smiles)
        fp_list.append(fp_array)
        valid_list.append(is_valid)
        canonical_list.append(canonical)

    fp_matrix = np.vstack(fp_list).astype(np.int8)
    fp_df = pd.DataFrame(fp_matrix, columns=RDKIT_COLUMNS)
    return fp_df, valid_list, canonical_list


def detect_rdkit_columns(df: pd.DataFrame):
    """Detect RDKit fingerprint columns flexibly."""
    candidates = [c for c in df.columns if str(c).lower().startswith("rdkit_")]
    if len(candidates) >= N_BITS:
        def key_func(x):
            try:
                return int(str(x).split("_")[-1])
            except Exception:
                return 10**9
        return sorted(candidates, key=key_func)[:N_BITS]

    # fallback for RDKit1, RDKit2 ... style
    candidates = [c for c in df.columns if str(c).lower().startswith("rdkit")]
    numbered = []
    for c in candidates:
        suffix = str(c).lower().replace("rdkit", "").replace("_", "")
        if suffix.isdigit():
            numbered.append(c)
    if len(numbered) >= N_BITS:
        return sorted(numbered, key=lambda x: int(str(x).lower().replace("rdkit", "").replace("_", "")))[:N_BITS]

    return []
