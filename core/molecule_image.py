from __future__ import annotations

import os


def save_molecule_image(smiles: str, output_png: str, size=(320, 240)) -> str:
    from rdkit import Chem
    from rdkit.Chem import Draw

    mol = Chem.MolFromSmiles(str(smiles or ""))
    if mol is None:
        raise ValueError("Invalid SMILES; cannot draw molecule.")
    os.makedirs(os.path.dirname(output_png), exist_ok=True)
    img = Draw.MolToImage(mol, size=size)
    img.save(output_png)
    return output_png
