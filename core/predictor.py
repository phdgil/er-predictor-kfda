from __future__ import annotations

import os
import tempfile
import zipfile
from dataclasses import dataclass
from typing import Optional

import numpy as np
import pandas as pd

from core.fingerprint import make_rdkit_fp_dataframe, rdkit_fp_from_smiles


def replace_labels(label):
    if pd.isna(label):
        return label
    if isinstance(label, str):
        v = label.strip().lower()
        if v == "active":
            return 1
        if v == "inactive":
            return 0
    return label


@dataclass
class PredictionResult:
    smiles: str
    canonical_smiles: str
    mol_valid: bool
    probability_inactive: float
    probability_active: float
    prediction: int
    prediction_label: str
    decision_rule: str


class KerasPredictor:
    def __init__(self):
        self.model = None
        self.model_path: Optional[str] = None
        self.model_name: str = ""
        self.input_shape = None

    def load_model(self, model_path: str):
        import tensorflow.keras as keras

        if not os.path.exists(model_path):
            raise FileNotFoundError(f"Model file not found: {model_path}")
        try:
            self.model = keras.models.load_model(model_path, compile=False)
        except Exception as first_error:
            self.model = self._load_erta_model_fallback(model_path, first_error)
        self.model_path = model_path
        self.model_name = os.path.splitext(os.path.basename(model_path))[0]
        self.input_shape = self.model.input_shape
        return self

    def _load_erta_model_fallback(self, model_path: str, first_error: Exception):
        """Load the bundled ERTA Keras 3 model in older tensorflow.keras runtimes.

        Some build environments deserialize the saved InputLayer with unsupported
        Keras 3 keys such as batch_shape/optional. The prediction architecture is
        fixed, so rebuild it and import the weights from model.weights.h5.
        """
        try:
            import h5py
            import tensorflow.keras as keras
        except Exception as import_error:
            raise RuntimeError(f"Model load failed: {first_error}") from import_error

        model = keras.Sequential(name="erta_dnn_binary")
        model.add(keras.layers.InputLayer(input_shape=(2048,), name="input_layer_14"))

        reg = keras.regularizers.l2(0.0001)
        for units, idx in [(256, 70), (128, 71), (64, 72), (32, 73)]:
            model.add(keras.layers.Dense(units, use_bias=False, kernel_regularizer=reg, name=f"dense_{idx}"))
            model.add(keras.layers.BatchNormalization(name=f"batch_normalization_{idx - 14}"))
            model.add(keras.layers.Activation("relu", name=f"activation_{idx - 14}"))
            model.add(keras.layers.Dropout(0.25, name=f"dropout_{idx - 14}"))
        model.add(keras.layers.Dense(2, activation="softmax", name="dense_74"))

        with tempfile.TemporaryDirectory() as tmp_dir:
            with zipfile.ZipFile(model_path, "r") as zf:
                zf.extract("model.weights.h5", tmp_dir)
            weights_path = os.path.join(tmp_dir, "model.weights.h5")
            try:
                model.load_weights(weights_path)
            except Exception:
                self._load_keras3_h5_weights_by_order(model, weights_path)

        return model

    @staticmethod
    def _read_h5_var(group, index: int):
        return group["vars"][str(index)][()]

    def _load_keras3_h5_weights_by_order(self, model, weights_path: str):
        import h5py

        with h5py.File(weights_path, "r") as h5:
            layer_groups = h5.get("layers")
            if layer_groups is None:
                raise RuntimeError("model.weights.h5 does not contain a 'layers' group.")

            dense_groups = sorted(
                [name for name in layer_groups.keys() if name.startswith("dense")],
                key=lambda x: (x != "dense", x),
            )
            bn_groups = sorted(
                [name for name in layer_groups.keys() if name.startswith("batch_normalization")],
                key=lambda x: (x != "batch_normalization", x),
            )

            dense_layers = [layer for layer in model.layers if layer.__class__.__name__ == "Dense"]
            bn_layers = [layer for layer in model.layers if layer.__class__.__name__ == "BatchNormalization"]

            if len(dense_groups) != len(dense_layers) or len(bn_groups) != len(bn_layers):
                raise RuntimeError("Saved model weights do not match the ERTA DNN architecture.")

            for layer, group_name in zip(dense_layers, dense_groups):
                group = layer_groups[group_name]
                weights = [self._read_h5_var(group, 0)]
                if layer.use_bias:
                    weights.append(self._read_h5_var(group, 1))
                layer.set_weights(weights)

            for layer, group_name in zip(bn_layers, bn_groups):
                group = layer_groups[group_name]
                layer.set_weights([
                    self._read_h5_var(group, 0),
                    self._read_h5_var(group, 1),
                    self._read_h5_var(group, 2),
                    self._read_h5_var(group, 3),
                ])

    def is_loaded(self) -> bool:
        return self.model is not None

    def _predict_array(self, x: np.ndarray):
        if self.model is None:
            raise RuntimeError("Model is not loaded.")
        expected_features = self.model.input_shape[-1]
        if expected_features is not None and expected_features != x.shape[1]:
            raise ValueError(
                f"Model expects {expected_features} features, but input has {x.shape[1]} features."
            )
        pred_prob = self.model.predict(x.astype(np.float32), batch_size=64, verbose=0)
        if pred_prob.ndim == 2 and pred_prob.shape[1] == 2:
            prob_neg = pred_prob[:, 0]
            prob_pos = pred_prob[:, 1]
        elif pred_prob.ndim == 2 and pred_prob.shape[1] == 1:
            prob_pos = pred_prob[:, 0]
            prob_neg = 1.0 - prob_pos
        else:
            raise ValueError(f"Unexpected prediction output shape: {pred_prob.shape}")
        return prob_neg, prob_pos

    def predict_smiles(self, smiles: str) -> PredictionResult:
        fp, valid, canonical = rdkit_fp_from_smiles(smiles)
        x = fp.reshape(1, -1).astype(np.float32)
        prob_neg, prob_pos = self._predict_array(x)
        pred = int(prob_pos[0] >= prob_neg[0])
        return PredictionResult(
            smiles=str(smiles or "").strip(),
            canonical_smiles=canonical,
            mol_valid=bool(valid),
            probability_inactive=float(prob_neg[0]),
            probability_active=float(prob_pos[0]),
            prediction=pred,
            prediction_label="Positive" if pred == 1 else "Negative",
            decision_rule="higher_probability",
        )

    def predict_dataframe(self, df: pd.DataFrame, smiles_col: str = "SMILES"):
        if smiles_col not in df.columns:
            raise ValueError(f"Input Excel must contain '{smiles_col}' column.")
        fp_df, mol_valid_list, canonical_list = make_rdkit_fp_dataframe(df, smiles_col=smiles_col)
        prob_neg, prob_pos = self._predict_array(fp_df.values.astype(np.float32))
        pred = (prob_pos >= prob_neg).astype(int)

        result_df = pd.DataFrame()
        for col in ["No.", "No", "CID", "CAS", "Chemical Name", "Name"]:
            if col in df.columns:
                result_df[col] = df[col]
        result_df["SMILES"] = df[smiles_col]
        result_df["Canonical_SMILES"] = canonical_list
        for col in ["label", "result"]:
            if col in df.columns:
                result_df[col] = [replace_labels(x) for x in df[col]]
        result_df["Mol_valid"] = mol_valid_list
        result_df["Probability_Negative_0"] = prob_neg
        result_df["Probability_Positive_1"] = prob_pos
        result_df["Prediction"] = pred
        result_df["Prediction_label"] = ["Positive" if x == 1 else "Negative" for x in pred]
        result_df["Decision_rule"] = "higher_probability"
        return result_df, fp_df
