from __future__ import annotations

import hashlib
import json
import os
import tempfile
import threading
from dataclasses import dataclass
from typing import Optional

import numpy as np
import pandas as pd
from sklearn.decomposition import PCA
from sklearn.neighbors import NearestNeighbors
from sklearn.preprocessing import StandardScaler

from core.fingerprint import detect_rdkit_columns

try:
    from rdkit import Chem, DataStructs
    from rdkit.Chem import rdMolDescriptors
    RDKIT_AVAILABLE = True
except Exception:
    RDKIT_AVAILABLE = False
    Chem = None
    DataStructs = None
    rdMolDescriptors = None


@dataclass
class ADResult:
    fitted: bool
    in_domain: Optional[bool] = None
    pca_in_domain: Optional[bool] = None
    distance_in_domain: Optional[bool] = None
    similarity_in_domain: Optional[bool] = None
    nearest_index: Optional[int] = None
    nearest_reference: Optional[dict] = None
    display_nearest_index: Optional[int] = None
    display_nearest_reference: Optional[dict] = None
    distance: Optional[float] = None
    threshold: Optional[float] = None
    max_similarity: Optional[float] = None
    display_similarity: Optional[float] = None
    similarity_threshold: Optional[float] = None
    pc1: Optional[float] = None
    pc2: Optional[float] = None
    message: str = ""


class ADCalculator:
    METHOD = "knn_mean_distance_p95_rdkit_ad_morgan_display"

    def __init__(
        self,
        coverage: float = 0.95,
        robust: bool = False,
        empirical: bool = True,
        similarity_threshold: float = 0.60,
        n_neighbors: int = 5,
        cache_root: str | None = None,
    ):
        self.coverage = coverage
        self.robust = robust
        self.empirical = empirical
        self.similarity_threshold = similarity_threshold
        self.n_neighbors = n_neighbors
        self.cache_root = cache_root
        self.pca = None
        self.scaler = None
        self.nn = None
        self.threshold = None
        self.train_z = None
        self.train_x_scaled = None
        self.train_fp_bool = None
        self.train_bit_counts = None
        self.train_morgan_fps = []
        self.train_morgan_fp_bool = None
        self.train_metadata = []
        self.fitted = False
        self.reference_path = None
        self.last_cache_path = None
        self.last_cache_status = ""
        self._load_condition = threading.Condition()
        self._load_in_progress = False
        self._load_error = None

    CACHE_SCHEMA_VERSION = 2

    def default_cache_path(self, reference_path: str) -> str:
        if not self.cache_root:
            return self.adjacent_cache_path(reference_path)
        reference = os.path.abspath(reference_path)
        digest = hashlib.sha256(reference.encode("utf-8")).hexdigest()[:16]
        return os.path.join(self.cache_root, f"erta_ad_{digest}.npz")

    @staticmethod
    def adjacent_cache_path(reference_path: str) -> str:
        base, _ = os.path.splitext(reference_path)
        return f"{base}_ad_cache.npz"

    @staticmethod
    def _source_sha256(reference_path: str) -> str:
        digest = hashlib.sha256()
        with open(reference_path, "rb") as source:
            for block in iter(lambda: source.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()

    def _append_cache_status(self, message: str) -> None:
        self.last_cache_status = (
            f"{self.last_cache_status}; {message}" if self.last_cache_status else message
        )

    def _cache_metadata(self) -> dict:
        if not self.reference_path or not os.path.isfile(self.reference_path):
            raise RuntimeError("AD cache requires an existing reference source file.")
        return {
            "schema_version": self.CACHE_SCHEMA_VERSION,
            "method": self.METHOD,
            "coverage": self.coverage,
            "robust": self.robust,
            "empirical": self.empirical,
            "similarity_threshold": self.similarity_threshold,
            "n_neighbors": self.n_neighbors,
            "reference_path": self.reference_path,
            "source_sha256": self._source_sha256(self.reference_path),
        }

    @staticmethod
    def _cache_array(cache, name: str, ndim: int | None = None) -> np.ndarray:
        if name not in cache.files:
            raise ValueError(f"AD cache schema is missing {name}.")
        value = np.asarray(cache[name])
        if value.dtype.hasobject or (ndim is not None and value.ndim != ndim):
            raise ValueError(f"AD cache schema has an invalid {name}.")
        return value
    @staticmethod
    def _clean_matrix(x: np.ndarray) -> np.ndarray:
        return np.nan_to_num(x.astype(float), nan=0.0, posinf=0.0, neginf=0.0)

    def save_cache(self, cache_path: str):
        if not self.fitted:
            raise RuntimeError("Cannot cache AD before fitting.")
        if cache_path.lower().endswith(".pkl"):
            raise ValueError("Refusing legacy pickle AD cache path.")
        metadata = self._cache_metadata()
        if self.scaler is None or self.train_x_scaled is None or self.train_fp_bool is None:
            raise RuntimeError("Cannot cache incomplete AD state.")
        if self.train_morgan_fp_bool is None:
            raise RuntimeError("Cannot cache AD state without Morgan reference fingerprints.")

        directory = os.path.dirname(cache_path) or "."
        os.makedirs(directory, exist_ok=True)
        temp_path = ""
        try:
            with tempfile.NamedTemporaryFile(
                mode="wb", suffix=".npz", prefix=".ad-cache-", dir=directory, delete=False
            ) as temporary:
                temp_path = temporary.name
                np.savez_compressed(
                    temporary,
                    metadata=np.asarray(json.dumps(metadata, sort_keys=True)),
                    threshold=np.asarray(self.threshold, dtype=float),
                    train_x_scaled=np.asarray(self.train_x_scaled, dtype=float),
                    train_fp_bool=np.asarray(self.train_fp_bool, dtype=bool),
                    train_morgan_fp_bool=np.asarray(self.train_morgan_fp_bool, dtype=bool),
                    scaler_mean=np.asarray(self.scaler.mean_, dtype=float),
                    scaler_scale=np.asarray(self.scaler.scale_, dtype=float),
                    scaler_var=np.asarray(self.scaler.var_, dtype=float),
                    scaler_n_samples_seen=np.asarray(self.scaler.n_samples_seen_),
                    train_metadata=np.asarray(json.dumps(self.train_metadata, ensure_ascii=False)),
                )
            os.replace(temp_path, cache_path)
        except Exception:
            if temp_path and os.path.exists(temp_path):
                os.unlink(temp_path)
            raise

        self.last_cache_path = cache_path
        self._append_cache_status(f"AD cache saved: {cache_path}")
        return cache_path

    def load_cache(self, cache_path: str, reference_path: str = "", validate_source: bool = True):
        if not os.path.exists(cache_path):
            raise FileNotFoundError(cache_path)
        if cache_path.lower().endswith(".pkl"):
            raise ValueError("Unsafe legacy pickle AD cache rejected; refitting is required.")

        try:
            with np.load(cache_path, allow_pickle=False) as cache:
                metadata_value = self._cache_array(cache, "metadata", ndim=0)
                train_metadata_value = self._cache_array(cache, "train_metadata", ndim=0)
                try:
                    metadata = json.loads(str(metadata_value.item()))
                    train_metadata = json.loads(str(train_metadata_value.item()))
                except (TypeError, ValueError, json.JSONDecodeError) as exc:
                    raise ValueError("AD cache metadata is not valid JSON.") from exc

                required_metadata = {
                    "schema_version", "method", "coverage", "robust", "empirical",
                    "similarity_threshold", "n_neighbors", "source_sha256",
                }
                if (
                    not isinstance(metadata, dict)
                    or not required_metadata.issubset(metadata)
                    or metadata["schema_version"] != self.CACHE_SCHEMA_VERSION
                    or not isinstance(metadata["source_sha256"], str)
                    or len(metadata["source_sha256"]) != 64
                    or any(character not in "0123456789abcdef" for character in metadata["source_sha256"])
                    or metadata["method"] != self.METHOD
                    or metadata["coverage"] != self.coverage
                    or metadata["robust"] != self.robust
                    or metadata["empirical"] != self.empirical
                    or metadata["similarity_threshold"] != self.similarity_threshold
                    or metadata["n_neighbors"] != self.n_neighbors
                    or not isinstance(train_metadata, list)
                    or not all(isinstance(item, dict) for item in train_metadata)
                ):
                    raise ValueError("AD cache settings or schema do not match current k-NN settings.")

                threshold = self._cache_array(cache, "threshold", ndim=0).astype(float).item()
                train_x_scaled = self._cache_array(cache, "train_x_scaled", ndim=2).astype(float)
                train_fp_bool = self._cache_array(cache, "train_fp_bool", ndim=2).astype(bool)
                train_morgan_fp_bool = self._cache_array(
                    cache, "train_morgan_fp_bool", ndim=2
                ).astype(bool)
                scaler_mean = self._cache_array(cache, "scaler_mean", ndim=1).astype(float)
                scaler_scale = self._cache_array(cache, "scaler_scale", ndim=1).astype(float)
                scaler_var = self._cache_array(cache, "scaler_var", ndim=1).astype(float)
                n_samples_seen = self._cache_array(cache, "scaler_n_samples_seen")
        except (OSError, ValueError, TypeError, json.JSONDecodeError) as exc:
            raise ValueError(f"AD cache rejected: {exc}") from exc

        if (
            not np.isfinite(threshold)
            or not np.isfinite(train_x_scaled).all()
            or not np.isfinite(scaler_mean).all()
            or not np.isfinite(scaler_scale).all()
            or not np.isfinite(scaler_var).all()
            or train_x_scaled.shape[0] == 0
            or train_x_scaled.shape != train_fp_bool.shape
            or train_morgan_fp_bool.shape != (train_x_scaled.shape[0], 2048)
            or train_x_scaled.shape[1] != scaler_mean.size
            or scaler_mean.shape != scaler_scale.shape
            or scaler_mean.shape != scaler_var.shape
            or n_samples_seen.ndim > 1
        ):
            raise ValueError("AD cache schema contains invalid numeric state.")

        if validate_source:
            if not reference_path or not os.path.isfile(reference_path):
                raise ValueError("AD cache source validation requires an existing reference file.")
            if self._source_sha256(reference_path) != metadata["source_sha256"]:
                raise ValueError("AD cache source SHA-256 does not match the reference file.")

        self.scaler = StandardScaler()
        self.scaler.mean_ = scaler_mean
        self.scaler.scale_ = scaler_scale
        self.scaler.var_ = scaler_var
        self.scaler.n_features_in_ = scaler_mean.size
        self.scaler.n_samples_seen_ = n_samples_seen.item() if n_samples_seen.ndim == 0 else n_samples_seen
        self.train_x_scaled = train_x_scaled
        self.train_fp_bool = train_fp_bool
        self.train_bit_counts = self.train_fp_bool.sum(axis=1)
        self.train_metadata = train_metadata
        self.train_morgan_fps = []
        self.train_morgan_fp_bool = train_morgan_fp_bool
        self.threshold = float(threshold)
        self.nn = NearestNeighbors(
            n_neighbors=min(int(self.n_neighbors), len(self.train_x_scaled)), metric="euclidean"
        )
        self.nn.fit(self.train_x_scaled)
        self.pca = PCA(n_components=2, random_state=2026)
        self.train_z = self.pca.fit_transform(self.train_x_scaled)
        self.fitted = True
        self.reference_path = reference_path or metadata.get("reference_path")
        self.last_cache_path = cache_path
        self._append_cache_status(f"AD cache loaded: {cache_path}")
        return self

    def fit_from_excel_cached(self, reference_path: str, cache_path: str | None = None, force_refit: bool = False):
        cache_path = cache_path or self.default_cache_path(reference_path)
        adjacent_cache_path = self.adjacent_cache_path(reference_path)
        self.last_cache_path = cache_path
        self.last_cache_status = ""
        if not force_refit:
            try:
                return self.load_cache(cache_path, reference_path=reference_path)
            except Exception as e:
                self._append_cache_status(f"AD cache not used: {e}")
            if adjacent_cache_path != cache_path and os.path.isfile(adjacent_cache_path):
                try:
                    self.load_cache(adjacent_cache_path, reference_path=reference_path)
                    try:
                        self.save_cache(cache_path)
                    except Exception as e:
                        self._append_cache_status(f"AD state cache save failed: {cache_path} ({e})")
                    return self
                except Exception as e:
                    self._append_cache_status(f"Adjacent AD cache not used: {e}")

        self.fit_from_excel(reference_path)
        try:
            self.save_cache(cache_path)
        except Exception as e:
            self._append_cache_status(f"AD cache save failed: {cache_path} ({e})")
        return self

    def ensure_fitted_from_excel_cached(
        self,
        reference_path: str,
        cache_path: str | None = None,
        force_refit: bool = False,
    ):
        """Load or fit this calculator once while concurrent callers share the outcome."""
        waited_for_load = False
        with self._load_condition:
            while self._load_in_progress:
                waited_for_load = True
                self._load_condition.wait()

            if waited_for_load:
                if self.fitted:
                    return self
                if self._load_error is not None:
                    raise RuntimeError(f"AD reference load failed: {self._load_error}") from self._load_error
                raise RuntimeError("AD reference load did not produce a fitted calculator.")

            if self.fitted and not force_refit:
                return self

            self._load_in_progress = True
            self._load_error = None

        try:
            result = self.fit_from_excel_cached(
                reference_path,
                cache_path=cache_path,
                force_refit=force_refit,
            )
            if not self.fitted:
                raise RuntimeError("AD reference load did not produce a fitted calculator.")
        except Exception as exc:
            with self._load_condition:
                self._load_error = exc
                self._load_in_progress = False
                self._load_condition.notify_all()
            raise
        with self._load_condition:
            self._load_in_progress = False
            self._load_condition.notify_all()
        return result

    def fit_from_excel(self, reference_path: str):
        if not os.path.exists(reference_path):
            raise FileNotFoundError(f"Training reference file not found: {reference_path}")
        df = pd.read_excel(reference_path)
        cols = detect_rdkit_columns(df)
        if len(cols) < 2048:
            raise ValueError("Could not detect 2048 RDKit fingerprint columns in training reference file.")
        x_train = df[cols].replace([float("inf"), -float("inf")], 0).fillna(0).values.astype(float)
        metadata = self._extract_metadata(df, exclude_columns=set(cols))
        return self.fit(x_train, reference_path=reference_path, metadata=metadata)

    @staticmethod
    def _clean_metadata_value(value):
        if pd.isna(value):
            return ""
        text = str(value).strip()
        return "" if text.lower() == "nan" else text

    def _extract_metadata(self, df: pd.DataFrame, exclude_columns: set):
        preferred = [
            "CAS", "CAS No", "CAS No.", "CAS RN", "CASRN", "CAS Number", "CAS_Number",
            "SMILES", "Canonical_SMILES", "Canonical SMILES", "Isomeric_SMILES", "Isomeric SMILES",
            "CID", "PubChem_CID", "Name", "Chemical name", "Chemical_Name", "Substance", "DTXSID",
            "label", "Label", "Activity", "active", "inactive",
        ]
        normalized = {str(col).strip().lower(): col for col in df.columns if col not in exclude_columns}
        selected = []
        for name in preferred:
            col = normalized.get(name.strip().lower())
            if col is not None and col not in selected:
                selected.append(col)
        if not selected:
            selected = [col for col in df.columns if col not in exclude_columns][:6]

        metadata = []
        for idx, row in df.iterrows():
            item = {"Reference row": str(idx + 2)}
            for col in selected:
                value = self._clean_metadata_value(row.get(col, ""))
                if value:
                    item[str(col)] = value
            metadata.append(item)
        return metadata

    def fit(self, x_train: np.ndarray, reference_path: str = "", metadata=None):
        x_train = self._clean_matrix(x_train)
        if len(x_train) == 0:
            raise ValueError("AD reference data is empty.")

        self.train_fp_bool = x_train > 0
        self.train_bit_counts = self.train_fp_bool.sum(axis=1)
        self.train_metadata = metadata or []
        self.train_morgan_fps = self._fit_morgan_fps(self.train_metadata)
        self.train_morgan_fp_bool = self._morgan_bool_matrix(self.train_morgan_fps)

        self.scaler = StandardScaler()
        self.train_x_scaled = self.scaler.fit_transform(x_train)
        n_neighbors = min(int(self.n_neighbors), len(self.train_x_scaled))
        self.nn = NearestNeighbors(n_neighbors=n_neighbors, metric="euclidean")
        self.nn.fit(self.train_x_scaled)

        threshold_neighbors = min(int(self.n_neighbors) + 1, len(self.train_x_scaled))
        threshold_nn = NearestNeighbors(n_neighbors=threshold_neighbors, metric="euclidean")
        threshold_nn.fit(self.train_x_scaled)
        train_distances, _ = threshold_nn.kneighbors(self.train_x_scaled)
        if threshold_neighbors > 1:
            train_distances = train_distances[:, 1:]
        mean_distances_train = train_distances.mean(axis=1)
        self.threshold = float(np.percentile(mean_distances_train, self.coverage * 100.0))

        self.pca = PCA(n_components=2, random_state=2026)
        self.train_z = self.pca.fit_transform(self.train_x_scaled)
        self.fitted = True
        self.reference_path = reference_path
        return self

    def tanimoto_similarity(self, x: np.ndarray, return_index: bool = False):
        if self.train_fp_bool is None:
            raise RuntimeError("AD reference fingerprints are not fitted.")
        x_bool = self._clean_matrix(x) > 0
        max_values = []
        max_indices = []
        for row in x_bool:
            row_count = row.sum()
            intersection = np.logical_and(self.train_fp_bool, row).sum(axis=1)
            union = self.train_bit_counts + row_count - intersection
            sim = np.divide(intersection, union, out=np.zeros_like(intersection, dtype=float), where=union != 0)
            if sim.size:
                idx = int(np.argmax(sim))
                max_values.append(float(sim[idx]))
                max_indices.append(idx)
            else:
                max_values.append(0.0)
                max_indices.append(-1)
        values = np.asarray(max_values, dtype=float)
        indices = np.asarray(max_indices, dtype=int)
        return (values, indices) if return_index else values

    @staticmethod
    def _metadata_smiles(item: dict) -> str:
        for key in ["SMILES", "Canonical_SMILES", "Canonical SMILES", "Isomeric_SMILES", "Isomeric SMILES"]:
            value = item.get(key) if item else ""
            if value:
                return str(value).strip()
        return ""

    @staticmethod
    def _morgan_fp_from_smiles(smiles: str):
        if not RDKIT_AVAILABLE or not smiles:
            return None
        mol = Chem.MolFromSmiles(str(smiles).strip())
        if mol is None:
            return None
        return rdMolDescriptors.GetMorganFingerprintAsBitVect(mol, radius=2, nBits=2048)

    def _fit_morgan_fps(self, metadata):
        fps = []
        for item in metadata or []:
            fps.append(self._morgan_fp_from_smiles(self._metadata_smiles(item)))
        return fps

    @staticmethod
    def _morgan_bool_matrix(fingerprints) -> np.ndarray:
        matrix = np.zeros((len(fingerprints), 2048), dtype=bool)
        if not RDKIT_AVAILABLE:
            return matrix
        for index, fingerprint in enumerate(fingerprints):
            if fingerprint is None:
                continue
            row = np.zeros((2048,), dtype=np.uint8)
            DataStructs.ConvertToNumpyArray(fingerprint, row)
            matrix[index] = row.astype(bool)
        return matrix

    def morgan_nearest_reference(self, smiles: str):
        query_fp = self._morgan_fp_from_smiles(smiles)
        if (
            query_fp is None
            or self.train_morgan_fp_bool is None
            or not len(self.train_morgan_fp_bool)
        ):
            return None, None, None
        query = np.zeros((2048,), dtype=np.uint8)
        DataStructs.ConvertToNumpyArray(query_fp, query)
        query_bool = query.astype(bool)
        intersection = np.logical_and(self.train_morgan_fp_bool, query_bool).sum(axis=1)
        union = self.train_morgan_fp_bool.sum(axis=1) + query_bool.sum() - intersection
        similarities = np.divide(
            intersection,
            union,
            out=np.full_like(intersection, -1.0, dtype=float),
            where=union != 0,
        )
        if not len(similarities) or float(np.max(similarities)) < 0:
            return None, None, None
        best_idx = int(np.argmax(similarities))
        best_sim = float(similarities[best_idx])
        reference = self.train_metadata[best_idx] if best_idx < len(self.train_metadata) else {}
        return best_idx, reference, best_sim

    def predict(self, x: np.ndarray, input_smiles: str = "") -> ADResult:
        if not self.fitted:
            return ADResult(fitted=False, message="AD reference is not fitted.")
        z, mean_distances, in_domain_values, max_similarity_values, distance_in_domain_values, similarity_in_domain_values = self.transform_with_similarity(x)
        similarity_values, similarity_indices = self.tanimoto_similarity(x, return_index=True)
        max_similarity = float(similarity_values[0]) if len(similarity_values) else float(max_similarity_values[0])
        nearest_index = int(similarity_indices[0]) if len(similarity_indices) else -1
        nearest_reference = self.train_metadata[nearest_index] if 0 <= nearest_index < len(self.train_metadata) else {}
        display_nearest_index = nearest_index
        display_nearest_reference = nearest_reference
        display_similarity = max_similarity
        if input_smiles:
            morgan_index, morgan_reference, morgan_similarity = self.morgan_nearest_reference(input_smiles)
            if morgan_reference is not None:
                display_nearest_index = morgan_index
                display_nearest_reference = morgan_reference
                display_similarity = float(morgan_similarity)
        in_domain = bool(in_domain_values[0])
        distance_in_domain = bool(distance_in_domain_values[0])
        similarity_in_domain = bool(similarity_in_domain_values[0])
        return ADResult(
            fitted=True,
            in_domain=in_domain,
            pca_in_domain=distance_in_domain,
            distance_in_domain=distance_in_domain,
            similarity_in_domain=similarity_in_domain,
            nearest_index=nearest_index,
            nearest_reference=nearest_reference,
            display_nearest_index=display_nearest_index,
            display_nearest_reference=display_nearest_reference,
            distance=float(mean_distances[0]),
            threshold=float(self.threshold),
            max_similarity=max_similarity,
            display_similarity=display_similarity,
            similarity_threshold=float(self.similarity_threshold),
            pc1=float(z[0, 0]),
            pc2=float(z[0, 1]),
            message="In-domain" if in_domain else "Out-of-domain",
        )

    def transform(self, x: np.ndarray):
        if not self.fitted:
            raise RuntimeError("AD reference is not fitted.")
        z, mean_distances, in_domain, _, _, _ = self.transform_with_similarity(x)
        return z, mean_distances, in_domain

    def transform_with_similarity(self, x: np.ndarray):
        if not self.fitted:
            raise RuntimeError("AD reference is not fitted.")
        x_clean = self._clean_matrix(x)
        x_scaled = self.scaler.transform(x_clean)
        distances, _ = self.nn.kneighbors(x_scaled)
        mean_distances = distances.mean(axis=1)
        distance_in_domain = mean_distances <= self.threshold
        z = self.pca.transform(x_scaled)
        max_similarity = self.tanimoto_similarity(x_clean)
        similarity_in_domain = max_similarity >= self.similarity_threshold
        in_domain = distance_in_domain & similarity_in_domain
        return z, mean_distances, in_domain, max_similarity, distance_in_domain, similarity_in_domain
