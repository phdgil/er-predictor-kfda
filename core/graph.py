from __future__ import annotations

import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


def _ensure(out_dir):
    os.makedirs(out_dir, exist_ok=True)


def _positive_probability_column(result_df: pd.DataFrame) -> str:
    if "Probability_Positive_1" in result_df.columns:
        return "Probability_Positive_1"
    return "Probability_Active_1"


def _add_knn_domain_boundary(ax, ad_calculator, z_train, z_pred):
    if not ad_calculator or not ad_calculator.fitted or ad_calculator.pca is None or ad_calculator.nn is None:
        return

    all_z = np.vstack([z_train, z_pred])
    x_min, x_max = all_z[:, 0].min(), all_z[:, 0].max()
    y_min, y_max = all_z[:, 1].min(), all_z[:, 1].max()
    x_pad = max((x_max - x_min) * 0.12, 1.0)
    y_pad = max((y_max - y_min) * 0.12, 1.0)
    xx, yy = np.meshgrid(
        np.linspace(x_min - x_pad, x_max + x_pad, 55),
        np.linspace(y_min - y_pad, y_max + y_pad, 55),
    )
    grid_z = np.c_[xx.ravel(), yy.ravel()]
    grid_scaled = ad_calculator.pca.inverse_transform(grid_z)
    distances, _ = ad_calculator.nn.kneighbors(grid_scaled)
    mean_distance = distances.mean(axis=1).reshape(xx.shape)

    ax.contourf(xx, yy, mean_distance <= ad_calculator.threshold, levels=[0.5, 1.5], colors=["#d8f5df"], alpha=0.35)
    ax.contour(xx, yy, mean_distance, levels=[ad_calculator.threshold], colors=["#1a7f37"], linewidths=1.6)


def _add_ad_decision_scatter(ax, ad_calculator, mean_distance, max_similarity, in_domain, point_size=42, title="AD decision plot"):
    colors = np.where(in_domain, "#1a7f37", "#cf222e")
    ax.scatter(mean_distance, max_similarity, s=point_size, c=colors, alpha=0.82, edgecolors="white", linewidths=0.5)
    ax.axvline(
        ad_calculator.threshold,
        color="#6f42c1",
        linestyle="--",
        linewidth=1.5,
        label="kNN 95% distance threshold",
    )
    ax.axhline(
        ad_calculator.similarity_threshold,
        color="#d97706",
        linestyle="--",
        linewidth=1.5,
        label="Similarity cutoff",
    )
    ax.fill_betweenx([ad_calculator.similarity_threshold, 1.0], 0, ad_calculator.threshold, color="#d8f5df", alpha=0.35)
    ax.set_xlabel("kNN mean distance")
    ax.set_ylabel("Max similarity")
    ax.set_title(title)
    ax.set_ylim(-0.03, 1.03)
    ax.set_xlim(left=0)
    ax.grid(color="#d0d7de", linewidth=0.6, alpha=0.6)
    ax.legend(loc="best", fontsize=8)
    ax.text(
        0.5,
        -0.22,
        f"In-domain: distance <= {ad_calculator.threshold:.3f} and similarity >= {ad_calculator.similarity_threshold:.2f}",
        transform=ax.transAxes,
        fontsize=8,
        ha="center",
        va="top",
    )


def save_class_count_plot(result_df: pd.DataFrame, out_dir: str) -> str:
    _ensure(out_dir)
    counts = result_df["Prediction_label"].value_counts().reindex(["Negative", "Positive"]).fillna(0)
    fig, ax = plt.subplots(figsize=(5, 3.5))
    ax.bar(counts.index, counts.values)
    ax.set_ylabel("Count")
    ax.set_title("Prediction class count")
    for i, v in enumerate(counts.values):
        ax.text(i, v, str(int(v)), ha="center", va="bottom")
    fig.tight_layout()
    path = os.path.join(out_dir, "prediction_class_count.png")
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path


def save_probability_histogram(result_df: pd.DataFrame, out_dir: str) -> str:
    _ensure(out_dir)
    prob_col = _positive_probability_column(result_df)
    fig, ax = plt.subplots(figsize=(5, 3.5))
    ax.hist(result_df[prob_col].dropna(), bins=20)
    ax.set_xlabel("Probability_Positive_1")
    ax.set_ylabel("Count")
    ax.set_title("Positive probability distribution")
    fig.tight_layout()
    path = os.path.join(out_dir, "positive_probability_histogram.png")
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path


def save_top_positive_plot(result_df: pd.DataFrame, out_dir: str, top_n: int = 10) -> str:
    _ensure(out_dir)
    prob_col = _positive_probability_column(result_df)
    df = result_df.copy().sort_values(prob_col, ascending=False).head(top_n)
    labels = []
    for _, row in df.iterrows():
        label = str(row.get("CAS", "") or row.get("CID", "") or row.get("No.", ""))
        if not label or label == "nan":
            label = str(row.get("SMILES", ""))[:20]
        labels.append(label)
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.barh(labels[::-1], df[prob_col].values[::-1])
    ax.set_xlabel("Probability_Positive_1")
    ax.set_title(f"Top {len(df)} positive chemicals")
    ax.set_xlim(0, 1)
    fig.tight_layout()
    path = os.path.join(out_dir, "top_positive_chemicals.png")
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path


def save_ad_plot(ad_calculator, x_pred, result_df: pd.DataFrame, out_dir: str) -> str:
    _ensure(out_dir)
    if not ad_calculator or not ad_calculator.fitted:
        raise RuntimeError("AD calculator is not fitted.")
    z_pred, mean_distance, in_domain, _, _, _ = ad_calculator.transform_with_similarity(x_pred)
    z_train = ad_calculator.train_z
    fig, ax = plt.subplots(figsize=(5.5, 4))
    _add_knn_domain_boundary(ax, ad_calculator, z_train, z_pred)
    ax.scatter(z_train[:, 0], z_train[:, 1], s=8, alpha=0.32, c="#8c959f", label="Training reference")
    ax.scatter(z_pred[in_domain, 0], z_pred[in_domain, 1], s=34, marker="o", c="#1a7f37", edgecolors="white", linewidths=0.4, label="Batch predictions in-domain")
    ax.scatter(z_pred[~in_domain, 0], z_pred[~in_domain, 1], s=46, marker="x", c="#cf222e", linewidths=1.4, label="Batch predictions out-of-domain")
    ax.set_xlabel("PC1")
    ax.set_ylabel("PC2")
    ax.set_title(f"Batch PCA AD plot (n={len(z_pred)})")
    ax.legend(loc="best", fontsize=8)
    ax.text(0.01, 0.01, f"Mean distance range: {mean_distance.min():.3f}-{mean_distance.max():.3f}", transform=ax.transAxes, fontsize=8, va="bottom")
    fig.tight_layout()
    path = os.path.join(out_dir, "ad_pca_plot.png")
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path


def save_ad_decision_plot(ad_calculator, x_pred, result_df: pd.DataFrame, out_dir: str) -> str:
    _ensure(out_dir)
    if not ad_calculator or not ad_calculator.fitted:
        raise RuntimeError("AD calculator is not fitted.")
    _, mean_distance, in_domain, max_similarity, distance_in_domain, similarity_in_domain = ad_calculator.transform_with_similarity(x_pred)

    fig, ax = plt.subplots(figsize=(5.5, 4.3))
    _add_ad_decision_scatter(ax, ad_calculator, mean_distance, max_similarity, in_domain, point_size=42, title="AD decision plot")
    fig.tight_layout(rect=[0, 0.08, 1, 1])
    path = os.path.join(out_dir, "ad_decision_plot.png")
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path


def save_single_ad_plot(ad_calculator, x_pred, out_dir: str) -> str:
    _ensure(out_dir)
    if not ad_calculator or not ad_calculator.fitted:
        raise RuntimeError("AD calculator is not fitted.")
    z_pred, mean_distance, in_domain, max_similarity, distance_in_domain, similarity_in_domain = ad_calculator.transform_with_similarity(x_pred)
    z_train = ad_calculator.train_z
    fig, axes = plt.subplots(1, 2, figsize=(8.4, 4.1), gridspec_kw={"width_ratios": [1.15, 1.0]})
    ax = axes[0]
    _add_knn_domain_boundary(ax, ad_calculator, z_train, z_pred)
    ax.scatter(z_train[:, 0], z_train[:, 1], s=8, alpha=0.32, c="#8c959f", label="Training reference")
    color = "#1a7f37" if bool(in_domain[0]) else "#cf222e"
    marker = "o" if bool(in_domain[0]) else "x"
    label = "Single prediction in-domain" if bool(in_domain[0]) else "Single prediction out-of-domain"
    ax.scatter(z_pred[:, 0], z_pred[:, 1], s=90, marker=marker, c=color, edgecolors="white" if marker == "o" else color, linewidths=0.7 if marker == "o" else 1.5, label=label)
    ax.set_xlabel("PC1")
    ax.set_ylabel("PC2")
    ax.set_title(f"PCA AD plot: similarity >= {ad_calculator.similarity_threshold:.2f}")
    ax.legend(loc="best", fontsize=8)
    ax.text(0.01, 0.01, f"Mean distance={mean_distance[0]:.3f}", transform=ax.transAxes, fontsize=8, va="bottom")
    _add_ad_decision_scatter(
        axes[1],
        ad_calculator,
        mean_distance,
        max_similarity,
        in_domain,
        point_size=90,
        title="AD decision plot",
    )
    fig.tight_layout(rect=[0, 0.08, 1, 1])
    path = os.path.join(out_dir, "single_ad_pca_plot.png")
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path


def save_all_batch_graphs(result_df: pd.DataFrame, out_dir: str, ad_calculator=None, fp_df=None):
    paths = []
    paths.append(save_class_count_plot(result_df, out_dir))
    paths.append(save_probability_histogram(result_df, out_dir))
    paths.append(save_top_positive_plot(result_df, out_dir))
    if ad_calculator is not None and ad_calculator.fitted and fp_df is not None:
        paths.append(save_ad_plot(ad_calculator, fp_df.values.astype(float), result_df, out_dir))
        paths.append(save_ad_decision_plot(ad_calculator, fp_df.values.astype(float), result_df, out_dir))
    return paths
