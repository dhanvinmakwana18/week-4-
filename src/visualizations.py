from pathlib import Path
from typing import Dict, Any, List
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import pandas as pd


def set_plotting_style() -> None:
    sns.set_theme(style="whitegrid", font_scale=1.1)
    plt.rcParams["font.sans-serif"] = ["DejaVu Sans", "Arial", "Helvetica"]
    plt.rcParams["axes.edgecolor"] = "#333333"
    plt.rcParams["axes.linewidth"] = 0.8


def plot_confusion_matrix(
    confusion_mat: np.ndarray,
    class_names: List[str],
    model_name: str,
    output_path: Path
) -> None:
    set_plotting_style()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    fig, ax = plt.subplots(figsize=(6.5, 5.2), dpi=300)
    
    cm_sum = np.sum(confusion_mat, axis=1, keepdims=True)
    cm_perc = np.divide(
        confusion_mat.astype(float),
        cm_sum,
        out=np.zeros_like(confusion_mat, dtype=float),
        where=cm_sum != 0
    ) * 100
    
    annot = np.empty_like(confusion_mat, dtype=object)
    for i in range(confusion_mat.shape[0]):
        for j in range(confusion_mat.shape[1]):
            count = confusion_mat[i, j]
            perc = cm_perc[i, j]
            annot[i, j] = f"{count}\n({perc:.1f}%)"
            
    sns.heatmap(
        confusion_mat,
        annot=annot,
        fmt="",
        cmap="Blues",
        cbar=True,
        xticklabels=class_names,
        yticklabels=class_names,
        ax=ax,
        linewidths=1.0,
        linecolor="#e0e0e0",
        annot_kws={"size": 11, "weight": "bold"}
    )
    
    ax.set_title(f"Figure 1: Confusion Matrix – {model_name}", fontsize=13, weight="bold", pad=15)
    ax.set_xlabel("Predicted Class", fontsize=11, weight="bold", labelpad=10)
    ax.set_ylabel("True Class", fontsize=11, weight="bold", labelpad=10)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300)
    plt.close(fig)


def plot_model_comparison(
    summary_df: pd.DataFrame,
    output_path: Path
) -> None:
    set_plotting_style()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    metrics_to_plot = ["Train Accuracy", "Test Accuracy", "Precision (Macro)", "Recall (Macro)", "F1-Score (Macro)"]
    
    melted_df = summary_df.melt(
        id_vars=["Model"],
        value_vars=metrics_to_plot,
        var_name="Metric",
        value_name="Score"
    )
    
    fig, ax = plt.subplots(figsize=(10, 5.8), dpi=300)
    
    palette = ["#1f77b4", "#2ca02c", "#ff7f0e"]
    barplot = sns.barplot(
        data=melted_df,
        x="Metric",
        y="Score",
        hue="Model",
        palette=palette,
        ax=ax,
        edgecolor="#333333",
        linewidth=0.8
    )
    
    ax.set_title("Figure 2: Machine Learning Model Performance Comparison", fontsize=13, weight="bold", pad=15)
    ax.set_xlabel("Evaluation Metric", fontsize=11, weight="bold", labelpad=10)
    ax.set_ylabel("Score (0.0 to 1.0)", fontsize=11, weight="bold", labelpad=10)
    ax.set_ylim(0.80, 1.05)
    
    for p in barplot.patches:
        height = p.get_height()
        if not np.isnan(height) and height > 0:
            ax.annotate(
                f"{height:.3f}",
                (p.get_x() + p.get_width() / 2.0, height),
                ha="center",
                va="bottom",
                fontsize=8.5,
                weight="bold",
                xytext=(0, 3),
                textcoords="offset points"
            )
            
    ax.legend(title="Algorithm", frameon=True, framealpha=0.9, loc="lower right")
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300)
    plt.close(fig)


def plot_multiclass_roc_curve(
    roc_curve_data: Dict[str, Dict[str, Any]],
    macro_auc: float,
    model_name: str,
    output_path: Path
) -> None:
    set_plotting_style()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    fig, ax = plt.subplots(figsize=(7.5, 5.8), dpi=300)
    
    colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728"]
    
    for idx, (class_name, data) in enumerate(roc_curve_data.items()):
        fpr = data["fpr"]
        tpr = data["tpr"]
        auc_val = data["auc"]
        color = colors[idx % len(colors)]
        ax.plot(
            fpr,
            tpr,
            color=color,
            lw=2.2,
            label=f"{class_name} (AUC = {auc_val:.4f})"
        )
        
    ax.plot([0, 1], [0, 1], color="#888888", lw=1.5, linestyle="--", label="Random Chance (AUC = 0.5000)")
    
    ax.set_xlim([-0.02, 1.02])
    ax.set_ylim([-0.02, 1.05])
    ax.set_title(
        f"Figure 3: Multiclass One-vs-Rest ROC Curves – {model_name}\n(Macro-Average ROC-AUC = {macro_auc:.4f})",
        fontsize=12,
        weight="bold",
        pad=15
    )
    ax.set_xlabel("False Positive Rate (1 - Specificity)", fontsize=11, weight="bold", labelpad=10)
    ax.set_ylabel("True Positive Rate (Sensitivity)", fontsize=11, weight="bold", labelpad=10)
    ax.legend(loc="lower right", frameon=True, framealpha=0.95, fontsize=9.5)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300)
    plt.close(fig)


def plot_feature_importance(
    rf_pipeline: Any,
    feature_names: List[str],
    output_path: Path
) -> None:
    set_plotting_style()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    classifier = rf_pipeline.named_steps.get("classifier", rf_pipeline)
    importances = classifier.feature_importances_
    
    feat_df = pd.DataFrame({
        "Feature": feature_names,
        "Importance": importances
    }).sort_values("Importance", ascending=True)
    
    fig, ax = plt.subplots(figsize=(8, 6), dpi=300)
    bars = ax.barh(feat_df["Feature"], feat_df["Importance"], color="#2b5c8f", edgecolor="#1a365d")
    
    ax.set_title("Figure 4: Random Forest Feature Importance (Mean Decrease in Impurity)", fontsize=12, weight="bold", pad=12)
    ax.set_xlabel("Feature Importance Score", fontsize=10.5, weight="bold", labelpad=8)
    ax.set_ylabel("Physicochemical Feature", fontsize=10.5, weight="bold", labelpad=8)
    
    for bar in bars:
        width = bar.get_width()
        ax.text(
            width + 0.003,
            bar.get_y() + bar.get_height() / 2.0,
            f"{width:.3f}",
            ha="left",
            va="center",
            fontsize=8.5,
            weight="bold"
        )
        
    ax.set_xlim(0, max(feat_df["Importance"]) * 1.15)
    plt.tight_layout()
    plt.savefig(output_path, dpi=300)
    plt.close(fig)
