from pathlib import Path
from typing import Dict, Any, List, Tuple
import pandas as pd
import numpy as np
from sklearn.metrics import (
    accuracy_score,
    precision_score,
    recall_score,
    f1_score,
    confusion_matrix,
    classification_report,
    roc_auc_score,
    roc_curve
)
from sklearn.preprocessing import label_binarize
from sklearn.pipeline import Pipeline


def evaluate_single_model(
    pipeline: Pipeline,
    X_train: pd.DataFrame,
    y_train: pd.Series,
    X_test: pd.DataFrame,
    y_test: pd.Series,
    class_names: List[str]
) -> Dict[str, Any]:
    y_pred_train = pipeline.predict(X_train)
    y_pred_test = pipeline.predict(X_test)
    y_prob_test = pipeline.predict_proba(X_test)
    
    train_acc = float(accuracy_score(y_train, y_pred_train))
    test_acc = float(accuracy_score(y_test, y_pred_test))
    
    precision_macro = float(precision_score(y_test, y_pred_test, average="macro", zero_division=0))
    recall_macro = float(recall_score(y_test, y_pred_test, average="macro", zero_division=0))
    f1_macro = float(f1_score(y_test, y_pred_test, average="macro", zero_division=0))
    
    precision_weighted = float(precision_score(y_test, y_pred_test, average="weighted", zero_division=0))
    recall_weighted = float(recall_score(y_test, y_pred_test, average="weighted", zero_division=0))
    f1_weighted = float(f1_score(y_test, y_pred_test, average="weighted", zero_division=0))
    
    train_f1_macro = float(f1_score(y_train, y_pred_train, average="macro", zero_division=0))
    generalization_gap_acc = float(train_acc - test_acc)
    generalization_gap_f1 = float(train_f1_macro - f1_macro)
    
    cm = confusion_matrix(y_test, y_pred_test)
    
    report_dict = classification_report(
        y_test,
        y_pred_test,
        target_names=class_names,
        output_dict=True,
        zero_division=0
    )
    
    report_df = pd.DataFrame(report_dict).T
    
    num_classes = len(class_names)
    y_test_bin = label_binarize(y_test, classes=list(range(num_classes)))
    
    roc_auc_ovr_macro = float(roc_auc_score(y_test, y_prob_test, multi_class="ovr", average="macro"))
    roc_auc_ovr_weighted = float(roc_auc_score(y_test, y_prob_test, multi_class="ovr", average="weighted"))
    
    roc_curve_data = {}
    for i in range(num_classes):
        fpr, tpr, thresholds = roc_curve(y_test_bin[:, i], y_prob_test[:, i])
        class_auc = float(roc_auc_score(y_test_bin[:, i], y_prob_test[:, i]))
        roc_curve_data[class_names[i]] = {
            "fpr": fpr.tolist(),
            "tpr": tpr.tolist(),
            "auc": class_auc
        }
        
    return {
        "train_accuracy": train_acc,
        "test_accuracy": test_acc,
        "precision_macro": precision_macro,
        "recall_macro": recall_macro,
        "f1_macro": f1_macro,
        "precision_weighted": precision_weighted,
        "recall_weighted": recall_weighted,
        "f1_weighted": f1_weighted,
        "train_f1_macro": train_f1_macro,
        "generalization_gap_accuracy": generalization_gap_acc,
        "generalization_gap_f1": generalization_gap_f1,
        "confusion_matrix": cm,
        "classification_report_dict": report_dict,
        "classification_report_df": report_df,
        "roc_auc_ovr_macro": roc_auc_ovr_macro,
        "roc_auc_ovr_weighted": roc_auc_ovr_weighted,
        "roc_curve_data": roc_curve_data,
        "y_pred_test": y_pred_test,
        "y_prob_test": y_prob_test
    }


def evaluate_all_models(
    trained_models: Dict[str, Pipeline],
    X_train: pd.DataFrame,
    y_train: pd.Series,
    X_test: pd.DataFrame,
    y_test: pd.Series,
    class_names: List[str]
) -> Tuple[Dict[str, Dict[str, Any]], pd.DataFrame]:
    evaluation_results = {}
    summary_rows = []
    
    for model_name, pipeline in trained_models.items():
        results = evaluate_single_model(
            pipeline=pipeline,
            X_train=X_train,
            y_train=y_train,
            X_test=X_test,
            y_test=y_test,
            class_names=class_names
        )
        evaluation_results[model_name] = results
        
        summary_rows.append({
            "Model": model_name,
            "Train Accuracy": results["train_accuracy"],
            "Test Accuracy": results["test_accuracy"],
            "Precision (Macro)": results["precision_macro"],
            "Recall (Macro)": results["recall_macro"],
            "F1-Score (Macro)": results["f1_macro"],
            "F1-Score (Weighted)": results["f1_weighted"],
            "ROC-AUC (OvR Macro)": results["roc_auc_ovr_macro"],
            "Generalization Gap (Acc)": results["generalization_gap_accuracy"]
        })
        
    summary_df = pd.DataFrame(summary_rows)
    return evaluation_results, summary_df


def extract_misclassifications(
    pipeline: Pipeline,
    X_test: pd.DataFrame,
    y_test: pd.Series,
    class_names: List[str]
) -> pd.DataFrame:
    y_pred = pipeline.predict(X_test)
    y_prob = pipeline.predict_proba(X_test)
    
    misclassified_mask = (y_pred != y_test)
    if not misclassified_mask.any():
        return pd.DataFrame()
        
    error_indices = X_test[misclassified_mask].index
    error_df = X_test.loc[error_indices].copy()
    
    error_df["true_class_id"] = y_test.loc[error_indices]
    error_df["true_class_name"] = [class_names[i] for i in y_test.loc[error_indices]]
    error_df["predicted_class_id"] = y_pred[misclassified_mask]
    error_df["predicted_class_name"] = [class_names[i] for i in y_pred[misclassified_mask]]
    
    for i, name in enumerate(class_names):
        error_df[f"prob_{name}"] = y_prob[misclassified_mask, i]
        
    return error_df


def save_tables(
    descriptive_stats: pd.DataFrame,
    summary_df: pd.DataFrame,
    evaluation_results: Dict[str, Dict[str, Any]],
    output_dir: Path
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    
    descriptive_stats.to_csv(output_dir / "descriptive_statistics.csv")
    summary_df.to_csv(output_dir / "model_comparison_metrics.csv", index=False)
    
    for model_name, res in evaluation_results.items():
        clean_name = model_name.lower().replace(" ", "_")
        report_path = output_dir / f"classification_report_{clean_name}.csv"
        res["classification_report_df"].to_csv(report_path)
        
        cm_df = pd.DataFrame(res["confusion_matrix"])
        cm_path = output_dir / f"confusion_matrix_{clean_name}.csv"
        cm_df.to_csv(cm_path, index=False)
