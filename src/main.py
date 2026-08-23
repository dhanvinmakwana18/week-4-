from pathlib import Path
import sys
import pandas as pd

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.data_loader import load_raw_wine_data, inspect_data, save_dataset
from src.data_preprocessing import prepare_train_test_split
from src.model_training import initialize_candidate_models, train_models
from src.model_evaluation import evaluate_all_models, extract_misclassifications, save_tables
from src.visualizations import (
    plot_confusion_matrix,
    plot_model_comparison,
    plot_multiclass_roc_curve,
    plot_feature_importance
)
from src.report_generator import generate_word_report


def run_pipeline() -> None:
    print("================================================================================")
    print("      MACHINE LEARNING MODEL DEVELOPMENT & EVALUATION PIPELINE (WEEK 4)")
    print("================================================================================")
    
    # 1. Setup Directories
    data_processed_dir = PROJECT_ROOT / "data" / "processed"
    outputs_figures_dir = PROJECT_ROOT / "outputs" / "figures"
    outputs_tables_dir = PROJECT_ROOT / "outputs" / "tables"
    report_dir = PROJECT_ROOT / "report"
    
    for directory in [data_processed_dir, outputs_figures_dir, outputs_tables_dir, report_dir]:
        directory.mkdir(parents=True, exist_ok=True)
        
    # 2. Data Loading & Inspection
    print("\n[Step 1/7] Loading and Inspecting Wine Classification Dataset...")
    df, metadata = load_raw_wine_data()
    inspection_summary = inspect_data(df, target_col=metadata["target_name"])
    
    print(f" -> Dataset shape: {df.shape[0]} rows, {df.shape[1]} columns")
    print(f" -> Missing values: {inspection_summary['total_missing_values']}")
    print(f" -> Duplicate rows: {inspection_summary['duplicate_rows']}")
    print(f" -> Target distribution: {inspection_summary['target_distribution']}")
    
    # Save full dataset
    save_dataset(df, data_processed_dir / "wine_raw.csv")
    
    # 3. Data Preprocessing & Train/Test Split
    print("\n[Step 2/7] Executing Stratified Train/Test Split (80/20, random_state=42)...")
    X_train, X_test, y_train, y_test = prepare_train_test_split(
        df=df,
        target_col=metadata["target_name"],
        test_size=0.20,
        random_state=42
    )
    
    # Save train/test partitions
    train_df = pd.concat([X_train, y_train], axis=1)
    test_df = pd.concat([X_test, y_test], axis=1)
    save_dataset(train_df, data_processed_dir / "train_data.csv")
    save_dataset(test_df, data_processed_dir / "test_data.csv")
    print(f" -> Training samples: {len(X_train)} | Test samples: {len(X_test)}")
    
    # 4. Model Initialization & Training
    print("\n[Step 3/7] Initializing and Training Candidate Classification Models...")
    candidate_models = initialize_candidate_models(random_state=42)
    trained_models = train_models(candidate_models, X_train, y_train)
    for model_name in trained_models.keys():
        print(f" -> Trained: {model_name}")
        
    # 5. Model Evaluation & Comparison
    print("\n[Step 4/7] Evaluating Models on Train and Test Sets...")
    evaluation_results, summary_df = evaluate_all_models(
        trained_models=trained_models,
        X_train=X_train,
        y_train=y_train,
        X_test=X_test,
        y_test=y_test,
        class_names=metadata["target_classes"]
    )
    
    print("\n--- Summary Performance Metrics ---")
    print(summary_df.to_string(index=False))
    
    # Extract misclassifications for error analysis
    lr_errors = extract_misclassifications(
        trained_models["Logistic Regression"],
        X_test,
        y_test,
        metadata["target_classes"]
    )
    rf_errors = extract_misclassifications(
        trained_models["Random Forest"],
        X_test,
        y_test,
        metadata["target_classes"]
    )
    dt_errors = extract_misclassifications(
        trained_models["Decision Tree"],
        X_test,
        y_test,
        metadata["target_classes"]
    )
    
    print(f"\n -> Misclassifications (Test Set):")
    print(f"    * Logistic Regression: {len(lr_errors)} error(s)")
    print(f"    * Random Forest:       {len(rf_errors)} error(s)")
    print(f"    * Decision Tree:       {len(dt_errors)} error(s)")
    
    # 6. Save Tables
    print("\n[Step 5/7] Saving Results Tables to outputs/tables/...")
    save_tables(
        descriptive_stats=inspection_summary["descriptive_statistics"],
        summary_df=summary_df,
        evaluation_results=evaluation_results,
        output_dir=outputs_tables_dir
    )
    if not lr_errors.empty:
        lr_errors.to_csv(outputs_tables_dir / "misclassified_logistic_regression.csv")
    if not dt_errors.empty:
        dt_errors.to_csv(outputs_tables_dir / "misclassified_decision_tree.csv")
    print(f" -> Saved tabular CSV files to: {outputs_tables_dir}")
    
    # 7. Generate Visualizations
    print("\n[Step 6/7] Generating Visualizations in outputs/figures/...")
    fig1_path = outputs_figures_dir / "figure1_confusion_matrix.png"
    fig2_path = outputs_figures_dir / "figure2_model_performance_comparison.png"
    fig3_path = outputs_figures_dir / "figure3_multiclass_roc_curve.png"
    fig4_path = outputs_figures_dir / "figure4_feature_importance.png"
    
    # Figure 1: Confusion matrix of final model (Random Forest)
    plot_confusion_matrix(
        confusion_mat=evaluation_results["Random Forest"]["confusion_matrix"],
        class_names=metadata["target_classes"],
        model_name="Random Forest (Final Champion Model)",
        output_path=fig1_path
    )
    print(f" -> Generated: {fig1_path.name}")
    
    # Figure 2: Model performance comparison bar chart
    plot_model_comparison(
        summary_df=summary_df,
        output_path=fig2_path
    )
    print(f" -> Generated: {fig2_path.name}")
    
    # Figure 3: Multiclass ROC curves for Logistic Regression
    plot_multiclass_roc_curve(
        roc_curve_data=evaluation_results["Logistic Regression"]["roc_curve_data"],
        macro_auc=evaluation_results["Logistic Regression"]["roc_auc_ovr_macro"],
        model_name="Logistic Regression",
        output_path=fig3_path
    )
    print(f" -> Generated: {fig3_path.name}")
    
    # Figure 4: Random Forest Feature Importance
    plot_feature_importance(
        rf_pipeline=trained_models["Random Forest"],
        feature_names=metadata["feature_names"],
        output_path=fig4_path
    )
    print(f" -> Generated: {fig4_path.name}")
    
    figures_paths = {
        "figure1": fig1_path,
        "figure2": fig2_path,
        "figure3": fig3_path,
        "figure4": fig4_path
    }
    
    # 8. Generate Academic Word Report
    print("\n[Step 7/7] Generating Comprehensive DOCX Report in report/...")
    report_docx_path = report_dir / "Week_4_Machine_Learning_Model_Development_Evaluation.docx"
    generate_word_report(
        metadata=metadata,
        inspection_summary=inspection_summary,
        summary_df=summary_df,
        evaluation_results=evaluation_results,
        misclassified_df=lr_errors,
        figures_paths=figures_paths,
        output_path=report_docx_path
    )
    print(f" -> Generated Word Report: {report_docx_path.name}")
    
    print("\n================================================================================")
    print("                    PIPELINE EXECUTION COMPLETED SUCCESSFULLY                  ")
    print("================================================================================")


if __name__ == "__main__":
    run_pipeline()
