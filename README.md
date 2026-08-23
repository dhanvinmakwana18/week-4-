# Machine Learning Model Development and Evaluation: Wine Cultivar Classification Pipeline

## Project Overview
This repository contains an end-to-end, reproducible machine learning classification pipeline developed for the **Week 4 Data Science Project: Machine Learning Model Development and Evaluation**.

The objective is to predict the cultivar origin of Italian wines from 13 continuous chemical and physical measurements using supervised multiclass classification algorithms.

---

## Research Question & Problem Formulation
- **Research Question:** Can the cultivar origin (Class 0: Barolo, Class 1: Grignolino, Class 2: Barbera) of Italian wines be reliably classified from 13 continuous physicochemical measurements, and which classification architecture (Standardized Multinomial Logistic Regression vs. Bagged Random Forest vs. Single Decision Tree) achieves the most robust out-of-sample generalization?
- **Input Space ($\mathbf{X}$):** 13-dimensional continuous physicochemical feature vector $\mathbf{x} \in \mathbb{R}^{13}$.
- **Target Space ($Y$):** Discrete multiclass label $y \in \{0, 1, 2\}$.

---

## Dataset Description
- **Dataset Name:** Wine Recognition Dataset
- **Source:** UCI Machine Learning Repository / `scikit-learn` (`sklearn.datasets.load_wine`)
- **Original Reference:** Forina, M. et al. (1988), *Multivariate data analysis as a discriminating tool of the origin of wines*, Vitis.
- **Observations:** 178 complete chemical profiles
- **Missing Values:** 0 missing values across all features
- **Duplicate Records:** 0 duplicates
- **Target Classes:** 3 cultivars
  - Class 0 (Barolo): 59 observations (33.15%)
  - Class 1 (Grignolino): 71 observations (39.89%)
  - Class 2 (Barbera): 48 observations (26.97%)

### Physicochemical Features (13 continuous attributes)
1. `alcohol`: Ethanol content (% by volume)
2. `malic_acid`: Malic acid concentration (g/L)
3. `ash`: Inorganic mineral residue (g/L)
4. `alcalinity_of_ash`: Buffering capacity of ash
5. `magnesium`: Magnesium cation concentration (mg/L)
6. `total_phenols`: Aggregate phenolic content (g/L)
7. `flavanoids`: Flavanoid polyphenol concentration (g/L)
8. `nonflavanoid_phenols`: Nonflavanoid phenols concentration (g/L)
9. `proanthocyanins`: Condensed tannins concentration (g/L)
10. `color_intensity`: Colorimetric absorbance measurement
11. `hue`: Yellow-to-red absorbance ratio
12. `od280/od315_of_diluted_wines`: Spectrophotometric protein/phenolic ratio
13. `proline`: Proline amino acid concentration (mg/L)

---

## Project Structure
```text
week4-machine-learning-model-development/
│
├── data/
│   └── processed/
│       ├── train_data.csv
│       ├── test_data.csv
│       └── wine_raw.csv
│
├── outputs/
│   ├── figures/
│   │   ├── figure1_confusion_matrix.png
│   │   ├── figure2_model_performance_comparison.png
│   │   ├── figure3_multiclass_roc_curve.png
│   │   └── figure4_feature_importance.png
│   └── tables/
│       ├── descriptive_statistics.csv
│       ├── model_comparison_metrics.csv
│       ├── classification_report_logistic_regression.csv
│       ├── classification_report_random_forest.csv
│       ├── classification_report_decision_tree.csv
│       ├── confusion_matrix_logistic_regression.csv
│       ├── confusion_matrix_random_forest.csv
│       ├── confusion_matrix_decision_tree.csv
│       ├── feature_importances_random_forest.csv
│       ├── misclassified_decision_tree.csv
│       └── misclassified_logistic_regression.csv
│
├── report/
│   └── Week_4_Machine_Learning_Model_Development_Evaluation.docx
│
├── src/
│   ├── data_loader.py
│   ├── data_preprocessing.py
│   ├── model_training.py
│   ├── model_evaluation.py
│   ├── visualizations.py
│   ├── report_generator.py
│   └── main.py
│
├── README.md
└── requirements.txt
```

---

## Machine Learning Pipeline Architecture

1. **Data Inspection & Hygiene:**
   - Verifies shapes, missing values, duplicate rows, data types, and computes descriptive statistics.
2. **Stratified Splitting & Leakage Prevention:**
   - Partitions data into 80% training ($N = 142$) and 20% testing ($N = 36$) using stratified sampling (`random_state=42`).
   - Standardizes features via `StandardScaler` strictly encapsulated inside scikit-learn `Pipeline` objects to prevent data snooping.
3. **Model Development:**
   - **Logistic Regression (Baseline):** Multinomial linear classifier with L2 penalty, L-BFGS solver, and standard feature scaling.
   - **Random Forest Classifier (Ensemble):** 100 de-correlated bagged decision trees (`max_depth=4`, `random_state=42`).
   - **Decision Tree Classifier (Single Tree):** CART model with Gini impurity criterion (`max_depth=3`, `random_state=42`).
4. **Multiclass Evaluation:**
   - Computes Accuracy, Macro/Weighted Precision, Macro/Weighted Recall, Macro/Weighted F1-Score, and One-vs-Rest (OvR) ROC AUC.
   - Generates confusion matrices and per-class classification reports.
   - Clarifies that perfect ROC-AUC (1.0000) for Logistic Regression indicates correct probability ranking across thresholds, while the single classification error at the default threshold resulted in 97.22% accuracy.
5. **Error & Generalization Analysis:**
   - Inspects misclassifications, class boundaries, and empirical generalization gaps ($\Delta = \text{Train Acc} - \text{Test Acc}$).
   - Employs scientifically cautious wording regarding single-split generalization.
6. **Publication Figures & Word Report Generation:**
   - Automatically renders 300 DPI figures and compiles the full academic DOCX report dynamically with all numerical results populated from computed metrics.

---

## Empirical Evaluation Summary

| Model | Train Accuracy | Test Accuracy | Macro Precision | Macro Recall | Macro F1-Score | Weighted F1-Score | OvR ROC-AUC | Generalization Gap |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **Random Forest (Champion)** | **1.0000** | **1.0000** | **1.0000** | **1.0000** | **1.0000** | **1.0000** | **1.0000** | **0.00%** |
| **Logistic Regression (Baseline)** | 1.0000 | 0.9722 | 0.9778 | 0.9667 | 0.9710 | 0.9720 | 1.0000 | +2.78% |
| **Decision Tree (Single Tree)** | 0.9930 | 0.9444 | 0.9583 | 0.9389 | 0.9457 | 0.9450 | 0.9493 | +4.86% |

---

## Installation & Execution

### 1. Prerequisites
- Python 3.10+ (tested on Python 3.14)
- Virtual environment (`venv` or `conda`)

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Run Pipeline
From the project root directory, execute:
```bash
python src/main.py
```

Upon execution, the script will:
1. Load and inspect the dataset.
2. Generate train/test splits in `data/processed/`.
3. Train candidate models.
4. Compute metrics, confusion matrices, and ROC statistics in `outputs/tables/`.
5. Render figures in `outputs/figures/`.
6. Compile the complete academic document in `report/Week_4_Machine_Learning_Model_Development_Evaluation.docx`.
