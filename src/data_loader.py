from pathlib import Path
from typing import Dict, Any, Tuple
import pandas as pd
from sklearn.datasets import load_wine


def load_raw_wine_data() -> Tuple[pd.DataFrame, Dict[str, Any]]:
    wine_bunch = load_wine(as_frame=True)
    df = wine_bunch.frame.copy()
    
    metadata = {
        "dataset_name": "Wine Recognition Dataset",
        "source": "UCI Machine Learning Repository / scikit-learn",
        "original_authors": "Forina, M. et al., 1988",
        "num_observations": int(df.shape[0]),
        "num_features": int(len(wine_bunch.feature_names)),
        "feature_names": list(wine_bunch.feature_names),
        "target_name": "target",
        "target_classes": [str(c) for c in wine_bunch.target_names],
        "target_mapping": {i: name for i, name in enumerate(wine_bunch.target_names)},
        "feature_descriptions": {
            "alcohol": "Alcohol content (% by volume)",
            "malic_acid": "Malic acid concentration (g/L)",
            "ash": "Ash content after incineration (g/L)",
            "alcalinity_of_ash": "Alcalinity of ash",
            "magnesium": "Magnesium concentration (mg/L)",
            "total_phenols": "Total phenolic content (g/L)",
            "flavanoids": "Flavanoid polyphenol concentration (g/L)",
            "nonflavanoid_phenols": "Nonflavanoid phenols concentration (g/L)",
            "proanthocyanins": "Proanthocyanin content (g/L)",
            "color_intensity": "Color intensity measurement",
            "hue": "Hue ratio of light absorbance",
            "od280/od315_of_diluted_wines": "OD280/OD315 ratio (absorbance of diluted wines)",
            "proline": "Proline amino acid concentration (mg/L)"
        }
    }
    
    return df, metadata


def inspect_data(df: pd.DataFrame, target_col: str = "target") -> Dict[str, Any]:
    feature_cols = [c for c in df.columns if c != target_col]
    
    shape = df.shape
    missing_values = df.isnull().sum().to_dict()
    total_missing = int(df.isnull().sum().sum())
    duplicate_rows = int(df.duplicated().sum())
    dtypes = {col: str(dtype) for col, dtype in df.dtypes.items()}
    
    target_distribution = df[target_col].value_counts().sort_index().to_dict()
    target_proportions = (df[target_col].value_counts(normalize=True).sort_index()).to_dict()
    
    descriptive_stats = df[feature_cols].describe().T
    descriptive_stats = descriptive_stats[["count", "mean", "std", "min", "25%", "50%", "75%", "max"]]
    
    inspection_summary = {
        "num_rows": int(shape[0]),
        "num_columns": int(shape[1]),
        "num_features": len(feature_cols),
        "total_missing_values": total_missing,
        "missing_per_column": missing_values,
        "duplicate_rows": duplicate_rows,
        "dtypes": dtypes,
        "target_distribution": target_distribution,
        "target_proportions": target_proportions,
        "descriptive_statistics": descriptive_stats
    }
    
    return inspection_summary


def save_dataset(df: pd.DataFrame, file_path: Path) -> None:
    file_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(file_path, index=False)
