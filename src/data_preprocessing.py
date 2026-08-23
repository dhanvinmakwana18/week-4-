from typing import Tuple
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.base import BaseEstimator


def prepare_train_test_split(
    df: pd.DataFrame,
    target_col: str = "target",
    test_size: float = 0.20,
    random_state: int = 42
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.Series, pd.Series]:
    X = df.drop(columns=[target_col]).copy()
    y = df[target_col].copy()
    
    X_train, X_test, y_train, y_test = train_test_split(
        X,
        y,
        test_size=test_size,
        random_state=random_state,
        stratify=y
    )
    
    return X_train, X_test, y_train, y_test


def build_pipeline(classifier: BaseEstimator, scale_features: bool = True) -> Pipeline:
    if scale_features:
        return Pipeline([
            ("scaler", StandardScaler()),
            ("classifier", classifier)
        ])
    return Pipeline([
        ("classifier", classifier)
    ])
