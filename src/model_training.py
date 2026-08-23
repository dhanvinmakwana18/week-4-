from typing import Dict, Any
import pandas as pd
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import RandomForestClassifier
from sklearn.tree import DecisionTreeClassifier
from sklearn.pipeline import Pipeline

from src.data_preprocessing import build_pipeline


def initialize_candidate_models(random_state: int = 42) -> Dict[str, Pipeline]:
    models = {
        "Logistic Regression": build_pipeline(
            LogisticRegression(
                random_state=random_state,
                max_iter=1000,
                C=1.0,
                solver="lbfgs"
            ),
            scale_features=True
        ),
        "Random Forest": build_pipeline(
            RandomForestClassifier(
                n_estimators=100,
                max_depth=4,
                random_state=random_state
            ),
            scale_features=False
        ),
        "Decision Tree": build_pipeline(
            DecisionTreeClassifier(
                max_depth=3,
                random_state=random_state
            ),
            scale_features=False
        )
    }
    return models


def train_models(
    models: Dict[str, Pipeline],
    X_train: pd.DataFrame,
    y_train: pd.Series
) -> Dict[str, Pipeline]:
    trained_models = {}
    for model_name, pipeline in models.items():
        pipeline.fit(X_train, y_train)
        trained_models[model_name] = pipeline
    return trained_models
