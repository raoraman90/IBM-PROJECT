"""
train_model.py
==============
Trains a Random Forest Regressor on the House Price Prediction dataset.
Saves the trained pipeline and feature metadata to the model/ directory.

Usage:
    py -3 train_model.py
"""

import os
import pickle
import numpy as np
import pandas as pd
from sklearn.ensemble import RandomForestRegressor
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
DATA_PATH  = os.path.join(BASE_DIR, "archive (1)", "House Price Prediction Dataset.csv")
MODEL_DIR  = os.path.join(BASE_DIR, "model")
MODEL_PATH = os.path.join(MODEL_DIR, "house_price_model.pkl")
META_PATH  = os.path.join(MODEL_DIR, "metadata.pkl")

# ---------------------------------------------------------------------------
# Feature definitions
# ---------------------------------------------------------------------------
FEATURES_NUM = ["Area", "Bedrooms", "Bathrooms", "Floors", "YearBuilt"]
FEATURES_CAT = ["Location", "Condition", "Garage"]
TARGET       = "Price"

# ---------------------------------------------------------------------------
# Load dataset
# ---------------------------------------------------------------------------
print("[1/5] Loading dataset ...")
df = pd.read_csv(DATA_PATH)
print(f"      Rows: {len(df):,}  |  Columns: {list(df.columns)}")

X = df[FEATURES_NUM + FEATURES_CAT]
y = df[TARGET]

# ---------------------------------------------------------------------------
# Train / test split
# ---------------------------------------------------------------------------
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=42
)
print(f"[2/5] Split -> train={len(X_train):,}  test={len(X_test):,}")

# ---------------------------------------------------------------------------
# Preprocessing + model pipeline
# ---------------------------------------------------------------------------
preprocessor = ColumnTransformer(transformers=[
    ("num", StandardScaler(),                                      FEATURES_NUM),
    ("cat", OneHotEncoder(handle_unknown="ignore", sparse_output=False), FEATURES_CAT),
])

pipeline = Pipeline(steps=[
    ("preprocessor", preprocessor),
    ("regressor", RandomForestRegressor(
        n_estimators=200,
        max_depth=None,
        min_samples_split=2,
        random_state=42,
        n_jobs=-1,
    )),
])

# ---------------------------------------------------------------------------
# Train
# ---------------------------------------------------------------------------
print("[3/5] Training RandomForestRegressor (200 estimators) ...")
pipeline.fit(X_train, y_train)

# ---------------------------------------------------------------------------
# Evaluate
# ---------------------------------------------------------------------------
print("[4/5] Evaluating ...")
y_pred   = pipeline.predict(X_test)
mae      = mean_absolute_error(y_test, y_pred)
rmse     = float(np.sqrt(mean_squared_error(y_test, y_pred)))
r2       = r2_score(y_test, y_pred)
cv_r2    = cross_val_score(pipeline, X, y, cv=5, scoring="r2")

print(f"      MAE        : ${mae:,.2f}")
print(f"      RMSE       : ${rmse:,.2f}")
print(f"      R2 Score   : {r2:.4f}")
print(f"      CV R2 (5x) : {cv_r2.mean():.4f} +/- {cv_r2.std():.4f}")

# ---------------------------------------------------------------------------
# Feature importances
# ---------------------------------------------------------------------------
rf           = pipeline.named_steps["regressor"]
ohe_names    = pipeline.named_steps["preprocessor"] \
                        .named_transformers_["cat"] \
                        .get_feature_names_out(FEATURES_CAT).tolist()
all_features = FEATURES_NUM + ohe_names
feat_df      = pd.DataFrame({
    "Feature":    all_features,
    "Importance": rf.feature_importances_,
}).sort_values("Importance", ascending=False).reset_index(drop=True)

print("\n      Top 10 Feature Importances:")
print(feat_df.head(10).to_string(index=False))

# ---------------------------------------------------------------------------
# Save artifacts
# ---------------------------------------------------------------------------
print("\n[5/5] Saving model artifacts ...")
os.makedirs(MODEL_DIR, exist_ok=True)

with open(MODEL_PATH, "wb") as f:
    pickle.dump(pipeline, f)

metadata = {
    "features_num":      FEATURES_NUM,
    "features_cat":      FEATURES_CAT,
    "all_features":      all_features,
    "feature_importances": feat_df.to_dict(orient="records"),
    "location_options":  sorted(df["Location"].unique().tolist()),
    "condition_options": sorted(df["Condition"].unique().tolist()),
    "garage_options":    sorted(df["Garage"].unique().tolist()),
    "dataset_shape":     list(df.shape),
    "dataset_columns":   list(df.columns),
    "price_min":         float(df[TARGET].min()),
    "price_max":         float(df[TARGET].max()),
    "price_mean":        float(df[TARGET].mean()),
    "price_std":         float(df[TARGET].std()),
    "area_min":          int(df["Area"].min()),
    "area_max":          int(df["Area"].max()),
    "year_min":          int(df["YearBuilt"].min()),
    "year_max":          int(df["YearBuilt"].max()),
    "metrics": {
        "mae":        float(mae),
        "rmse":       float(rmse),
        "r2":         float(r2),
        "cv_r2_mean": float(cv_r2.mean()),
        "cv_r2_std":  float(cv_r2.std()),
    },
    "model_params": {
        "algorithm":    "RandomForestRegressor",
        "n_estimators": 200,
        "max_depth":    "None (unlimited)",
        "random_state": 42,
        "train_size":   len(X_train),
        "test_size":    len(X_test),
    },
}

with open(META_PATH, "wb") as f:
    pickle.dump(metadata, f)

print(f"      Model    -> {MODEL_PATH}")
print(f"      Metadata -> {META_PATH}")
print("\nTraining complete!")
