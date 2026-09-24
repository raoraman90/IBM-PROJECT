"""
app.py
======
Flask REST API for House Price Prediction.

Endpoints:
  GET  /health              -> service health check
  GET  /model-info          -> model metadata, features, training metrics
  GET  /dataset             -> dataset summary and sample rows
  POST /predict             -> predict price for a single house
  POST /batch-predict       -> predict prices for a list of houses

Usage:
    py -3 app.py
    API available at http://127.0.0.1:5000
"""

import os
import pickle
import pandas as pd
from flask import Flask, request, jsonify
from flask_cors import CORS

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
MODEL_PATH = os.path.join(BASE_DIR, "model", "house_price_model.pkl")
META_PATH  = os.path.join(BASE_DIR, "model", "metadata.pkl")
DATA_PATH  = os.path.join(BASE_DIR, "archive (1)", "House Price Prediction Dataset.csv")

# ---------------------------------------------------------------------------
# Load model artifacts
# ---------------------------------------------------------------------------
def _load_artifacts():
    if not os.path.exists(MODEL_PATH):
        raise FileNotFoundError(
            f"Model not found at {MODEL_PATH}. Run: py -3 train_model.py"
        )
    with open(MODEL_PATH, "rb") as f:
        model = pickle.load(f)
    with open(META_PATH, "rb") as f:
        meta = pickle.load(f)
    return model, meta

model, metadata = _load_artifacts()

FEATURES_NUM = metadata["features_num"]
FEATURES_CAT = metadata["features_cat"]

# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------
app = Flask(__name__)
CORS(app)

# ---------------------------------------------------------------------------
# Input validation helper
# ---------------------------------------------------------------------------
def _validate(data: dict):
    """Returns (cleaned_dict, error_string). error_string is None on success."""
    required = FEATURES_NUM + FEATURES_CAT
    missing  = [k for k in required if k not in data]
    if missing:
        return None, f"Missing fields: {missing}"

    try:
        cleaned = {
            "Area":      int(data["Area"]),
            "Bedrooms":  int(data["Bedrooms"]),
            "Bathrooms": int(data["Bathrooms"]),
            "Floors":    int(data["Floors"]),
            "YearBuilt": int(data["YearBuilt"]),
            "Location":  str(data["Location"]),
            "Condition": str(data["Condition"]),
            "Garage":    str(data["Garage"]),
        }
    except (ValueError, TypeError) as exc:
        return None, f"Type error: {exc}"

    if cleaned["Location"] not in metadata["location_options"]:
        return None, f"Invalid Location. Options: {metadata['location_options']}"
    if cleaned["Condition"] not in metadata["condition_options"]:
        return None, f"Invalid Condition. Options: {metadata['condition_options']}"
    if cleaned["Garage"] not in metadata["garage_options"]:
        return None, f"Invalid Garage. Options: {metadata['garage_options']}"

    return cleaned, None


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route("/health", methods=["GET"])
def health():
    """Service health check."""
    return jsonify({
        "status":    "ok",
        "model":     metadata["model_params"]["algorithm"],
        "api_version": "1.0.0",
    }), 200


@app.route("/model-info", methods=["GET"])
def model_info():
    """Return model metadata, features used, and training metrics."""
    return jsonify({
        "algorithm":         metadata["model_params"]["algorithm"],
        "n_estimators":      metadata["model_params"]["n_estimators"],
        "max_depth":         metadata["model_params"]["max_depth"],
        "random_state":      metadata["model_params"]["random_state"],
        "train_size":        metadata["model_params"]["train_size"],
        "test_size":         metadata["model_params"]["test_size"],
        "numeric_features":  metadata["features_num"],
        "categorical_features": metadata["features_cat"],
        "preprocessing":     {
            "numeric":     "StandardScaler",
            "categorical": "OneHotEncoder (handle_unknown=ignore)",
        },
        "metrics": metadata["metrics"],
        "top_features": metadata["feature_importances"][:10],
    }), 200


@app.route("/dataset", methods=["GET"])
def dataset():
    """Return dataset summary statistics and sample rows."""
    df = pd.read_csv(DATA_PATH)
    sample = df.sample(n=min(10, len(df)), random_state=42).to_dict(orient="records")
    stats  = df.describe().round(2).to_dict()
    loc_dist  = df["Location"].value_counts().to_dict()
    cond_dist = df["Condition"].value_counts().to_dict()
    gar_dist  = df["Garage"].value_counts().to_dict()

    return jsonify({
        "total_rows":    len(df),
        "total_columns": len(df.columns),
        "columns":       list(df.columns),
        "price_stats": {
            "min":  float(df["Price"].min()),
            "max":  float(df["Price"].max()),
            "mean": round(float(df["Price"].mean()), 2),
            "std":  round(float(df["Price"].std()),  2),
        },
        "location_distribution":  loc_dist,
        "condition_distribution": cond_dist,
        "garage_distribution":    gar_dist,
        "descriptive_stats":      stats,
        "sample_rows":            sample,
    }), 200


@app.route("/predict", methods=["POST"])
def predict():
    """Predict price for a single house."""
    body     = request.get_json(force=True)
    cleaned, error = _validate(body)
    if error:
        return jsonify({"error": error}), 400

    df    = pd.DataFrame([cleaned])
    price = float(model.predict(df)[0])

    return jsonify({
        "predicted_price": round(price, 2),
        "currency":        "USD",
        "input":           cleaned,
    }), 200


@app.route("/batch-predict", methods=["POST"])
def batch_predict():
    """Predict prices for a list of houses."""
    body = request.get_json(force=True)
    if not isinstance(body, list):
        return jsonify({"error": "Expected a JSON array of house objects."}), 400

    results = []
    for idx, item in enumerate(body):
        cleaned, error = _validate(item)
        if error:
            results.append({"index": idx, "error": error})
            continue
        price = float(model.predict(pd.DataFrame([cleaned]))[0])
        results.append({
            "index":           idx,
            "predicted_price": round(price, 2),
            "currency":        "USD",
            "input":           cleaned,
        })

    return jsonify(results), 200


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print("House Price Prediction API")
    print("Endpoints: /health  /model-info  /dataset  /predict  /batch-predict")
    print("Running on http://127.0.0.1:5000")
    app.run(debug=False, host="0.0.0.0", port=5000)
