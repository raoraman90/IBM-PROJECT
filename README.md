# 🏠 House Price Prediction

A full-stack Machine Learning project for predicting residential property prices, built entirely in Python.

---

## Project Structure

```
house_price_prediction/
│
├── archive (1)/
│   └── House Price Prediction Dataset.csv   ← 2,000 rows dataset
│
├── model/                                   ← generated after training
│   ├── house_price_model.pkl
│   └── metadata.pkl
│
├── reports/                                 ← generated Word report
│   └── house_price_report.docx
│
├── screenshots/                             ← generated matplotlib images
│   ├── fig_price_distribution.png
│   ├── fig_correlation_heatmap.png
│   ├── fig_location_analysis.png
│   ├── fig_feature_importance.png
│   ├── fig_predicted_vs_actual.png
│   ├── ui_predict_page.png
│   ├── ui_insights_page.png
│   └── ui_performance_page.png
│
├── train_model.py      ← Step 1: Train ML model
├── app.py              ← Step 2: Flask REST API
├── ui.py               ← Step 3: Streamlit frontend
├── generate_report.py  ← Step 4: Generate Word report
├── requirements.txt
└── README.md
```

---

## Quick Start

### 1. Install dependencies

```bash
cd house_price_prediction
py -3 -m pip install -r requirements.txt
```

### 2. Train the model

```bash
py -3 train_model.py
```

Saves `model/house_price_model.pkl` and `model/metadata.pkl`.

### 3. Start Flask API *(Terminal 1)*

```bash
py -3 app.py
```

API runs at **http://127.0.0.1:5000**

### 4. Launch Streamlit UI *(Terminal 2)*

```bash
py -3 -m streamlit run ui.py
```

UI opens at **http://localhost:8501**

### 5. Generate Word Report *(Optional)*

```bash
py -3 generate_report.py
```

Report saved to `reports/house_price_report.docx`

---

## API Reference

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/health` | Service health check |
| GET | `/model-info` | Model parameters and training metrics |
| GET | `/dataset` | Dataset summary, statistics, sample rows |
| POST | `/predict` | Predict price for a single house |
| POST | `/batch-predict` | Predict prices for multiple houses |

### Example: Single Prediction

```bash
curl -X POST http://127.0.0.1:5000/predict \
  -H "Content-Type: application/json" \
  -d '{"Area":2500,"Bedrooms":3,"Bathrooms":2,"Floors":2,"YearBuilt":2005,"Location":"Suburban","Condition":"Good","Garage":"Yes"}'
```

Response:
```json
{
  "predicted_price": 385420.50,
  "currency": "USD",
  "input": { ... }
}
```

---

## Dataset Features

| Feature | Type | Range / Values |
|---------|------|----------------|
| Area | Numeric | sq ft |
| Bedrooms | Numeric | 1 – 5 |
| Bathrooms | Numeric | 1 – 4 |
| Floors | Numeric | 1 – 3 |
| YearBuilt | Numeric | 1900 – 2023 |
| Location | Categorical | Downtown, Suburban, Urban, Rural |
| Condition | Categorical | Excellent, Good, Fair, Poor |
| Garage | Categorical | Yes, No |
| **Price** | **Target** | **USD** |

---

## ML Model

| Parameter | Value |
|-----------|-------|
| Algorithm | Random Forest Regressor |
| Estimators | 200 trees |
| Numeric Preprocessing | StandardScaler |
| Categorical Encoding | OneHotEncoder |
| Train / Test Split | 80% / 20% |
| Validation | 5-Fold Cross-Validation |

---

## Tech Stack

| Layer | Library |
|-------|---------|
| Frontend UI | Streamlit, Plotly |
| Backend API | Flask, Flask-CORS |
| ML Pipeline | scikit-learn |
| Data Processing | pandas, NumPy |
| Visualisation | matplotlib, seaborn |
| Report Generation | python-docx, Pillow |
