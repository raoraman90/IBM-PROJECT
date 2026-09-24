"""
generate_report.py
==================
Generates a comprehensive Word (.docx) project report for the
House Price Prediction system.

Includes:
  - Project overview & architecture
  - Dataset analysis with sample data table
  - Model information & training metrics
  - API documentation with sample outputs
  - All matplotlib screenshots embedded as figures
  - Batch prediction results table

Run with:
    py -3 generate_report.py
Output: reports/house_price_report.docx
"""

import os
import json
import pickle
from datetime import datetime

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import pandas as pd

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
SHOTS_DIR = os.path.join(BASE_DIR, "screenshots")
META_PATH = os.path.join(BASE_DIR, "model", "metadata.pkl")
DATA_PATH = os.path.join(BASE_DIR, "archive (1)", "House Price Prediction Dataset.csv")
API_PATH  = os.path.join(BASE_DIR, "api_outputs.json")
OUT_DIR   = os.path.join(BASE_DIR, "reports")
OUT_PATH  = os.path.join(OUT_DIR, "house_price_report.docx")
os.makedirs(OUT_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# Load artifacts
# ---------------------------------------------------------------------------
meta    = pickle.load(open(META_PATH, "rb"))
df      = pd.read_csv(DATA_PATH)
api_out = json.load(open(API_PATH))

# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------
def set_cell_bg(cell, hex_color: str):
    """Set table cell background colour."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def add_horizontal_rule(doc):
    """Add a thin horizontal line paragraph."""
    p    = doc.add_paragraph()
    pPr  = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    "4")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "CCCCCC")
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p

def style_heading(para, size_pt=14):
    para.runs[0].font.size = Pt(size_pt)
    para.runs[0].font.color.rgb = RGBColor(0x1d, 0x4e, 0xd8)

def add_kv_row(table, label, value, row_idx, bg="F7F8FA"):
    row = table.rows[row_idx]
    row.cells[0].text = label
    row.cells[1].text = str(value)
    row.cells[0].paragraphs[0].runs[0].font.bold = True
    row.cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(0x1f, 0x23, 0x28)
    if bg:
        set_cell_bg(row.cells[0], bg)

def add_figure(doc, img_file, caption, width_in=6.0):
    """Insert an image with a centred caption."""
    img_path = os.path.join(SHOTS_DIR, img_file)
    if not os.path.exists(img_path):
        doc.add_paragraph(f"[Image not found: {img_file}]")
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(img_path, width=Inches(width_in))
    cap = doc.add_paragraph(caption)
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.runs[0].font.size = Pt(9)
    cap.runs[0].font.italic = True
    cap.runs[0].font.color.rgb = RGBColor(0x57, 0x60, 0x6a)
    doc.add_paragraph()   # spacing

# ---------------------------------------------------------------------------
# Create document
# ---------------------------------------------------------------------------
doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# Default font
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)

# ===========================================================================
# COVER PAGE
# ===========================================================================
doc.add_paragraph()
doc.add_paragraph()
doc.add_paragraph()

title_para = doc.add_paragraph()
title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_run = title_para.add_run("House Price Prediction System")
title_run.font.size  = Pt(28)
title_run.font.bold  = True
title_run.font.color.rgb = RGBColor(0x1d, 0x4e, 0xd8)

sub_para = doc.add_paragraph()
sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub_run = sub_para.add_run("Full-Stack Machine Learning Project Report")
sub_run.font.size = Pt(16)
sub_run.font.color.rgb = RGBColor(0x57, 0x60, 0x6a)

doc.add_paragraph()
add_horizontal_rule(doc)
doc.add_paragraph()

info_para = doc.add_paragraph()
info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
info_para.add_run(f"Generated: {datetime.now().strftime('%B %d, %Y  %H:%M')}\n").font.size = Pt(11)
info_para.add_run("Technology: Python | Flask | Streamlit | scikit-learn\n").font.size = Pt(11)
info_para.add_run("Dataset: House Price Prediction Dataset (2,000 records)").font.size = Pt(11)

doc.add_page_break()

# ===========================================================================
# SECTION 1 — Executive Summary
# ===========================================================================
h = doc.add_heading("1. Executive Summary", level=1)
style_heading(h, 16)

doc.add_paragraph(
    "This report documents the complete House Price Prediction System — a full-stack "
    "Machine Learning application built entirely in Python. The project encompasses a "
    "trained Random Forest Regressor model, a Flask REST API backend, a Streamlit "
    "interactive frontend, and automated report generation."
)
doc.add_paragraph(
    "The system predicts residential house prices based on eight property features: "
    "area, bedrooms, bathrooms, floors, year built, location, condition, and garage "
    "availability. The dataset contains 2,000 records spanning properties in Downtown, "
    "Suburban, Urban, and Rural locations with prices ranging from $50,005 to $999,656."
)

# Key highlights table
doc.add_paragraph()
tbl = doc.add_table(rows=5, cols=2)
tbl.style = "Table Grid"
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
headers = [("Aspect", "Detail"),
           ("Dataset",    "2,000 rows | 9 features | Price range $50K-$1M"),
           ("Algorithm",  "Random Forest Regressor (200 estimators)"),
           ("API",        "Flask REST API — 5 endpoints on port 5000"),
           ("Frontend",   "Streamlit — 4 interactive pages on port 8501")]
for i, (k, v) in enumerate(headers):
    add_kv_row(tbl, k, v, i, bg="1D4ED8" if i == 0 else ("EFF6FF" if i % 2 else "F7F8FA"))
    if i == 0:
        for cell in tbl.rows[0].cells:
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            cell.paragraphs[0].runs[0].font.bold = True

doc.add_paragraph()
doc.add_page_break()

# ===========================================================================
# SECTION 2 — Project Structure
# ===========================================================================
h = doc.add_heading("2. Project Structure", level=1)
style_heading(h, 16)

doc.add_paragraph("The project follows a clean separation between data, model, backend, and frontend components:")
doc.add_paragraph()

struct_para = doc.add_paragraph()
struct_para.add_run(
    "house_price_prediction/\n"
    "  archive (1)/\n"
    "    House Price Prediction Dataset.csv   <- 2,000 rows dataset\n"
    "  model/\n"
    "    house_price_model.pkl                <- trained pipeline\n"
    "    metadata.pkl                         <- feature info & metrics\n"
    "  reports/\n"
    "    house_price_report.docx              <- this report\n"
    "  screenshots/\n"
    "    fig_*.png  /  ui_*.png               <- generated charts\n"
    "  train_model.py                         <- Step 1: Train ML model\n"
    "  app.py                                 <- Step 2: Flask REST API\n"
    "  ui.py                                  <- Step 3: Streamlit UI\n"
    "  generate_screenshots.py                <- Step 4: Chart generation\n"
    "  generate_report.py                     <- Step 5: Word report\n"
    "  requirements.txt\n"
    "  README.md"
).font.name = "Courier New"
struct_para.runs[0].font.size = Pt(9)
struct_para.paragraph_format.space_before = Pt(0)
struct_para.paragraph_format.space_after  = Pt(0)

doc.add_paragraph()

# Tech stack table
h2 = doc.add_heading("2.1 Technology Stack", level=2)
tbl2 = doc.add_table(rows=7, cols=2)
tbl2.style = "Table Grid"
stack = [("Layer", "Library / Technology"),
         ("Model Training",    "scikit-learn, pandas, NumPy"),
         ("Backend API",       "Flask 3.x, Flask-CORS"),
         ("Frontend UI",       "Streamlit 1.35+, Plotly"),
         ("Visualisation",     "matplotlib, seaborn"),
         ("Report Generation", "python-docx, Pillow"),
         ("Data Format",       "CSV (2,000 rows)")]
for i, (k, v) in enumerate(stack):
    add_kv_row(tbl2, k, v, i, bg="1D4ED8" if i == 0 else ("EFF6FF" if i % 2 else "F7F8FA"))
    if i == 0:
        for cell in tbl2.rows[0].cells:
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            cell.paragraphs[0].runs[0].font.bold = True

doc.add_page_break()

# ===========================================================================
# SECTION 3 — Dataset Analysis
# ===========================================================================
h = doc.add_heading("3. Dataset Analysis", level=1)
style_heading(h, 16)

doc.add_paragraph(
    "The dataset contains 2,000 residential property records with 10 columns. "
    "It includes both numerical and categorical features, with Price as the prediction target."
)

# Dataset stats table
h2 = doc.add_heading("3.1 Dataset Statistics", level=2)
tbl3 = doc.add_table(rows=7, cols=2)
tbl3.style = "Table Grid"
ds_info = api_out["dataset"]
stats_data = [
    ("Attribute", "Value"),
    ("Total Records",    f"{ds_info['total_rows']:,}"),
    ("Total Columns",    str(ds_info["total_columns"])),
    ("Price Range",      f"${ds_info['price_stats']['min']:,.0f}  —  ${ds_info['price_stats']['max']:,.0f}"),
    ("Average Price",    f"${ds_info['price_stats']['mean']:,.2f}"),
    ("Price Std Dev",    f"${ds_info['price_stats']['std']:,.2f}"),
    ("Feature Types",    "5 Numeric  +  3 Categorical"),
]
for i, (k, v) in enumerate(stats_data):
    add_kv_row(tbl3, k, v, i, bg="1D4ED8" if i == 0 else ("EFF6FF" if i % 2 else "F7F8FA"))
    if i == 0:
        for cell in tbl3.rows[0].cells:
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            cell.paragraphs[0].runs[0].font.bold = True

doc.add_paragraph()

# Feature descriptions
h2 = doc.add_heading("3.2 Feature Descriptions", level=2)
tbl4 = doc.add_table(rows=10, cols=3)
tbl4.style = "Table Grid"
feat_desc = [
    ("Feature",    "Type",        "Description / Values"),
    ("Area",       "Numeric",     "Property area in sq ft"),
    ("Bedrooms",   "Numeric",     "Number of bedrooms (1-5)"),
    ("Bathrooms",  "Numeric",     "Number of bathrooms (1-4)"),
    ("Floors",     "Numeric",     "Number of floors (1-3)"),
    ("YearBuilt",  "Numeric",     "Year of construction"),
    ("Location",   "Categorical", "Downtown / Suburban / Urban / Rural"),
    ("Condition",  "Categorical", "Excellent / Good / Fair / Poor"),
    ("Garage",     "Categorical", "Yes / No"),
    ("Price",      "Target",      "House price in USD (prediction target)"),
]
for i, (f, t, d) in enumerate(feat_desc):
    row = tbl4.rows[i]
    row.cells[0].text = f; row.cells[1].text = t; row.cells[2].text = d
    for j in range(3):
        row.cells[j].paragraphs[0].runs[0].font.bold = (i == 0)
        if i == 0:
            set_cell_bg(row.cells[j], "1D4ED8")
            row.cells[j].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        elif i == 9:
            set_cell_bg(row.cells[j], "EFF6FF")

doc.add_paragraph()

# Distribution tables
h2 = doc.add_heading("3.3 Category Distributions", level=2)
col_a, col_b = doc.add_paragraph(), doc.add_paragraph()

dist_tbl = doc.add_table(rows=6, cols=3)
dist_tbl.style = "Table Grid"
loc_d  = ds_info["location_distribution"]
cond_d = ds_info["condition_distribution"]
for i, (row_data) in enumerate([
    ("Location",    "Count",   "Condition"),
    ("Downtown",   str(loc_d.get("Downtown",0)),  "Excellent"),
    ("Suburban",   str(loc_d.get("Suburban",0)),  "Fair"),
    ("Urban",      str(loc_d.get("Urban",0)),      "Good"),
    ("Rural",      str(loc_d.get("Rural",0)),      "Poor"),
    ("",           "",                             ""),
]):
    r = dist_tbl.rows[i]
    r.cells[0].text = row_data[0]
    r.cells[1].text = row_data[1]
    r.cells[2].text = row_data[2]
    if i == 0:
        for j in range(3):
            set_cell_bg(r.cells[j], "1D4ED8")
            r.cells[j].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            r.cells[j].paragraphs[0].runs[0].font.bold = True

doc.add_paragraph()
add_figure(doc, "fig_price_distribution.png",
               "Figure 1: House Price Distribution and Price by Condition", width_in=6.2)
add_figure(doc, "fig_location_analysis.png",
               "Figure 2: Average Price by Location and Location Distribution", width_in=6.2)
add_figure(doc, "fig_correlation_heatmap.png",
               "Figure 3: Feature Correlation Heatmap", width_in=5.5)

doc.add_page_break()

# ===========================================================================
# SECTION 4 — Sample Dataset Rows
# ===========================================================================
h = doc.add_heading("4. Dataset Sample (10 rows)", level=1)
style_heading(h, 16)

sample = api_out["dataset"]["sample_rows"][:10]
cols   = ["Area", "Bedrooms", "Bathrooms", "Floors", "YearBuilt", "Location", "Condition", "Garage", "Price"]
tbl_s  = doc.add_table(rows=len(sample)+1, cols=len(cols))
tbl_s.style = "Table Grid"

# Header row
for j, col in enumerate(cols):
    cell = tbl_s.rows[0].cells[j]
    cell.text = col
    set_cell_bg(cell, "1D4ED8")
    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    cell.paragraphs[0].runs[0].font.bold = True
    cell.paragraphs[0].runs[0].font.size = Pt(8)

# Data rows
for i, row in enumerate(sample):
    for j, col in enumerate(cols):
        cell = tbl_s.rows[i+1].cells[j]
        val  = row.get(col, "")
        if col == "Price":
            cell.text = f"${val:,.0f}" if isinstance(val, (int, float)) else str(val)
        else:
            cell.text = str(val)
        cell.paragraphs[0].runs[0].font.size = Pt(8)
        if i % 2 == 0:
            set_cell_bg(cell, "F7F8FA")

doc.add_page_break()

# ===========================================================================
# SECTION 5 — ML Model
# ===========================================================================
h = doc.add_heading("5. Machine Learning Model", level=1)
style_heading(h, 16)

doc.add_paragraph(
    "The model is a scikit-learn Pipeline combining preprocessing transformers "
    "and a Random Forest Regressor. Numeric features are normalized with StandardScaler; "
    "categorical features are encoded with OneHotEncoder."
)

h2 = doc.add_heading("5.1 Model Parameters", level=2)
mi = api_out["model_info"]
tbl_m = doc.add_table(rows=9, cols=2)
tbl_m.style = "Table Grid"
model_rows = [
    ("Parameter",          "Value"),
    ("Algorithm",          mi["algorithm"]),
    ("N Estimators",       str(mi["n_estimators"])),
    ("Max Depth",          str(mi["max_depth"])),
    ("Random State",       str(mi["random_state"])),
    ("Train Size",         f"{mi['train_size']:,} rows"),
    ("Test Size",          f"{mi['test_size']:,} rows"),
    ("Numeric Preprocessing",   mi["preprocessing"]["numeric"]),
    ("Categorical Encoding",    mi["preprocessing"]["categorical"]),
]
for i, (k, v) in enumerate(model_rows):
    add_kv_row(tbl_m, k, v, i, bg="1D4ED8" if i == 0 else ("EFF6FF" if i % 2 else "F7F8FA"))
    if i == 0:
        for cell in tbl_m.rows[0].cells:
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            cell.paragraphs[0].runs[0].font.bold = True

doc.add_paragraph()
h2 = doc.add_heading("5.2 Training Metrics", level=2)
metrics = mi["metrics"]
tbl_met = doc.add_table(rows=6, cols=2)
tbl_met.style = "Table Grid"
met_rows = [
    ("Metric",     "Value"),
    ("MAE",        f"${metrics['mae']:,.2f}"),
    ("RMSE",       f"${metrics['rmse']:,.2f}"),
    ("R2 Score",   f"{metrics['r2']:.4f}"),
    ("CV R2 Mean", f"{metrics['cv_r2_mean']:.4f}"),
    ("CV R2 Std",  f"{metrics['cv_r2_std']:.4f}"),
]
for i, (k, v) in enumerate(met_rows):
    add_kv_row(tbl_met, k, v, i, bg="1D4ED8" if i == 0 else ("EFF6FF" if i % 2 else "F7F8FA"))
    if i == 0:
        for cell in tbl_met.rows[0].cells:
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            cell.paragraphs[0].runs[0].font.bold = True

doc.add_paragraph()
h2 = doc.add_heading("5.3 Feature Importances", level=2)
fi_data = mi["top_features"]
tbl_fi  = doc.add_table(rows=len(fi_data)+1, cols=2)
tbl_fi.style = "Table Grid"
for j, hdr in enumerate(["Feature", "Importance"]):
    tbl_fi.rows[0].cells[j].text = hdr
    set_cell_bg(tbl_fi.rows[0].cells[j], "1D4ED8")
    tbl_fi.rows[0].cells[j].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tbl_fi.rows[0].cells[j].paragraphs[0].runs[0].font.bold = True
for i, fi in enumerate(fi_data):
    tbl_fi.rows[i+1].cells[0].text = fi["Feature"]
    tbl_fi.rows[i+1].cells[1].text = f"{fi['Importance']:.4f}"
    if i % 2 == 0:
        set_cell_bg(tbl_fi.rows[i+1].cells[0], "EFF6FF")
        set_cell_bg(tbl_fi.rows[i+1].cells[1], "EFF6FF")

doc.add_paragraph()
add_figure(doc, "fig_feature_importance.png",
               "Figure 4: Top 10 Feature Importances", width_in=5.8)
add_figure(doc, "fig_predicted_vs_actual.png",
               "Figure 5: Predicted vs Actual Prices and Residuals Distribution", width_in=6.2)

doc.add_page_break()

# ===========================================================================
# SECTION 6 — Flask API
# ===========================================================================
h = doc.add_heading("6. Flask REST API", level=1)
style_heading(h, 16)

doc.add_paragraph(
    "The backend is a Flask application (app.py) that exposes five REST endpoints. "
    "It loads the trained scikit-learn pipeline at startup and validates all inputs "
    "before making predictions. Flask-CORS allows cross-origin requests from the Streamlit frontend."
)

h2 = doc.add_heading("6.1 API Endpoints", level=2)
tbl_api = doc.add_table(rows=6, cols=3)
tbl_api.style = "Table Grid"
ep_data = [
    ("Method + Endpoint",   "Description",                    "Response"),
    ("GET /health",         "Service health check",           "{'status':'ok', 'model':...}"),
    ("GET /model-info",     "Model parameters and metrics",   "Algorithm, features, metrics"),
    ("GET /dataset",        "Dataset statistics and sample",  "Stats, distributions, rows"),
    ("POST /predict",       "Single house prediction",        "{'predicted_price': ...}"),
    ("POST /batch-predict", "Batch of house predictions",     "[{'predicted_price':...}, ...]"),
]
for i, (m, d, r) in enumerate(ep_data):
    row = tbl_api.rows[i]
    row.cells[0].text = m; row.cells[1].text = d; row.cells[2].text = r
    for j in range(3):
        row.cells[j].paragraphs[0].runs[0].font.bold = (i == 0)
        row.cells[j].paragraphs[0].runs[0].font.size = Pt(9)
        if i == 0:
            set_cell_bg(row.cells[j], "1D4ED8")
            row.cells[j].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        elif i % 2 == 0:
            set_cell_bg(row.cells[j], "EFF6FF")

doc.add_paragraph()
h2 = doc.add_heading("6.2 /health Output", level=2)
health_json = json.dumps(api_out["health"], indent=2)
p = doc.add_paragraph(health_json)
p.runs[0].font.name = "Courier New"
p.runs[0].font.size = Pt(9)

doc.add_paragraph()
h2 = doc.add_heading("6.3 /model-info Output (excerpt)", level=2)
mi_excerpt = {
    "algorithm": api_out["model_info"]["algorithm"],
    "n_estimators": api_out["model_info"]["n_estimators"],
    "metrics": api_out["model_info"]["metrics"],
    "numeric_features": api_out["model_info"]["numeric_features"],
}
p = doc.add_paragraph(json.dumps(mi_excerpt, indent=2))
p.runs[0].font.name = "Courier New"
p.runs[0].font.size = Pt(9)

doc.add_paragraph()
h2 = doc.add_heading("6.4 /predict Output", level=2)
p = doc.add_paragraph(json.dumps(api_out["single_predict"], indent=2))
p.runs[0].font.name = "Courier New"
p.runs[0].font.size = Pt(9)

doc.add_page_break()

# ===========================================================================
# SECTION 7 — Batch Prediction Results
# ===========================================================================
h = doc.add_heading("7. Batch Prediction Results", level=1)
style_heading(h, 16)

doc.add_paragraph(
    "The /batch-predict endpoint accepts a JSON array of house objects and returns predictions "
    "for each. Below are the three sample predictions from the API verification run:"
)
doc.add_paragraph()

batch = api_out["batch_predict"]
tbl_b = doc.add_table(rows=len(batch)+1, cols=9)
tbl_b.style = "Table Grid"
b_cols = ["#", "Area", "Beds", "Baths", "Floors", "Location", "Condition", "Garage", "Predicted Price"]
for j, hdr in enumerate(b_cols):
    tbl_b.rows[0].cells[j].text = hdr
    set_cell_bg(tbl_b.rows[0].cells[j], "1D4ED8")
    tbl_b.rows[0].cells[j].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tbl_b.rows[0].cells[j].paragraphs[0].runs[0].font.bold = True
    tbl_b.rows[0].cells[j].paragraphs[0].runs[0].font.size = Pt(9)

for i, item in enumerate(batch):
    inp = item["input"]
    vals = [str(i+1), str(inp["Area"]), str(inp["Bedrooms"]), str(inp["Bathrooms"]),
            str(inp["Floors"]), inp["Location"], inp["Condition"], inp["Garage"],
            f"${item['predicted_price']:,.2f}"]
    for j, v in enumerate(vals):
        tbl_b.rows[i+1].cells[j].text = v
        tbl_b.rows[i+1].cells[j].paragraphs[0].runs[0].font.size = Pt(9)
        if i % 2 == 0:
            set_cell_bg(tbl_b.rows[i+1].cells[j], "EFF6FF")
        if j == 8:  # price column bold
            tbl_b.rows[i+1].cells[j].paragraphs[0].runs[0].font.bold = True
            tbl_b.rows[i+1].cells[j].paragraphs[0].runs[0].font.color.rgb = RGBColor(0x1d, 0x4e, 0xd8)

doc.add_page_break()

# ===========================================================================
# SECTION 8 — Frontend UI
# ===========================================================================
h = doc.add_heading("8. Streamlit Frontend UI", level=1)
style_heading(h, 16)

doc.add_paragraph(
    "The frontend (ui.py) is built with Streamlit and Plotly. It connects to the Flask API "
    "and provides four interactive pages: Predict Price, Data Insights, Model Performance, and About."
)

h2 = doc.add_heading("8.1 Pages Overview", level=2)
tbl_p = doc.add_table(rows=5, cols=2)
tbl_p.style = "Table Grid"
pages = [
    ("Page",              "Description"),
    ("Predict Price",     "Input form for house features. Displays predicted price, gauge chart, and input summary."),
    ("Data Insights",     "4-tab dashboard: price distribution, correlations, location analysis, raw data table."),
    ("Model Performance", "Metrics KPIs, predicted vs actual scatter, residuals histogram, feature importances."),
    ("About",             "Architecture overview, endpoint reference, tech stack, and run instructions."),
]
for i, (k, v) in enumerate(pages):
    add_kv_row(tbl_p, k, v, i, bg="1D4ED8" if i == 0 else ("EFF6FF" if i % 2 else "F7F8FA"))
    if i == 0:
        for cell in tbl_p.rows[0].cells:
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            cell.paragraphs[0].runs[0].font.bold = True

doc.add_paragraph()
add_figure(doc, "ui_predict_page.png",
               "Figure 6: Predict Price Page — Input Form and Prediction Result", width_in=6.2)
add_figure(doc, "ui_insights_page.png",
               "Figure 7: Data Insights Page — Charts Dashboard", width_in=6.2)
add_figure(doc, "ui_performance_page.png",
               "Figure 8: Model Performance Page — Metrics and Scatter Plot", width_in=6.2)

doc.add_page_break()

# ===========================================================================
# SECTION 9 — How to Run
# ===========================================================================
h = doc.add_heading("9. How to Run", level=1)
style_heading(h, 16)

steps = [
    ("Step 1 — Install Dependencies",
     "py -3 -m pip install -r requirements.txt",
     "Installs all required Python packages."),
    ("Step 2 — Train the Model",
     "py -3 train_model.py",
     "Trains the Random Forest model and saves artifacts to model/ directory."),
    ("Step 3 — Start Flask API (Terminal 1)",
     "py -3 app.py",
     "Starts the REST API at http://127.0.0.1:5000"),
    ("Step 4 — Launch Streamlit UI (Terminal 2)",
     "py -3 -m streamlit run ui.py",
     "Opens the interactive UI at http://localhost:8501"),
    ("Step 5 — Generate Report (Optional)",
     "py -3 generate_report.py",
     "Creates this Word report at reports/house_price_report.docx"),
]
for label, cmd, desc in steps:
    h3 = doc.add_heading(label, level=2)
    p  = doc.add_paragraph(cmd)
    p.runs[0].font.name = "Courier New"
    p.runs[0].font.size = Pt(10)
    p.runs[0].font.bold = True
    p.runs[0].font.color.rgb = RGBColor(0x1d, 0x4e, 0xd8)
    doc.add_paragraph(desc)

doc.add_page_break()

# ===========================================================================
# SECTION 10 — Conclusion
# ===========================================================================
h = doc.add_heading("10. Conclusion", level=1)
style_heading(h, 16)

doc.add_paragraph(
    "This project demonstrates a complete end-to-end Machine Learning pipeline using Python. "
    "The system successfully trains a Random Forest model on 2,000 house records, exposes "
    "predictions through a clean REST API, and provides an interactive web interface for "
    "real-time predictions and data exploration."
)
doc.add_paragraph(
    "Key achievements:"
)
for item in [
    "Clean, modular Python codebase with clear separation of concerns",
    "Flask REST API with 5 documented endpoints including batch prediction",
    "Streamlit frontend with 4 interactive pages and Plotly visualisations",
    "Automated screenshot generation using matplotlib (8 figures)",
    "Comprehensive Word report generated programmatically with python-docx",
    "Dataset analysis covering 2,000 records, 4 locations, 4 conditions",
    "Feature importance analysis identifying Area and YearBuilt as key predictors",
]:
    p = doc.add_paragraph(f"  • {item}")
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)

doc.add_paragraph()
add_horizontal_rule(doc)
doc.add_paragraph()

footer_para = doc.add_paragraph()
footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
footer_run = footer_para.add_run(
    f"House Price Prediction System  |  Generated {datetime.now().strftime('%Y-%m-%d')}  |  Python + Flask + Streamlit + scikit-learn"
)
footer_run.font.size  = Pt(9)
footer_run.font.color.rgb = RGBColor(0x57, 0x60, 0x6a)
footer_run.font.italic = True

# ===========================================================================
# Save
# ===========================================================================
doc.save(OUT_PATH)
print(f"Report saved -> {OUT_PATH}")
