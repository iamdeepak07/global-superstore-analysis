# Global Superstore Analysis

Interactive data analysis dashboards for the Global Superstore dataset using Streamlit and Plotly.

## Features

- Sales and profit analysis by category, region, segment, and shipping mode
- Year-over-year trends and comparison views
- Top-performing sub-category insights
- Browser-based dashboard UI built with Streamlit

## Main Files

- `streamlit_dashboard.py` - Primary Streamlit dashboard
- `streamlit_dashboard_v2.py` - Alternate dashboard version
- `app.py` - Additional Streamlit app entry
- `Sales_Prediction_Dashboard.py` - Sales prediction-focused dashboard
- `global_superstore (1).xlsx` - Main dataset

## Setup

1. Clone the repository:

```bash
git clone https://github.com/iamdeepak07/global-superstore-analysis.git
cd global-superstore-analysis
```

2. Create and activate a virtual environment:

```bash
python -m venv .venv
.venv\Scripts\activate
```

3. Install dependencies:

```bash
pip install streamlit pandas plotly openpyxl scikit-learn numpy
```

## Run The Dashboard

```bash
streamlit run "streamlit_dashboard.py"
```

Then open `http://localhost:8501` in your browser.

## Notes

- Some generated HTML chart exports are included in the repository.
- If port `8501` is busy, run Streamlit with another port:

```bash
streamlit run "streamlit_dashboard.py" --server.port 8502
```
