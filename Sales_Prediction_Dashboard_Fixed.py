import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import mean_squared_error, r2_score, mean_absolute_error
import warnings
warnings.filterwarnings('ignore')

# ==================== LOAD DATA ====================
file_path = r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\global_superstore (1).xlsx'
df = pd.read_excel(file_path)

print("=" * 60)
print("GLOBAL SUPERSTORE - SALES PREDICTION DASHBOARD")
print("=" * 60)
print(f"\nDataset loaded: {df.shape[0]} rows, {df.shape[1]} columns")

# ==================== DATA PREPROCESSING ====================
df['Order Date'] = pd.to_datetime(df['Order Date'])
df['Ship Date'] = pd.to_datetime(df['Ship Date'])
df['Year'] = df['Order Date'].dt.year
df['Month'] = df['Order Date'].dt.month
df['Quarter'] = df['Order Date'].dt.quarter
df['DayOfWeek'] = df['Order Date'].dt.dayofweek
df['Days_to_Ship'] = (df['Ship Date'] - df['Order Date']).dt.days

# Monthly aggregation for time series analysis
monthly_sales = df.groupby(df['Order Date'].dt.to_period('M')).agg({
    'Sales': 'sum',
    'Profit': 'sum',
    'Quantity': 'sum',
    'Order ID': 'count'
}).reset_index()
monthly_sales['Order Date'] = monthly_sales['Order Date'].dt.to_timestamp()
monthly_sales.columns = ['Order Date', 'Total Sales', 'Total Profit', 'Total Quantity', 'Orders']

print("\n" + "=" * 60)
print("KEY METRICS OVERVIEW")
print("=" * 60)
print(f"Total Sales: ${df['Sales'].sum():,.2f}")
print(f"Total Profit: ${df['Profit'].sum():,.2f}")
print(f"Average Order Value: ${df['Sales'].mean():,.2f}")
print(f"Profit Margin: {(df['Profit'].sum() / df['Sales'].sum() * 100):.2f}%")
print(f"Total Orders: {df['Order ID'].nunique():,}")
print(f"Date Range: {df['Order Date'].min().date()} to {df['Order Date'].max().date()}")

# ==================== VISUALIZATION 1: Sales Trend ====================
fig1 = go.Figure()
fig1.add_trace(go.Scatter(
    x=monthly_sales['Order Date'],
    y=monthly_sales['Total Sales'],
    name='Total Sales',
    line=dict(color='#2E86AB', width=3),
    fill='tozeroy',
    fillcolor='rgba(46, 134, 171, 0.2)'
))
fig1.update_layout(
    title='Monthly Sales Trend',
    xaxis_title='Date',
    yaxis_title='Sales ($)',
    template='plotly_white',
    height=500,
    hovermode='x unified'
)
fig1.write_html(r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\1_sales_trend.html')
print("\n[OK] Sales Trend chart saved")

# ==================== VISUALIZATION 2: Sales by Category ====================
category_sales = df.groupby('Category')['Sales'].sum().sort_values(ascending=False)
fig2 = px.bar(
    x=category_sales.values,
    y=category_sales.index,
    orientation='h',
    title='Total Sales by Category',
    labels={'x': 'Sales ($)', 'y': 'Category'},
    color=category_sales.values,
    color_continuous_scale='Blues'
)
fig2.update_layout(height=400, template='plotly_white')
fig2.write_html(r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\2_sales_by_category.html')
print("[OK] Sales by Category chart saved")

# ==================== VISUALIZATION 3: Sales by Region ====================
region_sales = df.groupby('Region').agg({'Sales': 'sum', 'Profit': 'sum'}).sort_values('Sales', ascending=False)
fig3 = make_subplots(
    rows=1, cols=2,
    specs=[[{'type': 'pie'}, {'type': 'bar'}]],
    subplot_titles=('Sales Distribution by Region', 'Profit by Region')
)
fig3.add_trace(
    go.Pie(labels=region_sales.index, values=region_sales['Sales'], name='Sales'),
    row=1, col=1
)
fig3.add_trace(
    go.Bar(x=region_sales.index, y=region_sales['Profit'], name='Profit', marker_color='#A23B72'),
    row=1, col=2
)
fig3.update_layout(height=450, template='plotly_white', showlegend=False)
fig3.write_html(r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\3_sales_by_region.html')
print("[OK] Sales by Region chart saved")

# ==================== VISUALIZATION 4: Segment Analysis ====================
segment_data = df.groupby('Segment').agg({
    'Sales': 'sum',
    'Profit': 'sum',
    'Order ID': 'count'
}).reset_index()
segment_data.columns = ['Segment', 'Sales', 'Profit', 'Orders']

fig4 = make_subplots(
    rows=1, cols=2,
    specs=[[{'type': 'bar'}, {'type': 'bar'}]],
    subplot_titles=('Sales by Segment', 'Profit by Segment')
)
fig4.add_trace(
    go.Bar(x=segment_data['Segment'], y=segment_data['Sales'], name='Sales', marker_color='#06A77D'),
    row=1, col=1
)
fig4.add_trace(
    go.Bar(x=segment_data['Segment'], y=segment_data['Profit'], name='Profit', marker_color='#D62828'),
    row=1, col=2
)
fig4.update_layout(height=400, template='plotly_white', showlegend=False)
fig4.write_html(r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\4_segment_analysis.html')
print("[OK] Segment Analysis chart saved")

# ==================== MACHINE LEARNING: SALES PREDICTION ====================
print("\n" + "=" * 60)
print("BUILDING SALES PREDICTION MODEL")
print("=" * 60)

# Prepare data for model
df_model = df.copy()
le_dict = {}

# Encode categorical variables
categorical_cols = ['Ship Mode', 'Segment', 'Country', 'Region', 'Category', 'Order Priority']
for col in categorical_cols:
    le = LabelEncoder()
    df_model[col + '_encoded'] = le.fit_transform(df_model[col])
    le_dict[col] = le

# Select features
feature_cols = ['Quantity', 'Discount', 'Shipping Cost', 'Year', 'Month', 'Quarter', 
                'DayOfWeek', 'Days_to_Ship', 'Ship Mode_encoded', 'Segment_encoded', 
                'Region_encoded', 'Category_encoded', 'Order Priority_encoded']

X = df_model[feature_cols]
y = df_model['Sales']

# Train-test split
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Train Random Forest model
model = RandomForestRegressor(n_estimators=100, max_depth=20, random_state=42, n_jobs=-1)
model.fit(X_train, y_train)

# Predictions
y_pred = model.predict(X_test)

# Model evaluation
mse = mean_squared_error(y_test, y_pred)
rmse = np.sqrt(mse)
mae = mean_absolute_error(y_test, y_pred)
r2 = r2_score(y_test, y_pred)

print(f"\nModel Performance:")
print(f"  R2 Score: {r2:.4f}")
print(f"  RMSE: ${rmse:,.2f}")
print(f"  MAE: ${mae:,.2f}")

# Feature importance
feature_importance = pd.DataFrame({
    'Feature': feature_cols,
    'Importance': model.feature_importances_
}).sort_values('Importance', ascending=False)

print(f"\nTop 5 Important Features:")
for idx, row in feature_importance.head(5).iterrows():
    print(f"  {row['Feature']}: {row['Importance']:.4f}")

# ==================== VISUALIZATION 5: Feature Importance ====================
fig5 = px.bar(
    feature_importance.head(10),
    x='Importance',
    y='Feature',
    orientation='h',
    title='Top 10 Features Influencing Sales',
    labels={'Importance': 'Importance Score', 'Feature': 'Features'},
    color='Importance',
    color_continuous_scale='Viridis'
)
fig5.update_layout(height=500, template='plotly_white')
fig5.write_html(r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\5_feature_importance.html')
print("\n[OK] Feature Importance chart saved")

# ==================== VISUALIZATION 6: Actual vs Predicted ====================
fig6 = go.Figure()
fig6.add_trace(go.Scatter(
    x=y_test.values[:1000],
    y=y_pred[:1000],
    mode='markers',
    name='Predictions',
    marker=dict(size=5, color='#E63946', opacity=0.6)
))
fig6.add_trace(go.Scatter(
    x=[y_test.values.min(), y_test.values.max()],
    y=[y_test.values.min(), y_test.values.max()],
    mode='lines',
    name='Perfect Prediction',
    line=dict(color='green', dash='dash')
))
fig6.update_layout(
    title=f'Actual vs Predicted Sales (R2 = {r2:.4f})',
    xaxis_title='Actual Sales ($)',
    yaxis_title='Predicted Sales ($)',
    height=500,
    template='plotly_white',
    hovermode='closest'
)
fig6.write_html(r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\6_actual_vs_predicted.html')
print("[OK] Actual vs Predicted chart saved")

# ==================== VISUALIZATION 7: Profit Margin by Category ====================
df['Profit_Margin'] = (df['Profit'] / df['Sales'] * 100).round(2)
margin_by_cat = df.groupby('Category')['Profit_Margin'].mean().sort_values(ascending=False)

fig7 = px.bar(
    x=margin_by_cat.index,
    y=margin_by_cat.values,
    title='Average Profit Margin by Category (%)',
    labels={'x': 'Category', 'y': 'Profit Margin (%)'},
    color=margin_by_cat.values,
    color_continuous_scale='RdYlGn'
)
fig7.update_layout(height=400, template='plotly_white')
fig7.write_html(r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\7_profit_margin.html')
print("[OK] Profit Margin chart saved")

# ==================== VISUALIZATION 8: Monthly Profit Trend ====================
fig8 = go.Figure()
fig8.add_trace(go.Scatter(
    x=monthly_sales['Order Date'],
    y=monthly_sales['Total Profit'],
    name='Monthly Profit',
    line=dict(color='#F77F00', width=3),
    fill='tozeroy',
    fillcolor='rgba(247, 127, 0, 0.2)'
))
fig8.update_layout(
    title='Monthly Profit Trend',
    xaxis_title='Date',
    yaxis_title='Profit ($)',
    template='plotly_white',
    height=500,
    hovermode='x unified'
)
fig8.write_html(r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\8_profit_trend.html')
print("[OK] Profit Trend chart saved")

# ==================== CREATE SUMMARY REPORT ====================
summary_html = f"""
<!DOCTYPE html>
<html>
<head>
    <title>Global Superstore - Sales Prediction Dashboard</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        }}
        .container {{
            background: white;
            border-radius: 10px;
            padding: 30px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.3);
        }}
        h1 {{
            color: #333;
            border-bottom: 3px solid #667eea;
            padding-bottom: 10px;
        }}
        h2 {{
            color: #667eea;
            margin-top: 30px;
        }}
        .metrics {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }}
        .metric-card {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
        }}
        .metric-value {{
            font-size: 28px;
            font-weight: bold;
            margin: 10px 0;
        }}
        .metric-label {{
            font-size: 14px;
            opacity: 0.9;
        }}
        .dashboard-link {{
            display: inline-block;
            background: #667eea;
            color: white;
            padding: 10px 20px;
            margin: 5px;
            border-radius: 5px;
            text-decoration: none;
            transition: background 0.3s;
        }}
        .dashboard-link:hover {{
            background: #764ba2;
        }}
        .model-stats {{
            background: #f8f9fa;
            padding: 20px;
            border-left: 4px solid #667eea;
            border-radius: 5px;
            margin: 20px 0;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }}
        th, td {{
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }}
        th {{
            background: #667eea;
            color: white;
        }}
        tr:hover {{
            background: #f5f5f5;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>GLOBAL SUPERSTORE - SALES PREDICTION DASHBOARD</h1>
        
        <h2>KEY PERFORMANCE INDICATORS</h2>
        <div class="metrics">
            <div class="metric-card">
                <div class="metric-label">Total Sales</div>
                <div class="metric-value">${df['Sales'].sum():,.0f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Profit</div>
                <div class="metric-value">${df['Profit'].sum():,.0f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Profit Margin</div>
                <div class="metric-value">{(df['Profit'].sum() / df['Sales'].sum() * 100):.2f}%</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Average Order Value</div>
                <div class="metric-value">${df['Sales'].mean():,.0f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Orders</div>
                <div class="metric-value">{df['Order ID'].nunique():,}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Data Period</div>
                <div class="metric-value">{df['Order Date'].dt.year.min()}-{df['Order Date'].dt.year.max()}</div>
            </div>
        </div>

        <h2>INTERACTIVE VISUALIZATIONS</h2>
        <p>Click below to view detailed dashboards:</p>
        <div>
            <a href="1_sales_trend.html" class="dashboard-link">Sales Trend</a>
            <a href="2_sales_by_category.html" class="dashboard-link">Sales by Category</a>
            <a href="3_sales_by_region.html" class="dashboard-link">Sales by Region</a>
            <a href="4_segment_analysis.html" class="dashboard-link">Segment Analysis</a>
            <a href="5_feature_importance.html" class="dashboard-link">Feature Importance</a>
            <a href="6_actual_vs_predicted.html" class="dashboard-link">Model Performance</a>
            <a href="7_profit_margin.html" class="dashboard-link">Profit Margin</a>
            <a href="8_profit_trend.html" class="dashboard-link">Profit Trend</a>
        </div>

        <h2>MACHINE LEARNING MODEL - SALES PREDICTION</h2>
        <div class="model-stats">
            <h3>Model Performance Metrics</h3>
            <table>
                <tr>
                    <th>Metric</th>
                    <th>Value</th>
                    <th>Interpretation</th>
                </tr>
                <tr>
                    <td>R2 Score</td>
                    <td><strong>{r2:.4f}</strong></td>
                    <td>Model explains {r2*100:.2f}% of sales variance</td>
                </tr>
                <tr>
                    <td>RMSE (Root Mean Squared Error)</td>
                    <td><strong>${rmse:,.2f}</strong></td>
                    <td>Average prediction error</td>
                </tr>
                <tr>
                    <td>MAE (Mean Absolute Error)</td>
                    <td><strong>${mae:,.2f}</strong></td>
                    <td>Average absolute prediction error</td>
                </tr>
                <tr>
                    <td>Test Set Size</td>
                    <td><strong>{len(X_test):,}</strong></td>
                    <td>Samples used for validation (20% of data)</td>
                </tr>
            </table>

            <h3>Top Features Influencing Sales</h3>
            <table>
                <tr>
                    <th>Rank</th>
                    <th>Feature</th>
                    <th>Importance Score</th>
                </tr>
"""

for i, (idx, row) in enumerate(feature_importance.head(10).iterrows(), 1):
    summary_html += f"""
                <tr>
                    <td>{i}</td>
                    <td>{row['Feature']}</td>
                    <td>{row['Importance']:.4f}</td>
                </tr>
"""

summary_html += f"""
            </table>
        </div>

        <h2>DATA SUMMARY</h2>
        <table>
            <tr>
                <th>Metric</th>
                <th>Value</th>
            </tr>
            <tr>
                <td>Total Records</td>
                <td>{len(df):,}</td>
            </tr>
            <tr>
                <td>Date Range</td>
                <td>{df['Order Date'].min().date()} to {df['Order Date'].max().date()}</td>
            </tr>
            <tr>
                <td>Countries</td>
                <td>{df['Country'].nunique()}</td>
            </tr>
            <tr>
                <td>Regions</td>
                <td>{df['Region'].nunique()}</td>
            </tr>
            <tr>
                <td>Categories</td>
                <td>{df['Category'].nunique()}</td>
            </tr>
            <tr>
                <td>Customer Segments</td>
                <td>{df['Segment'].nunique()}</td>
            </tr>
        </table>

        <h2>KEY INSIGHTS</h2>
        <ul>
            <li><strong>Top Category:</strong> {df.groupby('Category')['Sales'].sum().idxmax()} with ${df.groupby('Category')['Sales'].sum().max():,.0f} in sales</li>
            <li><strong>Best Performing Region:</strong> {df.groupby('Region')['Profit'].sum().idxmax()} with ${df.groupby('Region')['Profit'].sum().max():,.0f} in profit</li>
            <li><strong>Most Profitable Segment:</strong> {df.groupby('Segment')['Profit'].sum().idxmax()}</li>
            <li><strong>Average Order Value:</strong> ${df['Sales'].mean():,.2f}</li>
            <li><strong>Model Accuracy:</strong> The Random Forest model achieves {r2*100:.2f}% accuracy in predicting sales</li>
        </ul>

        <footer style="text-align: center; margin-top: 40px; color: #999; border-top: 1px solid #ddd; padding-top: 20px;">
            <p>Generated using Python - Data Analysis Dashboard</p>
        </footer>
    </div>
</body>
</html>
"""

with open(r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\INDEX.html', 'w') as f:
    f.write(summary_html)

print("\n[OK] Summary Dashboard (INDEX.html) created")

print("\n" + "=" * 60)
print("[COMPLETE] DASHBOARD CREATED SUCCESSFULLY!")
print("=" * 60)
print("\nAll files saved to: c:\\Users\\ASUS\\OneDrive\\Documents\\Data Analyst\\")
print("\nTo view the dashboard, open: INDEX.html")
