import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Load data
file_path = r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\global_superstore (1).xlsx'
df = pd.read_excel(file_path)
df['Order Date'] = pd.to_datetime(df['Order Date'])

# Get yearly sales
yearly_sales = df.groupby(df['Order Date'].dt.year).agg({
    'Sales': 'sum',
    'Profit': 'sum',
    'Order ID': 'count',
    'Quantity': 'sum'
}).reset_index()
yearly_sales.columns = ['Year', 'Total_Sales', 'Total_Profit', 'Orders', 'Quantity']

# Calculate YoY changes
yearly_sales['YoY_Change'] = yearly_sales['Total_Sales'].diff()
yearly_sales['YoY_Percent'] = yearly_sales['Total_Sales'].pct_change() * 100

print("\nYearly Sales Analysis:")
print(yearly_sales)

# Create responsive chart
fig = make_subplots(
    rows=2, cols=1,
    subplot_titles=('Yearly Sales Trend with Increase/Decrease', 'Year-over-Year Sales Change (%)'),
    specs=[[{'secondary_y': True}], [{'secondary_y': False}]],
    vertical_spacing=0.15,
    row_heights=[0.6, 0.4]
)

# Chart 1: Sales trend with bar colors
colors = ['green' if x > 0 else 'red' for x in yearly_sales['YoY_Change'].fillna(0)]
colors[0] = 'lightblue'  # First year neutral color

fig.add_trace(
    go.Bar(
        x=yearly_sales['Year'],
        y=yearly_sales['Total_Sales'],
        name='Total Sales',
        marker=dict(color=colors, line=dict(color='darkblue', width=2)),
        text=[f"${x:,.0f}" for x in yearly_sales['Total_Sales']],
        textposition='outside',
        hovertemplate='<b>Year: %{x}</b><br>Sales: $%{y:,.0f}<extra></extra>'
    ),
    row=1, col=1, secondary_y=False
)

# Add line chart for trend
fig.add_trace(
    go.Scatter(
        x=yearly_sales['Year'],
        y=yearly_sales['Total_Sales'],
        mode='lines+markers',
        name='Sales Trend',
        line=dict(color='darkblue', width=3),
        marker=dict(size=10, symbol='circle'),
        hovertemplate='<b>Year: %{x}</b><br>Sales: $%{y:,.0f}<extra></extra>'
    ),
    row=1, col=1, secondary_y=False
)

# Chart 2: YoY percentage change
yoy_colors = ['green' if x > 0 else 'red' for x in yearly_sales['YoY_Percent'].fillna(0)]

fig.add_trace(
    go.Bar(
        x=yearly_sales['Year'][1:],
        y=yearly_sales['YoY_Percent'][1:],
        name='YoY Change %',
        marker=dict(color=yoy_colors[1:]),
        text=[f"{x:.1f}%" for x in yearly_sales['YoY_Percent'][1:]],
        textposition='outside',
        hovertemplate='<b>Year: %{x}</b><br>Change: %{y:.1f}%<extra></extra>'
    ),
    row=2, col=1
)

# Update layout
fig.update_layout(
    title=dict(
        text='<b>YEARLY SALES TREND ANALYSIS</b><br><sub>Sales Growth and Decline Year-over-Year</sub>',
        x=0.5,
        xanchor='center',
        font=dict(size=24)
    ),
    height=800,
    showlegend=True,
    template='plotly_white',
    hovermode='x unified',
    font=dict(size=12),
)

# Update y axes
fig.update_yaxes(title_text='Sales ($)', row=1, col=1, tickformat='$,.0f')
fig.update_yaxes(title_text='YoY Change (%)', row=2, col=1)
fig.update_xaxes(title_text='Year', row=2, col=1)

fig.write_html(r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\9_yearly_sales_trend.html')
print("\n[OK] Yearly Sales Trend chart saved!")

# Create a detailed summary chart
fig2 = go.Figure()

# Add sales bars
fig2.add_trace(
    go.Bar(
        x=yearly_sales['Year'],
        y=yearly_sales['Total_Sales'],
        name='Total Sales',
        marker=dict(
            color=yearly_sales['Total_Sales'],
            colorscale='Blues',
            showscale=True,
            colorbar=dict(title="Sales ($)")
        ),
        text=[f"${x/1e6:.2f}M" for x in yearly_sales['Total_Sales']],
        textposition='outside',
        hovertemplate='<b>Year %{x}</b><br>Sales: $%{y:,.0f}<br><extra></extra>'
    )
)

# Add change indicators
change_text = []
for i in range(len(yearly_sales)):
    if i == 0:
        change_text.append('Start Year')
    else:
        change_pct = yearly_sales['YoY_Percent'].iloc[i]
        if change_pct > 0:
            change_text.append(f'▲ +{change_pct:.1f}%')
        elif change_pct < 0:
            change_text.append(f'▼ {change_pct:.1f}%')
        else:
            change_text.append('→ 0%')

fig2.update_layout(
    title=dict(
        text='<b>ANNUAL SALES PERFORMANCE</b><br><sub>Green = Increase | Red = Decrease from Previous Year</sub>',
        x=0.5,
        xanchor='center',
        font=dict(size=24)
    ),
    xaxis_title='Year',
    yaxis_title='Total Sales ($)',
    height=600,
    template='plotly_white',
    hovermode='x unified',
    font=dict(size=12),
)

fig2.update_yaxes(tickformat='$,.0f')

# Add annotations for change percentages
for i, year in enumerate(yearly_sales['Year']):
    fig2.add_annotation(
        x=year,
        y=yearly_sales['Total_Sales'].iloc[i],
        text=change_text[i],
        showarrow=True,
        arrowhead=2,
        arrowsize=1,
        arrowwidth=2,
        arrowcolor='darkblue',
        ax=0,
        ay=-40,
        font=dict(
            size=11,
            color='darkblue'
        ),
        bgcolor='rgba(255, 255, 255, 0.8)',
        bordercolor='darkblue',
        borderwidth=1
    )

fig2.write_html(r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\10_annual_sales_performance.html')
print("[OK] Annual Sales Performance chart saved!")

# Create comparison chart: Sales vs Profit vs Orders
fig3 = make_subplots(
    rows=1, cols=3,
    subplot_titles=('Total Sales by Year', 'Total Profit by Year', 'Total Orders by Year'),
    specs=[[{'secondary_y': False}, {'secondary_y': False}, {'secondary_y': False}]]
)

fig3.add_trace(
    go.Bar(x=yearly_sales['Year'], y=yearly_sales['Total_Sales'], name='Sales', marker_color='#636EFA'),
    row=1, col=1
)

fig3.add_trace(
    go.Bar(x=yearly_sales['Year'], y=yearly_sales['Total_Profit'], name='Profit', marker_color='#00CC96'),
    row=1, col=2
)

fig3.add_trace(
    go.Bar(x=yearly_sales['Year'], y=yearly_sales['Orders'], name='Orders', marker_color='#AB63FA'),
    row=1, col=3
)

fig3.update_yaxes(title_text='Sales ($)', tickformat='$,.0f', row=1, col=1)
fig3.update_yaxes(title_text='Profit ($)', tickformat='$,.0f', row=1, col=2)
fig3.update_yaxes(title_text='Order Count', row=1, col=3)

fig3.update_xaxes(title_text='Year', row=1, col=1)
fig3.update_xaxes(title_text='Year', row=1, col=2)
fig3.update_xaxes(title_text='Year', row=1, col=3)

fig3.update_layout(
    title_text='<b>YEARLY BUSINESS METRICS COMPARISON</b>',
    height=500,
    template='plotly_white',
    showlegend=False,
    font=dict(size=12)
)

fig3.write_html(r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\11_yearly_metrics_comparison.html')
print("[OK] Yearly Metrics Comparison chart saved!")

print("\n" + "=" * 60)
print("ALL YEARLY ANALYSIS CHARTS CREATED SUCCESSFULLY!")
print("=" * 60)
print("\nCharts saved:")
print("  1. 9_yearly_sales_trend.html - Dual chart with sales trend & YoY change")
print("  2. 10_annual_sales_performance.html - Annual sales with indicators")
print("  3. 11_yearly_metrics_comparison.html - Sales vs Profit vs Orders")
