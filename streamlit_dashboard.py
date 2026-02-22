import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import numpy as np

# Page configuration
st.set_page_config(
    page_title="Global Superstore Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
    <style>
        .main {
            padding-top: 2rem;
        }
        .metric-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
        }
        h1 {
            color: #667eea;
            text-align: center;
            margin-bottom: 30px;
        }
        h2 {
            color: #667eea;
            border-bottom: 3px solid #667eea;
            padding-bottom: 10px;
        }
    </style>
""", unsafe_allow_html=True)

# Load data
@st.cache_data
def load_data():
    file_path = r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\global_superstore (1).xlsx'
    df = pd.read_excel(file_path)
    df['Order Date'] = pd.to_datetime(df['Order Date'])
    return df

df = load_data()

# Title
st.markdown("<h1>📊 Global Superstore Sales Dashboard</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #666; font-size: 16px;'>Interactive Analytics & Sales Metrics (2012-2015)</p>", unsafe_allow_html=True)

# Sidebar filters
st.sidebar.header("🎛️ Filters & Controls")
selected_year = st.sidebar.multiselect(
    "Select Year(s):",
    options=sorted(df['Order Date'].dt.year.unique()),
    default=sorted(df['Order Date'].dt.year.unique())
)

selected_region = st.sidebar.multiselect(
    "Select Region(s):",
    options=df['Region'].unique(),
    default=df['Region'].unique()
)

selected_segment = st.sidebar.multiselect(
    "Select Segment(s):",
    options=df['Segment'].unique(),
    default=df['Segment'].unique()
)

# Filter data
filtered_df = df[
    (df['Order Date'].dt.year.isin(selected_year)) &
    (df['Region'].isin(selected_region)) &
    (df['Segment'].isin(selected_segment))
]

# KPI Row
st.markdown("## 📈 Key Performance Indicators")

# Check if filtered_df is empty
if len(filtered_df) == 0:
    st.warning("No data available for selected filters. Please adjust your selections.")
else:
    col1, col2, col3, col4, col5, col6 = st.columns(6)

    with col1:
        st.metric("Total Sales", f"${filtered_df['Sales'].sum():,.0f}")

    with col2:
        st.metric("Total Profit", f"${filtered_df['Profit'].sum():,.0f}")

    with col3:
        total_sales = filtered_df['Sales'].sum()
        if total_sales > 0:
            profit_margin = (filtered_df['Profit'].sum() / total_sales * 100)
        else:
            profit_margin = 0
        st.metric("Profit Margin", f"{profit_margin:.2f}%")

    with col4:
        avg_order = filtered_df['Sales'].mean()
        st.metric("Avg Order Value", f"${avg_order:,.0f}" if not pd.isna(avg_order) else "$0.00")

    with col5:
        total_orders = filtered_df['Order ID'].nunique()
        st.metric("Total Orders", f"{total_orders:,}")

    with col6:
        avg_quantity = filtered_df['Quantity'].mean()
        st.metric("Avg Quantity", f"{avg_quantity:.2f}" if not pd.isna(avg_quantity) else "0.00")

st.divider()

# Row 1: Sales by Category and Region
st.markdown("## 📊 Sales by Category & Region")

if len(filtered_df) > 0:
    col1, col2 = st.columns(2)

    with col1:
        # Sales by Category - Bar Chart
        category_sales = filtered_df.groupby('Category')['Sales'].sum().sort_values(ascending=False)
        if len(category_sales) > 0:
            fig1 = px.bar(
                x=category_sales.index,
                y=category_sales.values,
                title="Sales by Category",
                labels={'x': 'Category', 'y': 'Sales ($)'},
                color=category_sales.values,
                color_continuous_scale='Blues',
                text=[f"${x:,.0f}" for x in category_sales.values]
            )
            fig1.update_layout(height=400, hovermode='x unified', showlegend=False)
            st.plotly_chart(fig1, use_container_width=True)
        else:
            st.info("No data available for Sales by Category")

    with col2:
        # Sales by Region - Bar Chart
        region_sales = filtered_df.groupby('Region')['Sales'].sum().sort_values(ascending=False)
        if len(region_sales) > 0:
            fig2 = px.bar(
                x=region_sales.index,
                y=region_sales.values,
                title="Sales by Region",
                labels={'x': 'Region', 'y': 'Sales ($)'},
                color=region_sales.values,
                color_continuous_scale='Greens',
                text=[f"${x:,.0f}" for x in region_sales.values]
            )
            fig2.update_layout(height=400, hovermode='x unified', showlegend=False)
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("No data available for Sales by Region")
else:
    st.warning("No data available for selected filters. Please adjust your selections.")

st.divider()

# Row 2: Profit Analysis
st.markdown("## 💰 Profit Analysis")
col1, col2 = st.columns(2)

with col1:
    # Profit by Category
    category_profit = filtered_df.groupby('Category')['Profit'].sum().sort_values(ascending=False)
    fig3 = px.bar(
        x=category_profit.index,
        y=category_profit.values,
        title="Profit by Category",
        labels={'x': 'Category', 'y': 'Profit ($)'},
        color=category_profit.values,
        color_continuous_scale='Reds',
        text=[f"${x:,.0f}" for x in category_profit.values]
    )
    fig3.update_layout(height=400, hovermode='x unified', showlegend=False)
    st.plotly_chart(fig3, use_container_width=True)

with col2:
    # Profit by Region
    region_profit = filtered_df.groupby('Region')['Profit'].sum().sort_values(ascending=False)
    fig4 = px.bar(
        x=region_profit.index,
        y=region_profit.values,
        title="Profit by Region",
        labels={'x': 'Region', 'y': 'Profit ($)'},
        color=region_profit.values,
        color_continuous_scale='Purples',
        text=[f"${x:,.0f}" for x in region_profit.values]
    )
    fig4.update_layout(height=400, hovermode='x unified', showlegend=False)
    st.plotly_chart(fig4, use_container_width=True)

st.divider()

# Row 3: Segment & Ship Mode Analysis
st.markdown("## 🎯 Segment & Ship Mode Analysis")
col1, col2 = st.columns(2)

with col1:
    # Sales by Segment
    segment_sales = filtered_df.groupby('Segment')['Sales'].sum().sort_values(ascending=False)
    fig5 = px.bar(
        x=segment_sales.index,
        y=segment_sales.values,
        title="Sales by Customer Segment",
        labels={'x': 'Segment', 'y': 'Sales ($)'},
        color=segment_sales.values,
        color_continuous_scale='Viridis',
        text=[f"${x:,.0f}" for x in segment_sales.values]
    )
    fig5.update_layout(height=400, hovermode='x unified', showlegend=False)
    st.plotly_chart(fig5, use_container_width=True)

with col2:
    # Sales by Ship Mode
    shipmode_sales = filtered_df.groupby('Ship Mode')['Sales'].sum().sort_values(ascending=False)
    fig6 = px.bar(
        x=shipmode_sales.index,
        y=shipmode_sales.values,
        title="Sales by Ship Mode",
        labels={'x': 'Ship Mode', 'y': 'Sales ($)'},
        color=shipmode_sales.values,
        color_continuous_scale='Plasma',
        text=[f"${x:,.0f}" for x in shipmode_sales.values]
    )
    fig6.update_layout(height=400, hovermode='x unified', showlegend=False)
    st.plotly_chart(fig6, use_container_width=True)

st.divider()

# Row 4: Yearly Analysis
st.markdown("## 📅 Yearly Sales Analysis")
col1, col2 = st.columns(2)

with col1:
    # Yearly Sales
    yearly_sales = filtered_df.groupby(filtered_df['Order Date'].dt.year).agg({
        'Sales': 'sum',
        'Profit': 'sum'
    }).reset_index()
    yearly_sales.columns = ['Year', 'Sales', 'Profit']
    
    fig7 = px.bar(
        x=yearly_sales['Year'],
        y=yearly_sales['Sales'],
        title="Total Sales by Year",
        labels={'x': 'Year', 'y': 'Sales ($)'},
        color=yearly_sales['Sales'],
        color_continuous_scale='Turbo',
        text=[f"${x:,.0f}" for x in yearly_sales['Sales']]
    )
    fig7.update_layout(height=400, hovermode='x unified', showlegend=False)
    st.plotly_chart(fig7, use_container_width=True)

with col2:
    # Yearly Profit
    fig8 = px.bar(
        x=yearly_sales['Year'],
        y=yearly_sales['Profit'],
        title="Total Profit by Year",
        labels={'x': 'Year', 'y': 'Profit ($)'},
        color=yearly_sales['Profit'],
        color_continuous_scale='RdYlGn',
        text=[f"${x:,.0f}" for x in yearly_sales['Profit']]
    )
    fig8.update_layout(height=400, hovermode='x unified', showlegend=False)
    st.plotly_chart(fig8, use_container_width=True)

st.divider()

# Row 5: Sub-Category Performance
st.markdown("## 🏷️ Top Sub-Categories by Sales")
col1 = st.columns(1)[0]

top_subcategories = filtered_df.groupby('Sub-Category')['Sales'].sum().sort_values(ascending=False).head(12)
fig9 = px.bar(
    x=top_subcategories.values,
    y=top_subcategories.index,
    orientation='h',
    title="Top 12 Sub-Categories by Sales",
    labels={'x': 'Sales ($)', 'y': 'Sub-Category'},
    color=top_subcategories.values,
    color_continuous_scale='Sunset',
    text=[f"${x:,.0f}" for x in top_subcategories.values]
)
fig9.update_layout(height=500, hovermode='y unified', showlegend=False)
st.plotly_chart(fig9, use_container_width=True)

st.divider()

# Row 6: Orders & Discount Analysis
st.markdown("## 📦 Orders & Discount Analysis")
col1, col2 = st.columns(2)

with col1:
    # Orders by Category
    orders_category = filtered_df.groupby('Category')['Order ID'].nunique().sort_values(ascending=False)
    fig10 = px.bar(
        x=orders_category.index,
        y=orders_category.values,
        title="Number of Orders by Category",
        labels={'x': 'Category', 'y': 'Order Count'},
        color=orders_category.values,
        color_continuous_scale='Inferno',
        text=orders_category.values
    )
    fig10.update_layout(height=400, hovermode='x unified', showlegend=False)
    st.plotly_chart(fig10, use_container_width=True)

with col2:
    # Average Discount by Category
    avg_discount = filtered_df.groupby('Category')['Discount'].mean().sort_values(ascending=False)
    fig11 = px.bar(
        x=avg_discount.index,
        y=avg_discount.values * 100,
        title="Average Discount by Category (%)",
        labels={'x': 'Category', 'y': 'Discount (%)'},
        color=avg_discount.values * 100,
        color_continuous_scale='RdYlBu_r',
        text=[f"{x:.1f}%" for x in avg_discount.values * 100]
    )
    fig11.update_layout(height=400, hovermode='x unified', showlegend=False)
    st.plotly_chart(fig11, use_container_width=True)

st.divider()

# Data Table
st.markdown("## 📋 Detailed Data View")
with st.expander("View Raw Data", expanded=False):
    st.dataframe(
        filtered_df[['Order Date', 'Category', 'Sub-Category', 'Sales', 'Profit', 'Region', 'Segment']].sort_values('Order Date', ascending=False),
        use_container_width=True,
        height=400
    )

# Summary Statistics
st.markdown("## 📊 Summary Statistics")
col1, col2, col3 = st.columns(3)

with col1:
    st.write("**Sales Statistics**")
    stats_sales = filtered_df['Sales'].describe()
    st.write(f"Mean: ${stats_sales['mean']:,.2f}")
    st.write(f"Std Dev: ${stats_sales['std']:,.2f}")
    st.write(f"Min: ${stats_sales['min']:,.2f}")
    st.write(f"Max: ${stats_sales['max']:,.2f}")

with col2:
    st.write("**Profit Statistics**")
    stats_profit = filtered_df['Profit'].describe()
    st.write(f"Mean: ${stats_profit['mean']:,.2f}")
    st.write(f"Std Dev: ${stats_profit['std']:,.2f}")
    st.write(f"Min: ${stats_profit['min']:,.2f}")
    st.write(f"Max: ${stats_profit['max']:,.2f}")

with col3:
    st.write("**Order Statistics**")
    st.write(f"Total Orders: {filtered_df['Order ID'].nunique():,}")
    st.write(f"Avg Quantity: {filtered_df['Quantity'].mean():.2f}")
    st.write(f"Total Quantity: {filtered_df['Quantity'].sum():,}")
    st.write(f"Unique Customers: {filtered_df['Customer ID'].nunique():,}")

st.divider()

# Footer
st.markdown("""
    <div style='text-align: center; margin-top: 50px; padding: 20px; color: #999; border-top: 1px solid #ddd;'>
        <p><strong>Global Superstore Sales Dashboard</strong></p>
        <p>Built with Streamlit | Data: 2012-2015 | Last Updated: February 2026</p>
    </div>
""", unsafe_allow_html=True)
