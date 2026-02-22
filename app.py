import streamlit as st
import pandas as pd
import plotly.express as px
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title="Superstore Dashboard", layout="wide")

st.markdown("<h1 style='text-align: center; color: #667eea;'>📊 Global Superstore Dashboard</h1>", unsafe_allow_html=True)

# Load data
try:
    df = pd.read_excel(r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\global_superstore (1).xlsx')
    df['Order Date'] = pd.to_datetime(df['Order Date'])
    st.success("✅ Data loaded successfully!")
except Exception as e:
    st.error(f"❌ Error loading data: {e}")
    st.stop()

# Sidebar filters
st.sidebar.header("Filters")
years = sorted(df['Order Date'].dt.year.unique())
selected_year = st.sidebar.multiselect("Year", years, default=years)

regions = sorted(df['Region'].unique())
selected_region = st.sidebar.multiselect("Region", regions, default=regions)

# Filter
filtered_df = df[(df['Order Date'].dt.year.isin(selected_year)) & (df['Region'].isin(selected_region))]

if len(filtered_df) == 0:
    st.warning("No data for selected filters")
    st.stop()

# KPIs
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Sales", f"${filtered_df['Sales'].sum():,.0f}")
col2.metric("Total Profit", f"${filtered_df['Profit'].sum():,.0f}")
col3.metric("Orders", f"{filtered_df['Order ID'].nunique():,}")
col4.metric("Avg Order", f"${filtered_df['Sales'].mean():,.0f}")

st.divider()

# Charts
st.subheader("Sales Analysis")
col1, col2 = st.columns(2)

with col1:
    cat_sales = filtered_df.groupby('Category')['Sales'].sum().sort_values(ascending=False)
    fig1 = px.bar(x=cat_sales.index, y=cat_sales.values, title="Sales by Category", 
                  labels={'x': 'Category', 'y': 'Sales'}, color=cat_sales.values, 
                  color_continuous_scale='Blues')
    st.plotly_chart(fig1, use_container_width=True)

with col2:
    reg_sales = filtered_df.groupby('Region')['Sales'].sum().sort_values(ascending=False)
    fig2 = px.bar(x=reg_sales.index, y=reg_sales.values, title="Sales by Region",
                  labels={'x': 'Region', 'y': 'Sales'}, color=reg_sales.values,
                  color_continuous_scale='Greens')
    st.plotly_chart(fig2, use_container_width=True)

st.divider()

st.subheader("Profit Analysis")
col1, col2 = st.columns(2)

with col1:
    cat_profit = filtered_df.groupby('Category')['Profit'].sum().sort_values(ascending=False)
    fig3 = px.bar(x=cat_profit.index, y=cat_profit.values, title="Profit by Category",
                  labels={'x': 'Category', 'y': 'Profit'}, color=cat_profit.values,
                  color_continuous_scale='Reds')
    st.plotly_chart(fig3, use_container_width=True)

with col2:
    reg_profit = filtered_df.groupby('Region')['Profit'].sum().sort_values(ascending=False)
    fig4 = px.bar(x=reg_profit.index, y=reg_profit.values, title="Profit by Region",
                  labels={'x': 'Region', 'y': 'Profit'}, color=reg_profit.values,
                  color_continuous_scale='Purples')
    st.plotly_chart(fig4, use_container_width=True)

st.divider()

st.subheader("Segment Analysis")
col1, col2 = st.columns(2)

with col1:
    seg_sales = filtered_df.groupby('Segment')['Sales'].sum().sort_values(ascending=False)
    fig5 = px.bar(x=seg_sales.index, y=seg_sales.values, title="Sales by Segment",
                  labels={'x': 'Segment', 'y': 'Sales'}, color=seg_sales.values,
                  color_continuous_scale='Viridis')
    st.plotly_chart(fig5, use_container_width=True)

with col2:
    ship_sales = filtered_df.groupby('Ship Mode')['Sales'].sum().sort_values(ascending=False)
    fig6 = px.bar(x=ship_sales.index, y=ship_sales.values, title="Sales by Ship Mode",
                  labels={'x': 'Ship Mode', 'y': 'Sales'}, color=ship_sales.values,
                  color_continuous_scale='Plasma')
    st.plotly_chart(fig6, use_container_width=True)

st.divider()

st.subheader("Yearly Trend")
col1, col2 = st.columns(2)

with col1:
    yearly_sales = filtered_df.groupby(filtered_df['Order Date'].dt.year)['Sales'].sum()
    fig7 = px.bar(x=yearly_sales.index, y=yearly_sales.values, title="Sales by Year",
                  labels={'x': 'Year', 'y': 'Sales'}, color=yearly_sales.values,
                  color_continuous_scale='Turbo')
    st.plotly_chart(fig7, use_container_width=True)

with col2:
    yearly_profit = filtered_df.groupby(filtered_df['Order Date'].dt.year)['Profit'].sum()
    fig8 = px.bar(x=yearly_profit.index, y=yearly_profit.values, title="Profit by Year",
                  labels={'x': 'Year', 'y': 'Profit'}, color=yearly_profit.values,
                  color_continuous_scale='RdYlGn')
    st.plotly_chart(fig8, use_container_width=True)

st.divider()

st.subheader("Top Sub-Categories")
top_sub = filtered_df.groupby('Sub-Category')['Sales'].sum().sort_values(ascending=False).head(10)
fig9 = px.bar(y=top_sub.index, x=top_sub.values, orientation='h', 
              title="Top 10 Sub-Categories", labels={'x': 'Sales', 'y': 'Sub-Category'},
              color=top_sub.values, color_continuous_scale='Sunset')
st.plotly_chart(fig9, use_container_width=True)

st.divider()

with st.expander("View Data"):
    st.dataframe(filtered_df, use_container_width=True)
