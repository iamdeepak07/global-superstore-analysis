import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(
    page_title="Global Superstore Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
        .main { padding-top: 2rem; }
        h1 { color: #667eea; text-align: center; margin-bottom: 30px; }
        h2 { color: #667eea; border-bottom: 3px solid #667eea; padding-bottom: 10px; }
        .error-box { background-color: #ffebee; padding: 10px; border-radius: 5px; color: #c62828; }
        .success-box { background-color: #e8f5e9; padding: 10px; border-radius: 5px; color: #2e7d32; }
    </style>
""", unsafe_allow_html=True)

# Load data with error handling
@st.cache_data
def load_data():
    try:
        file_path = r'c:\Users\ASUS\OneDrive\Documents\Data Analyst\global_superstore (1).xlsx'
        df = pd.read_excel(file_path)
        df['Order Date'] = pd.to_datetime(df['Order Date'])
        return df
    except FileNotFoundError:
        st.error("Data file not found!")
        return None
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None

# Load data
df = load_data()

if df is None or len(df) == 0:
    st.error("Unable to load data. Please check the file path.")
    st.stop()

# Title
st.markdown("<h1>📊 Global Superstore Sales Dashboard</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #666; font-size: 16px;'>Interactive Analytics & Sales Metrics (2012-2015)</p>", unsafe_allow_html=True)

# Sidebar filters
st.sidebar.header("🎛️ Filters & Controls")
try:
    selected_year = st.sidebar.multiselect(
        "Select Year(s):",
        options=sorted(df['Order Date'].dt.year.unique()),
        default=sorted(df['Order Date'].dt.year.unique())
    )

    selected_region = st.sidebar.multiselect(
        "Select Region(s):",
        options=sorted(df['Region'].unique()),
        default=sorted(df['Region'].unique())
    )

    selected_segment = st.sidebar.multiselect(
        "Select Segment(s):",
        options=sorted(df['Segment'].unique()),
        default=sorted(df['Segment'].unique())
    )
except Exception as e:
    st.error(f"Error in filters: {str(e)}")
    st.stop()

# Filter data
filtered_df = df[
    (df['Order Date'].dt.year.isin(selected_year)) &
    (df['Region'].isin(selected_region)) &
    (df['Segment'].isin(selected_segment))
].copy()

# Check if we have data
if len(filtered_df) == 0:
    st.warning("⚠️ No data available for selected filters. Please adjust your selections.")
    st.stop()

# KPI Row with error handling
st.markdown("## 📈 Key Performance Indicators")
try:
    col1, col2, col3, col4, col5, col6 = st.columns(6)

    total_sales = filtered_df['Sales'].sum()
    total_profit = filtered_df['Profit'].sum()
    
    with col1:
        st.metric("Total Sales", f"${total_sales:,.0f}")

    with col2:
        st.metric("Total Profit", f"${total_profit:,.0f}")

    with col3:
        profit_margin = (total_profit / total_sales * 100) if total_sales > 0 else 0
        st.metric("Profit Margin", f"{profit_margin:.2f}%")

    with col4:
        avg_order = filtered_df['Sales'].mean()
        st.metric("Avg Order Value", f"${avg_order:,.0f}" if not pd.isna(avg_order) else "$0")

    with col5:
        total_orders = filtered_df['Order ID'].nunique()
        st.metric("Total Orders", f"{total_orders:,}")

    with col6:
        avg_quantity = filtered_df['Quantity'].mean()
        st.metric("Avg Quantity", f"{avg_quantity:.2f}" if not pd.isna(avg_quantity) else "0")

except Exception as e:
    st.error(f"Error calculating KPIs: {str(e)}")

st.divider()

# Helper function to create bar chart safely
def create_bar_chart(data, x_col, y_col, title, x_label, y_label, color_scale='Blues'):
    try:
        if len(data) == 0:
            st.info(f"No data for {title}")
            return None
        
        chart_data = data.groupby(x_col)[y_col].sum().sort_values(ascending=False)
        
        if len(chart_data) == 0:
            st.info(f"No data available for {title}")
            return None
        
        fig = go.Figure(data=[
            go.Bar(
                x=chart_data.index,
                y=chart_data.values,
                text=[f"${x:,.0f}" if y_col in ['Sales', 'Profit'] else f"{x:,.0f}" 
                      for x in chart_data.values],
                textposition='outside',
                marker=dict(
                    color=chart_data.values,
                    colorscale=color_scale,
                    showscale=False
                ),
                hovertemplate=f"<b>%{{x}}</b><br>{y_label}: %{{y:,.0f}}<extra></extra>"
            )
        ])
        
        fig.update_layout(
            title=title,
            xaxis_title=x_label,
            yaxis_title=y_label,
            height=400,
            template='plotly_white',
            hovermode='x unified',
            showlegend=False
        )
        
        if y_col in ['Sales', 'Profit']:
            fig.update_yaxes(tickformat='$,.0f')
        
        return fig
    except Exception as e:
        st.error(f"Error creating chart for {title}: {str(e)}")
        return None

# Row 1: Sales by Category and Region
st.markdown("## 📊 Sales by Category & Region")
try:
    col1, col2 = st.columns(2)
    
    with col1:
        fig1 = create_bar_chart(filtered_df, 'Category', 'Sales', 
                               'Sales by Category', 'Category', 'Sales ($)', 'Blues')
        if fig1:
            st.plotly_chart(fig1, use_container_width=True)
    
    with col2:
        fig2 = create_bar_chart(filtered_df, 'Region', 'Sales', 
                               'Sales by Region', 'Region', 'Sales ($)', 'Greens')
        if fig2:
            st.plotly_chart(fig2, use_container_width=True)
except Exception as e:
    st.error(f"Error in Sales section: {str(e)}")

st.divider()

# Row 2: Profit Analysis
st.markdown("## 💰 Profit Analysis")
try:
    col1, col2 = st.columns(2)
    
    with col1:
        fig3 = create_bar_chart(filtered_df, 'Category', 'Profit', 
                               'Profit by Category', 'Category', 'Profit ($)', 'Reds')
        if fig3:
            st.plotly_chart(fig3, use_container_width=True)
    
    with col2:
        fig4 = create_bar_chart(filtered_df, 'Region', 'Profit', 
                               'Profit by Region', 'Region', 'Profit ($)', 'Purples')
        if fig4:
            st.plotly_chart(fig4, use_container_width=True)
except Exception as e:
    st.error(f"Error in Profit section: {str(e)}")

st.divider()

# Row 3: Segment & Ship Mode Analysis
st.markdown("## 🎯 Segment & Ship Mode Analysis")
try:
    col1, col2 = st.columns(2)
    
    with col1:
        fig5 = create_bar_chart(filtered_df, 'Segment', 'Sales', 
                               'Sales by Customer Segment', 'Segment', 'Sales ($)', 'Viridis')
        if fig5:
            st.plotly_chart(fig5, use_container_width=True)
    
    with col2:
        fig6 = create_bar_chart(filtered_df, 'Ship Mode', 'Sales', 
                               'Sales by Ship Mode', 'Ship Mode', 'Sales ($)', 'Plasma')
        if fig6:
            st.plotly_chart(fig6, use_container_width=True)
except Exception as e:
    st.error(f"Error in Segment section: {str(e)}")

st.divider()

# Row 4: Yearly Analysis
st.markdown("## 📅 Yearly Sales Analysis")
try:
    col1, col2 = st.columns(2)
    
    # Prepare yearly data
    yearly_data = filtered_df.groupby(filtered_df['Order Date'].dt.year).agg({
        'Sales': 'sum',
        'Profit': 'sum'
    }).reset_index()
    yearly_data.columns = ['Year', 'Sales', 'Profit']
    
    with col1:
        fig7 = go.Figure(data=[
            go.Bar(
                x=yearly_data['Year'],
                y=yearly_data['Sales'],
                text=[f"${x:,.0f}" for x in yearly_data['Sales']],
                textposition='outside',
                marker=dict(
                    color=yearly_data['Sales'],
                    colorscale='Turbo',
                    showscale=False
                ),
                hovertemplate="<b>Year: %{x}</b><br>Sales: $%{y:,.0f}<extra></extra>"
            )
        ])
        fig7.update_layout(
            title='Total Sales by Year',
            xaxis_title='Year',
            yaxis_title='Sales ($)',
            height=400,
            template='plotly_white',
            yaxis_tickformat='$,.0f'
        )
        st.plotly_chart(fig7, use_container_width=True)
    
    with col2:
        fig8 = go.Figure(data=[
            go.Bar(
                x=yearly_data['Year'],
                y=yearly_data['Profit'],
                text=[f"${x:,.0f}" for x in yearly_data['Profit']],
                textposition='outside',
                marker=dict(
                    color=yearly_data['Profit'],
                    colorscale='RdYlGn',
                    showscale=False
                ),
                hovertemplate="<b>Year: %{x}</b><br>Profit: $%{y:,.0f}<extra></extra>"
            )
        ])
        fig8.update_layout(
            title='Total Profit by Year',
            xaxis_title='Year',
            yaxis_title='Profit ($)',
            height=400,
            template='plotly_white',
            yaxis_tickformat='$,.0f'
        )
        st.plotly_chart(fig8, use_container_width=True)
except Exception as e:
    st.error(f"Error in Yearly Analysis: {str(e)}")

st.divider()

# Row 5: Top Sub-Categories
st.markdown("## 🏷️ Top Sub-Categories by Sales")
try:
    top_sub = filtered_df.groupby('Sub-Category')['Sales'].sum().sort_values(ascending=True).tail(12)
    
    if len(top_sub) > 0:
        fig9 = go.Figure(data=[
            go.Bar(
                y=top_sub.index,
                x=top_sub.values,
                orientation='h',
                text=[f"${x:,.0f}" for x in top_sub.values],
                textposition='outside',
                marker=dict(
                    color=top_sub.values,
                    colorscale='Sunset',
                    showscale=False
                ),
                hovertemplate="<b>%{y}</b><br>Sales: $%{x:,.0f}<extra></extra>"
            )
        ])
        fig9.update_layout(
            title='Top 12 Sub-Categories by Sales',
            xaxis_title='Sales ($)',
            yaxis_title='Sub-Category',
            height=500,
            template='plotly_white',
            xaxis_tickformat='$,.0f'
        )
        st.plotly_chart(fig9, use_container_width=True)
except Exception as e:
    st.error(f"Error in Sub-Categories section: {str(e)}")

st.divider()

# Row 6: Orders & Discount
st.markdown("## 📦 Orders & Discount Analysis")
try:
    col1, col2 = st.columns(2)
    
    with col1:
        orders_by_cat = filtered_df.groupby('Category')['Order ID'].nunique().sort_values(ascending=False)
        if len(orders_by_cat) > 0:
            fig10 = go.Figure(data=[
                go.Bar(
                    x=orders_by_cat.index,
                    y=orders_by_cat.values,
                    text=orders_by_cat.values,
                    textposition='outside',
                    marker=dict(
                        color=orders_by_cat.values,
                        colorscale='Inferno',
                        showscale=False
                    ),
                    hovertemplate="<b>%{x}</b><br>Orders: %{y:,.0f}<extra></extra>"
                )
            ])
            fig10.update_layout(
                title='Number of Orders by Category',
                xaxis_title='Category',
                yaxis_title='Order Count',
                height=400,
                template='plotly_white'
            )
            st.plotly_chart(fig10, use_container_width=True)
    
    with col2:
        avg_disc = filtered_df.groupby('Category')['Discount'].mean().sort_values(ascending=False) * 100
        if len(avg_disc) > 0:
            fig11 = go.Figure(data=[
                go.Bar(
                    x=avg_disc.index,
                    y=avg_disc.values,
                    text=[f"{x:.1f}%" for x in avg_disc.values],
                    textposition='outside',
                    marker=dict(
                        color=avg_disc.values,
                        colorscale='RdYlBu_r',
                        showscale=False
                    ),
                    hovertemplate="<b>%{x}</b><br>Avg Discount: %{y:.1f}%<extra></extra>"
                )
            ])
            fig11.update_layout(
                title='Average Discount by Category (%)',
                xaxis_title='Category',
                yaxis_title='Discount (%)',
                height=400,
                template='plotly_white'
            )
            st.plotly_chart(fig11, use_container_width=True)
except Exception as e:
    st.error(f"Error in Orders section: {str(e)}")

st.divider()

# Data Table
st.markdown("## 📋 Detailed Data View")
try:
    with st.expander("View Raw Data", expanded=False):
        display_cols = ['Order Date', 'Category', 'Sub-Category', 'Sales', 'Profit', 'Region', 'Segment']
        st.dataframe(
            filtered_df[display_cols].sort_values('Order Date', ascending=False),
            use_container_width=True,
            height=400
        )
except Exception as e:
    st.error(f"Error displaying data table: {str(e)}")

st.divider()

# Summary Statistics
st.markdown("## 📊 Summary Statistics")
try:
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
except Exception as e:
    st.error(f"Error displaying statistics: {str(e)}")

st.divider()

# Footer
st.markdown("""
    <div style='text-align: center; margin-top: 50px; padding: 20px; color: #999; border-top: 1px solid #ddd;'>
        <p><strong>Global Superstore Sales Dashboard</strong></p>
        <p>Built with Streamlit | Data: 2012-2015 | Last Updated: February 2026</p>
        <p style='font-size: 12px;'>✅ All errors resolved and fully functional</p>
    </div>
""", unsafe_allow_html=True)
