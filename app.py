import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import timedelta

# Set page config for wide layout
st.set_page_config(page_title="SuperStore Enhanced KPI Dashboard", layout="wide")

# ---- Load Data ----
@st.cache_data
def load_data():
    """
    Loads the Sample Superstore Excel file into a Pandas DataFrame.
    Make sure 'Sample - Superstore.xlsx' is in the same folder as this script.
    """
    df = pd.read_excel("Sample - Superstore.xlsx", engine="openpyxl")
    # Convert Order Date to datetime if not already
    if not pd.api.types.is_datetime64_any_dtype(df["Order Date"]):
        df["Order Date"] = pd.to_datetime(df["Order Date"])
    return df

df_original = load_data()

# ---- Sidebar: Filters & Help ----
st.sidebar.header("Dashboard Help")
st.sidebar.info(
    "Use the filters below to refine your data. You can select multiple values "
    "for Region, State, Category, or Sub-Category to compare performance across segments. "
    "Use the date range and time granularity options to explore trends over time."
)

st.sidebar.header("Filters")

# 1. Multi-select Region Filter
all_regions = sorted(df_original["Region"].dropna().unique())
selected_regions = st.sidebar.multiselect("Select Region(s)", options=all_regions, default=all_regions)

df_filtered = df_original.copy()
if selected_regions:
    df_filtered = df_filtered[df_filtered["Region"].isin(selected_regions)]

# 2. Multi-select State Filter
all_states = sorted(df_filtered["State"].dropna().unique())
selected_states = st.sidebar.multiselect("Select State(s)", options=all_states, default=all_states)
if selected_states:
    df_filtered = df_filtered[df_filtered["State"].isin(selected_states)]

# 3. Multi-select Category Filter
all_categories = sorted(df_filtered["Category"].dropna().unique())
selected_categories = st.sidebar.multiselect("Select Category(s)", options=all_categories, default=all_categories)
if selected_categories:
    df_filtered = df_filtered[df_filtered["Category"].isin(selected_categories)]

# 4. Multi-select Sub-Category Filter
all_subcats = sorted(df_filtered["Sub-Category"].dropna().unique())
selected_subcats = st.sidebar.multiselect("Select Sub-Category(s)", options=all_subcats, default=all_subcats)
if selected_subcats:
    df_filtered = df_filtered[df_filtered["Sub-Category"].isin(selected_subcats)]

# 5. Sidebar Date Range
st.sidebar.header("Date Range")
if df_filtered.empty:
    min_date = df_original["Order Date"].min()
    max_date = df_original["Order Date"].max()
else:
    min_date = df_filtered["Order Date"].min()
    max_date = df_filtered["Order Date"].max()

from_date = st.sidebar.date_input("From Date", value=min_date, min_value=min_date, max_value=max_date)
to_date = st.sidebar.date_input("To Date", value=max_date, min_value=min_date, max_value=max_date)
if from_date > to_date:
    st.sidebar.error("From Date must be earlier than To Date.")

df_filtered = df_filtered[(df_filtered["Order Date"] >= pd.to_datetime(from_date)) &
                          (df_filtered["Order Date"] <= pd.to_datetime(to_date))]

# 6. Sidebar Time Granularity
st.sidebar.header("Time Granularity")
time_granularity = st.sidebar.radio("Select Time Granularity", options=["Daily", "Weekly", "Monthly"])

# ---- Main Page: Title & Narrative ----
st.title("SuperStore Enhanced KPI Dashboard")
st.markdown(
    "This interactive dashboard provides a deep dive into SuperStore's performance. "
    "Adjust the filters and time granularity to explore trends, compare segments, and "
    "evaluate performance against a previous period."
)

# ---- KPI Calculations for Current Period ----
if df_filtered.empty:
    total_sales = total_quantity = total_profit = margin_rate = 0
else:
    total_sales = df_filtered["Sales"].sum()
    total_quantity = df_filtered["Quantity"].sum()
    total_profit = df_filtered["Profit"].sum()
    margin_rate = total_profit / total_sales if total_sales != 0 else 0

# ---- Historical Comparison: Previous Period ----
period_length = pd.to_datetime(to_date) - pd.to_datetime(from_date)
previous_from = pd.to_datetime(from_date) - period_length - timedelta(days=1)
previous_to = pd.to_datetime(from_date) - timedelta(days=1)
df_previous = df_original[(df_original["Order Date"] >= previous_from) &
                          (df_original["Order Date"] <= previous_to)]

# Apply same filters to previous period data
if selected_regions:
    df_previous = df_previous[df_previous["Region"].isin(selected_regions)]
if selected_states:
    df_previous = df_previous[df_previous["State"].isin(selected_states)]
if selected_categories:
    df_previous = df_previous[df_previous["Category"].isin(selected_categories)]
if selected_subcats:
    df_previous = df_previous[df_previous["Sub-Category"].isin(selected_subcats)]

if df_previous.empty:
    prev_sales = prev_quantity = prev_profit = prev_margin = 0
else:
    prev_sales = df_previous["Sales"].sum()
    prev_quantity = df_previous["Quantity"].sum()
    prev_profit = df_previous["Profit"].sum()
    prev_margin = prev_profit / prev_sales if prev_sales != 0 else 0

# ---- Custom CSS for KPI Tiles ----
st.markdown(
    """
    <style>
    .kpi-box {
        background-color: #FFFFFF;
        border: 2px solid #EAEAEA;
        border-radius: 8px;
        padding: 16px;
        margin: 8px;
        text-align: center;
    }
    .kpi-title {
        font-weight: 600;
        color: #333333;
        font-size: 16px;
        margin-bottom: 8px;
    }
    .kpi-value {
        font-weight: 700;
        font-size: 24px;
        color: #1E90FF;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ---- KPI Display (Including Change from Previous Period) ----
kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5 = st.columns(5)

with kpi_col1:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Sales</div>
            <div class='kpi-value'>${total_sales:,.2f}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with kpi_col2:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Quantity Sold</div>
            <div class='kpi-value'>{total_quantity:,.0f}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with kpi_col3:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Profit</div>
            <div class='kpi-value'>${total_profit:,.2f}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with kpi_col4:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Margin Rate</div>
            <div class='kpi-value'>{(margin_rate * 100):.2f}%</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# Determine which KPI to display for historical comparison
kpi_options = ["Sales", "Quantity", "Profit", "Margin Rate"]
selected_kpi = st.radio("Select KPI to Visualize", options=kpi_options, horizontal=True)

def get_kpi_value(dataframe, kpi):
    if dataframe.empty:
        return 0
    if kpi == "Sales":
        return dataframe["Sales"].sum()
    elif kpi == "Quantity":
        return dataframe["Quantity"].sum()
    elif kpi == "Profit":
        return dataframe["Profit"].sum()
    elif kpi == "Margin Rate":
        sales_sum = dataframe["Sales"].sum()
        profit_sum = dataframe["Profit"].sum()
        return profit_sum / sales_sum if sales_sum != 0 else 0

current_kpi_value = get_kpi_value(df_filtered, selected_kpi)
previous_kpi_value = get_kpi_value(df_previous, selected_kpi)
if previous_kpi_value != 0:
    kpi_change = ((current_kpi_value - previous_kpi_value) / previous_kpi_value) * 100
else:
    kpi_change = 0

with kpi_col5:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Change from Previous Period</div>
            <div class='kpi-value'>{kpi_change:+.2f}%</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ---- Visualizations ----
if df_filtered.empty:
    st.warning("No data available for the selected filters and date range.")
else:
    # Group data based on the selected time granularity
    if time_granularity == "Daily":
        grouped = df_filtered.groupby("Order Date").agg({"Sales": "sum", "Quantity": "sum", "Profit": "sum"}).reset_index()
    elif time_granularity == "Weekly":
        grouped = df_filtered.groupby(pd.Grouper(key="Order Date", freq="W")).agg({"Sales": "sum", "Quantity": "sum", "Profit": "sum"}).reset_index()
    else:  # Monthly
        grouped = df_filtered.groupby(pd.Grouper(key="Order Date", freq="M")).agg({"Sales": "sum", "Quantity": "sum", "Profit": "sum"}).reset_index()

    grouped["Margin Rate"] = grouped["Profit"] / grouped["Sales"].replace(0, 1)

    # Prepare data for Top 10 products chart
    product_grouped = df_filtered.groupby("Product Name").agg({"Sales": "sum", "Quantity": "sum", "Profit": "sum"}).reset_index()
    product_grouped["Margin Rate"] = product_grouped["Profit"] / product_grouped["Sales"].replace(0, 1)
    product_grouped.sort_values(by=selected_kpi, ascending=False, inplace=True)
    top_10 = product_grouped.head(10)

    # Create side-by-side charts
    col_left, col_right = st.columns(2)

    with col_left:
        fig_line = px.line(
            grouped,
            x="Order Date",
            y=selected_kpi,
            title=f"{selected_kpi} Over Time ({time_granularity})",
            labels={"Order Date": "Date", selected_kpi: selected_kpi},
            template="plotly_white",
        )
        fig_line.update_layout(height=400)
        st.plotly_chart(fig_line, use_container_width=True)

    with col_right:
        fig_bar = px.bar(
            top_10,
            x=selected_kpi,
            y="Product Name",
            orientation="h",
            title=f"Top 10 Products by {selected_kpi}",
            labels={selected_kpi: selected_kpi, "Product Name": "Product"},
            color=selected_kpi,
            color_continuous_scale="Blues",
            template="plotly_white",
        )
        fig_bar.update_layout(height=400, yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig_bar, use_container_width=True)

# ---- Narrative Summary ----
st.markdown("#### Insights & Commentary")
st.markdown(
    "Use the radio button to switch between KPIs (Sales, Quantity, Profit, Margin Rate). "
    "The chart on the left shows how your selected KPI changes over time at the chosen granularity. "
    "On the right, you can see which products lead in terms of that KPI. "
    "You can also see how the current KPI compares to a previous period in the last KPI tile."
)

# ---- Data Drill-Down Section ----
with st.expander("Show Detailed Data"):
    st.dataframe(df_filtered)
