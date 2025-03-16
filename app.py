import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="SuperStore Enhanced KPI Dashboard", layout="wide")

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

st.title("SuperStore Enhanced KPI Dashboard")
st.markdown(
    "This interactive dashboard provides a deep dive into SuperStore's performance. "
    "Adjust the filters and time granularity to explore trends and compare segments."
)

if df_filtered.empty:
    total_sales = total_quantity = total_profit = margin_rate = 0
else:
    total_sales = df_filtered["Sales"].sum()
    total_quantity = df_filtered["Quantity"].sum()
    total_profit = df_filtered["Profit"].sum()
    margin_rate = total_profit / total_sales if total_sales != 0 else 0

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

kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)

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

kpi_options = ["Sales", "Quantity", "Profit", "Margin Rate"]
selected_kpi = st.radio("Select KPI to Visualize", options=kpi_options, horizontal=True)

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

st.markdown("#### Insights & Commentary")
st.markdown(
    "Use the radio button above to switch between KPIs (Sales, Quantity, Profit, Margin Rate). "
    "The chart on the left shows how your selected KPI changes over time at the chosen granularity. "
    "On the right, you can see which products lead in terms of that KPI."
)
with st.expander("Show Detailed Data"):
    st.dataframe(df_filtered)
