# @Email:  contact@pythonandvba.com
# @Website:  https://pythonandvba.com
# @YouTube:  https://youtube.com/c/CodingIsFun
# @Project:  Sales Dashboard w/ Streamlit



import pandas as pd  # pip install pandas openpyxl
import plotly.express as px  # pip install plotly-express
import streamlit as st  # pip install streamlit

# emojis: https://www.webfx.com/tools/emoji-cheat-sheet/
st.set_page_config(page_title="Sales Dashboard", page_icon=":bar_chart:", layout="wide")

# ---- READ EXCEL ----
@st.cache_data
def get_data_from_excel():
    df = pd.read_excel(
        io="supermarkt_sales.xlsx",
        engine="openpyxl",
        sheet_name="Sales",
        skiprows=3,
        usecols="B:R",
        nrows=1000,
    )
    # Add 'hour' column to dataframe
    df["hour"] = pd.to_datetime(df["Time"], format="%H:%M:%S").dt.hour
    return df

df = get_data_from_excel()

# ---- SIDEBAR ----
st.sidebar.header("Please Filter Here:")
city = st.sidebar.multiselect(
    "Select the City:",
    options=df["city"].unique(),
    default=df["city"].unique()
)

customer_type = st.sidebar.multiselect(
    "Select the Customer Type:",
    options=df["Customer_type"].unique(),
    default=df["Customer_type"].unique()
)

gender = st.sidebar.multiselect(
    "Select the Gender:",
    options=df["Gender"].unique(),
    default[df["Gender"].unique()
)

df_selection = df.query(
    "city == @city & Customer_type == @customer_type & Gender == @gender"
)

# Check if the dataframe is empty:
if df_selection.empty:
    st.warning("No data available based on the current filter settings!")
    st.stop()  # This will halt the app from further execution.

# ---- MAINPAGE ----
st.title(":bar_chart: Sales Dashboard")
st.markdown("##")

# TOP KPI's
total_sales = int(df_selection["Total"].sum())
average_rating = round(df_selection["Rating"].mean(), 1)
star_rating = ":star:" * int(round(average_rating, 0))
average_sale_by_transaction = round(df_selection["Total"].mean(), 2)

left_column, middle_column, right_column = st.columns(3)
with left_column:
    st.subheader("Total Sales:")
    st.subheader(f"US $ {total_sales:,}")
with middle_column:
    st.subheader("Average Rating:")
    st.subheader(f"{average_rating} {star_rating}")
with right_column:
    st.subheader("Average Sales Per Transaction:")
    st.subheader(f"US $ {average_sale_by_transaction}")

st.markdown("""---""")

# SALES BY PRODUCT LINE [BAR CHART]
sales_by_product_line = df_selection.groupby(by=["Product line"])[["Total"]].sum().sort_values(by="Total")
fig_product_sales = px.bar(
    sales_by_product_line,
    x="Total",
    y=sales_by_product_line.index,
    orientation="h",
    title="<b>Sales by Product Line</b>",
    color_discrete_sequence=["#0083B8"] * len(sales_by_product_line),
    template="plotly_white",
)
fig_product_sales.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=(dict(showgrid=False))
)

# SALES BY HOUR [BAR CHART]
sales_by_hour = df_selection.groupby(by=["hour"])[["Total"]].sum()
fig_hourly_sales = px.bar(
    sales_by_hour,
    x=sales_by_hour.index,
    y="Total",
    title="<b>Sales by hour</b>",
    color_discrete_sequence=["#0083B8"] * len(sales_by_hour),
    template="plotly_white",
)
fig_hourly_sales.update_layout(
    xaxis=dict(tickmode="linear"),
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=(dict(showgrid=False)),
)

# SALES BY CITY [BAR CHART]
sales_by_city = df_selection.groupby(by=["city"])[["Total"]].sum().sort_values(by="Total")
fig_city_sales = px.bar(
    sales_by_city,
    x=sales_by_city.index,
    y="Total",
    title="<b>Sales by City</b>",
    color_discrete_sequence=["#0083B8"] * len(sales_by_city),
    template="plotly_white",
)
fig_city_sales.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=(dict(showgrid=False)),
)

# GROSS INCOME BY PRODUCT LINE [LINE CHART]
gross_income_by_product_line = df_selection.groupby(by=["Product line"])[["gross income"]].sum().sort_values(by="gross income")
fig_gross_income = px.line(
    gross_income_by_product_line,
    x=gross_income_by_product_line.index,
    y="gross income",
    title="<b>Gross Income by Product Line</b>",
    markers=True,
    color_discrete_sequence=["#0083B8"] * len(gross_income_by_product_line),
    template="plotly_white",
)
fig_gross_income.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=dict(showgrid=False),
    yaxis=dict(showgrid=False),
)

# SALES BY PAYMENT METHOD [PIE CHART]
sales_by_payment = df_selection.groupby(by=["Payment"])[["Total"]].sum()
fig_payment_sales = px.pie(
    sales_by_payment,
    values="Total",
    names=sales_by_payment.index,
    title="<b>Sales by Payment Method</b>",
    color_discrete_sequence=px.colors.sequential.RdBu,
    template="plotly_white",
)

left_column, right_column = st.columns(2)
left_column.plotly_chart(fig_hourly_sales, use_container_width=True)
right_column.plotly_chart(fig_product_sales, use_container_width=True)

# Add sales by city chart
st.plotly_chart(fig_city_sales, use_container_width=True)

# Add gross income by product line chart below the sales by city chart
st.plotly_chart(fig_gross_income, use_container_width=True)

# Add sales by payment method pie chart
st.plotly_chart(fig_payment_sales, use_container_width=True)

# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)
