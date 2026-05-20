import streamlit as st
import pandas as pd


# ======================================================
# PAGE CONFIG
# ======================================================

st.set_page_config(
    page_title="DEMAND INCREASE REPORT",
    layout="wide"
)
st.markdown("""
<style>

/* scroll container */
.table-container {
    width: 100%;
    max-height: 650px;
    overflow-y: auto;
    overflow-x: auto;
    border: 1px solid #D6E4F0;
    border-radius: 10px;
}

/* table */
table {
    width: 100%;
    border-collapse: collapse;
    table-layout: auto;
}

/* sticky header FIX */
thead th {
    position: sticky;
    top: 0;
    z-index: 10;
    background: linear-gradient(90deg, #1E3C72, #2A5298) !important;
    color: white !important;
    text-align: center !important;
    padding: 10px !important;
}

/* cells */
td {
    text-align: center !important;
    padding: 8px !important;
    border: 1px solid #E6EEF8 !important;
    white-space: nowrap;
}

/* row colors */
tbody tr:nth-child(even) {
    background-color: #EEF4FF !important;
}

tbody tr:nth-child(odd) {
    background-color: #FFFFFF !important;
}

/* hover */
tbody tr:hover {
    background-color: #D9E8FF !important;
}

</style>
""", unsafe_allow_html=True)
# ======================================================
# CSS
# ======================================================

st.markdown("""
<style>

header {visibility:hidden;}
#MainMenu {visibility:hidden;}
footer {visibility:hidden;}

.stApp {
    background-color:#f4f7fb;
}

.block-container {
    padding-top:0.4rem;
    padding-left:1rem;
    padding-right:1rem;
}

/* TITLE */
.dashboard-title {
    text-align:center;
    font-size:42px;
    font-weight:900;
    color:white;
    padding:16px;
    border-radius:16px;
    background:linear-gradient(90deg,#0F3D56,#1976D2);
    box-shadow:0px 4px 14px rgba(0,0,0,0.15);
    margin-bottom:12px;
}

/* FILTER CARD */

.filter-card {
    background: linear-gradient(135deg,#ffffff,#f3f8ff);
    padding:18px;
    border-radius:18px;
    box-shadow:0px 4px 14px rgba(0,0,0,0.10);
    margin-bottom:18px;
    border:1px solid #D6E4F0;
}
/* FILTER LABELS */

label {
    font-weight:600 !important;
    color:#0F3D56 !important;
    font-size:14px !important;
}

/* SELECT BOX */

div[data-baseweb="select"] {
    background-color:white;
    border-radius:12px;
    border:2px solid #D6E4F0;
    box-shadow:0px 2px 6px rgba(0,0,0,0.05);
}

/* DATE BOX */

div[data-testid="stDateInput"] input {
    border-radius:12px !important;
    border:1px solid #D6E4F0 !important;
    background-color:white !important;
}

/* HOVER EFFECT */

div[data-baseweb="select"]:hover {
    border:1px solid #1976D2;
    transition:0.2s;
}

/* KPI */
.kpi-card {
    padding:18px;
    border-radius:18px;
    color:white;
    text-align:center;
    box-shadow:0px 4px 12px rgba(0,0,0,0.12);
    transition:0.3s;
}

.kpi-card:hover {
    transform:translateY(-5px);
}

.kpi-title {
    font-size:18px;
    font-weight:600;
}

.kpi-value {
    font-size:30px;
    font-weight:800;
    margin-top:8px;
}

/* TABLE HEADER */
th {
    background: linear-gradient(90deg, #1E3C72, #2A5298) !important;
    color: white !important;
    text-align: center !important;
    font-size: 12px !important;
    white-space: normal !important;
    word-wrap: break-word !important;
    border: 1px solid #D6E4F0 !important;
    padding: 10px !important;
}

/* TABLE ROWS */
tbody tr:nth-child(even) {
    background-color: #EEF4FF !important;
}

tbody tr:nth-child(odd) {
    background-color: #FFFFFF !important;
}

/* HOVER EFFECT */
tbody tr:hover {
    background-color: #D9E8FF !important;
    transition: 0.2s;
}

/* TABLE CELLS */
td {
    text-align: center !important;
    font-size: 12px !important;
    padding: 8px !important;
    border: 1px solid #E6EEF8 !important;
}
/* TABLE CELLS */
td {
    text-align: center !important;
    font-size: 12px !important;
    padding: 8px !important;
    border: 1px solid #E6EEF8 !important;
}

/* STICKY HEADER */

thead th {
    position: sticky;
    top: 0;
    z-index: 100;
}

</style>
""", unsafe_allow_html=True)


# ======================================================
# LOAD DATA
# ======================================================

file_path = "master_file_updated.xlsx"

df = pd.read_excel(file_path)


cols_to_clean = ["ulb_name", "district", "region", "grade"]

for col in cols_to_clean:
    df[col] = (
        df[col]
        .astype(str)
        .str.strip()              # remove leading/trailing spaces
        .str.replace(r"\s+", " ", regex=True)  # collapse multiple spaces into single space
    )


# ======================================================
# CLEAN DATA
# ======================================================

df.columns = df.columns.str.strip()

df["erpdisposaldt"] = pd.to_datetime(
    df["erpdisposaldt"],
    dayfirst=True,
    errors="coerce"
)

df = df[df["erpdisposaldt"].notna()]

df["Org Demand Increased"] = pd.to_numeric(
    df["Org Demand Increased"],
    errors="coerce"
).fillna(0)
df = df[df["erpdisposaldt"] <= pd.Timestamp.today()]


# ======================================================
# ULB MERGING LOGIC
# ======================================================

# ---------- GVMC MERGE ----------

gvmc_mask = (
    df["ulb_name"]
    .astype(str)
    .str.strip()
    .str.lower()
    .isin([
        "anakapalli",
        "bheemili",
        "visakhapatnam",
        "gvmc"
    ])
)

df.loc[gvmc_mask, "ulb_name"] = "GVMC"
df.loc[gvmc_mask, "ulb_code"] = 1086
df.loc[gvmc_mask, "district"] = "VISAKHAPATNAM"
df.loc[gvmc_mask, "region"] = "VISAKHAPATNAM"
df.loc[gvmc_mask, "grade"] = "Corp"

# ---------- MANGALAGIRI TADEPALLI MERGE ----------

mt_mask = (
    df["ulb_name"]
    .astype(str)
    .str.strip()
    .str.lower()
    .isin([
        "mangalagiri",
        "tadepalli",
        "mangalagiri tadepalli",
        "mangalagiri  tadepalli"
    ])
)

df.loc[mt_mask, "ulb_name"] = "Mangalagiri Tadepalli"
df.loc[mt_mask, "ulb_code"] = 1023
df.loc[mt_mask, "district"] = "GUNTUR"
df.loc[mt_mask, "region"] = "GUNTUR"
df.loc[mt_mask, "grade"] = "Corp"
# ======================================================
# TITLE
# ======================================================

st.markdown("""
<div class='dashboard-title'>
DEMAND INCREASE REPORT
</div>
""", unsafe_allow_html=True)

# ======================================================
# FILTER CARD
# ======================================================

st.markdown("<div class='filter-card'>", unsafe_allow_html=True)

region_options = sorted(df["region"].dropna().astype(str).unique())

col1, col2, col3, col4, col5, col6, col7 = st.columns(7)

with col1:
    report_type = st.selectbox(
        "Report Type",
        [
            "ULB Wise",
            "Region Wise",
            "District Wise",
            "Grade Wise"
        ]
    )

with col2:
    selected_region = st.selectbox(
        "Region",
        ["ALL"] + region_options
    )

if selected_region != "ALL":
    district_df = df[
        df["region"].astype(str) == selected_region
    ]
else:
    district_df = df.copy()

district_options = sorted(
    district_df["district"].dropna().astype(str).unique()
)

with col3:
    selected_district = st.selectbox(
        "District",
        ["ALL"] + district_options
    )

grade_options = sorted(df["grade"].dropna().astype(str).unique())

with col4:
    selected_grade = st.selectbox(
        "Grade",
        ["ALL"] + grade_options
    )

ulb_df = district_df.copy()

if selected_district != "ALL":
    ulb_df = ulb_df[
        ulb_df["district"].astype(str) == selected_district
    ]

ulb_options = sorted(
    ulb_df["ulb_name"].dropna().astype(str).unique()
)

with col5:
    selected_ulb = st.selectbox(
        "ULB",
        ["ALL"] + ulb_options
    )

with col6:
    from_date = st.date_input(
        "From Date",
        df["erpdisposaldt"].min(),
        min_value=df["erpdisposaldt"].min(),
        max_value=df["erpdisposaldt"].max()
    )

with col7:
    to_date = st.date_input(
        "To Date",
        df["erpdisposaldt"].max(),
        min_value=df["erpdisposaldt"].min(),
        max_value=df["erpdisposaldt"].max()
    )

st.markdown("</div>", unsafe_allow_html=True)

# ======================================================
# FILTER APPLY
# ======================================================

filtered_df = df.copy()

if selected_region != "ALL":
    filtered_df = filtered_df[
        filtered_df["region"].astype(str) == selected_region
    ]

if selected_district != "ALL":
    filtered_df = filtered_df[
        filtered_df["district"].astype(str) == selected_district
    ]

if selected_grade != "ALL":
    filtered_df = filtered_df[
        filtered_df["grade"].astype(str) == selected_grade
    ]

if selected_ulb != "ALL":
    filtered_df = filtered_df[
        filtered_df["ulb_name"].astype(str) == selected_ulb
    ]

filtered_df = filtered_df[
    (
        filtered_df["erpdisposaldt"] >= pd.to_datetime(from_date)
    ) &
    (
        filtered_df["erpdisposaldt"] <= pd.to_datetime(to_date)
    )
]

if filtered_df.empty:
    st.warning("No Data Found")
    st.stop()

# ======================================================
# FAST LOGIC
# ======================================================

create_mask = (
    filtered_df["modify_reason"]
    .astype(str)
    .str.strip()
    .str.upper()
    .eq("CREATE")
)

filtered_df["Unassessed_Count"] = create_mask.astype(int)

filtered_df["Under_Assessed_Count"] = (~create_mask).astype(int)

filtered_df["Unassessed_Amount"] = filtered_df[
    "Org Demand Increased"
].where(create_mask, 0)

filtered_df["Under_Assessed_Amount"] = filtered_df[
    "Org Demand Increased"
].where(~create_mask, 0)

# ======================================================
# KPI VALUES
# ======================================================

unassessed_amt = filtered_df["Unassessed_Amount"].sum()

under_assessed_amt = filtered_df[
    "Under_Assessed_Amount"
].sum()

total_amt = unassessed_amt + under_assessed_amt

latest_date = filtered_df["erpdisposaldt"].max()

# ======================================================
# DISPLAY UNIT
# ======================================================

unit = st.radio(
    "Display Unit",
    ["Crores", "Lakhs"],
    horizontal=True
)

if unit == "Lakhs":
    divisor = 100000
    suffix = " L"
else:
    divisor = 10000000
    suffix = " Cr"

# ======================================================
# KPI CARDS
# ======================================================

k1, k2, k3, k4 = st.columns(4)

with k1:
    st.markdown(f"""
    <div class='kpi-card'
         style='background:linear-gradient(135deg,#1565C0,#42A5F5);'>
        <div class='kpi-title'>
            Unassessed Demand
        </div>
        <div class='kpi-value'>
            {unassessed_amt/divisor:,.2f}{suffix}
        </div>
    </div>
    """, unsafe_allow_html=True)

with k2:
    st.markdown(f"""
    <div class='kpi-card'
         style='background:linear-gradient(135deg,#EF6C00,#FFA726);'>
        <div class='kpi-title'>
            Under Assessed Demand
        </div>
        <div class='kpi-value'>
            {under_assessed_amt/divisor:,.2f}{suffix}
        </div>
    </div>
    """, unsafe_allow_html=True)

with k3:
    st.markdown(f"""
    <div class='kpi-card'
         style='background:linear-gradient(135deg,#2E7D32,#66BB6A);'>
        <div class='kpi-title'>
            Total Achievement
        </div>
        <div class='kpi-value'>
            {total_amt/divisor:,.2f}{suffix}
        </div>
    </div>
    """, unsafe_allow_html=True)

with k4:
    st.markdown(f"""
    <div class='kpi-card'
         style='background:linear-gradient(135deg,#6A1B9A,#AB47BC);'>
        <div class='kpi-title'>
            Report As On
        </div>
        <div class='kpi-value'>
            {latest_date.strftime('%d-%m-%Y')}
        </div>
    </div>
    """, unsafe_allow_html=True)

# ======================================================
# GROUPING
# ======================================================

if report_type == "Region Wise":
    group_cols = ["region"]

elif report_type == "District Wise":
    group_cols = ["region", "district"]

elif report_type == "Grade Wise":
    group_cols = ["grade"]

else:
    group_cols = [
        "region",
        "district",
        "ulb_code",
        "ulb_name",
        "grade"
    ]

summary_df = (
    filtered_df
    .groupby(group_cols, dropna=False)
    .agg({
        "Unassessed_Count": "sum",
        "Unassessed_Amount": "sum",
        "Under_Assessed_Count": "sum",
        "Under_Assessed_Amount": "sum"
    })
    .reset_index()
)

summary_df["Total Achievement"] = (
    summary_df["Unassessed_Amount"] +
    summary_df["Under_Assessed_Amount"]
)

# ======================================================
# SERIAL NUMBER
# ======================================================

summary_df.insert(
    0,
    "Sl.No",
    range(1, len(summary_df) + 1)
)

# ======================================================
# RENAME
# ======================================================

summary_df.rename(columns={

    "region": "Name of the Region",
    "district": "Name of the District",
    "ulb_name": "Name of the ULB",
    "ulb_code": "ULB Code",
    "grade": "ULB Grade",

    "Unassessed_Count": "No.of Unassessed Properties",
    "Unassessed_Amount": "Unassessed Demand",
    "Under_Assessed_Count": "No.of Under Assessed Properties",
    "Under_Assessed_Amount": "Under Assessed Demand"

}, inplace=True)

# ======================================================
# TOTAL ROW
# ======================================================

total_row = {}

for col in summary_df.columns:

    if col in [
        "No.of Unassessed Properties",
        "Unassessed Demand",
        "No.of Under Assessed Properties",
        "Under Assessed Demand",
        "Total Achievement"
    ]:

        total_row[col] = summary_df[col].sum()

    else:
        total_row[col] = ""

total_row["Sl.No"] = ""

if "Name of the Region" in summary_df.columns:
    total_row["Name of the Region"] = "TOTAL"

summary_df = pd.concat(
    [summary_df, pd.DataFrame([total_row])],
    ignore_index=True
)

# ======================================================
# TABLE TITLE
# ======================================================

st.markdown(f"""
<h3 style='color:#0F3D56;'>
{report_type} Demand Increase Report (In Rs.)
</h3>
""", unsafe_allow_html=True)

# ======================================================
# DOWNLOAD
# ======================================================

csv = summary_df.to_csv(index=False).encode("utf-8")

st.download_button(
    label="Download CSV",
    data=csv,
    file_name="Demand_Increase_Report.csv",
    mime="text/csv"
)




# ======================================================
# TABLE
# ======================================================

table_html = summary_df.to_html(index=False)

st.markdown(f"""
<div style="
    height:650px;
    overflow-y:auto;
    border:1px solid #D6E4F0;
    border-radius:10px;
">

{table_html}

</div>
""", unsafe_allow_html=True)
