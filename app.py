# =====================================================
# ================== IMPORTS ==========================
# =====================================================
import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill

# =====================================================
# ============== APP CONFIG ===========================
# =====================================================
st.set_page_config(
    page_title="Sigma Consultants",
    layout="wide"
)

# =====================================================
# ============== SESSION STATE ========================
# =====================================================
if "page" not in st.session_state:
    st.session_state.page = "Dashboard"

# =====================================================
# ============== DATA LOADING =========================
# =====================================================
# ⚠️ Replace this block with your real data source
proposals_df = pd.DataFrame({
    "Client_Name": ["Client A", "Client B", "Client A", "Client C"],
    "Start_Date": pd.to_datetime(["2024-01-01", "2024-01-10", "2024-02-01", "2024-02-15"]),
    "End_Date": pd.to_datetime(["2024-06-01", "2024-06-01", "2024-07-01", "2024-07-01"]),
    "Proposal_Cost": [100000, 200000, 150000, 120000],
    "Rate": [10, 10, 12, 10],
    "Final_Cost": [110000, 220000, 168000, 132000],
    "Profit": [10000, 20000, 18000, 12000],
    "Status": ["Open", "Closed", "Open", "Open"]
})

# =====================================================
# ============== SIDEBAR NAVIGATION ===================
# =====================================================
st.sidebar.title("📂 Sigma Consultants")

if st.sidebar.button("🏠 Dashboard"):
    st.session_state.page = "Dashboard"

if st.sidebar.button("📊 Summary"):
    st.session_state.page = "Summary"

if st.sidebar.button("📤 Export Data"):
    st.session_state.page = "Export Data"

st.sidebar.markdown("---")
st.sidebar.caption(f"Current Page: {st.session_state.page}")

# =====================================================
# ================== DASHBOARD ========================
# =====================================================
if st.session_state.page == "Dashboard":

    st.header("🏠 Dashboard")

    total_inv = proposals_df["Proposal_Cost"].sum()
    total_profit = proposals_df["Profit"].sum()
    open_cnt = len(proposals_df[proposals_df["Status"] == "Open"])

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Investment", f"₹ {total_inv:,.2f}")
    c2.metric("Total Profit", f"₹ {total_profit:,.2f}")
    c3.metric("Open Proposals", open_cnt)

# =====================================================
# ================== SUMMARY ==========================
# =====================================================
elif st.session_state.page == "Summary":

    st.header("📊 Summary")

    if proposals_df.empty:
        st.info("No data available")
        st.stop()

    st.dataframe(proposals_df, use_container_width=True)

# =====================================================
# ================= EXPORT DATA =======================
# =====================================================
import io
from openpyxl.styles import Font, PatternFill

st.markdown("---")
st.subheader("📤 Export Data")

# ---------- SAFETY CHECK ----------
if proposals_df.empty:
    st.info("No data available for export")
    st.stop()

# ---------- DATE CONVERSION ----------
proposals_df["Start_Date"] = pd.to_datetime(
    proposals_df["Start_Date"], errors="coerce"
)
proposals_df["End_Date"] = pd.to_datetime(
    proposals_df["End_Date"], errors="coerce"
)

# =====================================================
# ================= EXPORT FILTERS ====================
# =====================================================
st.markdown("### 🔎 Export Filters")

# 1️⃣ STATUS FILTER
status_options = ["All"] + sorted(
    proposals_df["Status"].dropna().unique().tolist()
)

export_status = st.selectbox("Select Status", status_options)

if export_status != "All":
    export_df = proposals_df[
        proposals_df["Status"] == export_status
    ].copy()
else:
    export_df = proposals_df.copy()

if export_df.empty:
    st.warning("No data for selected status")
    st.stop()

# 2️⃣ DATE RANGE FILTER
min_date = export_df["Start_Date"].min().date()
max_date = export_df["End_Date"].max().date()

date_range = st.date_input(
    "Select Date Range (Start Date → End Date)",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)

if len(date_range) != 2:
    st.warning("Please select a valid date range")
    st.stop()

start_date, end_date = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])

export_df = export_df[
    (export_df["Start_Date"] >= start_date) &
    (export_df["End_Date"] <= end_date)
].copy()

if export_df.empty:
    st.warning("No data for selected date range")
    st.stop()

# 3️⃣ CLIENT FILTER
client_options = ["All"] + sorted(
    export_df["Client_Name"].dropna().unique().tolist()
)

export_client = st.selectbox("Select Client", client_options)

if export_client != "All":
    export_df = export_df[
        export_df["Client_Name"] == export_client
    ].copy()

if export_df.empty:
    st.warning("No data for selected client")
    st.stop()

# =====================================================
# ========== PREPARE EXPORT DATA ======================
# =====================================================
export_df["Rate"] = export_df["Rate"].round(0).astype(int)

final_export_df = export_df[
    [
        "Client_Name",
        "Start_Date",
        "Proposal_Cost",
        "Rate",
        "Final_Cost",
        "Profit"
    ]
].copy()

final_export_df["Start_Date"] = final_export_df["Start_Date"].dt.strftime(
    "%d-%m-%Y"
)

# =====================================================
# ========== GRAND TOTAL ROW ===========================
# =====================================================
grand_total = {
    "Client_Name": "GRAND TOTAL",
    "Start_Date": "",
    "Proposal_Cost": final_export_df["Proposal_Cost"].sum(),
    "Rate": "",
    "Final_Cost": final_export_df["Final_Cost"].sum(),
    "Profit": final_export_df["Profit"].sum()
}

final_export_df = pd.concat(
    [final_export_df, pd.DataFrame([grand_total])],
    ignore_index=True
)

# =====================================================
# ========== SUMMARY SHEET =============================
# =====================================================
summary_df = pd.DataFrame({
    "Metric": [
        "Total Proposals",
        "Total Investment",
        "Total Final Amount",
        "Total Profit"
    ],
    "Value": [
        len(export_df),
        export_df["Proposal_Cost"].sum(),
        export_df["Final_Cost"].sum(),
        export_df["Profit"].sum()
    ]
})

# =====================================================
# ============== EXCEL EXPORT FUNCTION =================
# =====================================================
def export_excel(data_df, summary_df):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        data_df.to_excel(writer, index=False, sheet_name="Data")
        summary_df.to_excel(writer, index=False, sheet_name="Summary")

        wb = writer.book
        ws = writer.sheets["Data"]

        # ---------- STYLE GRAND TOTAL ROW ----------
        last_row = ws.max_row
        bold_font = Font(bold=True)
        fill = PatternFill(
            start_color="FFF4CCCC",
            end_color="FFF4CCCC",
            fill_type="solid"
        )

        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=last_row, column=col)
            cell.font = bold_font
            cell.fill = fill

    output.seek(0)
    return output

# =====================================================
# ============== DOWNLOAD BUTTON =======================
# =====================================================
st.download_button(
    label="⬇️ Download Excel (Advanced)",
    data=export_excel(final_export_df, summary_df),
    file_name="Sigma_Advanced_Export.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
