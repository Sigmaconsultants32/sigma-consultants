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
elif st.session_state.page == "Export Data":

    st.header("📤 Export Data")

    if proposals_df.empty:
        st.info("No data available for export")
        st.stop()

    # ---------- STATUS FILTER ----------
    status = st.selectbox(
        "Select Status",
        ["All"] + sorted(proposals_df["Status"].unique().tolist())
    )

    if status != "All":
        export_df = proposals_df[proposals_df["Status"] == status].copy()
    else:
        export_df = proposals_df.copy()

    if export_df.empty:
        st.warning("No data after status filter")
        st.stop()

    # ---------- DATE RANGE FILTER ----------
    min_date = export_df["Start_Date"].min().date()
    max_date = export_df["End_Date"].max().date()

    date_range = st.date_input(
        "Select Date Range",
        value=(min_date, max_date)
    )

    if len(date_range) != 2:
        st.stop()

    start_date = pd.to_datetime(date_range[0])
    end_date = pd.to_datetime(date_range[1])

    export_df = export_df[
        (export_df["Start_Date"] >= start_date) &
        (export_df["End_Date"] <= end_date)
    ].copy()

    if export_df.empty:
        st.warning("No data for selected date range")
        st.stop()

    # ---------- CLIENT FILTER ----------
    client = st.selectbox(
        "Select Client",
        ["All"] + sorted(export_df["Client_Name"].unique().tolist())
    )

    if client != "All":
        export_df = export_df[export_df["Client_Name"] == client].copy()

    if export_df.empty:
        st.warning("No data for selected client")
        st.stop()

    # ---------- PREPARE DATA SHEET ----------
    export_df["Rate"] = export_df["Rate"].round(0).astype(int)

    data_sheet = export_df[
        ["Client_Name", "Start_Date", "Proposal_Cost", "Rate", "Final_Cost", "Profit"]
    ].copy()

    data_sheet["Start_Date"] = data_sheet["Start_Date"].dt.strftime("%d-%m-%Y")

    # ---------- GRAND TOTAL ----------
    grand_total = {
        "Client_Name": "GRAND TOTAL",
        "Start_Date": "",
        "Proposal_Cost": data_sheet["Proposal_Cost"].sum(),
        "Rate": "",
        "Final_Cost": data_sheet["Final_Cost"].sum(),
        "Profit": data_sheet["Profit"].sum()
    }

    data_sheet = pd.concat(
        [data_sheet, pd.DataFrame([grand_total])],
        ignore_index=True
    )

    # ---------- SUMMARY SHEET ----------
    summary_sheet = pd.DataFrame({
        "Metric": [
            "Total Records",
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

    # ---------- EXPORT FUNCTION ----------
    def export_excel(data_df, summary_df):
        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            data_df.to_excel(writer, index=False, sheet_name="Data")
            summary_df.to_excel(writer, index=False, sheet_name="Summary")

            ws = writer.sheets["Data"]
            last_row = ws.max_row

            bold = Font(bold=True)
            fill = PatternFill(
                start_color="FFF4CCCC",
                end_color="FFF4CCCC",
                fill_type="solid"
            )

            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=last_row, column=col)
                cell.font = bold
                cell.fill = fill

        output.seek(0)
        return output

    # ---------- DOWNLOAD ----------
    st.download_button(
        "⬇️ Download Excel (Data + Summary)",
        data=export_excel(data_sheet, summary_sheet),
        file_name="Sigma_Export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
