import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Sigma Consultants", layout="wide")

# ---------- PAGE STATE ----------
if "page" not in st.session_state:
    st.session_state.page = "Dashboard"

# ---------- SAMPLE DATA ----------
# (Your real data will already exist – this is only for testing)
proposals_df = pd.DataFrame({
    "Client_Name": ["Client A", "Client B", "Client C"],
    "Start_Date": pd.to_datetime(["2024-01-01", "2024-02-01", "2024-03-01"]),
    "End_Date": pd.to_datetime(["2024-06-01", "2024-07-01", "2024-08-01"]),
    "Proposal_Cost": [100000, 200000, 150000],
    "Rate": [10, 12, 10],
    "Final_Cost": [110000, 224000, 165000],
    "Profit": [10000, 24000, 15000],
    "Status": ["Open", "Closed", "Open"]
})

# ---------- SIDEBAR ----------
st.sidebar.title("📂 Sigma Consultants")

if st.sidebar.button("Dashboard"):
    st.session_state.page = "Dashboard"

if st.sidebar.button("Export Data"):
    st.session_state.page = "Export Data"

# ---------- DASHBOARD ----------
if st.session_state.page == "Dashboard":
    st.header("🏠 Dashboard")
    st.write("This page is working.")

# ---------- EXPORT DATA ----------
if st.session_state.page == "Export Data":

    st.header("📤 Export Data")

    if proposals_df.empty:
        st.info("No data available")
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

    # ---------- PREPARE EXPORT ----------
    export_df["Rate"] = export_df["Rate"].round(0).astype(int)

    final_export_df = export_df[
        ["Client_Name", "Start_Date", "Proposal_Cost", "Rate", "Final_Cost", "Profit"]
    ].copy()

    final_export_df["Start_Date"] = final_export_df["Start_Date"].dt.strftime("%d-%m-%Y")

    # ---------- GRAND TOTAL ----------
    total_row = {
        "Client_Name": "GRAND TOTAL",
        "Start_Date": "",
        "Proposal_Cost": final_export_df["Proposal_Cost"].sum(),
        "Rate": "",
        "Final_Cost": final_export_df["Final_Cost"].sum(),
        "Profit": final_export_df["Profit"].sum()
    }

    final_export_df = pd.concat(
        [final_export_df, pd.DataFrame([total_row])],
        ignore_index=True
    )

    # ---------- EXPORT FUNCTION ----------
    def export_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        output.seek(0)
        return output

    # ---------- DOWNLOAD ----------
    st.download_button(
        "⬇️ Download Excel",
        data=export_excel(final_export_df),
        file_name="Sigma_Filtered_Export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
