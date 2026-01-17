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

    st.write("This page is working.")

    status = st.selectbox(
        "Select Status",
        ["All"] + proposals_df["Status"].unique().tolist()
    )

    if status != "All":
        export_df = proposals_df[proposals_df["Status"] == status]
    else:
        export_df = proposals_df.copy()

    export_df = export_df[
        ["Client_Name", "Start_Date", "Proposal_Cost", "Rate", "Final_Cost", "Profit"]
    ].copy()

    export_df["Start_Date"] = export_df["Start_Date"].dt.strftime("%d-%m-%Y")

    total_row = {
        "Client_Name": "GRAND TOTAL",
        "Start_Date": "",
        "Proposal_Cost": export_df["Proposal_Cost"].sum(),
        "Rate": "",
        "Final_Cost": export_df["Final_Cost"].sum(),
        "Profit": export_df["Profit"].sum()
    }

    export_df = pd.concat([export_df, pd.DataFrame([total_row])], ignore_index=True)

    def export_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        output.seek(0)
        return output

    st.download_button(
        "⬇️ Download Excel",
        data=export_excel(export_df),
        file_name="Sigma_Export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
