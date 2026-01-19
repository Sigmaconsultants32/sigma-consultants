# =====================================================
# Sigma Consultants – Full CRM (Proposal-ID Corrected)
# =====================================================

import streamlit as st
import pandas as pd
from datetime import datetime
import os, io

# ---------------- PAGE CONFIG ----------------
st.set_page_config(page_title="Sigma Consultants", layout="wide")

# ---------------- MOBILE TOGGLE ----------------
is_mobile = st.toggle("📱 Mobile View", value=True)

def card(title, value):
    st.markdown(
        f"""
        <div style="padding:14px;border-radius:12px;
        background:#ffffff;margin-bottom:10px;
        box-shadow:0 3px 8px rgba(0,0,0,0.08)">
        <div style="font-size:13px;color:#555">{title}</div>
        <div style="font-size:22px;font-weight:600">{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ---------------- LOGIN ----------------
PASSWORD = "sigma123"

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔐 Sigma Consultants Login")
    pwd = st.text_input("Enter Password", type="password")

    if st.button("Login", use_container_width=True):
        if pwd == PASSWORD:
            st.session_state.auth = True
            st.session_state.page = "Summary"
            st.rerun()
        else:
            st.error("Incorrect password")
    st.stop()

# ---------------- FILES ----------------
CLIENT_FILE = "clients.xlsx"
PROPOSAL_FILE = "proposals.xlsx"

# ---------------- LOADERS ----------------
def load_clients():
    if os.path.exists(CLIENT_FILE):
        return pd.read_excel(CLIENT_FILE)
    df = pd.DataFrame(columns=["Client_ID","Client_Name","Created_Date","Is_Archived"])
    df.to_excel(CLIENT_FILE,index=False)
    return df

def load_proposals():
    if os.path.exists(PROPOSAL_FILE):
        df = pd.read_excel(PROPOSAL_FILE)
    else:
        df = pd.DataFrame(columns=[
            "Proposal_ID","Client_ID","Client_Name",
            "Proposal_Cost","Rate",
            "Start_Date","End_Date","Status"
        ])
        df.to_excel(PROPOSAL_FILE,index=False)

    df["Start_Date"] = pd.to_datetime(df["Start_Date"])
    df["End_Date"] = pd.to_datetime(df["End_Date"])
    df["Proposal_Cost"] = df["Proposal_Cost"].astype(float)
    df["Rate"] = df["Rate"].astype(float)
    return df

clients_df = load_clients()
proposals_df = load_proposals()

def save_clients(): clients_df.to_excel(CLIENT_FILE,index=False)
def save_proposals(): proposals_df.to_excel(PROPOSAL_FILE,index=False)

# ---------------- CORE CALCULATION (ONLY SOURCE OF TRUTH) ----------------
def proposal_summary(df):
    g = df.groupby("Proposal_ID").agg({
        "Start_Date":"first",
        "End_Date":"first",
        "Rate":"first",
        "Proposal_Cost":"sum"
    }).reset_index()

    g["Days"] = (g["End_Date"] - g["Start_Date"]).dt.days
    g["Profit"] = g["Proposal_Cost"] * g["Rate"] * g["Days"] / 36500
    g["Final_Amount"] = g["Proposal_Cost"] + g["Profit"]
    return g.round(2)

# ---------------- PAGE STATE ----------------
if "page" not in st.session_state:
    st.session_state.page = "Summary"

# ---------------- SIDEBAR ----------------
with st.sidebar:
    st.markdown("## 📂 Sigma Consultants")
    for p, label in [
        ("Summary","🏠 Summary"),
        ("Add","➕ Add Proposal"),
        ("Find","🔍 Find Details"),
        ("Clients","👤 Clients"),
        ("Export","📤 Export Data")
    ]:
        if st.button(label, use_container_width=True):
            st.session_state.page = p
            st.rerun()

# =====================================================
# ================= SUMMARY ===========================
# =====================================================
if st.session_state.page == "Summary":

    st.header("📊 Proposal Summary")

    if proposals_df.empty:
        st.info("No proposals available")
        st.stop()

    ps = proposal_summary(proposals_df)

    if is_mobile:
        card("Investment", f"₹ {ps['Proposal_Cost'].sum():,.2f}")
        card("Profit", f"₹ {ps['Profit'].sum():,.2f}")
        card("Final", f"₹ {ps['Final_Amount'].sum():,.2f}")
    else:
        c1,c2,c3 = st.columns(3)
        c1.metric("Investment", f"₹ {ps['Proposal_Cost'].sum():,.2f}")
        c2.metric("Profit", f"₹ {ps['Profit'].sum():,.2f}")
        c3.metric("Final Amount", f"₹ {ps['Final_Amount'].sum():,.2f}")

    st.dataframe(ps, use_container_width=True)

# =====================================================
# ================= ADD PROPOSAL ======================
# =====================================================
if st.session_state.page == "Add":

    st.header("➕ Add Client Investment to Proposal")

    client = st.selectbox("Client", clients_df["Client_Name"].unique())
    pid = st.text_input("Proposal ID (same for same dates & rate)")
    cost = st.number_input("Amount", min_value=0.0, step=1000.0)
    rate = st.number_input("Annual Rate (%)", min_value=0.0)
    sd = st.date_input("Start Date")
    ed = st.date_input("End Date")

    if st.button("Save"):
        new = {
            "Proposal_ID": pid,
            "Client_ID": clients_df.loc[clients_df["Client_Name"]==client,"Client_ID"].values[0],
            "Client_Name": client,
            "Proposal_Cost": cost,
            "Rate": rate,
            "Start_Date": sd,
            "End_Date": ed,
            "Status": "Open"
        }
        proposals_df.loc[len(proposals_df)] = new
        save_proposals()
        st.success("Saved")
        st.rerun()

# =====================================================
# ================= FIND DETAILS ======================
# =====================================================
if st.session_state.page == "Find":

    pid = st.selectbox("Select Proposal ID", proposals_df["Proposal_ID"].unique())
    df = proposals_df[proposals_df["Proposal_ID"] == pid]

    st.subheader("Client-wise Investments")
    st.dataframe(df, use_container_width=True)

    st.subheader("Proposal Summary")
    st.dataframe(proposal_summary(df), use_container_width=True)

# =====================================================
# ================= EXPORT DATA =======================
# =====================================================
if st.session_state.page == "Export":

    def export_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            proposals_df.to_excel(writer, index=False, sheet_name="Raw_Data")
            proposal_summary(proposals_df).to_excel(
                writer, index=False, sheet_name="Proposal_Summary"
            )
        output.seek(0)
        return output

    st.download_button(
        "⬇️ Download Excel",
        data=export_excel(),
        file_name="sigma_consultants.xlsx"
    )
