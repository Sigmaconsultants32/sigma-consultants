# =====================================================
# Sigma Consultants – Production Build v2.0
# Password: sigma123
# =====================================================

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os

# ---------------- PASSWORD -----------------
PASSWORD = "sigma123"
if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔐 Sigma Consultants Login")
    pwd = st.text_input("Enter Password", type="password")
    if st.button("Login"):
        if pwd == PASSWORD:
            st.session_state.auth = True
            st.rerun()
        else:
            st.error("Incorrect password")
    st.stop()

# ---------------- CONFIG -----------------
st.set_page_config(page_title="Sigma Consultants", layout="wide")

st.markdown("""
<style>
.header {border-bottom:3px solid #1f77ff;margin-bottom:20px;padding-bottom:6px;}
.btn-blue button {background:#1f77ff !important;color:white;}
</style>
""", unsafe_allow_html=True)

# ---------------- STATE -----------------
if "page" not in st.session_state:
    st.session_state.page = "Home"

# ---------------- FILES -----------------
CLIENT_FILE = "clients.xlsx"
PROPOSAL_FILE = "proposals.xlsx"

# ---------------- LOAD -----------------
def load_clients():
    if os.path.exists(CLIENT_FILE):
        return pd.read_excel(CLIENT_FILE)
    df = pd.DataFrame(columns=["Client_ID","Client_Name","Created_Date"])
    df.to_excel(CLIENT_FILE,index=False)
    return df

def load_proposals():
    if os.path.exists(PROPOSAL_FILE):
        df = pd.read_excel(PROPOSAL_FILE)
        if "Closing_Date" not in df.columns:
            df["Closing_Date"]=""
        return df
    df = pd.DataFrame(columns=["Proposal_ID","Client_ID","Client_Name","Proposal_Cost","Rate","Final_Cost","Profit","Start_Date","End_Date","Status","Closing_Date"])
    df.to_excel(PROPOSAL_FILE,index=False)
    return df

clients_df = load_clients()
proposals_df = load_proposals()

for c in ["Start_Date","End_Date","Closing_Date"]:
    if c in proposals_df.columns:
        proposals_df[c] = pd.to_datetime(proposals_df[c], errors="coerce")

def save_clients(df): df.to_excel(CLIENT_FILE,index=False)
def save_proposals(df): df.to_excel(PROPOSAL_FILE,index=False)

# ---------------- UTIL -----------------
def new_client_id(): return f"SIG-C-{len(clients_df)+1:03d}"
def new_proposal_id(): return f"SIG-P-{len(proposals_df)+1:03d}"

def calc(cost,rate,days):
    months = days/30
    profit = cost*(rate/100)*months
    return round(cost+profit,2), round(profit,2)

# ---------------- HOME -----------------
if st.session_state.page=="Home":
    st.markdown("<h2 class='header'>Sigma Consultants</h2>",unsafe_allow_html=True)
    c1,c2=st.columns(2)
    with c1:
        if st.button("➕ Add Client"): st.session_state.page="Add Client"
        if st.button("📄 Add Proposal"): st.session_state.page="Add Proposal"
        if st.button("✏️ Edit / Delete"): st.session_state.page="Edit"
    with c2:
        if st.button("🔍 Find Details"): st.session_state.page="Find"
        if st.button("📊 Summary"): st.session_state.page="Summary"
    st.stop()
# ================= ADD CLIENT =====================
if st.session_state.page=="Add Client":
    st.markdown("<h2 class='header'>Add Client</h2>",unsafe_allow_html=True)
    with st.form("c"):
        name = st.text_input("Client Name")
        submit = st.form_submit_button("Save Client")
        if submit:
            if name.strip()=="":
                st.error("Name required")
            elif name.lower() in clients_df["Client_Name"].str.lower().values:
                st.warning("Client already exists")
            else:
                clients_df.loc[len(clients_df)] = [new_client_id(), name.strip(), datetime.today().date()]
                save_clients(clients_df)
                st.success("Client added successfully")
    if st.button("⬅️ Back"): st.session_state.page="Home"

# ================= ADD PROPOSAL ===================
if st.session_state.page=="Add Proposal":
    st.markdown("<h2 class='header'>Add Proposal</h2>",unsafe_allow_html=True)
    if clients_df.empty:
        st.warning("Add client first")
    else:
        cmap = dict(zip(clients_df["Client_Name"], clients_df["Client_ID"]))
        with st.form("p"):
            cname = st.selectbox("Client", sorted(cmap.keys()))
            cost = st.number_input("Proposal Cost",0.0,step=1000.0)
            rate = st.selectbox("Rate (% per month)", list(range(5,16)))
            start = st.date_input("Start Date", datetime.today())
            days = st.number_input("Duration (days)",1)
            submit = st.form_submit_button("Save Proposal")
            if submit:
                final,profit = calc(cost,rate,days)
                proposals_df.loc[len(proposals_df)] = [
                    new_proposal_id(), cmap[cname], cname, cost, rate, final, profit,
                    start, start+timedelta(days=int(days)), "Open", ""
                ]
                save_proposals(proposals_df)
                st.success("Proposal added successfully")
    if st.button("⬅️ Back"): st.session_state.page="Home"

# ================= EDIT / DELETE ==================
if st.session_state.page=="Edit":
    st.markdown("<h2 class='header'>Edit / Delete Proposals</h2>",unsafe_allow_html=True)
    if not proposals_df.empty:
        pid = st.selectbox("Select Proposal", proposals_df["Proposal_ID"])
        row = proposals_df[proposals_df["Proposal_ID"]==pid].iloc[0]

        with st.form("e"):
            cost = st.number_input("Proposal Cost", float(row["Proposal_Cost"]))
            rate = st.selectbox("Rate", list(range(5,16)), index=int(row["Rate"])-5)
            days = (row["End_Date"] - row["Start_Date"]).days
            status = st.selectbox("Status", ["Open","Closed"], 0 if row["Status"]=="Open" else 1)
            submit = st.form_submit_button("Update")
            if submit:
                final,profit = calc(cost,rate,days)
                proposals_df.loc[proposals_df["Proposal_ID"]==pid,
                    ["Proposal_Cost","Rate","Final_Cost","Profit","Status"]] = [cost,rate,final,profit,status]
                save_proposals(proposals_df)
                st.success("Updated")

        st.markdown("---")
        dels = st.multiselect("Select proposals to delete", proposals_df["Proposal_ID"])
        if st.button("Delete Selected"):
            proposals_df = proposals_df[~proposals_df["Proposal_ID"].isin(dels)]
            save_proposals(proposals_df)
            st.success("Deleted")
    if st.button("⬅️ Back"): st.session_state.page="Home"

# ================= FIND DETAILS ===================
if st.session_state.page=="Find":
    st.markdown("<h2 class='header'>Find Details</h2>",unsafe_allow_html=True)
    f = proposals_df.copy()
    client = st.selectbox("Client", ["All"]+sorted(f["Client_Name"].unique().tolist()))
    if client!="All": f = f[f["Client_Name"]==client]

    st.markdown("### Grand Totals")
    st.metric("Total Profit", f["Profit"].sum())
    st.metric("Total Turnover", f["Final_Cost"].sum())

    st.dataframe(f,use_container_width=True)
    if st.button("⬅️ Back"): st.session_state.page="Home"

# ================= SUMMARY ========================
if st.session_state.page=="Summary":
    st.markdown("<h2 class='header'>Summary</h2>",unsafe_allow_html=True)
    st.metric("Total Clients", len(clients_df))
    st.metric("Total Proposals", len(proposals_df))
    st.metric("Total Turnover", proposals_df["Final_Cost"].sum())
    st.metric("Total Profit", proposals_df["Profit"].sum())

    st.markdown("### Client-wise Contribution")
    if not proposals_df.empty:
        table = proposals_df.groupby("Client_Name").agg(
            Turnover=("Final_Cost","sum"),
            Profit=("Profit","sum")
        ).reset_index().sort_values("Turnover",ascending=False)
        st.dataframe(table,use_container_width=True)

    if st.button("⬅️ Back"): st.session_state.page="Home"
