# =====================================================
# Sigma Consultants – Production Build v2.0
# Password: sigma123
# =====================================================

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import streamlit as st

# Temporary toggle (useful for testing)
is_mobile = st.toggle("📱 Mobile View", value=True)

def card(title, value):
    st.markdown(
        f"""
        <div style="
        padding:16px;
        border-radius:14px;
        background:#ffffff;
        margin-bottom:12px;
        box-shadow:0 4px 10px rgba(0,0,0,0.08)">
        <div style="font-size:14px;color:#555">{title}</div>
        <div style="font-size:26px;font-weight:600">₹ {value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

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
    st.markdown("<h2 class='header'>Sigma Consultants</h2>", unsafe_allow_html=True)

    # --------- DASHBOARD BUTTONS ---------
    c1, c2 = st.columns(2)
    with c1:
        if st.button("➕ Add Client", use_container_width=True):
            st.session_state.page = "Add Client"
        if st.button("📄 Add Proposal", use_container_width=True):
            st.session_state.page = "Add Proposal"
        if st.button("✏️ Edit / Delete", use_container_width=True):
            st.session_state.page = "Edit"

    with c2:
        if st.button("🔍 Find Details", use_container_width=True):
            st.session_state.page = "Find"
        if st.button("📊 Summary", use_container_width=True):
            st.session_state.page = "Summary"

    st.markdown("---")

    # --------- BACKUP SECTION ---------
    import zipfile
    import io

    st.subheader("📦 Data Backup")

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        if os.path.exists("clients.xlsx"):
            zipf.write("clients.xlsx")
        if os.path.exists("proposals.xlsx"):
            zipf.write("proposals.xlsx")

    st.download_button(
        "⬇ Download Backup",
        data=buffer.getvalue(),
        file_name="Sigma_Consultants_Backup.zip",
        mime="application/zip",
        key="backup_download"
    )

    st.stop()

    # --------- GOOGLE DRIVE SECTION ---------
    if "drive_creds" not in st.session_state:
        st.session_state.drive_creds = None
import zipfile
import io

st.subheader("📦 Data Backup")

buffer = io.BytesIO()
with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
    if os.path.exists("clients.xlsx"):
        zipf.write("clients.xlsx")
    if os.path.exists("proposals.xlsx"):
        zipf.write("proposals.xlsx")

st.download_button(
    "⬇ Download Backup",
    data=buffer.getvalue(),
    file_name="Sigma_Consultants_Backup.zip",
    mime="application/zip"
)

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
st.header("✏️ Edit Client Details")

client_name = st.selectbox("Select Client", df["Client Name"].unique())

client = df[df["Client Name"] == client_name].iloc[0]

st.markdown("---")

if is_mobile:
    # MOBILE LAYOUT (VERTICAL)
    amount = st.number_input("Investment Amount (₹)", value=int(client["Amount"]), step=1000)
    rate = st.number_input("Monthly Profit Rate (%)", value=float(client["Rate"]), step=0.1)
    duration = st.number_input("Duration (Months)", value=int(client["Duration"]))
    status = st.selectbox("Status", ["Active", "Closed"], index=0 if client["Status"]=="Active" else 1)

    with st.expander("Additional Details"):
        remarks = st.text_area("Remarks", value=client.get("Remarks", ""))

else:
    # DESKTOP LAYOUT
    col1, col2 = st.columns(2)
    amount = col1.number_input("Investment Amount (₹)", value=int(client["Amount"]), step=1000)
    rate = col2.number_input("Monthly Profit Rate (%)", value=float(client["Rate"]), step=0.1)

    col3, col4 = st.columns(2)
    duration = col3.number_input("Duration (Months)", value=int(client["Duration"]))
    status = col4.selectbox("Status", ["Active", "Closed"], index=0 if client["Status"]=="Active" else 1)

    remarks = st.text_area("Remarks", value=client.get("Remarks", ""))

st.markdown("---")

if st.button("💾 Save Changes", use_container_width=True):
    df.loc[df["Client Name"] == client_name, "Amount"] = amount
    df.loc[df["Client Name"] == client_name, "Rate"] = rate
    df.loc[df["Client Name"] == client_name, "Duration"] = duration
    df.loc[df["Client Name"] == client_name, "Status"] = status
    df.loc[df["Client Name"] == client_name, "Remarks"] = remarks

    st.success("Client details updated successfully")

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
st.header("📊 Summary")

total_investment = df["Amount"].sum()
total_profit = df["Profit"].sum()
active_clients = len(df[df["Status"] == "Active"])

st.markdown("---")

if is_mobile:
    card("Total Investment", f"{total_investment:,.0f}")
    card("Total Profit", f"{total_profit:,.0f}")
    card("Active Clients", active_clients)
else:
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Investment", f"₹ {total_investment:,.0f}")
    c2.metric("Total Profit", f"₹ {total_profit:,.0f}")
    c3.metric("Active Clients", active_clients)

st.markdown("---")

with st.expander("📄 View Detailed Report"):
    st.dataframe(df, use_container_width=True)













