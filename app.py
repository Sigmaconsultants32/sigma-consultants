# =====================================================
# Sigma Consultants – Production Build v2.3
# =====================================================

import streamlit as st
import pandas as pd
from datetime import datetime
import os

# ---------------- PAGE CONFIG ----------------
st.set_page_config(page_title="Sigma Consultants", layout="wide")

# ---------------- MOBILE TOGGLE ----------------
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
        <div style="font-size:22px;font-weight:600">{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ---------------- PASSWORD ----------------
PASSWORD = "sigma123"
if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔐 Sigma Consultants Login")
    pwd = st.text_input("Enter Password", type="password")
    if st.button("Login", use_container_width=True):
        if pwd == PASSWORD:
            st.session_state.auth = True
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
    df = pd.DataFrame(columns=["Client_ID","Client_Name","Created_Date"])
    df.to_excel(CLIENT_FILE,index=False)
    return df

def load_proposals():
    if os.path.exists(PROPOSAL_FILE):
        df = pd.read_excel(PROPOSAL_FILE)
    else:
        df = pd.DataFrame(columns=[
            "Proposal_ID","Client_ID","Client_Name",
            "Proposal_Cost","Rate","Final_Cost","Profit",
            "Start_Date","End_Date","Status","Closing_Date"
        ])
        df.to_excel(PROPOSAL_FILE,index=False)

    for c in ["Start_Date","End_Date","Closing_Date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

# ---------------- SESSION STATE ----------------
if "clients_df" not in st.session_state:
    st.session_state.clients_df = load_clients()

if "proposals_df" not in st.session_state:
    st.session_state.proposals_df = load_proposals()

clients_df = st.session_state.clients_df
proposals_df = st.session_state.proposals_df

def save_clients(): st.session_state.clients_df.to_excel(CLIENT_FILE,index=False)
def save_proposals(): st.session_state.proposals_df.to_excel(PROPOSAL_FILE,index=False)

# ---------------- UTIL ----------------
def new_client_id(): return f"SIG-C-{len(clients_df)+1:03d}"
def new_proposal_id(): return f"SIG-P-{len(proposals_df)+1:03d}"

# ---------------- PAGE STATE ----------------
if "page" not in st.session_state:
    st.session_state.page = "Summary"

# ---------------- SIDEBAR ----------------
with st.sidebar:
    st.markdown("## 📂 Sigma Consultants")

    if st.button("🏠 Home / Summary", use_container_width=True):
        st.session_state.page = "Summary"; st.rerun()

    if st.button("➕ Add Proposal", use_container_width=True):
        st.session_state.page = "AddProposal"; st.rerun()
    
    if st.button("🔍 Find Details", use_container_width=True):
        st.session_state.page = "Find"; st.rerun()

    if st.button("✏️ Edit Proposal", use_container_width=True):
        st.session_state.page = "Edit"; st.rerun()

    if st.button("👤 Clients", use_container_width=True):
        st.session_state.page = "Clients"; st.rerun()

    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.auth = False; st.rerun()

# =====================================================
# ================= SUMMARY ===========================
# =====================================================
if st.session_state.page == "Summary":

    st.header("📊 Summary")

    if proposals_df.empty:
        st.info("No proposals available")
    else:
        total_investment = proposals_df["Proposal_Cost"].sum()
        total_profit = proposals_df["Profit"].sum()
        active = len(proposals_df[proposals_df["Status"]=="Active"])

        if is_mobile:
            card("Total Investment", f"₹ {total_investment:,.0f}")
            card("Total Profit", f"₹ {total_profit:,.0f}")
            card("Active Proposals", active)
        else:
            c1,c2,c3 = st.columns(3)
            c1.metric("Total Investment", f"₹ {total_investment:,.0f}")
            c2.metric("Total Profit", f"₹ {total_profit:,.0f}")
            c3.metric("Active Proposals", active)

# =====================================================
# ================= FIND DETAILS (FINAL) ==============
# =====================================================
if st.session_state.page == "Find":

    st.header("🔍 Find Proposal Details")

    if proposals_df.empty:
        st.warning("No proposal data available")
        st.stop()

    # ---------- FILTERS ----------
    col1, col2 = st.columns(2)
    client = col1.selectbox(
        "Client Name",
        ["All"] + sorted(proposals_df["Client_Name"].dropna().unique())
    )

    status = col2.selectbox(
        "Status",
        ["All", "Open", "Closed"]
    )

    col3, col4 = st.columns(2)
    start_date = col3.date_input(
        "Start Date (From)",
        value=None
    )

    end_date = col4.date_input(
        "End Date (To)",
        value=None
    )

    # ---------- APPLY FILTERS ----------
    result = proposals_df.copy()

    if client != "All":
        result = result[result["Client_Name"] == client]

    if status != "All":
        result = result[result["Status"] == status]

    if start_date:
        result = result[result["Start_Date"] >= pd.to_datetime(start_date)]

    if end_date:
        result = result[result["End_Date"] <= pd.to_datetime(end_date)]

    st.markdown("---")
    st.subheader("📋 Follow-up Details")

    if result.empty:
        st.info("No records found for selected criteria")
        st.stop()

    # ---------- MOBILE VIEW (VERTICAL FOLLOW-UP CARDS) ----------
    if is_mobile:
        for _, r in result.sort_values("Start_Date").iterrows():
            st.markdown(
                f"""
                <div style="
                border:1px solid #ddd;
                border-radius:12px;
                padding:14px;
                margin-bottom:12px;
                background:#fafafa">

                <b>Client:</b> {r['Client_Name']}<br>
                <b>Proposal ID:</b> {r['Proposal_ID']}<br>
                <b>Status:</b> {r['Status']}<br>
                <b>Rate (%):</b> {float(r['Rate']):.2f}%<br>
                <b>Start Date:</b> {r['Start_Date'].date() if pd.notna(r['Start_Date']) else "-"}<br>
                <b>End Date:</b> {r['End_Date'].date() if pd.notna(r['End_Date']) else "-"}<br>
                <b>Proposal Amount:</b> ₹ {float(r['Proposal_Cost']):,.2f}<br>
                <b>Final Amount:</b> ₹ {float(r['Final_Cost']):,.2f}<br>
                <b>Profit:</b> ₹ {float(r['Profit']):,.2f}

                </div>
                """,
                unsafe_allow_html=True
            )

    # ---------- DESKTOP VIEW (FOLLOW-UP TABLE) ----------
    else:
        display_df = result[[
            "Client_Name",
            "Proposal_ID",
            "Rate",
            "Start_Date",
            "End_Date",
            "Proposal_Cost",
            "Final_Cost",
            "Profit",
            "Status"
        ]].sort_values("Start_Date")

        display_df["Rate"] = display_df["Rate"].astype(float).round(2)
        display_df["Proposal_Cost"] = display_df["Proposal_Cost"].astype(float).round(2)
        display_df["Final_Cost"] = display_df["Final_Cost"].astype(float).round(2)
        display_df["Profit"] = display_df["Profit"].astype(float).round(2)
        display_df["Start_Date"] = display_df["Start_Date"].dt.date
        display_df["End_Date"] = display_df["End_Date"].dt.date

        st.dataframe(display_df, use_container_width=True)

# =====================================================
# ================= EDIT PROPOSAL =====================
# =====================================================
if st.session_state.page == "Edit":

    st.header("✏️ Edit Proposal")

    if proposals_df.empty:
        st.warning("No proposals to edit")
        st.stop()

    pid = st.selectbox("Select Proposal ID", proposals_df["Proposal_ID"].unique())
    row = proposals_df[proposals_df["Proposal_ID"] == pid].iloc[0]

    cost = st.number_input("Proposal Cost (₹)", value=int(row["Proposal_Cost"]), step=1000)
    rate = st.number_input("Rate (%)", value=float(row["Rate"]), step=0.1)
    status = st.selectbox("Status", ["Active","Closed"],
                          index=0 if row["Status"]=="Active" else 1)

    if st.button("💾 Save Changes", use_container_width=True):
        proposals_df.loc[proposals_df["Proposal_ID"]==pid, ["Proposal_Cost","Rate","Status"]] = [cost,rate,status]
        save_proposals()
        st.success("Proposal updated successfully")

# =====================================================
# ================= CLIENTS ===========================
# =====================================================
if st.session_state.page == "Clients":

    st.header("👤 Clients")

    if clients_df.empty:
        st.info("No clients available")
    else:
        st.dataframe(clients_df, use_container_width=True)

    with st.expander("➕ Add New Client"):
        cname = st.text_input("Client Name")
        if st.button("Add Client", use_container_width=True):
            if cname.strip():
                clients_df.loc[len(clients_df)] = [new_client_id(), cname, datetime.now()]
                save_clients()
                st.success("Client added successfully")
                st.rerun()

# =====================================================
# ================= ADD NEW PROPOSAL ==================
# =====================================================
if st.session_state.page == "AddProposal":

    st.header("➕ Add New Proposal")

    if clients_df.empty:
        st.warning("Please add a client before creating a proposal.")
        st.stop()

    # ----------- INPUTS -----------
    client_name = st.selectbox(
        "Client Name",
        sorted(clients_df["Client_Name"].unique())
    )

    start_date = st.date_input(
        "Start Date",
        value=datetime.today()
    )

    end_date = st.date_input(
        "End Date",
        value=datetime.today()
    )

    proposal_cost = st.number_input(
        "Proposal Amount (₹)",
        min_value=0.0,
        step=1000.0,
        format="%.2f"
    )

    rate = st.number_input(
        "Monthly Rate (%)",
        min_value=0.0,
        step=0.10,
        format="%.2f"
    )

    # ----------- CALCULATIONS -----------
    days = (end_date - start_date).days

    if days < 0:
        st.error("End Date must be after Start Date")
        st.stop()

    months = days / 30
    profit = proposal_cost * (rate / 100) * months
    final_cost = proposal_cost + profit

    st.markdown("---")
    st.subheader("💰 Proposal Details")

    if is_mobile:
        st.markdown(f"""
        <div style="border:1px solid #ddd;border-radius:12px;padding:14px;background:#fafafa">
        <b>Duration:</b> {days} days ({months:.2f} months)<br>
        <b>Profit:</b> ₹ {profit:,.2f}<br>
        <b>Final Amount:</b> ₹ {final_cost:,.2f}
        </div>
        """, unsafe_allow_html=True)
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Duration (Months)", f"{months:.2f}")
        c2.metric("Profit (₹)", f"{profit:,.2f}")
        c3.metric("Final Amount (₹)", f"{final_cost:,.2f}")

    st.markdown("---")

    # ----------- SAVE -----------
    if st.button("💾 Save Proposal", use_container_width=True):

        new_row = {
            "Proposal_ID": new_proposal_id(),
            "Client_ID": clients_df.loc[
                clients_df["Client_Name"] == client_name, "Client_ID"
            ].values[0],
            "Client_Name": client_name,
            "Proposal_Cost": round(proposal_cost, 2),
            "Rate": round(rate, 2),
            "Final_Cost": round(final_cost, 2),
            "Profit": round(profit, 2),
            "Start_Date": pd.to_datetime(start_date),
            "End_Date": pd.to_datetime(end_date),
            "Status": "Open",
            "Closing_Date": ""
        }

        st.session_state.proposals_df = pd.concat(
            [proposals_df, pd.DataFrame([new_row])],
            ignore_index=True
        )

        save_proposals()

        st.success("✅ Proposal added successfully")
        st.session_state.page = "Summary"
        st.rerun()






