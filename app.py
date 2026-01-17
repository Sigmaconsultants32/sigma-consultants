# =====================================================
# Sigma Consultants – Full CRM (Final Stable Build)
# =====================================================

import streamlit as st
import pandas as pd
from datetime import datetime
import os, io

# ---------------- PAGE CONFIG ----------------
st.set_page_config(page_title="Sigma Consultants", layout="wide")

# ---------------- MOBILE TOGGLE (STABLE) ----------------
if "is_mobile" not in st.session_state:
    st.session_state.is_mobile = True

st.session_state.is_mobile = st.toggle(
    "📱 Mobile View",
    value=st.session_state.is_mobile
)
is_mobile = st.session_state.is_mobile

# ---------------- CARD UI ----------------
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
            st.session_state.page = "Welcome"
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
        df = pd.read_excel(CLIENT_FILE)
    else:
        df = pd.DataFrame(columns=[
            "Client_ID","Client_Name","Created_Date",
            "Is_Archived","Notes"
        ])
        df.to_excel(CLIENT_FILE, index=False)

    if "Is_Archived" not in df.columns:
        df["Is_Archived"] = False
    if "Notes" not in df.columns:
        df["Notes"] = ""

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
        df.to_excel(PROPOSAL_FILE, index=False)

    for c in ["Start_Date","End_Date","Closing_Date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    return df

# ---------------- SESSION DATA ----------------
if "clients_df" not in st.session_state:
    st.session_state.clients_df = load_clients()

if "proposals_df" not in st.session_state:
    st.session_state.proposals_df = load_proposals()

clients_df = st.session_state.clients_df
proposals_df = st.session_state.proposals_df

def save_clients(): clients_df.to_excel(CLIENT_FILE, index=False)
def save_proposals(): proposals_df.to_excel(PROPOSAL_FILE, index=False)

def new_client_id(): return f"SIG-C-{len(clients_df)+1:03d}"
def new_proposal_id(): return f"SIG-P-{len(proposals_df)+1:03d}"

# ---------------- PAGE STATE ----------------
if "page" not in st.session_state:
    st.session_state.page = "Welcome"

# ---------------- SIDEBAR ----------------
with st.sidebar:
    st.markdown("## 📂 Sigma Consultants")

    menu = [
        ("🏠 Summary","Summary"),
        ("➕ Add Proposal","AddProposal"),
        ("🔍 Find Details","Find"),
        ("✏️ Edit Proposal","Edit"),
        ("👤 Clients","Clients"),
        ("📊 Client Dashboard","ClientDashboard"),
        ("📥 Export Data","Export"),
    ]

    for label, page in menu:
        if st.button(label, use_container_width=True):
            st.session_state.page = page
            st.rerun()

    st.markdown("---")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.auth = False
        st.session_state.page = "Welcome"
        st.rerun()

# =====================================================
# ================= WELCOME ===========================
# =====================================================
if st.session_state.page == "Welcome":

    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        if os.path.exists("sigma_logo.png"):
            st.image("sigma_logo.png", use_container_width=True)
        else:
            st.markdown("<h2 style='text-align:center;'>Sigma Consultants</h2>",
                        unsafe_allow_html=True)

        st.markdown(
            """
            <div style="text-align:center;padding:20px;
            background:#ffffff;border-radius:16px;
            box-shadow:0 4px 12px rgba(0,0,0,0.08)">
            <h3>Welcome</h3>
            <p>Manage clients, proposals, profits and follow-ups in one place.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button("🚀 Go to Dashboard", use_container_width=True):
            st.session_state.page = "Summary"
            st.rerun()

# =====================================================
# ================= SUMMARY ===========================
# =====================================================
if st.session_state.page == "Summary":

    st.header("📊 Summary")

    if proposals_df.empty:
        st.info("No proposals available")
        st.stop()

    total_inv = proposals_df["Proposal_Cost"].sum()
    total_profit = proposals_df["Profit"].sum()
    open_cnt = len(proposals_df[proposals_df["Status"]=="Open"])

    if is_mobile:
        card("Total Investment", f"₹ {total_inv:,.2f}")
        card("Total Profit", f"₹ {total_profit:,.2f}")
        card("Open Proposals", open_cnt)
    else:
        c1,c2,c3 = st.columns(3)
        c1.metric("Total Investment", f"₹ {total_inv:,.2f}")
        c2.metric("Total Profit", f"₹ {total_profit:,.2f}")
        c3.metric("Open Proposals", open_cnt)

    st.markdown("---")
    st.subheader("📅 End Date Based Summary")

    status = st.selectbox(
        "Status",
        ["All"] + sorted(proposals_df["Status"].dropna().unique())
    )

    df = proposals_df.copy()
    if status != "All":
        df = df[df["Status"] == status]

    end_dates = sorted(df["End_Date"].dropna().unique())
    if not end_dates:
        st.info("No End Dates available")
        st.stop()

    end_date = st.selectbox(
        "End Date",
        end_dates,
        format_func=lambda x: x.strftime("%d-%m-%Y")
    )

    df = df[df["End_Date"] == end_date].copy()
    df["Rate_Int"] = df["Rate"].round(0).astype(int)

    summary_df = (
        df.groupby(["Rate_Int","Start_Date"], as_index=False)
        .agg({
            "Proposal_Cost":"sum",
            "Final_Cost":"sum",
            "Profit":"sum"
        })
        .sort_values(["Rate_Int","Start_Date"])
    )

    for _,r in summary_df.iterrows():
        st.markdown(
            f"""
**Start Date** : {r['Start_Date'].strftime('%d-%m-%Y')}  
**Rate** : {r['Rate_Int']} %  
**Investment** : ₹ {r['Proposal_Cost']:,.2f}  
**Final Amount** : ₹ {r['Final_Cost']:,.2f}  
**Profit** : 🟢 ₹ {r['Profit']:,.2f}
"""
        )
        st.markdown("---")

# =====================================================
# ================= EXPORT ============================
# =====================================================
if st.session_state.page == "Export":

    st.header("📥 Export Data")

    def export_excel():
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            clients_df.to_excel(w, sheet_name="Clients", index=False)
            proposals_df.to_excel(w, sheet_name="Proposals", index=False)
        out.seek(0)
        return out

    st.download_button(
        "📥 Download Excel Backup",
        data=export_excel(),
        file_name="Sigma_Consultants_Data.xlsx",
        use_container_width=True
    )
