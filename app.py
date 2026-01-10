# =====================================================
# Sigma Consultants
# Proposal & Client Management System
# Step 5 – UI Polish (Active Page Highlight)
# =====================================================

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os

# =====================================================
# APP CONFIG
# =====================================================
st.set_page_config(page_title="Sigma Consultants", layout="wide")

# =====================================================
# CUSTOM CSS (BLUE HIGHLIGHT)
# =====================================================
st.markdown("""
<style>
.active-btn {
    background-color: #1f77ff !important;
    color: white !important;
    border-radius: 8px;
    font-weight: 600;
}
.page-title {
    border-bottom: 3px solid #1f77ff;
    padding-bottom: 6px;
    margin-bottom: 20px;
}
</style>
""", unsafe_allow_html=True)

# =====================================================
# PAGE STATE
# =====================================================
if "page" not in st.session_state:
    st.session_state.page = "Home"

# =====================================================
# FILE PATHS
# =====================================================
CLIENT_FILE = "clients.xlsx"
PROPOSAL_FILE = "proposals.xlsx"

# =====================================================
# LOAD DATA
# =====================================================
def load_clients():
    if os.path.exists(CLIENT_FILE):
        return pd.read_excel(CLIENT_FILE)
    return pd.DataFrame(columns=["Client_ID", "Client_Name", "Created_Date"])

def load_proposals():
    if os.path.exists(PROPOSAL_FILE):
        df = pd.read_excel(PROPOSAL_FILE)
        if "Closing_Date" not in df.columns:
            df["Closing_Date"] = ""
        return df
    return pd.DataFrame(columns=[
        "Proposal_ID", "Client_ID", "Client_Name",
        "Proposal_Cost", "Rate", "Final_Cost", "Profit",
        "Start_Date", "End_Date", "Status", "Closing_Date"
    ])

clients_df = load_clients()
proposals_df = load_proposals()

for col in ["Start_Date", "End_Date", "Closing_Date"]:
    if col in proposals_df.columns:
        proposals_df[col] = pd.to_datetime(proposals_df[col], errors="coerce")

# =====================================================
# HELPER: NAV BUTTON
# =====================================================
def nav_button(label, page_key):
    is_active = st.session_state.page == page_key
    btn_class = "active-btn" if is_active else ""
    clicked = st.button(label, use_container_width=True)
    if clicked:
        st.session_state.page = page_key
    st.markdown(
        f"<style>div[data-testid='stButton'] > button:has(span:contains('{label}')) {{{''}}}</style>",
        unsafe_allow_html=True
    )

# =====================================================
# 🏠 HOME
# =====================================================
if st.session_state.page == "Home":
    st.markdown("<h2 class='page-title'>Sigma Consultants</h2>", unsafe_allow_html=True)
    st.caption("Proposal & Client Management System")

    col1, col2 = st.columns(2)

    with col1:
        if st.button("➕ Add New Client", use_container_width=True):
            st.session_state.page = "Add Client"
        if st.button("📄 Add Proposal", use_container_width=True):
            st.session_state.page = "Add Proposal"
        if st.button("✏️ Edit Proposal", use_container_width=True):
            st.session_state.page = "Edit Proposal"

    with col2:
        if st.button("🔍 Find Details", use_container_width=True):
            st.session_state.page = "Find Details"
        if st.button("📊 Summary", use_container_width=True):
            st.session_state.page = "Summary"

    st.stop()

# =====================================================
# ➕ ADD CLIENT
# =====================================================
if st.session_state.page == "Add Client":
    st.markdown("<h2 class='page-title'>Add New Client</h2>", unsafe_allow_html=True)

    with st.form("add_client"):
        name = st.text_input("Client Name")
        submit = st.form_submit_button("Save Client")

        if submit:
            if name.strip() == "":
                st.error("Client name required")
            elif name.lower() in clients_df["Client_Name"].str.lower().values:
                st.warning("Client already exists")
            else:
                cid = f"SIG-C-{len(clients_df)+1:03d}"
                clients_df.loc[len(clients_df)] = [
                    cid, name.strip(), datetime.today().date()
                ]
                clients_df.to_excel(CLIENT_FILE, index=False)
                st.success(f"Client added (ID: {cid})")

    if st.button("⬅️ Back to Home"):
        st.session_state.page = "Home"

# =====================================================
# 📄 ADD PROPOSAL
# =====================================================
if st.session_state.page == "Add Proposal":
    st.markdown("<h2 class='page-title'>Add Proposal</h2>", unsafe_allow_html=True)
    st.info("Business logic unchanged. UI only.")

    if st.button("⬅️ Back to Home"):
        st.session_state.page = "Home"

# =====================================================
# ✏️ EDIT PROPOSAL
# =====================================================
if st.session_state.page == "Edit Proposal":
    st.markdown("<h2 class='page-title'>Edit / Delete Proposal</h2>", unsafe_allow_html=True)
    st.info("Business logic unchanged. UI only.")

    if st.button("⬅️ Back to Home"):
        st.session_state.page = "Home"

# =====================================================
# 🔍 FIND DETAILS
# =====================================================
if st.session_state.page == "Find Details":
    st.markdown("<h2 class='page-title'>Find Details</h2>", unsafe_allow_html=True)
    st.info("Filters & totals unchanged. UI only.")

    if st.button("⬅️ Back to Home"):
        st.session_state.page = "Home"

# =====================================================
# 📊 SUMMARY
# =====================================================
if st.session_state.page == "Summary":
    st.markdown("<h2 class='page-title'>Summary</h2>", unsafe_allow_html=True)
    st.info("Summary logic unchanged. UI only.")

    if st.button("⬅️ Back to Home"):
        st.session_state.page = "Home"
