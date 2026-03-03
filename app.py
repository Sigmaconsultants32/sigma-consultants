# =====================================================
# Sigma Consultants – Full CRM Production Build
# =====================================================

import streamlit as st
import pandas as pd
from datetime import datetime
import os, io

# ---------------- PAGE CONFIG ----------------
st.set_page_config(page_title="Sigma Consultants", layout="wide")

# ================= BRAND THEME =================
PRIMARY_COLOR = "#9DD3DB"     # Soft teal (main brand)
SECONDARY_COLOR = "#FF6800"   # Orange (energy)
ACCENT_COLOR = "#EE2929"      # Red (attention)

st.markdown(
    f"""
    <style>

    /* ---------- APP BACKGROUND ---------- */
    .stApp {{
        background: linear-gradient(
            135deg,
            {PRIMARY_COLOR} 0%,
            #ffffff 65%
        );
    }}

    /* ---------- PAGE CONTAINERS (CARDS) ---------- */
    section[data-testid="stVerticalBlock"] > div {{
        background: rgba(255,255,255,0.95);
        border-radius: 18px;
        padding: 18px;
        margin-bottom: 18px;
        box-shadow: 0 12px 28px rgba(0,0,0,0.08);
        border-left: 6px solid {SECONDARY_COLOR};
    }}

    /* ---------- HEADERS ---------- */
    h1, h2, h3 {{
        color: #1f2937; /* dark text for readability */
        font-weight: 700;
    }}

    /* ---------- PRIMARY BUTTONS ---------- */
    button[kind="primary"] {{
        background: {SECONDARY_COLOR};
        border-radius: 12px;
        font-weight: 600;
        border: none;
        color: white;
    }}

    button[kind="primary"]:hover {{
        background: {ACCENT_COLOR};
        color: white;
    }}

    /* ---------- SECONDARY BUTTONS ---------- */
    button {{
        border-radius: 10px;
    }}

    /* ---------- METRICS ---------- */
    div[data-testid="metric-container"] {{
        border-radius: 14px;
        padding: 14px;
        background: linear-gradient(
            145deg,
            #ffffff,
            {PRIMARY_COLOR}
        );
        border: 1px solid #e5e7eb;
    }}

    /* ---------- SIDEBAR (MATCHING BRAND) ---------- */
    section[data-testid="stSidebar"] {{
        background: linear-gradient(
            180deg,
            {SECONDARY_COLOR} 0%,
            {ACCENT_COLOR} 100%
        );
    }}

    </style>
    """,
    unsafe_allow_html=True
)

# ================= SIDEBAR STYLING =================
st.markdown(
    """
    <style>
    /* ---------- SIDEBAR BACKGROUND ---------- */
    section[data-testid="stSidebar"] {
        background: linear-gradient(
            180deg,
            #FFBFAD 0%,
            #FFA08A 100%
        );
        padding-top: 20px;
    }

    /* ---------- SIDEBAR TITLE ---------- */
    .sidebar-title {
        color: #4A1F14;
        font-size: 22px;
        font-weight: 700;
        text-align: center;
        margin-bottom: 20px;
    }

    /* ---------- MENU BUTTON BASE ---------- */
    .sidebar-btn button {
        width: 100%;
        background: transparent;
        color: #4A1F14;
        border: none;
        text-align: left;
        padding: 12px 16px;
        font-size: 15px;
        border-radius: 12px;
        margin-bottom: 8px;
        transition: all 0.2s ease-in-out;
        font-weight: 500;
    }

    /* ---------- HOVER EFFECT ---------- */
    .sidebar-btn button:hover {
        background: rgba(255,255,255,0.45);
        transform: translateX(4px);
    }

    /* ---------- ACTIVE PAGE ---------- */
    .sidebar-active button {
        background: white;
        color: #FF6800;
        font-weight: 700;
        box-shadow: 0 6px 14px rgba(0,0,0,0.12);
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ---------------- MOBILE TOGGLE ----------------
if "is_mobile" not in st.session_state:
    st.session_state.is_mobile = False

st.session_state.is_mobile = st.toggle(
    "📱 Mobile View",
    value=st.session_state.is_mobile
)

is_mobile = st.session_state.is_mobile


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
PASSWORD = os.getenv("SIGMA_PASSWORD", "sigma123")

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:

    st.markdown("<br><br>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:
        st.markdown("""
        <div class="login-box ui-card">
            <h3 style="text-align:center;margin-bottom:10px;">🔐 Sigma Consultants</h3>
        """, unsafe_allow_html=True)

        pwd = st.text_input("Password", type="password")

        if st.button("Login", use_container_width=True, type="primary"):
            if pwd == PASSWORD:
                st.session_state.auth = True
                st.session_state.page = "Welcome"
                st.rerun()
            else:
                st.error("Incorrect password")

        st.markdown("</div>", unsafe_allow_html=True)

    st.stop()

# ---------------- STYLES ----------------
st.markdown("""
<style>

/* Header underline */
.header {
    border-bottom: 3px solid #1f77ff;
    margin-bottom: 20px;
    padding-bottom: 6px;
}

/* Primary buttons */
button[kind="primary"] {
    background-color: #1f77ff !important;
    border-radius: 8px !important;
}

/* Card style */
.ui-card {
    border: 1px solid #e0e0e0;
    border-radius: 14px;
    padding: 14px;
    margin-bottom: 12px;
    background: #fafafa;
}

/* SAFE content padding (does not affect layout engine) */
section.main > div {
    padding-top: 1.5rem;
    padding-bottom: 2rem;
}

/* Login box width */
.login-box {
    max-width: 420px;
    margin: auto;
}

</style>
""", unsafe_allow_html=True)

# ---------------- FILES ----------------
CLIENT_FILE = "clients.xlsx"
PROPOSAL_FILE = "proposals.xlsx"

# ---------------- LOADERS ----------------
def load_clients():
    if os.path.exists(CLIENT_FILE):
        df = pd.read_excel(CLIENT_FILE)
    else:
        df = pd.DataFrame(columns=[
            "Client_ID",
            "Client_Name",
            "Created_Date",
            "Is_Archived",
            "Notes"
        ])
        df.to_excel(CLIENT_FILE, index=False)

    # ---- schema safety for old files ----
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
            "Proposal_ID",
            "Client_ID",
            "Client_Name",
            "Proposal_Cost",
            "Rate",
            "Final_Cost",
            "Profit",
            "Start_Date",
            "End_Date",
            "Status",
            "Closing_Date"
        ])
        df.to_excel(PROPOSAL_FILE, index=False)

    # ---- schema safety for numeric & status columns ----
    default_cols = {
        "Proposal_Cost": 0.0,
        "Rate": 0.0,
        "Final_Cost": 0.0,
        "Profit": 0.0,
        "Status": ""
    }

    for col, val in default_cols.items():
        if col not in df.columns:
            df[col] = val

    # ---- date normalization ----
    for c in ["Start_Date", "End_Date", "Closing_Date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    return df

# ---------------- SESSION STATE ----------------

# ---- Load data once per session ----
if "clients_df" not in st.session_state:
    st.session_state.clients_df = load_clients()

if "proposals_df" not in st.session_state:
    st.session_state.proposals_df = load_proposals()

# ---- Aliases for convenience (read/write safe) ----
clients_df = st.session_state.clients_df
proposals_df = st.session_state.proposals_df


# ---- Save helpers (always save from session_state) ----
def save_clients():
    st.session_state.clients_df.to_excel(CLIENT_FILE, index=False)

def save_proposals():
    st.session_state.proposals_df.to_excel(PROPOSAL_FILE, index=False)

# ---------------- UTIL ----------------

def new_client_id():
    if clients_df.empty or "Client_ID" not in clients_df.columns:
        return "SIG-C-001"

    last_num = (
        clients_df["Client_ID"]
        .str.replace("SIG-C-", "", regex=False)
        .astype(int)
        .max()
    )
    return f"SIG-C-{last_num + 1:03d}"


def new_proposal_id():
    if proposals_df.empty or "Proposal_ID" not in proposals_df.columns:
        return "SIG-P-001"

    last_num = (
        proposals_df["Proposal_ID"]
        .str.replace("SIG-P-", "", regex=False)
        .astype(int)
        .max()
    )
    return f"SIG-P-{last_num + 1:03d}"


def calc(cost, rate, days):
    try:
        cost = float(cost)
        rate = float(rate)
        days = float(days)
    except (TypeError, ValueError):
        return 0.0, 0.0

    months = days / 30
    profit = cost * (rate / 100) * months
    final_amount = cost + profit

    return round(final_amount, 2), round(profit, 2)

# ---------------- PAGE STATE ----------------
if "page" not in st.session_state:
    st.session_state.page = "Welcome"

# ---------------- SIDEBAR ----------------
with st.sidebar:
    st.markdown("## 📂 Sigma Consultants")

    def nav(label, page_key):
        if st.button(label, use_container_width=True):
            st.session_state.page = page_key

    nav("🏠 Summary", "Summary")
    nav("➕ Add Proposal", "AddProposal")
    nav("🔍 Find Details", "Find")
    nav("✏️ Edit Proposal", "Edit")
    nav("👤 Clients", "Clients")
    nav("📊 Client Dashboard", "ClientDashboard")

    st.markdown("---")

    nav("📥 Export Data", "Export")

    st.markdown("---")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.auth = False
        st.session_state.page = "Summary"
    
# =====================================================
# ================= WELCOME SCREEN ====================
# =====================================================
if st.session_state.page == "Welcome":

    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:

        logo_path = "sigma_logo.png"

        if os.path.exists(logo_path):
            st.image(logo_path, width=180)
        else:
            st.markdown("<h3 style='text-align:center;'>Sigma Consultants</h3>", unsafe_allow_html=True)

        st.markdown("""
        <div class="ui-card" style="text-align:center;">
            <p style="color:#555;">
                Manage clients, proposals, profits & follow-ups
            </p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("➕ Add Client", use_container_width=True, type="primary"):
            st.session_state.page = "Clients"

        if st.button("📄 Add Proposal", use_container_width=True):
            st.session_state.page = "AddProposal"

        if st.button("🔍 Find Details", use_container_width=True):
            st.session_state.page = "Find"

# =====================================================
# ================= SUMMARY ===========================
# =====================================================
if st.session_state.page == "Summary":
    st.header("📊 Summary")

    # ---------- SAFETY CHECK ----------
    if proposals_df.empty:
        st.info("No proposals available")
        st.stop()

    # ---------- WORKING COPY (DO NOT MUTATE SESSION DATA) ----------
    df = proposals_df.copy()

    # ---------- BASIC SUMMARY ----------
    total_inv = df["Proposal_Cost"].sum()
    total_profit = df["Profit"].sum()
    open_cnt = len(df[df["Status"] == "Open"])

    if is_mobile:
        card("Investment", f"₹ {total_inv:,.2f}")
        card("Profit", f"₹ {total_profit:,.2f}")
        card("Open", open_cnt)
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Investment", f"₹ {total_inv:,.2f}")
        c2.metric("Total Profit", f"₹ {total_profit:,.2f}")
        c3.metric("Open Proposals", open_cnt)

    # =====================================================
    # ============ DATE BASED SUMMARY =====================
    # =====================================================
    st.markdown("---")
    st.subheader("📅 Date Based Summary")
    st.caption("ℹ️ Amounts shown in actual ₹ (rupees)")

    # ---------- DATE NORMALIZATION ----------
    for c in ["Start_Date", "End_Date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    # =====================================================
    # ================= STATUS FILTER =====================
    # =====================================================
    status_options = ["All"] + sorted(df["Status"].dropna().unique().tolist())
    selected_status = st.selectbox("Select Proposal Status", status_options)

    if selected_status != "All":
        df = df[df["Status"] == selected_status].copy()

    if df.empty:
        st.warning("No data for selected status")
        st.stop()

    # =====================================================
    # ============ START / END DATE SELECTOR ===============
    # =====================================================
    date_type = st.radio(
        "Select Date Type",
        ["Start Date", "End Date"],
        horizontal=True
    )

    date_col = "Start_Date" if date_type == "Start Date" else "End_Date"

    df = df.dropna(subset=[date_col])

    if df.empty:
        st.info("No dates available")
        st.stop()

    df["DateOnly"] = df[date_col].dt.date

    available_dates = sorted(df["DateOnly"].unique())
    date_options = ["All"] + available_dates

    selected_dates = st.multiselect(
        f"Select {date_type}(s)",
        date_options,
        default="All",
        format_func=lambda x: x if x == "All" else x.strftime("%d-%m-%Y")
    )

    if "All" not in selected_dates:
        df = df[df["DateOnly"].isin(selected_dates)]

    if df.empty:
        st.warning("No data for selected date(s)")
        st.stop()

    # =====================================================
    # ================ SAFE RATE NORMALIZATION =============
    # =====================================================
    df["Rate_Int"] = (
        pd.to_numeric(df["Rate"], errors="coerce")
        .fillna(0)
        .round(0)
        .astype(int)
    )

    # =====================================================
    # ================= PROPOSAL STATS ====================
    # =====================================================
    total_inv = df["Proposal_Cost"].sum()
    total_final = df["Final_Cost"].sum()
    total_profit = df["Profit"].sum()

    st.markdown("### 📊 Proposal Stats")

    if is_mobile:
        card("Investment", f"₹ {total_inv:,.2f}")
        card("Final Amount", f"₹ {total_final:,.2f}")
        card("Profit", f"₹ {total_profit:,.2f}")
    else:
        p1, p2, p3 = st.columns(3)
        p1.metric("Total Investment", f"₹ {total_inv:,.2f}")
        p2.metric("Total Final Amount", f"₹ {total_final:,.2f}")
        p3.metric("Total Profit", f"₹ {total_profit:,.2f}")

    st.markdown("---")

    # =====================================================
    # ========= GROUP BY RATE + DATE (SAFE) ================
    # =====================================================
    summary_df = (
        df.groupby(["Rate_Int", "DateOnly"], as_index=False)
        .agg({
            "Proposal_Cost": "sum",
            "Final_Cost": "sum",
            "Profit": "sum"
        })
        .sort_values(["Rate_Int", "DateOnly"])
    )

# =====================================================
# ============ COMPACT SUMMARY DISPLAY ================
# =====================================================
if (
    st.session_state.page == "Summary"
    and "summary_df" in st.session_state
    and not st.session_state.summary_df.empty
):

    summary_df = st.session_state.summary_df
    
    for _, row in summary_df.iterrows():

        profit_color = "🟢" if row["Profit"] >= 0 else "🔴"

        date_str = (
            row["DateOnly"].strftime("%d-%m-%Y")
            if pd.notna(row["DateOnly"])
            else "—"
        )

        rate = int(row["Rate_Int"]) if "Rate_Int" in row else int(round(row["Rate"], 0))

        if is_mobile:
            st.markdown(
                f"""
**📅 {date_str} | {rate} %**  
💰 Invested : ₹ {row['Proposal_Cost']:,.2f}  
📈 Final : ₹ {row['Final_Cost']:,.2f}  
{profit_color} Profit : ₹ {row['Profit']:,.2f}
"""
            )
        else:
            st.markdown(
                f"""
**{date_str} | Rate {rate} %**  
Investment : ₹ {row['Proposal_Cost']:,.2f}  
Final Amount : ₹ {row['Final_Cost']:,.2f}  
Profit : ₹ {row['Profit']:,.2f}
"""
            )

        st.markdown("---")

# ===================================================== 
# ================= ADD NEW PROPOSAL ================== 
# =====================================================
# =====================================================
# ================= ADD NEW PROPOSAL ==================
# =====================================================

if st.session_state.page == "AddProposal":

    st.header("➕ Add New Proposal")

    if "clients_df" not in st.session_state or st.session_state.clients_df.empty:
        st.warning("Please add a client first.")
        st.stop()

    clients_df = st.session_state.clients_df
    proposals_df = st.session_state.proposals_df

    # ---------- Draft Storage ----------
    if "proposal_clients" not in st.session_state:
        st.session_state.proposal_clients = []

    # ---------- HEADER FIELDS ----------
    start_date = st.date_input(
        "Start Date",
        value=datetime.today().date()
    )

    months = st.selectbox(
        "Duration (Months)",
        [1, 2, 3]
    )

    end_date = (pd.to_datetime(start_date) + pd.DateOffset(months=months)).date()
    st.info(f"End Date: {end_date}")

    rate = st.number_input(
        "Monthly Rate (%)",
        min_value=0.0,
        step=0.10,
        format="%.2f"
    )

    st.markdown("---")
    st.subheader("Add Clients To Proposal")

    # ---------- CLIENT FILTER ----------
    active_clients = clients_df[clients_df["Is_Archived"] == False]

    client_options = sorted(
        active_clients["Client_Name"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    col1, col2 = st.columns(2)

    with col1:
        selected_client = st.selectbox(
            "Client Name",
            ["Select Client"] + client_options
        )

    with col2:
        principal = st.number_input(
            "Proposal Amount (₹)",
            min_value=0.0,
            step=1000.0,
            format="%.2f"
        )

    # ---------- ADD CLIENT ----------
    if st.button("➕ Add Client"):

        if selected_client == "Select Client" or principal <= 0:
            st.warning("Select valid client and enter valid amount.")
            st.stop()

        # Prevent duplicate
        existing_clients = [c["Client_Name"] for c in st.session_state.proposal_clients]
        if selected_client in existing_clients:
            st.warning("Client already added.")
            st.stop()

        # Finance Formula (Month Based)
        profit = principal * rate * months / 100
        final_amount = principal + profit

        st.session_state.proposal_clients.append({
            "Client_Name": selected_client,
            "Principal": principal,
            "Profit": profit,
            "Final_Amount": final_amount
        })

        st.success(f"{selected_client} added successfully.")

    # ---------- AUTO RECALCULATE IF RATE/MONTH CHANGES ----------
    for item in st.session_state.proposal_clients:
        p = item["Principal"]
        item["Profit"] = p * rate * months / 100
        item["Final_Amount"] = p + item["Profit"]

    # ---------- DISPLAY ADDED CLIENTS ----------
    if st.session_state.proposal_clients:

        st.markdown("### Added Clients")

        for i, item in enumerate(st.session_state.proposal_clients):

            col1, col2, col3, col4, col5 = st.columns([2,2,2,2,1])

            col1.write(item["Client_Name"])
            col2.write(f"₹ {item['Principal']:,.2f}")
            col3.write(f"₹ {item['Profit']:,.2f}")
            col4.write(f"₹ {item['Final_Amount']:,.2f}")

            if col5.button("❌", key=f"delete_{i}"):
                st.session_state.proposal_clients.pop(i)
                st.rerun()

        df_preview = pd.DataFrame(st.session_state.proposal_clients)

        total_principal = df_preview["Principal"].sum()
        total_profit = df_preview["Profit"].sum()
        total_final = df_preview["Final_Amount"].sum()

        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Principal", f"₹ {total_principal:,.2f}")
        c2.metric("Total Profit", f"₹ {total_profit:,.2f}")
        c3.metric("Total Final Amount", f"₹ {total_final:,.2f}")

    else:
        total_principal = 0
        total_profit = 0
        total_final = 0

    st.markdown("---")

    # ---------- SAVE COMPLETE PROPOSAL ----------
    if st.button("💾 Save Proposal"):

        if not st.session_state.proposal_clients:
            st.warning("Add at least one client.")
            st.stop()

        proposal_id = new_proposal_id()

        rows = []

        for item in st.session_state.proposal_clients:

            client_row = clients_df.loc[
                clients_df["Client_Name"] == item["Client_Name"]
            ].iloc[0]

            rows.append({
                "Proposal_ID": proposal_id,
                "Client_ID": client_row["Client_ID"],
                "Client_Name": item["Client_Name"],
                "Proposal_Cost": round(item["Principal"], 2),
                "Rate": round(rate, 2),
                "Profit": round(item["Profit"], 2),
                "Final_Cost": round(item["Final_Amount"], 2),
                "Start_Date": pd.to_datetime(start_date),
                "End_Date": pd.to_datetime(end_date),
                "Status": "Open",
                "Closing_Date": pd.NaT
            })

        st.session_state.proposals_df = pd.concat(
            [proposals_df, pd.DataFrame(rows)],
            ignore_index=True
        )

        save_proposals()

        # Reset Draft
        st.session_state.proposal_clients = []

        st.success(f"✅ Proposal {proposal_id} created successfully.")
        st.session_state.page = "Summary"

# =====================================================
# ================= EDIT PROPOSAL =====================
# =====================================================
if st.session_state.page == "Edit":

    st.header("✏️ Edit Proposal")

    if proposals_df.empty:
        st.warning("No proposals available")
        st.stop()

    df = proposals_df.copy()

    # =================================================
    # STEP 1: STATUS FILTER (WITH SELECT)
    # =================================================
    status_filter = st.selectbox(
        "Select Proposal Status",
        ["Select", "Open", "Closed"]
    )

    if status_filter == "Select":
        st.info("Please select Proposal Status")
        st.stop()

    df = df[df["Status"] == status_filter]

    if df.empty:
        st.info(f"No {status_filter} proposals found")
        st.stop()

    # =================================================
    # STEP 2: PROPOSAL ID FILTER (WITH SELECT)
    # =================================================
    proposal_id_list = ["Select Proposal ID"] + sorted(df["Proposal_ID"].unique())

    proposal_id = st.selectbox(
        "Select Proposal ID",
        proposal_id_list
    )

    if proposal_id == "Select Proposal ID":
        st.info("Please select a Proposal ID")
        st.stop()

    proposal_df = df[df["Proposal_ID"] == proposal_id]

    # =================================================
    # STEP 3: READ-ONLY MASTER DATA
    # =================================================
    start_date = proposal_df["Start_Date"].iloc[0]
    end_date = proposal_df["End_Date"].iloc[0]
    rate_master = proposal_df["Rate"].iloc[0]

    c1, c2, c3 = st.columns(3)
    c1.text_input("Start Date", start_date.strftime("%d-%m-%Y"), disabled=True)
    c2.text_input("End Date", end_date.strftime("%d-%m-%Y"), disabled=True)
    c3.text_input("Rate (%)", int(round(rate_master, 0)), disabled=True)

    # =================================================
    # STEP 4: CLIENT SELECTION (WITH SELECT)
    # =================================================
    client_list = ["Select"] + sorted(proposal_df["Client_Name"].unique())

    client_name = st.selectbox(
        "Select Client Included in Proposal",
        client_list
    )

    if client_name == "Select":
        st.info("Please select a Client")
        st.stop()

    row = proposal_df[
        proposal_df["Client_Name"] == client_name
    ].iloc[0]

    st.markdown("---")

    # =================================================
    # STEP 5: CURRENT DETAILS
    # =================================================
    if is_mobile:
        st.markdown(
            f"""
            <div style="border:1px solid #ddd;
            border-radius:12px;
            padding:14px;
            background:#fafafa">

            <b>Client:</b> {row['Client_Name']}<br>
            <b>Status:</b> {row['Status']}<br>
            <b>Start Date:</b> {row['Start_Date'].date()}<br>
            <b>End Date:</b> {row['End_Date'].date()}<br>
            <b>Proposal Amount:</b> ₹ {row['Proposal_Cost']:,.2f}<br>
            <b>Rate:</b> {int(round(row['Rate'],0))} %<br>
            <b>Final Amount:</b> ₹ {row['Final_Cost']:,.2f}<br>
            <b>Profit:</b> ₹ {row['Profit']:,.2f}

            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        st.dataframe(
            pd.DataFrame([row])[[ 
                "Client_Name",
                "Proposal_ID",
                "Start_Date",
                "End_Date",
                "Proposal_Cost",
                "Rate",
                "Final_Cost",
                "Profit",
                "Status"
            ]],
            use_container_width=True
        )

    # =================================================
    # STEP 6: EDIT MODE
    # =================================================
    st.markdown("---")
    edit_mode = st.checkbox("✏️ Edit Proposal")

    if edit_mode:

        st.warning("Editing will recalculate Final Amount & Profit")

        if is_mobile:
            proposal_cost = st.number_input(
                "Proposal Amount (₹)",
                value=float(row["Proposal_Cost"]),
                step=1000.0
            )
            rate = st.number_input(
                "Rate (%)",
                value=float(row["Rate"]),
                step=0.5
            )
            new_start = st.date_input(
                "Start Date",
                value=row["Start_Date"].date()
            )
            new_end = st.date_input(
                "End Date",
                value=row["End_Date"].date()
            )
            new_status = st.selectbox(
                "Status",
                ["Open", "Closed"],
                index=0 if row["Status"] == "Open" else 1
            )
        else:
            c1, c2, c3, c4, c5 = st.columns(5)
            proposal_cost = c1.number_input(
                "Proposal Amount (₹)",
                value=float(row["Proposal_Cost"]),
                step=1000.0
            )
            rate = c2.number_input(
                "Rate (%)",
                value=float(row["Rate"]),
                step=0.5
            )
            new_start = c3.date_input(
                "Start Date",
                value=row["Start_Date"].date()
            )
            new_end = c4.date_input(
                "End Date",
                value=row["End_Date"].date()
            )
            new_status = c5.selectbox(
                "Status",
                ["Open", "Closed"],
                index=0 if row["Status"] == "Open" else 1
            )

        # =================================================
        # APPLY TO ALL CLIENTS OPTION
        # =================================================
        apply_all = False
        if len(client_list) > 2:
            apply_all = st.checkbox(
                "Apply changes to all clients in this proposal",
                help="Applies Start Date, End Date, Rate and Status to all clients"
            )

        # =================================================
        # STEP 7: SAFE AUTO CALCULATION
        # =================================================
        days = (pd.to_datetime(new_end) - pd.to_datetime(new_start)).days

        if days < 0:
            st.error("End Date must be after Start Date")
            st.stop()

        final_cost, profit = calc(proposal_cost, rate, days)
        months = days / 30

        st.markdown("---")
        st.subheader("💰 Auto Calculated")

        if is_mobile:
            st.markdown(
                f"""
                <div style="border:1px solid #ddd;
                border-radius:12px;
                padding:14px;
                background:#fafafa">

                <b>Duration:</b> {days} days ({months:.2f} months)<br>
                <b>Profit:</b> ₹ {profit:,.2f}<br>
                <b>Final Amount:</b> ₹ {final_cost:,.2f}

                </div>
                """,
                unsafe_allow_html=True
            )
        else:
            d1, d2, d3 = st.columns(3)
            d1.metric("Duration (Months)", f"{months:.2f}")
            d2.metric("Profit (₹)", f"{profit:,.2f}")
            d3.metric("Final Amount (₹)", f"{final_cost:,.2f}")

        # =================================================
        # STEP 8: SAVE
        # =================================================
        if st.button("💾 Save Changes", use_container_width=True):

            if apply_all:
                mask = (
                    st.session_state.proposals_df["Proposal_ID"] == proposal_id
                )
            else:
                mask = (
                    (st.session_state.proposals_df["Proposal_ID"] == proposal_id) &
                    (st.session_state.proposals_df["Client_Name"] == client_name)
                )

            st.session_state.proposals_df.loc[mask, [
                "Rate",
                "Start_Date",
                "End_Date",
                "Status",
                "Closing_Date"
            ]] = [
                round(rate, 2),
                pd.to_datetime(new_start),
                pd.to_datetime(new_end),
                new_status,
                pd.Timestamp.today() if new_status == "Closed" else pd.NaT
            ]

            st.session_state.proposals_df.loc[
                (st.session_state.proposals_df["Proposal_ID"] == proposal_id) &
                (st.session_state.proposals_df["Client_Name"] == client_name),
                ["Proposal_Cost", "Final_Cost", "Profit"]
            ] = [
                round(proposal_cost, 2),
                round(final_cost, 2),
                round(profit, 2)
            ]

            save_proposals()

            st.success("✅ Proposal updated successfully")
            st.session_state.page = "Summary"

        # =================================================
        # DELETE OPTIONS
        # =================================================
        st.markdown("---")
        st.subheader("⚠️ Delete Options")

        col_del1, col_del2 = st.columns(2)

        # -----------------------------
        # DELETE SELECTED CLIENT
        # -----------------------------
        with col_del1:
            if st.button("🗑 Delete This Client", use_container_width=True):

                mask = (
                    (st.session_state.proposals_df["Proposal_ID"] == proposal_id) &
                    (st.session_state.proposals_df["Client_Name"] == client_name)
                )

                st.session_state.proposals_df = (
                    st.session_state.proposals_df.loc[~mask]
                )

                save_proposals()

                st.success(f"Client '{client_name}' deleted from proposal.")
                st.session_state.page = "Summary"
                st.rerun()


        # -----------------------------
        # DELETE ENTIRE PROPOSAL
        # -----------------------------
        with col_del2:
            if st.button("❌ Delete Entire Proposal", use_container_width=True):

                mask = (
                    st.session_state.proposals_df["Proposal_ID"] == proposal_id
                )

                st.session_state.proposals_df = (
                    st.session_state.proposals_df.loc[~mask]
                )

                save_proposals()

                st.success(f"Proposal '{proposal_id}' deleted completely.")
                st.session_state.page = "Summary"
                st.rerun()

# =====================================================
# ================= FIND DETAILS PAGE =================
# =====================================================

# ================= IMPORTS =================
import pandas as pd
import streamlit as st

# ================= SAFETY INITIALIZATION =================
if "page" not in st.session_state:
    st.session_state.page = "Find"

is_mobile = st.session_state.get("is_mobile", False)

# ================= UTILITIES =================
def get_master_df(proposals_df):
    df = proposals_df.copy()
    for c in ["Start_Date", "End_Date", "Closing_Date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

# ================= COMPONENT : GRAND TOTAL =================
def render_grand_total(total_invest, total_final, total_profit, is_mobile):
    if is_mobile:
        st.markdown(
            f"""
            <div style="
                background:#e3f2fd;
                border-radius:14px;
                padding:14px;
                margin:15px 0;
                border:1px solid #90caf9">
                <b>💠 Grand Total</b><br><br>
                <b>Investment:</b> ₹ {total_invest:,.2f}<br>
                <b>Final Amount:</b> ₹ {total_final:,.2f}<br>
                <b>Profit:</b> ₹ {total_profit:,.2f}
            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Investment", f"₹ {total_invest:,.2f}")
        c2.metric("Final Amount", f"₹ {total_final:,.2f}")
        c3.metric("Profit", f"₹ {total_profit:,.2f}")

# ================= MODE SELECTOR =================
def render_find_mode_selector():
    return st.radio(
        "Find Details Mode",
        ["By Proposal", "By Client Name", "By Start / End Date"],
        horizontal=True
    )

# ================= BY PROPOSAL =================
def render_by_proposal(df_master, is_mobile):

    st.subheader("📄 Find Details By Proposal")

    status = st.selectbox("Status", ["All", "Open", "Closed"])
    df = df_master if status == "All" else df_master[df_master["Status"] == status]

    if df.empty:
        st.info("No proposals found")
        return

    proposal_options = ["-- Select Proposal --"] + sorted(df["Proposal_ID"].unique())
    proposal_id = st.selectbox("Select Proposal ID", proposal_options)

    if proposal_id == "-- Select Proposal --":
        st.info("Please select a Proposal ID")
        return

    proposal_df = df[df["Proposal_ID"] == proposal_id]

    c1, c2, c3 = st.columns(3)
    c1.text_input("Start Date", proposal_df["Start_Date"].iloc[0].strftime("%d-%m-%Y"), disabled=True)
    c2.text_input("End Date", proposal_df["End_Date"].iloc[0].strftime("%d-%m-%Y"), disabled=True)
    c3.text_input("Rate (%)", int(round(proposal_df["Rate"].iloc[0])), disabled=True)

    client = st.selectbox(
        "Select Client",
        ["All"] + sorted(proposal_df["Client_Name"].unique())
    )

    result = proposal_df if client == "All" else proposal_df[proposal_df["Client_Name"] == client]

    if result.empty:
        st.info("No data found")
        return

    render_grand_total(
        result["Proposal_Cost"].sum(),
        result["Final_Cost"].sum(),
        result["Profit"].sum(),
        is_mobile
    )

    st.dataframe(
        result[["Client_Name", "Proposal_Cost", "Final_Cost", "Profit"]].round(2),
        use_container_width=True,
        hide_index=True
    )

# ================= BY CLIENT NAME =================
def render_by_client(df_master, is_mobile):

    st.subheader("🔎 Find Details Using Client Name")

    # ---------- Status Filter ----------
    status = st.selectbox("Status", ["All", "Open", "Closed"])
    df = df_master if status == "All" else df_master[df_master["Status"] == status]

    if df.empty:
        st.info("No records found")
        return

    # ---------- Client Name Selector ----------
    client_list = sorted(df["Client_Name"].dropna().unique().tolist())
    client_list.insert(0, "Select Client Name")

    client = st.selectbox("Client Name", client_list)

    if client == "Select Client Name":
        st.info("Please select a client name")
        return

    df = df[df["Client_Name"] == client]

    # ---------- Date Type ----------
    date_type = st.radio("Date Type", ["Start Date", "End Date"], horizontal=True)
    date_col = "Start_Date" if date_type == "Start Date" else "End_Date"

    df = df.dropna(subset=[date_col])
    df["DateOnly"] = df[date_col].dt.date

    if df.empty:
        st.info("No dates available")
        return

    # ---------- Date Selector ----------
    date_list = sorted(df["DateOnly"].unique().tolist())
    date_list.insert(0, "Select Date")

    selected_date = st.selectbox(
        "Select Date",
        date_list,
        format_func=lambda x: x if isinstance(x, str) else x.strftime("%d-%m-%Y")
    )

    if selected_date == "Select Date":
        st.info("Please select a date")
        return

    result = df[df["DateOnly"] == selected_date]

    # ---------- Grand Total ----------
    render_grand_total(
        result["Proposal_Cost"].sum(),
        result["Final_Cost"].sum(),
        result["Profit"].sum(),
        is_mobile
    )

    # ---------- Table ----------
    st.dataframe(
        result[["Proposal_Cost", "Rate", "Final_Cost", "Profit"]].round(2),
        use_container_width=True,
        hide_index=True
    )
    
# ================= BY DATE =================
def render_by_date(df_master, is_mobile):

    st.subheader("📅 Find Details Using Start / End Date")

    # -----------------------------
    # STATUS
    # -----------------------------
    status = st.selectbox(
        "Status",
        ["All", "Open", "Closed"],
        key="date_status"
    )

    df = df_master if status == "All" else df_master[df_master["Status"] == status]

    if df.empty:
        st.info("No records found")
        return

    # -----------------------------
    # INIT BLOCKS (ONCE)
    # -----------------------------
    if "date_blocks" not in st.session_state:
        st.session_state.date_blocks = [1]

    if "remove_block" not in st.session_state:
        st.session_state.remove_block = None

    # -----------------------------
    # ADD BLOCK
    # -----------------------------
    if st.button("➕ Check Another Date"):
        st.session_state.date_blocks.append(
            max(st.session_state.date_blocks) + 1
        )

    st.divider()

    # -----------------------------
    # RENDER BLOCKS
    # -----------------------------
    for i, block_id in enumerate(list(st.session_state.date_blocks)):

        col1, col2 = st.columns([6, 1])

        with col1:
            st.markdown(f"### 🔍 Date Filter {block_id}")

        with col2:
            if i != 0:
                if st.button("❌", key=f"remove_{block_id}"):
                    st.session_state.remove_block = block_id

        # -------------------------
        # DATE TYPE (NO "Select")
        # -------------------------
        date_type = st.radio(
            "Date Type",
            ["Start Date", "End Date"],
            horizontal=True,
            key=f"date_type_{block_id}"
        )

        date_col = "Start_Date" if date_type == "Start Date" else "End_Date"

        df_local = df.dropna(subset=[date_col]).copy()
        df_local["DateOnly"] = df_local[date_col].dt.date

        if df_local.empty:
            st.info("No dates available")
            st.divider()
            continue

        # -------------------------
        # DATE SELECT (SAFE)
        # -------------------------
        date_options = [None] + sorted(df_local["DateOnly"].unique())

        selected_date = st.selectbox(
            f"Select {date_type}",
            date_options,
            key=f"date_select_{block_id}",
            format_func=lambda x: "Select" if x is None else x.strftime("%d-%m-%Y")
        )

        if selected_date is None:
            st.divider()
            continue

        # -------------------------
        # RESULT
        # -------------------------
        result = df_local[df_local["DateOnly"] == selected_date]

        if result.empty:
            st.info("No records found")
            st.divider()
            continue

        # -------------------------
        # GRAND TOTAL
        # -------------------------
        render_grand_total(
            result["Proposal_Cost"].sum(),
            result["Final_Cost"].sum(),
            result["Profit"].sum(),
            is_mobile
        )

        st.dataframe(
            result[
                ["Client_Name", "Proposal_Cost", "Rate", "Final_Cost", "Profit"]
            ].round(2),
            use_container_width=True,
            hide_index=True
        )

        st.divider()

    # -----------------------------
    # REMOVE BLOCK (AFTER RENDER)
    # -----------------------------
    if st.session_state.remove_block:
        if st.session_state.remove_block in st.session_state.date_blocks:
            st.session_state.date_blocks.remove(
                st.session_state.remove_block
            )
        st.session_state.remove_block = None

# ================= PAGE CONTROLLER =================
if st.session_state.page == "Find":

    st.header("🔍 Find Proposal Details")

    if "proposals_df" not in globals() or proposals_df.empty:
        st.warning("No proposal data available")
        st.stop()

    df_master = get_master_df(proposals_df)

    find_mode = render_find_mode_selector()

    if find_mode == "By Proposal":
        render_by_proposal(df_master, is_mobile)
    elif find_mode == "By Client Name":
        render_by_client(df_master, is_mobile)
    elif find_mode == "By Start / End Date":
        render_by_date(df_master, is_mobile)


# =====================================================
# ================= CLIENTS ===========================
# =====================================================
if st.session_state.page == "Clients":

    st.header("👤 Clients Management")

    clients_df = st.session_state.clients_df
    proposals_df = st.session_state.proposals_df

    # -------------------------------------------------
    # ADD NEW CLIENT
    # -------------------------------------------------
    with st.expander("➕ Add New Client"):
        cname = st.text_input("Client Name", placeholder="Enter client name")

        if st.button("Add Client", use_container_width=True):
            cname = cname.strip()

            if not cname:
                st.error("Client name cannot be empty")
                st.stop()

            if cname.lower() in clients_df["Client_Name"].str.lower().tolist():
                st.error("Client already exists")
                st.stop()

            new_row = {
                "Client_ID": new_client_id(),
                "Client_Name": cname,
                "Created_Date": pd.Timestamp.now(),
                "Is_Archived": False,
                "Notes": ""
            }

            st.session_state.clients_df = pd.concat(
                [clients_df, pd.DataFrame([new_row])],
                ignore_index=True
            )

            save_clients()
            st.success("✅ Client added successfully")

    # -------------------------------------------------
    # CLIENT SEARCH (ACTIVE ONLY)
    # -------------------------------------------------
    search = st.text_input("🔍 Search Client", placeholder="Type client name")

    active_clients = clients_df[clients_df["Is_Archived"] == False].copy()

    if search.strip():
        active_clients = active_clients[
            active_clients["Client_Name"]
            .str.contains(search, case=False, na=False)
        ]

    # -------------------------------------------------
    # CLIENT LIST WITH ACTIONS
    # -------------------------------------------------
    with st.expander("📋 Client List", expanded=True):

        if active_clients.empty:
            st.info("No active clients found")
            st.stop()

        for _, client in active_clients.iterrows():

            client_id = client["Client_ID"]
            client_name = client["Client_Name"]

            proposal_count = len(
                proposals_df[proposals_df["Client_ID"] == client_id]
            )

            st.markdown(
                f"""
                <div style="border:1px solid #ddd;border-radius:12px;
                padding:12px;margin-bottom:8px;background:#fafafa">
                <b>{client_name}</b><br>
                <span style="color:#555">Client ID:</span> {client_id}<br>
                <span style="color:#555">Proposals:</span> {proposal_count}
                </div>
                """,
                unsafe_allow_html=True
            )

            col1, col2, col3 = st.columns([4, 2, 2])

            # -------- RENAME CLIENT (SAFE) --------
            with col1:
                new_name = st.text_input(
                    "Rename Client",
                    value=client_name,
                    key=f"rename_{client_id}"
                )

                if st.button("💾 Save Name", key=f"save_{client_id}"):
                    new_name = new_name.strip()

                    if not new_name:
                        st.error("Name cannot be empty")
                        st.stop()

                    st.session_state.clients_df.loc[
                        st.session_state.clients_df["Client_ID"] == client_id,
                        "Client_Name"
                    ] = new_name

                    st.session_state.proposals_df.loc[
                        st.session_state.proposals_df["Client_ID"] == client_id,
                        "Client_Name"
                    ] = new_name

                    save_clients()
                    save_proposals()
                    st.success("Client name updated")

            # -------- VIEW PROPOSALS --------
            with col2:
                if st.button("📄 View Proposals", key=f"view_{client_id}"):
                    st.session_state.find_client = client_name
                    st.session_state.page = "Find"

            # -------- ARCHIVE CLIENT --------
            with col3:
                if st.button("📦 Archive", key=f"arc_{client_id}"):
                    st.session_state.clients_df.loc[
                        st.session_state.clients_df["Client_ID"] == client_id,
                        "Is_Archived"
                    ] = True

                    save_clients()
                    st.success("Client archived (data preserved)")

# =====================================================
# ================= CLIENT DASHBOARD ==================
# =====================================================
if st.session_state.page == "ClientDashboard":

    st.header("📊 Client Summary Dashboard")

    # ---------- ACTIVE CLIENTS ONLY ----------
    active_clients = st.session_state.clients_df[
        st.session_state.clients_df["Is_Archived"] == False
    ].copy()

    if active_clients.empty:
        st.info("No active clients")
        st.stop()

    # =================================================
    # ============== CLIENT FILTER ====================
    # =================================================
    client_options = ["Select"] + sorted(active_clients["Client_Name"].unique())

    client = st.selectbox("Select Client", client_options)

    if client == "Select":
        st.info("Please select a client to continue")
        st.stop()

    client_row = active_clients[
        active_clients["Client_Name"] == client
    ].iloc[0]

    client_id = client_row["Client_ID"]

    # ---------- WORKING COPY ----------
    client_df = st.session_state.proposals_df[
        st.session_state.proposals_df["Client_ID"] == client_id
    ].copy()

    if client_df.empty:
        st.info("No proposals available for this client")
        st.stop()

    # ---------- DATE NORMALIZATION ----------
    for c in ["Start_Date", "End_Date"]:
        client_df[c] = pd.to_datetime(client_df[c], errors="coerce")

    # =================================================
    # ============== STATUS FILTER ====================
    # =================================================
    status = st.selectbox(
        "Select Status",
        ["Select", "All", "Open", "Closed"]
    )

    if status == "Select":
        st.info("Please select status to view data")
        st.stop()

    if status != "All":
        client_df = client_df[client_df["Status"] == status]

    if client_df.empty:
        st.info("No records found for selected filters")
        st.stop()

    # =================================================
    # ============== GRAND TOTALS =====================
    # =================================================
    total_invest = client_df["Proposal_Cost"].sum()
    total_final = client_df["Final_Cost"].sum()
    total_profit = client_df["Profit"].sum()
    open_props = (client_df["Status"] == "Open").sum()
    closed_props = (client_df["Status"] == "Closed").sum()

    if is_mobile:
        card("Investment", f"₹ {total_invest:,.2f}")
        card("Final Amount", f"₹ {total_final:,.2f}")
        card("Profit", f"₹ {total_profit:,.2f}")
        card("Open", open_props)
        card("Closed", closed_props)
    else:
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Investment", f"₹ {total_invest:,.2f}")
        c2.metric("Total Final Amount", f"₹ {total_final:,.2f}")
        c3.metric("Total Profit", f"₹ {total_profit:,.2f}")
        c4.metric("Open Proposals", open_props)
        c5.metric("Closed Proposals", closed_props)

    st.markdown("---")

    # =================================================
    # ============== DISPLAY DATA =====================
    # =================================================
    display_df = client_df.sort_values("Start_Date").copy()

    display_df["Start_Date"] = display_df["Start_Date"].dt.strftime("%d/%m/%Y")
    display_df["End_Date"] = display_df["End_Date"].dt.strftime("%d/%m/%Y")
    display_df["Rate"] = (
        pd.to_numeric(display_df["Rate"], errors="coerce")
        .fillna(0)
        .round(0)
        .astype(int)
    )

    # ---------------- MOBILE VIEW -------------------
    if is_mobile:
        for _, r in display_df.iterrows():
            st.markdown(
                f"""
                <div style="border:1px solid #ddd;
                border-radius:12px;
                padding:14px;
                margin-bottom:12px;
                background:#fafafa">

                <b>Start Date:</b> {r['Start_Date']}<br>
                <b>End Date:</b> {r['End_Date']}<br>
                <b>Proposal Cost:</b> ₹ {r['Proposal_Cost']:,.2f}<br>
                <b>Final Cost:</b> ₹ {r['Final_Cost']:,.2f}<br>
                <b>Profit:</b> ₹ {r['Profit']:,.2f}<br>
                <b>Status:</b> {r['Status']}

                </div>
                """,
                unsafe_allow_html=True
            )

    # ---------------- DESKTOP VIEW ------------------
    else:
        st.dataframe(
            display_df[[
                "Proposal_ID",
                "Start_Date",
                "End_Date",
                "Proposal_Cost",
                "Final_Cost",
                "Profit",
                "Status"
            ]],
            use_container_width=True
        )

# =====================================================
# ================= EXPORT DATA =======================
# =====================================================
if st.session_state.page == "Export":

    st.header("📤 Export Data")

    proposals_df = st.session_state.proposals_df
    clients_df = st.session_state.clients_df

    # ---------- SAFETY CHECK ----------
    if proposals_df.empty and clients_df.empty:
        st.warning("No data available to export.")
        st.stop()

    # ---------- EXPORT FUNCTION ----------
    def export_excel():
        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            if not clients_df.empty:
                clients_df.to_excel(
                    writer,
                    index=False,
                    sheet_name="Clients"
                )

            if not proposals_df.empty:
                proposals_df.to_excel(
                    writer,
                    index=False,
                    sheet_name="Proposals"
                )

        output.seek(0)
        return output

    # ---------- GENERATE FILE ----------
    excel_file = export_excel()

    # ---------- DOWNLOAD BUTTON ----------
    st.download_button(
        label="⬇️ Download Excel (Clients + Proposals)",
        data=excel_file,
        file_name="sigma_consultants_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ---------- OPTIONAL CSV BACKUP ----------
    st.markdown("### 🔁 Alternative Format")

    if not proposals_df.empty:
        st.download_button(
            label="⬇️ Download Proposals CSV",
            data=proposals_df.to_csv(index=False),
            file_name="sigma_proposals.csv",
            mime="text/csv"
        )

    if not clients_df.empty:
        st.download_button(
            label="⬇️ Download Clients CSV",
            data=clients_df.to_csv(index=False),
            file_name="sigma_clients.csv",
            mime="text/csv"
        )










































