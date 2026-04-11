# =====================================================
# Sigma Consultants – CRM
# =====================================================

import streamlit as st
import pandas as pd
from datetime import datetime
import os
import io


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

# =====================================================
# LOAD DATA
# =====================================================

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

    if "Is_Archived" not in df.columns:
        df["Is_Archived"] = False

    return df

def load_proposals():

    required_cols = [
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
    ]

    # ---------- LOAD OR CREATE FILE ----------
    if os.path.exists(PROPOSAL_FILE):
        df = pd.read_excel(PROPOSAL_FILE)
    else:
        df = pd.DataFrame(columns=required_cols)
        df.to_excel(PROPOSAL_FILE, index=False)

    # ---------- ADD MISSING COLUMNS ----------
    for col in required_cols:
        if col not in df.columns:
            df[col] = None

    # ---------- DATE SAFETY ----------
    date_cols = ["Start_Date", "End_Date", "Closing_Date"]
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    # ---------- STATUS SAFETY ----------
    df["Status"] = df["Status"].fillna("Open")

    return df
    
# =====================================================
# SESSION STATE
# =====================================================

if "clients_df" not in st.session_state:
    st.session_state.clients_df = load_clients()

if "proposals_df" not in st.session_state:
    st.session_state.proposals_df = load_proposals()

clients_df = st.session_state.clients_df
proposals_df = st.session_state.proposals_df

# =====================================================
# SAVE DATA
# =====================================================

def save_clients():

    st.session_state.clients_df.to_excel(
        CLIENT_FILE,
        index=False
    )


def save_proposals():

    st.session_state.proposals_df.to_excel(
        PROPOSAL_FILE,
        index=False
    )


def new_client_id():

    if clients_df.empty:
        return "SIG-C-001"

    last = (
        clients_df["Client_ID"]
        .str.replace("SIG-C-","",regex=False)
        .astype(int)
        .max()
    )

    return f"SIG-C-{last+1:03d}"


def new_proposal_id():

    if proposals_df.empty:
        return "SIG-P-001"

    last = (
        proposals_df["Proposal_ID"]
        .str.replace("SIG-P-","",regex=False)
        .astype(int)
        .max()
    )

    return f"SIG-P-{last+1:03d}"

# =====================================================
# CALCULATION FUNCTION (ADD HERE)
# =====================================================

def calc(principal, rate, days):
    profit = principal * rate * (days / 30) / 100
    final = principal + profit
    return final, profit

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

    if proposals_df.empty:
        st.info("No proposals available")
        st.stop()

    df = proposals_df.copy()

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

    st.markdown("---")
    st.subheader("📅 Date Based Summary")

    for c in ["Start_Date", "End_Date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    status_options = ["All"] + sorted(
        df["Status"].dropna().unique().tolist()
    )

    selected_status = st.selectbox(
        "Select Proposal Status",
        status_options
    )

    if selected_status != "All":
        df = df[df["Status"] == selected_status]

    if df.empty:
        st.warning("No data for selected status")
        st.stop()

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

    selected_dates = st.multiselect(
        f"Select {date_type}(s)",
        ["All"] + sorted(df["DateOnly"].unique()),
        default=["All"],
        format_func=lambda x:
            x if x == "All"
            else x.strftime("%d-%m-%Y")
    )

    if "All" not in selected_dates:
        df = df[df["DateOnly"].isin(selected_dates)]

    if df.empty:
        st.warning("No data for selected date(s)")
        st.stop()

    df["Rate_Int"] = (
        pd.to_numeric(df["Rate"], errors="coerce")
        .fillna(0)
        .round(0)
        .astype(int)
    )

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

    st.session_state.summary_df = (
        df.groupby(["Rate_Int", "DateOnly"], as_index=False)
        .agg({
            "Proposal_Cost": "sum",
            "Final_Cost": "sum",
            "Profit": "sum"
        })
        .sort_values(["Rate_Int", "DateOnly"])
    )

    summary_df = st.session_state.summary_df

    if summary_df.empty:
        st.stop()

    for _, row in summary_df.iterrows():

        profit_color = "🟢" if row["Profit"] >= 0 else "🔴"

        date_str = row["DateOnly"].strftime("%d-%m-%Y")

        rate = int(row["Rate_Int"])

        st.markdown(
            f"""
**{date_str} | Rate {rate}%**  
Investment : ₹ {row['Proposal_Cost']:,.2f}  
Final Amount : ₹ {row['Final_Cost']:,.2f}  
{profit_color} Profit : ₹ {row['Profit']:,.2f}
"""
        )

        st.markdown("---")

# =====================================================
# ================= ADD NEW PROPOSAL ==================
# =====================================================

if st.session_state.page == "AddProposal":

    st.header("➕ Add New Proposal")

    if st.session_state.clients_df.empty:
        st.warning("Please add a client first.")
        st.stop()

    if "proposal_clients" not in st.session_state:
        st.session_state.proposal_clients = []

    start_date = st.date_input("Start Date")

    days = st.selectbox(
        "Duration (Days)",
        [15, 20, 30, 45, 60, 90]
    )

    end_date = (
        pd.to_datetime(start_date)
        + pd.Timedelta(days=days)
    ).date()

    rate = st.number_input("Monthly Rate (%)", min_value=0.0)

    active_clients = st.session_state.clients_df[
        st.session_state.clients_df["Is_Archived"] == False
    ]

    selected_client = st.selectbox(
        "Client Name",
        ["Select Client"] + sorted(active_clients["Client_Name"])
    )

    principal = st.number_input("Proposal Amount (₹)", min_value=0.0)

    if st.button("➕ Add Client"):

        if selected_client == "Select Client":
            st.warning("Select client")
            st.stop()

        final_amount, profit = calc(principal, rate, days)

        st.session_state.proposal_clients.append({
            "Client_Name": selected_client,
            "Principal": principal,
            "Profit": profit,
            "Final_Amount": final_amount
        })

    if st.session_state.proposal_clients:

        df_preview = pd.DataFrame(st.session_state.proposal_clients)

        st.dataframe(df_preview)

    if st.button("💾 Save Proposal"):

        proposal_id = new_proposal_id()

        rows = []

        for item in st.session_state.proposal_clients:

            client_row = active_clients[
                active_clients["Client_Name"]
                == item["Client_Name"]
            ].iloc[0]

            rows.append({
                "Proposal_ID": proposal_id,
                "Client_ID": client_row["Client_ID"],
                "Client_Name": item["Client_Name"],
                "Proposal_Cost": item["Principal"],
                "Rate": rate,
                "Profit": item["Profit"],
                "Final_Cost": item["Final_Amount"],
                "Start_Date": start_date,
                "End_Date": end_date,
                "Status": "Open",
                "Closing_Date": pd.NaT
            })

        st.session_state.proposals_df = pd.concat(
            [
                st.session_state.proposals_df,
                pd.DataFrame(rows)
            ],
            ignore_index=True
        )

        save_proposals()

        st.session_state.proposal_clients = []

        st.success("Proposal created")

        st.session_state.page = "Summary"
        st.rerun()


# =====================================================
# ================= EDIT PROPOSAL =====================
# =====================================================

if st.session_state.page == "Edit":

    st.header("✏️ Edit Proposal")

    if proposals_df.empty:
        st.stop()

    status_filter = st.selectbox(
        "Select Proposal Status",
        ["Select", "Open", "Closed"]
    )

    if status_filter == "Select":
        st.stop()

    df = proposals_df[
        proposals_df["Status"] == status_filter
    ]

    proposal_id = st.selectbox(
        "Select Proposal ID",
        ["Select"] + sorted(df["Proposal_ID"].unique())
    )

    if proposal_id == "Select":
        st.stop()

    proposal_df = df[df["Proposal_ID"] == proposal_id]

    client_name = st.selectbox(
        "Select Client",
        proposal_df["Client_Name"].unique()
    )

    row = proposal_df[
        proposal_df["Client_Name"] == client_name
    ].iloc[0]

    proposal_cost = st.number_input(
        "Proposal Amount",
        value=float(row["Proposal_Cost"])
    )

    rate = st.number_input(
        "Rate (%)",
        value=float(row["Rate"])
    )

    new_start = st.date_input(
        "Start Date",
        value=row["Start_Date"].date()
    )

    new_end = st.date_input(
        "End Date",
        value=row["End_Date"].date()
    )

    days = (new_end - new_start).days

    final_cost, profit = calc(
        proposal_cost,
        rate,
        days
    )

    if st.button("💾 Save Changes"):

        mask = (
            (st.session_state.proposals_df["Proposal_ID"] == proposal_id)
            &
            (st.session_state.proposals_df["Client_Name"] == client_name)
        )

        st.session_state.proposals_df.loc[
            mask,
            ["Proposal_Cost", "Rate",
             "Start_Date", "End_Date",
             "Final_Cost", "Profit"]
        ] = [
            proposal_cost,
            rate,
            new_start,
            new_end,
            final_cost,
            profit
        ]

        save_proposals()

        st.success("Updated")

        st.session_state.page = "Summary"
        st.rerun()


# =====================================================
# ================= FIND DETAILS PAGE =================
# =====================================================

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

    cname = st.text_input("Client Name")

    if st.button("Add Client"):

        if cname.strip() == "":
            st.stop()

        st.session_state.clients_df = pd.concat(
            [
                st.session_state.clients_df,
                pd.DataFrame([{
                    "Client_ID": new_client_id(),
                    "Client_Name": cname.strip(),
                    "Created_Date": pd.Timestamp.now(),
                    "Is_Archived": False,
                    "Notes": ""
                }])
            ],
            ignore_index=True
        )

        save_clients()

        st.success("Client added")

        st.rerun()

    st.dataframe(st.session_state.clients_df)

# =====================================================
# ================= CLIENT DASHBOARD ==================
# =====================================================

if st.session_state.page == "ClientDashboard":

    st.header("📊 Client Summary Dashboard")

    # ---------- SAFETY CHECK ----------
    if "clients_df" not in st.session_state or st.session_state.clients_df.empty:
        st.info("No clients available")
        st.stop()

    if "proposals_df" not in st.session_state:
        st.info("No proposal data available")
        st.stop()

    clients_df = st.session_state.clients_df.copy()
    proposals_df = st.session_state.proposals_df.copy()

    # ---------- ACTIVE CLIENTS ONLY ----------
    active_clients = clients_df[
        clients_df["Is_Archived"] == False
    ].copy()

    if active_clients.empty:
        st.info("No active clients")
        st.stop()

    # =================================================
    # CLIENT FILTER
    # =================================================
    client_options = ["Select"] + sorted(
        active_clients["Client_Name"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    client = st.selectbox("Select Client", client_options)

    if client == "Select":
        st.info("Please select a client to continue")
        st.stop()

    client_match = active_clients[
        active_clients["Client_Name"] == client
    ]

    if client_match.empty:
        st.error("Client not found")
        st.stop()

    client_id = client_match.iloc[0]["Client_ID"]

    # ---------- FILTER CLIENT PROPOSALS ----------
    client_df = proposals_df[
        proposals_df["Client_ID"] == client_id
    ].copy()

    if client_df.empty:
        st.info("No proposals available for this client")
        st.stop()

    # ---------- DATE NORMALIZATION ----------
    for col in ["Start_Date", "End_Date", "Closing_Date"]:
        if col in client_df.columns:
            client_df[col] = pd.to_datetime(
                client_df[col],
                errors="coerce"
            )

    # =================================================
    # STATUS FILTER
    # =================================================
    status = st.selectbox(
        "Select Status",
        ["Select", "All", "Open", "Closed"]
    )

    if status == "Select":
        st.info("Please select status to view data")
        st.stop()

    if status != "All":
        client_df = client_df[
            client_df["Status"] == status
        ]

    if client_df.empty:
        st.info("No records found for selected filters")
        st.stop()

    # =================================================
    # GRAND TOTALS
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
    # DISPLAY DATA TABLE
    # =================================================
    display_df = client_df.sort_values(
        by="Start_Date",
        ascending=True
    ).copy()

    # SAFE DATE FORMAT
    for col in ["Start_Date", "End_Date"]:
        display_df[col] = display_df[col].dt.strftime("%d/%m/%Y")

    # SAFE RATE FORMAT
    display_df["Rate"] = (
        pd.to_numeric(display_df["Rate"], errors="coerce")
        .fillna(0)
        .round(0)
        .astype(int)
    )

    # ---------------- MOBILE VIEW -------------------
    if is_mobile:

        for _, row in display_df.iterrows():

            st.markdown(
                f"""
<div style="border:1px solid #ddd;
border-radius:12px;
padding:14px;
margin-bottom:12px;
background:#fafafa">

<b>Start Date:</b> {row['Start_Date']}<br>
<b>End Date:</b> {row['End_Date']}<br>
<b>Proposal Cost:</b> ₹ {row['Proposal_Cost']:,.2f}<br>
<b>Final Cost:</b> ₹ {row['Final_Cost']:,.2f}<br>
<b>Profit:</b> ₹ {row['Profit']:,.2f}<br>
<b>Status:</b> {row['Status']}

</div>
""",
                unsafe_allow_html=True
            )

    # ---------------- DESKTOP VIEW ------------------
    else:

        st.dataframe(
            display_df[
                [
                    "Proposal_ID",
                    "Start_Date",
                    "End_Date",
                    "Proposal_Cost",
                    "Final_Cost",
                    "Profit",
                    "Status"
                ]
            ],
            use_container_width=True
        )

# =====================================================
# ================= EXPORT DATA =======================
# =====================================================

if st.session_state.page == "Export":

    st.header("📤 Export Data")

    # ---------- SAFETY CHECK ----------
    if "proposals_df" not in st.session_state:
        st.session_state.proposals_df = pd.DataFrame()

    if "clients_df" not in st.session_state:
        st.session_state.clients_df = pd.DataFrame()

    proposals_df = st.session_state.proposals_df.copy()
    clients_df = st.session_state.clients_df.copy()

    if proposals_df.empty and clients_df.empty:
        st.warning("No data available to export.")
        st.stop()

    # ---------- EXPORT FUNCTION ----------
    @st.cache_data(show_spinner=False)
    def export_excel(clients, proposals):

        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            if not clients.empty:
                clients.to_excel(
                    writer,
                    index=False,
                    sheet_name="Clients"
                )

            if not proposals.empty:
                proposals.to_excel(
                    writer,
                    index=False,
                    sheet_name="Proposals"
                )

        output.seek(0)
        return output

    excel_file = export_excel(
        clients_df,
        proposals_df
    )

    # ---------- TIMESTAMPED FILE NAME ----------
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")

    file_name = f"sigma_consultants_data_{timestamp}.xlsx"

    # ---------- DOWNLOAD BUTTON ----------
    st.download_button(
        label="⬇️ Download Excel (Clients + Proposals)",
        data=excel_file,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ---------- OPTIONAL CSV BACKUP ----------
    st.markdown("### 🔁 Alternative Format")

    if not proposals_df.empty:

        st.download_button(
            label="⬇️ Download Proposals CSV",
            data=proposals_df.to_csv(
                index=False,
                encoding="utf-8-sig"
            ),
            file_name=f"sigma_proposals_{timestamp}.csv",
            mime="text/csv"
        )

    if not clients_df.empty:

        st.download_button(
            label="⬇️ Download Clients CSV",
            data=clients_df.to_csv(
                index=False,
                encoding="utf-8-sig"
            ),
            file_name=f"sigma_clients_{timestamp}.csv",
            mime="text/csv"
        )
















