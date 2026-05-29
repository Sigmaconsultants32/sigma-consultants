# =====================================================
# Sigma Consultants – CRM (Modern UI Base)
# =====================================================

import streamlit as st
import pandas as pd
from datetime import datetime
import os
import io
import base64

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="Sigma Consultants CRM",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =====================================================
# BRAND COLORS
# =====================================================

PRIMARY_COLOR = "#0F172A"
SECONDARY_COLOR = "#2563EB"
ACCENT_COLOR = "#14B8A6"
BG_COLOR = "#F5F7FB"
CARD_BG = "#FFFFFF"
TEXT_COLOR = "#0F172A"
MUTED_COLOR = "#64748B"
BORDER_COLOR = "#E2E8F0"

# =====================================================
# GLOBAL STYLE
# =====================================================

st.markdown(
    f"""
<style>
:root {{
    --sigma-primary: {PRIMARY_COLOR};
    --sigma-secondary: {SECONDARY_COLOR};
    --sigma-accent: {ACCENT_COLOR};
    --sigma-bg: {BG_COLOR};
    --sigma-card: {CARD_BG};
    --sigma-text: {TEXT_COLOR};
    --sigma-muted: {MUTED_COLOR};
    --sigma-border: {BORDER_COLOR};
}}

.stApp {{
    background: linear-gradient(180deg, #F8FAFC 0%, #EEF4FB 100%);
    color: var(--sigma-text);
}}

html, body, [class*="css"] {{
    font-family: "Inter", "Segoe UI", system-ui, -apple-system, BlinkMacSystemFont, "Helvetica Neue", Arial, sans-serif;
}}

div.block-container {{
    padding-top: 1.1rem;
    padding-bottom: 1.4rem;
    max-width: 1320px;
}}

section[data-testid="stVerticalBlock"] {{
    gap: 0.6rem;
}}

section[data-testid="stVerticalBlock"] > div {{
    background: rgba(255, 255, 255, 0.96);
    border: 1px solid rgba(226, 232, 240, 0.9);
    border-radius: 18px;
    padding: 14px 16px;
    margin-bottom: 12px;
    box-shadow: 0 10px 30px rgba(15, 23, 42, 0.05);
    backdrop-filter: blur(6px);
}}

h1, h2, h3, h4, h5 {{
    color: var(--sigma-text);
    font-weight: 700;
    letter-spacing: -0.02em;
}}

p, span, label, div {{
    color: var(--sigma-text);
}}

[data-testid="stMetric"] {{
    background: #ffffff;
    border: 1px solid var(--sigma-border);
    border-radius: 16px;
    padding: 12px 14px;
    box-shadow: 0 8px 20px rgba(15, 23, 42, 0.04);
}}

[data-testid="stMetricLabel"] p {{
    color: var(--sigma-muted);
    font-size: 0.9rem;
}}

[data-testid="stMetricValue"] {{
    color: var(--sigma-text);
    font-weight: 800;
}}

button[kind="primary"] {{
    background: linear-gradient(135deg, var(--sigma-secondary) 0%, var(--sigma-accent) 100%);
    border: none;
    border-radius: 12px;
    color: white;
    font-weight: 700;
    box-shadow: 0 10px 20px rgba(37, 99, 235, 0.18);
}}

button[kind="primary"]:hover {{
    filter: brightness(0.97);
    transform: translateY(-1px);
}}

button {{
    border-radius: 12px;
}}

[data-testid="stDataFrame"], .stDataFrame {{
    border: 1px solid var(--sigma-border);
    border-radius: 14px;
    overflow: hidden;
    background: #fff;
}}

[data-baseweb="select"], [data-baseweb="input"] {{
    border-radius: 12px !important;
}}

[data-baseweb="select"] > div, [data-baseweb="input"] > div {{
    min-height: 42px;
}}

section[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, #0B1120 0%, #111827 100%);
    padding-top: 18px;
    border-right: 1px solid rgba(255,255,255,0.06);
}}

section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] p {{
    color: #F8FAFC;
}}

.sidebar-btn button {{
    background: rgba(255,255,255,0.04);
    color: #E5E7EB;
    border-radius: 12px;
    font-weight: 600;
    border: 1px solid rgba(255,255,255,0.08);
    padding: 0.6rem 0.85rem;
    margin-bottom: 6px;
    transition: all 0.18s ease-in-out;
    box-shadow: none;
}}

.sidebar-btn button:hover {{
    background: rgba(255,255,255,0.09);
    transform: translateX(2px);
}}

.sidebar-active button {{
    background: linear-gradient(135deg, var(--sigma-secondary), var(--sigma-accent)) !important;
    color: white !important;
    font-weight: 700;
    border-radius: 12px;
    box-shadow: 0 10px 20px rgba(37,99,235,0.2);
    border: 1px solid rgba(255,255,255,0.08);
}}

section[data-testid="stSidebar"] hr {{
    margin-top: 12px;
    margin-bottom: 12px;
    border-color: rgba(255,255,255,0.08);
}}

.stToggle label {{
    font-weight: 600;
}}

/* Compact mobile-friendly spacing */
@media (max-width: 768px) {{
    div.block-container {{
        padding-left: 0.7rem;
        padding-right: 0.7rem;
        padding-top: 0.8rem;
    }}

    section[data-testid="stVerticalBlock"] > div {{
        padding: 12px 12px;
        border-radius: 16px;
    }}
}}

</style>
""",
    unsafe_allow_html=True
)

# =====================================================
# SIDEBAR STYLE
# =====================================================

st.markdown(
    """
<style>
section[data-testid="stSidebar"] {
    padding-left: 0.6rem;
    padding-right: 0.6rem;
}

section[data-testid="stSidebar"] button[kind="secondary"] {
    background: rgba(255,255,255,0.08);
    color: #F8FAFC;
    border: 1px solid rgba(255,255,255,0.10);
    border-radius: 12px;
    font-weight: 600;
}
</style>
""",
    unsafe_allow_html=True
)

# =====================================================
# MOBILE MODE
# =====================================================

if "is_mobile" not in st.session_state:
    st.session_state.is_mobile = False

st.session_state.is_mobile = st.toggle(
    "📱 Mobile View",
    value=st.session_state.is_mobile,
    help="Switch between compact mobile-style cards and desktop table view."
)

is_mobile = st.session_state.is_mobile

# =====================================================
# CARD COMPONENT
# =====================================================

def card(title, value):

    st.markdown(
        f"""
<div style="
padding:12px 14px;
border-radius:14px;
background:linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(248,250,252,0.98) 100%);
margin-bottom:10px;
box-shadow:0 8px 20px rgba(15,23,42,0.05);
border:1px solid rgba(226,232,240,0.95);
border-left:5px solid {SECONDARY_COLOR};
">

<div style="font-size:12px;color:{MUTED_COLOR};text-transform:uppercase;letter-spacing:0.08em;font-weight:700;">{title}</div>
<div style="font-size:20px;font-weight:800;color:{TEXT_COLOR};margin-top:4px;">{value}</div>

</div>
""",
        unsafe_allow_html=True
    )

# =====================================================
# LOGIN SYSTEM
# =====================================================

PASSWORD = os.getenv("SIGMA_PASSWORD", "sigma123")

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:

    col1, col2, col3 = st.columns([1,2,1])

    with col2:

        st.markdown("### 🔐 Sigma Consultants Login")

        pwd = st.text_input("Password", type="password")

        if st.button(
            "Login",
            use_container_width=True,
            type="primary"
        ):

            if pwd == PASSWORD:

                st.session_state.auth = True
                st.session_state.page = "Welcome"
                st.rerun()

            else:

                st.error("Incorrect password")

    st.stop()

# =====================================================
# FILE PATHS
# =====================================================

CLIENT_FILE = "clients.xlsx"
PROPOSAL_FILE = "proposals.xlsx"

# =====================================================
# LOAD CLIENTS
# =====================================================

def load_clients():

    required_cols = [
        "Client_ID",
        "Client_Name",
        "Created_Date",
        "Is_Archived",
        "Notes"
    ]

    if os.path.exists(CLIENT_FILE):

        df = pd.read_excel(CLIENT_FILE)

    else:

        df = pd.DataFrame(columns=required_cols)
        df.to_excel(CLIENT_FILE, index=False)

    for col in required_cols:

        if col not in df.columns:
            df[col] = None

    df["Is_Archived"] = df["Is_Archived"].fillna(False)

    return df

# =====================================================
# LOAD PROPOSALS
# =====================================================

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

    if os.path.exists(PROPOSAL_FILE):

        df = pd.read_excel(PROPOSAL_FILE)

    else:

        df = pd.DataFrame(columns=required_cols)
        df.to_excel(PROPOSAL_FILE, index=False)

    for col in required_cols:

        if col not in df.columns:
            df[col] = None

    numeric_cols = [
        "Proposal_Cost",
        "Rate",
        "Final_Cost",
        "Profit"
    ]

    for col in numeric_cols:

        df[col] = pd.to_numeric(
            df[col],
            errors="coerce"
        ).fillna(0)

    for col in ["Start_Date", "End_Date", "Closing_Date"]:

        df[col] = pd.to_datetime(
            df[col],
            errors="coerce"
        )

    df["Status"] = df["Status"].apply(normalize_status)

    return df

# =====================================================
# SESSION INIT
# =====================================================

if "clients_df" not in st.session_state:
    st.session_state.clients_df = load_clients()

if "proposals_df" not in st.session_state:
    st.session_state.proposals_df = load_proposals()

clients_df = st.session_state.clients_df
proposals_df = st.session_state.proposals_df

# =====================================================
# SAVE FUNCTIONS
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

# =====================================================
# SAFE ID GENERATORS
# =====================================================

def new_client_id():

    if clients_df.empty:
        return "SIG-C-001"

    last = (
        clients_df["Client_ID"]
        .astype(str)
        .str.replace("SIG-C-", "", regex=False)
        .fillna("0")
        .astype(int)
        .max()
    )

    return f"SIG-C-{last+1:03d}"


def new_proposal_id():

    if proposals_df.empty:
        return "SIG-P-001"

    last = (
        proposals_df["Proposal_ID"]
        .astype(str)
        .str.replace("SIG-P-", "", regex=False)
        .fillna("0")
        .astype(int)
        .max()
    )

    return f"SIG-P-{last+1:03d}"

# =====================================================
# INTEREST CALC ENGINE
# =====================================================

def calc(principal, rate, days):

    profit = principal * rate * (days / 30) / 100
    final = principal + profit

    return final, profit

# =====================================================
# STATUS NORMALIZATION
# =====================================================

def normalize_status(value):

    if pd.isna(value):
        return "Open"

    status = str(value).strip().lower()

    if status in {"close", "closed", "done", "complete", "completed"}:
        return "Close"

    if status == "open":
        return "Open"

    return str(value).strip().title()

# =====================================================
# PAGE STATE
# =====================================================

if "page" not in st.session_state:
    st.session_state.page = "Welcome"

# =====================================================
# SIDEBAR NAVIGATION
# =====================================================

with st.sidebar:

    st.markdown("## 📂 Sigma Consultants")

    current_page = st.session_state.page


    def nav(label, page_key):

        is_active = current_page == page_key

        btn_class = "sidebar-active" if is_active else "sidebar-btn"

        st.markdown(f'<div class="{btn_class}">', unsafe_allow_html=True)

        if st.button(label, use_container_width=True):
            st.session_state.page = page_key
            st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)


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
        st.session_state.page = "Welcome"
        st.rerun()
        
# =====================================================
# WELCOME SCREEN
# =====================================================

if st.session_state.page == "Welcome":

    st.markdown("<br>", unsafe_allow_html=True)

    center_col = st.columns([1,2,1])[1]

    with center_col:

        logo_path = "sigma_logo.png"

        if os.path.exists(logo_path):

            st.markdown(
                f"""
                <div style="text-align:center;">
                    <img src="data:image/png;base64,{base64.b64encode(open(logo_path, "rb").read()).decode()}"
                    width="200">
                </div>
                """,
                unsafe_allow_html=True
            )

        else:

            st.markdown(
                "<h2 style='text-align:center;'>Sigma Consultants</h2>",
                unsafe_allow_html=True
            )

        st.markdown(
            """
<div style="
border-radius:14px;
padding:18px;
margin-top:12px;
background:#ffffff;
box-shadow:0 6px 18px rgba(0,0,0,0.08);
text-align:center;
max-width:520px;
margin-left:auto;
margin-right:auto;
">

<h4>Welcome to Sigma Consultants CRM</h4>

<p style="color:#555;font-size:14px;">
Manage clients, proposals, investments, profits,
and maturity tracking from one dashboard.
</p>

</div>
""",
            unsafe_allow_html=True
        )

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("➕ Add Client", use_container_width=True, type="primary"):
            st.session_state.page = "Clients"
            st.rerun()

        if st.button("📄 Add Proposal", use_container_width=True):
            st.session_state.page = "AddProposal"
            st.rerun()

        if st.button("🔍 Find Details", use_container_width=True):
            st.session_state.page = "Find"
            st.rerun()

        if st.button("📊 View Summary", use_container_width=True):
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
        ["Select", "Open", "Close"]
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
<div style="background:#e3f2fd;border-radius:14px;
padding:14px;margin:15px 0;border:1px solid #90caf9">

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

    status = st.selectbox("Status", ["All", "Open", "Close"])

    df = df_master if status == "All" else df_master[
        df_master["Status"] == status
    ]

    if df.empty:
        st.info("No proposals found")
        return

    proposal_options = ["Select Proposal"] + sorted(
        df["Proposal_ID"].dropna().unique().tolist()
    )

    proposal_id = st.selectbox(
        "Select Proposal ID",
        proposal_options
    )

    if proposal_id == "Select Proposal":
        return

    proposal_df = df[df["Proposal_ID"] == proposal_id]

    if proposal_df.empty:
        st.warning("Proposal not found")
        return

    start_date = proposal_df["Start_Date"].iloc[0]
    end_date = proposal_df["End_Date"].iloc[0]
    rate_val = proposal_df["Rate"].iloc[0]

    c1, c2, c3 = st.columns(3)

    c1.text_input(
        "Start Date",
        start_date.strftime("%d-%m-%Y") if pd.notna(start_date) else "—",
        disabled=True
    )

    c2.text_input(
        "End Date",
        end_date.strftime("%d-%m-%Y") if pd.notna(end_date) else "—",
        disabled=True
    )

    c3.text_input(
        "Rate (%)",
        int(round(rate_val)) if pd.notna(rate_val) else 0,
        disabled=True
    )

    client = st.selectbox(
        "Select Client",
        ["All"] + sorted(
            proposal_df["Client_Name"].dropna().unique().tolist()
        )
    )

    result = (
        proposal_df
        if client == "All"
        else proposal_df[
            proposal_df["Client_Name"] == client
        ]
    )

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
        result[
            ["Client_Name", "Proposal_Cost", "Final_Cost", "Profit"]
        ].round(2),
        use_container_width=True,
        hide_index=True
    )


# ================= BY CLIENT NAME =================
def render_by_client(df_master, is_mobile):

    st.subheader("🔎 Find Details Using Client Name")

    status = st.selectbox("Status", ["All", "Open", "Close"])

    df = df_master if status == "All" else df_master[
        df_master["Status"] == status
    ]

    if df.empty:
        return

    client_list = ["Select Client Name"] + sorted(
        df["Client_Name"].dropna().unique().tolist()
    )

    client = st.selectbox("Client Name", client_list)

    if client == "Select Client Name":
        return

    df = df[df["Client_Name"] == client]

    date_type = st.radio(
        "Date Type",
        ["Start Date", "End Date"],
        horizontal=True
    )

    date_col = (
        "Start_Date"
        if date_type == "Start Date"
        else "End_Date"
    )

    df = df.dropna(subset=[date_col]).copy()

    if df.empty:
        return

    df["DateOnly"] = df[date_col].dt.date

    date_list = ["Select Date"] + sorted(
        df["DateOnly"].unique().tolist()
    )

    selected_date = st.selectbox(
        "Select Date",
        date_list,
        format_func=lambda x:
            x if isinstance(x, str)
            else x.strftime("%d-%m-%Y")
    )

    if selected_date == "Select Date":
        return

    result = df[df["DateOnly"] == selected_date]

    render_grand_total(
        result["Proposal_Cost"].sum(),
        result["Final_Cost"].sum(),
        result["Profit"].sum(),
        is_mobile
    )

    st.dataframe(
        result[
            ["Proposal_Cost", "Rate", "Final_Cost", "Profit"]
        ].round(2),
        use_container_width=True,
        hide_index=True
    )


# ================= BY DATE =================
def render_by_date(df_master, is_mobile):

    st.subheader("📅 Find Details Using Start / End Date")

    status = st.selectbox(
        "Status",
        ["All", "Open", "Close"],
        key="date_status"
    )

    df = df_master if status == "All" else df_master[
        df_master["Status"] == status
    ]

    if df.empty:
        return

    if "date_blocks" not in st.session_state:
        st.session_state.date_blocks = [1]

    if "remove_block" not in st.session_state:
        st.session_state.remove_block = None

    if st.button("➕ Check Another Date"):
        st.session_state.date_blocks.append(
            max(st.session_state.date_blocks) + 1
        )

    st.divider()

    for i, block_id in enumerate(
        list(st.session_state.date_blocks)
    ):

        col1, col2 = st.columns([6, 1])

        with col1:
            st.markdown(f"### 🔍 Date Filter {block_id}")

        with col2:
            if i != 0:
                if st.button("❌", key=f"remove_{block_id}"):
                    st.session_state.remove_block = block_id

        date_type = st.radio(
            "Date Type",
            ["Start Date", "End Date"],
            horizontal=True,
            key=f"date_type_{block_id}"
        )

        date_col = (
            "Start_Date"
            if date_type == "Start Date"
            else "End_Date"
        )

        df_local = df.dropna(subset=[date_col]).copy()

        if df_local.empty:
            continue

        df_local["DateOnly"] = df_local[date_col].dt.date

        selected_date = st.selectbox(
            f"Select {date_type}",
            [None] + sorted(df_local["DateOnly"].unique()),
            key=f"date_select_{block_id}",
            format_func=lambda x:
                "Select"
                if x is None
                else x.strftime("%d-%m-%Y")
        )

        if selected_date is None:
            continue

        result = df_local[
            df_local["DateOnly"] == selected_date
        ]

        render_grand_total(
            result["Proposal_Cost"].sum(),
            result["Final_Cost"].sum(),
            result["Profit"].sum(),
            is_mobile
        )

        st.dataframe(
            result[
                [
                    "Client_Name",
                    "Proposal_Cost",
                    "Rate",
                    "Final_Cost",
                    "Profit"
                ]
            ].round(2),
            use_container_width=True,
            hide_index=True
        )

        st.divider()

    if st.session_state.remove_block:

        if (
            st.session_state.remove_block
            in st.session_state.date_blocks
        ):
            st.session_state.date_blocks.remove(
                st.session_state.remove_block
            )

        st.session_state.remove_block = None


# ================= PAGE CONTROLLER =================
if st.session_state.page == "Find":

    st.header("🔍 Find Proposal Details")

    if "proposals_df" not in st.session_state:
        st.warning("No proposal data available")
        st.stop()

    if st.session_state.proposals_df.empty:
        st.warning("No proposal data available")
        st.stop()

    df_master = get_master_df(
        st.session_state.proposals_df
    )

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

    # ---------- DATE / STATUS NORMALIZATION ----------
    for col in ["Start_Date", "End_Date", "Closing_Date"]:
        if col in proposals_df.columns:
            proposals_df[col] = pd.to_datetime(proposals_df[col], errors="coerce")

    if "Status" in proposals_df.columns:
        proposals_df["Status"] = proposals_df["Status"].apply(normalize_status)
    else:
        proposals_df["Status"] = "Open"

    # =================================================
    # VIEW MODE
    # =================================================
    view_mode = st.radio(
        "Select View",
        ["Single Client", "All Clients"],
        horizontal=True
    )

    # =================================================
    # SINGLE CLIENT VIEW
    # =================================================
    if view_mode == "Single Client":

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

        client_df = proposals_df[
            proposals_df["Client_ID"] == client_id
        ].copy()

        if client_df.empty:
            st.info("No proposals available for this client")
            st.stop()

        status = st.selectbox(
            "Select Proposal Status",
            ["All", "Open", "Close"]
        )

        if status != "All":
            client_df = client_df[client_df["Status"] == status]

        if client_df.empty:
            st.info("No records found for selected filters")
            st.stop()

        total_invest = client_df["Proposal_Cost"].sum()
        total_final = client_df["Final_Cost"].sum()
        total_profit = client_df["Profit"].sum()

        open_props = (client_df["Status"] == "Open").sum()
        closed_props = (client_df["Status"] == "Close").sum()

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
        st.subheader("📄 Proposal Details")

        display_df = client_df.sort_values(
            by=["Start_Date", "Proposal_ID"],
            ascending=True
        ).copy()

        display_df.insert(0, "Sr. No.", range(1, len(display_df) + 1))

        for col in ["Start_Date", "End_Date"]:
            if col in display_df.columns:
                display_df[col] = pd.to_datetime(display_df[col], errors="coerce").dt.strftime("%d/%m/%Y")

        for col in ["Proposal_Cost", "Final_Cost", "Profit"]:
            display_df[col] = pd.to_numeric(display_df[col], errors="coerce").fillna(0).round(2)

        st.dataframe(
            display_df[
                [
                    "Sr. No.",
                    "Proposal_ID",
                    "Client_Name",
                    "Start_Date",
                    "End_Date",
                    "Proposal_Cost",
                    "Final_Cost",
                    "Profit",
                    "Status"
                ]
            ],
            use_container_width=True,
            hide_index=True
        )

    # =================================================
    # ALL CLIENTS VIEW
    # =================================================
    else:

        status = st.selectbox(
            "Select Proposal Status",
            ["All", "Open", "Close"]
        )

        all_df = proposals_df.copy()

        if all_df.empty:
            st.info("No proposal data available")
            st.stop()

        if status != "All":
            all_df = all_df[all_df["Status"] == status]

        if all_df.empty:
            st.info("No records found for selected filters")
            st.stop()

        summary_df = (
            all_df.groupby(["Client_Name"], as_index=False)
            .agg(
                **{
                    "Total of Proposal price": ("Proposal_Cost", "sum"),
                    "Total of Profit amount": ("Profit", "sum")
                }
            )
        )

        summary_df["Proposal Status"] = status
        summary_df["Total of Proposal Price and Profit amount"] = (
            summary_df["Total of Proposal price"] + summary_df["Total of Profit amount"]
        )

        summary_df = summary_df[
            [
                "Client_Name",
                "Proposal Status",
                "Total of Proposal price",
                "Total of Profit amount",
                "Total of Proposal Price and Profit amount"
            ]
        ].sort_values("Client_Name").reset_index(drop=True)

        summary_df.insert(0, "Sr. No.", range(1, len(summary_df) + 1))

        summary_df = summary_df.rename(
            columns={
                "Client_Name": "Client Name"
            }
        )

        for col in [
            "Total of Proposal price",
            "Total of Profit amount",
            "Total of Proposal Price and Profit amount"
        ]:
            summary_df[col] = pd.to_numeric(summary_df[col], errors="coerce").fillna(0).round(2)

        if is_mobile:
            table_df = summary_df
            st.dataframe(table_df, use_container_width=True, hide_index=True)
        else:
            st.dataframe(summary_df, use_container_width=True, hide_index=True)

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















