# =====================================================
# Sigma Consultants – CRM (fixed)
# =====================================================
from __future__ import annotations

import io
import os
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# =====================================================
# PATHS & CONSTANTS
# =====================================================

BASE_DIR = Path(__file__).resolve().parent
CLIENT_FILE = BASE_DIR / "clients.xlsx"
PROPOSAL_FILE = BASE_DIR / "proposals.xlsx"
LOGO_FILE = BASE_DIR / "sigma_logo.png"

CLIENT_ID_RE = re.compile(r"SIG-C-(\d+)", re.IGNORECASE)
PROPOSAL_ID_RE = re.compile(r"SIG-P-(\d+)", re.IGNORECASE)

CLIENT_COLS = ["Client_ID", "Client_Name", "Created_Date", "Is_Archived", "Notes"]
PROPOSAL_COLS = [
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
    "Closing_Date",
]
NUMERIC_PROPOSAL_COLS = ["Proposal_Cost", "Rate", "Final_Cost", "Profit"]
DATE_PROPOSAL_COLS = ["Start_Date", "End_Date", "Closing_Date"]
DURATION_DAYS = [15, 20, 30, 45, 60, 90]

PRIMARY_COLOR = "#0F172A"
SECONDARY_COLOR = "#2563EB"
ACCENT_COLOR = "#14B8A6"
BG_COLOR = "#F5F7FB"
CARD_BG = "#FFFFFF"
TEXT_COLOR = "#0F172A"
MUTED_COLOR = "#64748B"
BORDER_COLOR = "#E2E8F0"


# =====================================================
# PAGE CONFIG (must be first Streamlit call)
# =====================================================

st.set_page_config(
    page_title="Sigma Consultants CRM",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =====================================================
# HELPERS
# =====================================================

def normalize_status(value) -> str:
    """Normalize legacy and mixed proposal status labels."""
    if pd.isna(value):
        return "Open"

    status = str(value).strip().lower()
    if status in {"close", "closed", "done", "complete", "completed"}:
        return "Close"
    if status in {"open", "active"}:
        return "Open"
    return str(value).strip().title()


def as_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    if pd.isna(value):
        return False
    return str(value).strip().lower() in {"true", "1", "yes", "y"}


def fmt_money(value) -> str:
    return f"₹ {float(value):,.2f}"


def fmt_date(value) -> str:
    if pd.isna(value):
        return "—"
    ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return "—"
    return ts.strftime("%d-%m-%Y")


def to_date(value, fallback: date | None = None) -> date:
    ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return fallback or date.today()
    return ts.date()


def calc(principal: float, rate: float, days: int) -> tuple[float, float]:
    """30-day month interest: profit = principal * rate% * (days / 30)."""
    days = max(int(days), 0)
    profit = principal * rate * (days / 30.0) / 100.0
    return principal + profit, profit


def next_id(series: pd.Series, prefix: str, pattern: re.Pattern) -> str:
    numbers = []
    for raw in series.dropna().astype(str):
        match = pattern.search(raw.strip())
        if match:
            numbers.append(int(match.group(1)))
    nxt = (max(numbers) + 1) if numbers else 1
    return f"{prefix}{nxt:03d}"


def new_client_id() -> str:
    return next_id(st.session_state.clients_df["Client_ID"], "SIG-C-", CLIENT_ID_RE)


def new_proposal_id() -> str:
    return next_id(st.session_state.proposals_df["Proposal_ID"], "SIG-P-", PROPOSAL_ID_RE)


def atomic_to_excel(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".xlsx.tmp")
    df.to_excel(tmp, index=False)
    os.replace(tmp, path)


def ensure_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    for col in columns:
        if col not in df.columns:
            df[col] = pd.NA
    return df[columns].copy()


def load_clients() -> pd.DataFrame:
    if CLIENT_FILE.exists():
        df = pd.read_excel(CLIENT_FILE)
    else:
        df = pd.DataFrame(columns=CLIENT_COLS)
        atomic_to_excel(df, CLIENT_FILE)

    df = ensure_columns(df, CLIENT_COLS)
    df["Client_ID"] = df["Client_ID"].astype(str).replace({"nan": "", "<NA>": ""})
    df["Client_Name"] = df["Client_Name"].fillna("").astype(str).str.strip()
    df["Notes"] = df["Notes"].fillna("").astype(str)
    df["Is_Archived"] = df["Is_Archived"].map(as_bool)
    df["Created_Date"] = pd.to_datetime(df["Created_Date"], errors="coerce")
    return df


def load_proposals() -> pd.DataFrame:
    if PROPOSAL_FILE.exists():
        df = pd.read_excel(PROPOSAL_FILE)
    else:
        df = pd.DataFrame(columns=PROPOSAL_COLS)
        atomic_to_excel(df, PROPOSAL_FILE)

    df = ensure_columns(df, PROPOSAL_COLS)
    for col in NUMERIC_PROPOSAL_COLS:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    for col in DATE_PROPOSAL_COLS:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    df["Status"] = df["Status"].apply(normalize_status)
    df["Client_Name"] = df["Client_Name"].fillna("").astype(str).str.strip()
    df["Proposal_ID"] = df["Proposal_ID"].astype(str)
    df["Client_ID"] = df["Client_ID"].astype(str)
    return df


def save_clients() -> None:
    atomic_to_excel(st.session_state.clients_df, CLIENT_FILE)


def save_proposals() -> None:
    atomic_to_excel(st.session_state.proposals_df, PROPOSAL_FILE)


def active_clients_df() -> pd.DataFrame:
    clients = st.session_state.clients_df
    return clients.loc[~clients["Is_Archived"]].copy()


def unique_sorted(series: pd.Series) -> list:
    return sorted(series.dropna().unique().tolist())


# =====================================================
# STYLE
# =====================================================

def inject_css() -> None:
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

h1, h2, h3, h4, h5 {{
    color: var(--sigma-text);
    font-weight: 700;
    letter-spacing: -0.02em;
}}

[data-testid="stMetric"] {{
    background: #ffffff;
    border: 1px solid var(--sigma-border);
    border-radius: 16px;
    padding: 12px 14px;
    box-shadow: 0 8px 20px rgba(15, 23, 42, 0.04);
}}

[data-testid="stMetricLabel"] p {{ color: var(--sigma-muted); font-size: 0.9rem; }}
[data-testid="stMetricValue"] {{ color: var(--sigma-text); font-weight: 800; }}

button[kind="primary"] {{
    background: linear-gradient(135deg, var(--sigma-secondary) 0%, var(--sigma-accent) 100%);
    border: none;
    border-radius: 12px;
    color: white;
    font-weight: 700;
    box-shadow: 0 10px 20px rgba(37, 99, 235, 0.18);
}}

[data-testid="stDataFrame"], .stDataFrame {{
    border: 1px solid var(--sigma-border);
    border-radius: 14px;
    overflow: hidden;
    background: #fff;
}}

section[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, #0B1120 0%, #111827 100%);
    padding: 18px 0.6rem 1rem;
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
    margin-bottom: 6px;
}}

.sidebar-btn button:hover {{
    background: rgba(255,255,255,0.09);
}}

.sidebar-active button {{
    background: linear-gradient(135deg, var(--sigma-secondary), var(--sigma-accent)) !important;
    color: white !important;
    font-weight: 700;
    border-radius: 12px;
    border: 1px solid rgba(255,255,255,0.08);
}}

section[data-testid="stSidebar"] hr {{
    border-color: rgba(255,255,255,0.08);
}}

@media (max-width: 768px) {{
    div.block-container {{
        padding-left: 0.7rem;
        padding-right: 0.7rem;
        padding-top: 0.8rem;
    }}
}}
</style>
""",
        unsafe_allow_html=True,
    )


def card(title, value) -> None:
    st.markdown(
        f"""
<div style="
padding:12px 14px;border-radius:14px;
background:linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(248,250,252,0.98) 100%);
margin-bottom:10px;box-shadow:0 8px 20px rgba(15,23,42,0.05);
border:1px solid rgba(226,232,240,0.95);border-left:5px solid {SECONDARY_COLOR};">
<div style="font-size:12px;color:{MUTED_COLOR};text-transform:uppercase;letter-spacing:0.08em;font-weight:700;">{title}</div>
<div style="font-size:20px;font-weight:800;color:{TEXT_COLOR};margin-top:4px;">{value}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def metrics_row(items: list[tuple[str, object]], mobile: bool) -> None:
    if mobile:
        for title, value in items:
            card(title, value)
        return
    cols = st.columns(len(items))
    for col, (title, value) in zip(cols, items):
        col.metric(title, value)


def render_grand_total(total_invest, total_final, total_profit, is_mobile: bool) -> None:
    metrics_row(
        [
            ("Investment", fmt_money(total_invest)),
            ("Final Amount", fmt_money(total_final)),
            ("Profit", fmt_money(total_profit)),
        ],
        is_mobile,
    )


# =====================================================
# SESSION / AUTH
# =====================================================

def init_session() -> None:
    if "auth" not in st.session_state:
        st.session_state.auth = False
    if "page" not in st.session_state:
        st.session_state.page = "Welcome"
    if "is_mobile" not in st.session_state:
        st.session_state.is_mobile = False
    if "proposal_clients" not in st.session_state:
        st.session_state.proposal_clients = []
    if "date_blocks" not in st.session_state:
        st.session_state.date_blocks = [1]
    if "clients_df" not in st.session_state:
        st.session_state.clients_df = load_clients()
    if "proposals_df" not in st.session_state:
        st.session_state.proposals_df = load_proposals()


def expected_password() -> str:
    secrets_pwd = None
    try:
        secrets_pwd = st.secrets.get("SIGMA_PASSWORD")
    except Exception:
        secrets_pwd = None
    return os.getenv("SIGMA_PASSWORD") or secrets_pwd or "sigma123"


def render_login() -> None:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("### 🔐 Sigma Consultants Login")
        with st.form("login_form"):
            pwd = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Login", use_container_width=True, type="primary")
        if submitted:
            if pwd == expected_password():
                st.session_state.auth = True
                st.session_state.page = "Welcome"
                st.rerun()
            st.error("Incorrect password")


def go(page: str) -> None:
    st.session_state.page = page
    st.rerun()


# =====================================================
# NAV
# =====================================================

def sidebar_nav() -> None:
    with st.sidebar:
        st.markdown("## 📂 Sigma Consultants")
        st.session_state.is_mobile = st.toggle(
            "📱 Mobile View",
            value=st.session_state.is_mobile,
            help="Switch between compact cards and desktop metrics.",
        )
        st.markdown("---")
        current = st.session_state.page

        def nav(label: str, page_key: str) -> None:
            css = "sidebar-active" if current == page_key else "sidebar-btn"
            st.markdown(f'<div class="{css}">', unsafe_allow_html=True)
            if st.button(label, use_container_width=True, key=f"nav_{page_key}"):
                go(page_key)
            st.markdown("</div>", unsafe_allow_html=True)

        nav("🏠 Welcome", "Welcome")
        nav("📊 Summary", "Summary")
        nav("➕ Add Proposal", "AddProposal")
        nav("🔍 Find Details", "Find")
        nav("✏️ Edit Proposal", "Edit")
        nav("👤 Clients", "Clients")
        nav("📊 Client Dashboard", "ClientDashboard")
        st.markdown("---")
        nav("📥 Export Data", "Export")
        st.markdown("---")
        if st.button("🔄 Reload Excel files", use_container_width=True):
            st.session_state.clients_df = load_clients()
            st.session_state.proposals_df = load_proposals()
            st.success("Reloaded from disk")
            st.rerun()
        if st.button("🚪 Logout", use_container_width=True):
            st.session_state.auth = False
            st.session_state.page = "Welcome"
            st.rerun()


# =====================================================
# PAGES
# =====================================================

def page_welcome() -> None:
    _, center, _ = st.columns([1, 2, 1])
    with center:
        if LOGO_FILE.exists():
            st.image(str(LOGO_FILE), width=200)
        else:
            st.markdown("<h2 style='text-align:center;'>Sigma Consultants</h2>", unsafe_allow_html=True)

        st.markdown(
            """
<div style="border-radius:14px;padding:18px;margin-top:12px;background:#ffffff;
box-shadow:0 6px 18px rgba(0,0,0,0.08);text-align:center;max-width:520px;
margin-left:auto;margin-right:auto;">
<h4>Welcome to Sigma Consultants CRM</h4>
<p style="color:#555;font-size:14px;">
Manage clients, proposals, investments, profits, and maturity tracking from one dashboard.
</p>
</div>
""",
            unsafe_allow_html=True,
        )
        st.write("")
        if st.button("➕ Add Client", use_container_width=True, type="primary"):
            go("Clients")
        if st.button("📄 Add Proposal", use_container_width=True):
            go("AddProposal")
        if st.button("🔍 Find Details", use_container_width=True):
            go("Find")
        if st.button("📊 View Summary", use_container_width=True):
            go("Summary")


def page_summary(is_mobile: bool) -> None:
    st.header("📊 Summary")
    df = st.session_state.proposals_df.copy()
    if df.empty:
        st.info("No proposals available")
        return

    metrics_row(
        [
            ("Total Investment", fmt_money(df["Proposal_Cost"].sum())),
            ("Total Profit", fmt_money(df["Profit"].sum())),
            ("Open Proposals", int((df["Status"] == "Open").sum())),
        ],
        is_mobile,
    )

    st.markdown("---")
    st.subheader("📅 Date Based Summary")

    status_options = ["All"] + unique_sorted(df["Status"])
    selected_status = st.selectbox("Select Proposal Status", status_options, key="summary_status")
    if selected_status != "All":
        df = df[df["Status"] == selected_status]
    if df.empty:
        st.warning("No data for selected status")
        return

    date_type = st.radio("Select Date Type", ["Start Date", "End Date"], horizontal=True, key="summary_date_type")
    date_col = "Start_Date" if date_type == "Start Date" else "End_Date"
    df = df.dropna(subset=[date_col]).copy()
    if df.empty:
        st.info("No dates available")
        return

    df["DateOnly"] = df[date_col].dt.date
    date_choices = unique_sorted(df["DateOnly"])
    selected_dates = st.multiselect(
        f"Select {date_type}(s)",
        ["All"] + date_choices,
        default=["All"],
        format_func=lambda x: x if x == "All" else x.strftime("%d-%m-%Y"),
        key="summary_dates",
    )
    if selected_dates and "All" not in selected_dates:
        df = df[df["DateOnly"].isin(selected_dates)]
    if df.empty:
        st.warning("No data for selected date(s)")
        return

    df["Rate_Int"] = pd.to_numeric(df["Rate"], errors="coerce").fillna(0).round(0).astype(int)

    st.markdown("### 📊 Proposal Stats")
    render_grand_total(df["Proposal_Cost"].sum(), df["Final_Cost"].sum(), df["Profit"].sum(), is_mobile)
    st.markdown("---")

    summary_df = (
        df.groupby(["Rate_Int", "DateOnly"], as_index=False)
        .agg({"Proposal_Cost": "sum", "Final_Cost": "sum", "Profit": "sum"})
        .sort_values(["Rate_Int", "DateOnly"])
    )
    for _, row in summary_df.iterrows():
        profit_color = "🟢" if row["Profit"] >= 0 else "🔴"
        st.markdown(
            f"""
**{row['DateOnly'].strftime('%d-%m-%Y')} | Rate {int(row['Rate_Int'])}%**  
Investment : {fmt_money(row['Proposal_Cost'])}  
Final Amount : {fmt_money(row['Final_Cost'])}  
{profit_color} Profit : {fmt_money(row['Profit'])}
"""
        )
        st.markdown("---")


def page_add_proposal() -> None:
    st.header("➕ Add New Proposal")
    active = active_clients_df()
    if active.empty:
        st.warning("Please add an active client first.")
        if st.button("Go to Clients"):
            go("Clients")
        return

    start_date = st.date_input("Start Date", value=date.today(), key="add_start")
    days = st.selectbox("Duration (Days)", DURATION_DAYS, key="add_days")
    end_date = (pd.Timestamp(start_date) + pd.Timedelta(days=int(days))).date()
    st.caption(f"End date: **{end_date.strftime('%d-%m-%Y')}**")

    rate = st.number_input("Monthly Rate (%)", min_value=0.0, step=0.25, key="add_rate")
    names = unique_sorted(active["Client_Name"])
    selected_client = st.selectbox("Client Name", ["Select Client"] + names, key="add_client")
    principal = st.number_input("Proposal Amount (₹)", min_value=0.0, step=1000.0, key="add_principal")

    c1, c2, c3 = st.columns(3)
    with c1:
        add_clicked = st.button("➕ Add to Proposal", use_container_width=True)
    with c2:
        clear_clicked = st.button("Clear draft", use_container_width=True)
    if clear_clicked:
        st.session_state.proposal_clients = []
        st.rerun()

    if add_clicked:
        if selected_client == "Select Client":
            st.warning("Select a client")
        elif principal <= 0:
            st.warning("Enter a proposal amount greater than 0")
        elif any(item["Client_Name"] == selected_client for item in st.session_state.proposal_clients):
            st.warning("That client is already on this proposal")
        else:
            final_amount, profit = calc(principal, rate, int(days))
            st.session_state.proposal_clients.append(
                {
                    "Client_Name": selected_client,
                    "Principal": principal,
                    "Profit": profit,
                    "Final_Amount": final_amount,
                }
            )
            st.success(f"Added {selected_client}")

    if st.session_state.proposal_clients:
        st.dataframe(pd.DataFrame(st.session_state.proposal_clients), use_container_width=True, hide_index=True)

    if st.button("💾 Save Proposal", type="primary"):
        if not st.session_state.proposal_clients:
            st.warning("Add at least one client to the proposal before saving")
            return

        proposal_id = new_proposal_id()
        rows = []
        for item in st.session_state.proposal_clients:
            matches = active[active["Client_Name"] == item["Client_Name"]]
            if matches.empty:
                st.error(f"Client not found: {item['Client_Name']}")
                return
            client_row = matches.iloc[0]
            final_amount, profit = calc(item["Principal"], rate, int(days))
            rows.append(
                {
                    "Proposal_ID": proposal_id,
                    "Client_ID": client_row["Client_ID"],
                    "Client_Name": item["Client_Name"],
                    "Proposal_Cost": item["Principal"],
                    "Rate": rate,
                    "Profit": profit,
                    "Final_Cost": final_amount,
                    "Start_Date": pd.Timestamp(start_date),
                    "End_Date": pd.Timestamp(end_date),
                    "Status": "Open",
                    "Closing_Date": pd.NaT,
                }
            )

        st.session_state.proposals_df = pd.concat(
            [st.session_state.proposals_df, pd.DataFrame(rows)],
            ignore_index=True,
        )
        save_proposals()
        st.session_state.proposal_clients = []
        st.success(f"Proposal {proposal_id} created")
        go("Summary")


def page_edit() -> None:
    st.header("✏️ Edit Proposal")
    proposals = st.session_state.proposals_df
    if proposals.empty:
        st.info("No proposals to edit")
        return

    status_filter = st.selectbox("Select Proposal Status", ["Open", "Close"], key="edit_status_filter")
    df = proposals[proposals["Status"] == status_filter]
    if df.empty:
        st.info(f"No {status_filter} proposals")
        return

    proposal_id = st.selectbox(
        "Select Proposal ID",
        ["Select"] + unique_sorted(df["Proposal_ID"]),
        key="edit_proposal_id",
    )
    if proposal_id == "Select":
        return

    proposal_df = df[df["Proposal_ID"] == proposal_id]
    client_name = st.selectbox("Select Client", unique_sorted(proposal_df["Client_Name"]), key="edit_client")
    matches = proposal_df[proposal_df["Client_Name"] == client_name]
    if matches.empty:
        st.error("Row not found")
        return
    row = matches.iloc[0]
    row_index = matches.index[0]

    proposal_cost = st.number_input("Proposal Amount", min_value=0.0, value=float(row["Proposal_Cost"]), key="edit_cost")
    rate = st.number_input("Rate (%)", min_value=0.0, value=float(row["Rate"]), key="edit_rate")
    new_start = st.date_input("Start Date", value=to_date(row["Start_Date"]), key="edit_start")
    new_end = st.date_input("End Date", value=to_date(row["End_Date"], fallback=new_start), key="edit_end")
    new_status = st.selectbox("Status", ["Open", "Close"], index=0 if row["Status"] == "Open" else 1, key="edit_new_status")

    days = (new_end - new_start).days
    if days < 0:
        st.error("End date cannot be before start date")
        return

    final_cost, profit = calc(proposal_cost, rate, days)
    st.caption(f"Duration: {days} days · Final: {fmt_money(final_cost)} · Profit: {fmt_money(profit)}")

    if st.button("💾 Save Changes", type="primary"):
        closing = pd.Timestamp(date.today()) if new_status == "Close" else pd.NaT
        if new_status == "Close" and pd.notna(row["Closing_Date"]):
            closing = row["Closing_Date"]

        st.session_state.proposals_df.at[row_index, "Proposal_Cost"] = proposal_cost
        st.session_state.proposals_df.at[row_index, "Rate"] = rate
        st.session_state.proposals_df.at[row_index, "Start_Date"] = pd.Timestamp(new_start)
        st.session_state.proposals_df.at[row_index, "End_Date"] = pd.Timestamp(new_end)
        st.session_state.proposals_df.at[row_index, "Final_Cost"] = final_cost
        st.session_state.proposals_df.at[row_index, "Profit"] = profit
        st.session_state.proposals_df.at[row_index, "Status"] = new_status
        st.session_state.proposals_df.at[row_index, "Closing_Date"] = closing
        save_proposals()
        st.success("Updated")
        go("Summary")


def render_by_proposal(df_master: pd.DataFrame, is_mobile: bool) -> None:
    st.subheader("📄 Find Details By Proposal")
    status = st.selectbox("Status", ["All", "Open", "Close"], key="find_prop_status")
    df = df_master if status == "All" else df_master[df_master["Status"] == status]
    if df.empty:
        st.info("No proposals found")
        return

    proposal_id = st.selectbox(
        "Select Proposal ID",
        ["Select Proposal"] + unique_sorted(df["Proposal_ID"]),
        key="find_prop_id",
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
    c1.text_input("Start Date", fmt_date(start_date), disabled=True)
    c2.text_input("End Date", fmt_date(end_date), disabled=True)
    c3.text_input("Rate (%)", str(int(round(rate_val)) if pd.notna(rate_val) else 0), disabled=True)

    client = st.selectbox(
        "Select Client",
        ["All"] + unique_sorted(proposal_df["Client_Name"]),
        key="find_prop_client",
    )
    result = proposal_df if client == "All" else proposal_df[proposal_df["Client_Name"] == client]
    if result.empty:
        st.info("No data found")
        return

    render_grand_total(result["Proposal_Cost"].sum(), result["Final_Cost"].sum(), result["Profit"].sum(), is_mobile)
    st.dataframe(
        result[["Client_Name", "Proposal_Cost", "Final_Cost", "Profit"]].round(2),
        use_container_width=True,
        hide_index=True,
    )


def render_by_client(df_master: pd.DataFrame, is_mobile: bool) -> None:
    st.subheader("🔎 Find Details Using Client Name")
    status = st.selectbox("Status", ["All", "Open", "Close"], key="find_client_status")
    df = df_master if status == "All" else df_master[df_master["Status"] == status]
    if df.empty:
        st.info("No proposals found")
        return

    client = st.selectbox(
        "Client Name",
        ["Select Client Name"] + unique_sorted(df["Client_Name"]),
        key="find_client_name",
    )
    if client == "Select Client Name":
        return

    df = df[df["Client_Name"] == client]
    date_type = st.radio("Date Type", ["Start Date", "End Date"], horizontal=True, key="find_client_date_type")
    date_col = "Start_Date" if date_type == "Start Date" else "End_Date"
    df = df.dropna(subset=[date_col]).copy()
    if df.empty:
        st.info("No dates for this client")
        return

    df["DateOnly"] = df[date_col].dt.date
    selected_date = st.selectbox(
        "Select Date",
        ["Select Date"] + unique_sorted(df["DateOnly"]),
        format_func=lambda x: x if isinstance(x, str) else x.strftime("%d-%m-%Y"),
        key="find_client_date",
    )
    if selected_date == "Select Date":
        return

    result = df[df["DateOnly"] == selected_date]
    render_grand_total(result["Proposal_Cost"].sum(), result["Final_Cost"].sum(), result["Profit"].sum(), is_mobile)
    st.dataframe(
        result[["Proposal_ID", "Proposal_Cost", "Rate", "Final_Cost", "Profit", "Status"]].round(2),
        use_container_width=True,
        hide_index=True,
    )


def render_by_date(df_master: pd.DataFrame, is_mobile: bool) -> None:
    st.subheader("📅 Find Details Using Start / End Date")
    status = st.selectbox("Status", ["All", "Open", "Close"], key="date_status")
    df = df_master if status == "All" else df_master[df_master["Status"] == status]
    if df.empty:
        st.info("No proposals found")
        return

    if st.button("➕ Check Another Date"):
        st.session_state.date_blocks.append(max(st.session_state.date_blocks) + 1)
        st.rerun()

    st.divider()
    for i, block_id in enumerate(list(st.session_state.date_blocks)):
        col1, col2 = st.columns([6, 1])
        with col1:
            st.markdown(f"### 🔍 Date Filter {block_id}")
        with col2:
            if i != 0 and st.button("❌", key=f"remove_{block_id}"):
                st.session_state.date_blocks.remove(block_id)
                st.rerun()

        date_type = st.radio(
            "Date Type",
            ["Start Date", "End Date"],
            horizontal=True,
            key=f"date_type_{block_id}",
        )
        date_col = "Start_Date" if date_type == "Start Date" else "End_Date"
        df_local = df.dropna(subset=[date_col]).copy()
        if df_local.empty:
            continue

        df_local["DateOnly"] = df_local[date_col].dt.date
        selected_date = st.selectbox(
            f"Select {date_type}",
            [None] + unique_sorted(df_local["DateOnly"]),
            key=f"date_select_{block_id}",
            format_func=lambda x: "Select" if x is None else x.strftime("%d-%m-%Y"),
        )
        if selected_date is None:
            continue

        result = df_local[df_local["DateOnly"] == selected_date]
        render_grand_total(result["Proposal_Cost"].sum(), result["Final_Cost"].sum(), result["Profit"].sum(), is_mobile)
        st.dataframe(
            result[["Client_Name", "Proposal_Cost", "Rate", "Final_Cost", "Profit"]].round(2),
            use_container_width=True,
            hide_index=True,
        )
        st.divider()


def page_find(is_mobile: bool) -> None:
    st.header("🔍 Find Proposal Details")
    proposals = st.session_state.proposals_df
    if proposals.empty:
        st.warning("No proposal data available")
        return

    df_master = proposals.copy()
    for col in DATE_PROPOSAL_COLS:
        df_master[col] = pd.to_datetime(df_master[col], errors="coerce")

    find_mode = st.radio(
        "Find Details Mode",
        ["By Proposal", "By Client Name", "By Start / End Date"],
        horizontal=True,
        key="find_mode",
    )
    if find_mode == "By Proposal":
        render_by_proposal(df_master, is_mobile)
    elif find_mode == "By Client Name":
        render_by_client(df_master, is_mobile)
    else:
        render_by_date(df_master, is_mobile)


def page_clients() -> None:
    st.header("👤 Clients Management")
    cname = st.text_input("Client Name", key="new_client_name")
    notes = st.text_input("Notes (optional)", key="new_client_notes")

    if st.button("Add Client", type="primary"):
        name = cname.strip()
        if not name:
            st.warning("Enter a client name")
        elif name.lower() in st.session_state.clients_df["Client_Name"].str.lower().tolist():
            st.warning("A client with that name already exists")
        else:
            new_row = pd.DataFrame(
                [
                    {
                        "Client_ID": new_client_id(),
                        "Client_Name": name,
                        "Created_Date": pd.Timestamp.now(),
                        "Is_Archived": False,
                        "Notes": notes.strip(),
                    }
                ]
            )
            st.session_state.clients_df = pd.concat(
                [st.session_state.clients_df, new_row],
                ignore_index=True,
            )
            save_clients()
            st.success("Client added")
            st.rerun()

    st.subheader("Directory")
    show_archived = st.toggle("Show archived clients", value=False)
    view = st.session_state.clients_df.copy()
    if not show_archived:
        view = view.loc[~view["Is_Archived"]]

    if view.empty:
        st.info("No clients to display")
        return

    st.dataframe(view, use_container_width=True, hide_index=True)

    target = st.selectbox("Archive / restore client", unique_sorted(st.session_state.clients_df["Client_Name"]))
    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("Archive"):
            mask = st.session_state.clients_df["Client_Name"] == target
            st.session_state.clients_df.loc[mask, "Is_Archived"] = True
            save_clients()
            st.rerun()
    with col_b:
        if st.button("Restore"):
            mask = st.session_state.clients_df["Client_Name"] == target
            st.session_state.clients_df.loc[mask, "Is_Archived"] = False
            save_clients()
            st.rerun()


def page_client_dashboard(is_mobile: bool) -> None:
    st.header("📊 Client Summary Dashboard")
    clients_df = st.session_state.clients_df.copy()
    proposals_df = st.session_state.proposals_df.copy()

    if clients_df.empty:
        st.info("No clients available")
        return

    active_clients = clients_df.loc[~clients_df["Is_Archived"]].copy()
    if active_clients.empty:
        st.info("No active clients")
        return

    for col in DATE_PROPOSAL_COLS:
        proposals_df[col] = pd.to_datetime(proposals_df[col], errors="coerce")
    proposals_df["Status"] = proposals_df["Status"].apply(normalize_status)

    view_mode = st.radio("Select View", ["Single Client", "All Clients"], horizontal=True, key="dash_view")

    if view_mode == "Single Client":
        client = st.selectbox(
            "Select Client",
            ["Select"] + unique_sorted(active_clients["Client_Name"]),
            key="dash_client",
        )
        if client == "Select":
            st.info("Please select a client to continue")
            return

        client_match = active_clients[active_clients["Client_Name"] == client]
        client_id = client_match.iloc[0]["Client_ID"]
        client_df = proposals_df[proposals_df["Client_ID"] == client_id].copy()
        if client_df.empty:
            st.info("No proposals available for this client")
            return

        status = st.selectbox("Select Proposal Status", ["All", "Open", "Close"], key="dash_status")
        if status != "All":
            client_df = client_df[client_df["Status"] == status]
        if client_df.empty:
            st.info("No records found for selected filters")
            return

        metrics_row(
            [
                ("Total Investment", fmt_money(client_df["Proposal_Cost"].sum())),
                ("Total Final Amount", fmt_money(client_df["Final_Cost"].sum())),
                ("Total Profit", fmt_money(client_df["Profit"].sum())),
                ("Open Proposals", int((client_df["Status"] == "Open").sum())),
                ("Close Proposals", int((client_df["Status"] == "Close").sum())),
            ],
            is_mobile,
        )

        st.markdown("---")
        st.subheader("📄 Proposal Details")
        display_df = client_df.sort_values(by=["Start_Date", "Proposal_ID"], ascending=True).copy()
        display_df.insert(0, "Sr. No.", range(1, len(display_df) + 1))
        for col in ["Start_Date", "End_Date"]:
            display_df[col] = display_df[col].map(fmt_date)
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
                    "Status",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )
        return

    status = st.selectbox("Select Proposal Status", ["All", "Open", "Close"], key="dash_all_status")
    all_df = proposals_df.copy()
    if all_df.empty:
        st.info("No proposal data available")
        return
    if status != "All":
        all_df = all_df[all_df["Status"] == status]
    if all_df.empty:
        st.info("No records found for selected filters")
        return

    summary_df = (
        all_df.groupby(["Client_ID", "Client_Name"], as_index=False)
        .agg(
            **{
                "Total of Proposal price": ("Proposal_Cost", "sum"),
                "Total of Profit amount": ("Profit", "sum"),
            }
        )
        .sort_values("Client_Name")
        .reset_index(drop=True)
    )
    summary_df["Proposal Status"] = status
    summary_df["Total of Proposal Price and Profit amount"] = (
        summary_df["Total of Proposal price"] + summary_df["Total of Profit amount"]
    )
    summary_df.insert(0, "Sr. No.", range(1, len(summary_df) + 1))
    summary_df = summary_df[
        [
            "Sr. No.",
            "Client_Name",
            "Proposal Status",
            "Total of Proposal price",
            "Total of Profit amount",
            "Total of Proposal Price and Profit amount",
        ]
    ].rename(columns={"Client_Name": "Client Name"})
    for col in [
        "Total of Proposal price",
        "Total of Profit amount",
        "Total of Proposal Price and Profit amount",
    ]:
        summary_df[col] = pd.to_numeric(summary_df[col], errors="coerce").fillna(0).round(2)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)


def build_excel_bytes(clients: pd.DataFrame, proposals: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if not clients.empty:
            clients.to_excel(writer, index=False, sheet_name="Clients")
        if not proposals.empty:
            proposals.to_excel(writer, index=False, sheet_name="Proposals")
    return output.getvalue()


def page_export() -> None:
    st.header("📤 Export Data")
    proposals_df = st.session_state.proposals_df.copy()
    clients_df = st.session_state.clients_df.copy()
    if proposals_df.empty and clients_df.empty:
        st.warning("No data available to export.")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    excel_bytes = build_excel_bytes(clients_df, proposals_df)
    st.download_button(
        label="⬇️ Download Excel (Clients + Proposals)",
        data=excel_bytes,
        file_name=f"sigma_consultants_data_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="export_xlsx",
    )

    st.markdown("### 🔁 Alternative Format")
    if not proposals_df.empty:
        st.download_button(
            label="⬇️ Download Proposals CSV",
            data=proposals_df.to_csv(index=False, encoding="utf-8-sig"),
            file_name=f"sigma_proposals_{timestamp}.csv",
            mime="text/csv",
            key="export_proposals_csv",
        )
    if not clients_df.empty:
        st.download_button(
            label="⬇️ Download Clients CSV",
            data=clients_df.to_csv(index=False, encoding="utf-8-sig"),
            file_name=f"sigma_clients_{timestamp}.csv",
            mime="text/csv",
            key="export_clients_csv",
        )


# =====================================================
# MAIN
# =====================================================

inject_css()
init_session()

if not st.session_state.auth:
    render_login()
    st.stop()

sidebar_nav()
is_mobile = bool(st.session_state.is_mobile)

page = st.session_state.page
if page == "Welcome":
    page_welcome()
elif page == "Summary":
    page_summary(is_mobile)
elif page == "AddProposal":
    page_add_proposal()
elif page == "Edit":
    page_edit()
elif page == "Find":
    page_find(is_mobile)
elif page == "Clients":
    page_clients()
elif page == "ClientDashboard":
    page_client_dashboard(is_mobile)
elif page == "Export":
    page_export()
else:
    st.session_state.page = "Welcome"
    st.rerun()
    
