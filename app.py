# =====================================================
# Sigma Consultants – CRM (fixed)
# =====================================================
from __future__ import annotations

import base64
import io
import os
import re
import shutil
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
BACKUP_DIR = BASE_DIR / "backups"
LOGO_FILE = BASE_DIR / "sigma_logo.jpg"
LOGO_FALLBACK = Path(
    r"C:\Users\Smart\.cursor\projects\empty-window\assets"
    r"\c__Users_Smart_AppData_Roaming_Cursor_User_workspaceStorage_empty-window_images_"
    r"Sigma_Con_Logo-e928b1a9-6cdc-44e2-837a-3d8331acffba.jpg"
)

LOGO_WIDTH = {
    "sidebar": 210,
    "login": 320,
    "welcome": 360,
}

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
SECONDARY_DARK = "#1D4ED8"
ACCENT_COLOR = "#059669"
BG_COLOR = "#EEF2F7"
CARD_BG = "#FFFFFF"
TEXT_COLOR = "#0F172A"
MUTED_COLOR = "#5B6B7C"
BORDER_COLOR = "#D8DEE6"
SURFACE_ALT = "#F4F6F9"


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


def auto_backup() -> None:
    """Keep timestamped Excel copies on every save."""
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    try:
        if CLIENT_FILE.exists():
            shutil.copy2(CLIENT_FILE, BACKUP_DIR / f"clients_{stamp}.xlsx")
        if PROPOSAL_FILE.exists():
            shutil.copy2(PROPOSAL_FILE, BACKUP_DIR / f"proposals_{stamp}.xlsx")
    except OSError:
        pass


def save_clients() -> None:
    atomic_to_excel(st.session_state.clients_df, CLIENT_FILE)
    auto_backup()


def save_proposals() -> None:
    atomic_to_excel(st.session_state.proposals_df, PROPOSAL_FILE)
    auto_backup()


def enrich_proposals(df: pd.DataFrame) -> pd.DataFrame:
    """Add days-left and maturity bucket for open proposals."""
    out = df.copy()
    for col in DATE_PROPOSAL_COLS:
        out[col] = pd.to_datetime(out[col], errors="coerce")
    out["Status"] = out["Status"].apply(normalize_status)
    today = date.today()

    def days_left(end) -> int | None:
        if pd.isna(end):
            return None
        return (pd.Timestamp(end).date() - today).days

    def maturity_bucket(row) -> str:
        if row["Status"] != "Open":
            return "Closed"
        left = row["Days_Left"]
        if left is None:
            return "No end date"
        if left < 0:
            return "Overdue"
        if left == 0:
            return "Due today"
        if left <= 7:
            return "Due in 7 days"
        if left <= 15:
            return "Due in 15 days"
        if left <= 30:
            return "Due in 30 days"
        return "Later"

    out["Days_Left"] = out["End_Date"].apply(days_left)
    out["Maturity_Bucket"] = out.apply(maturity_bucket, axis=1)
    return out


def filter_by_maturity_bucket(df: pd.DataFrame, bucket: str) -> pd.DataFrame:
    open_df = df[df["Status"] == "Open"].copy()
    if bucket == "All open":
        return open_df.sort_values(["End_Date", "Proposal_ID", "Client_Name"])
    if bucket == "Overdue":
        return open_df[open_df["Days_Left"] < 0].sort_values("End_Date")
    if bucket == "Due today":
        return open_df[open_df["Days_Left"] == 0].sort_values("Proposal_ID")
    if bucket == "Next 7 days":
        return open_df[open_df["Days_Left"].between(0, 7)].sort_values("End_Date")
    if bucket == "Next 15 days":
        return open_df[open_df["Days_Left"].between(0, 15)].sort_values("End_Date")
    if bucket == "Next 30 days":
        return open_df[open_df["Days_Left"].between(0, 30)].sort_values("End_Date")
    return open_df.sort_values("End_Date")


def view_proposal(proposal_id: str) -> None:
    st.session_state.selected_proposal_id = proposal_id
    go("ProposalDetail")


def close_proposal(proposal_id: str) -> None:
    mask = st.session_state.proposals_df["Proposal_ID"] == proposal_id
    st.session_state.proposals_df.loc[mask, "Status"] = "Close"
    st.session_state.proposals_df.loc[mask, "Closing_Date"] = pd.Timestamp(date.today())
    save_proposals()


def renew_proposal(proposal_id: str, new_start: date, duration_days: int) -> str:
    source = st.session_state.proposals_df[
        st.session_state.proposals_df["Proposal_ID"] == proposal_id
    ].copy()
    if source.empty:
        raise ValueError("Proposal not found")

    new_id = new_proposal_id()
    new_end = pd.Timestamp(new_start) + pd.Timedelta(days=int(duration_days))
    rows = []
    for _, row in source.iterrows():
        final_cost, profit = calc(float(row["Proposal_Cost"]), float(row["Rate"]), duration_days)
        rows.append(
            {
                "Proposal_ID": new_id,
                "Client_ID": row["Client_ID"],
                "Client_Name": row["Client_Name"],
                "Proposal_Cost": row["Proposal_Cost"],
                "Rate": row["Rate"],
                "Profit": profit,
                "Final_Cost": final_cost,
                "Start_Date": pd.Timestamp(new_start),
                "End_Date": new_end,
                "Status": "Open",
                "Closing_Date": pd.NaT,
            }
        )
    st.session_state.proposals_df = pd.concat(
        [st.session_state.proposals_df, pd.DataFrame(rows)],
        ignore_index=True,
    )
    save_proposals()
    return new_id


def proposal_display_table(df: pd.DataFrame) -> pd.DataFrame:
    """Format proposal rows for readable tables."""
    show = df.copy()
    show["End_Date"] = show["End_Date"].map(fmt_date)
    show["Start_Date"] = show["Start_Date"].map(fmt_date)
    show["Rate"] = pd.to_numeric(show["Rate"], errors="coerce").round(2)
    for col in ["Proposal_Cost", "Final_Cost", "Profit"]:
        show[col] = pd.to_numeric(show[col], errors="coerce").round(2)
    return show


RAW_IMPORT_ALIASES = {
    "start date": "Start_Date",
    "start_date": "Start_Date",
    "client name": "Client_Name",
    "client_name": "Client_Name",
    "initial amount": "Proposal_Cost",
    "proposal amount": "Proposal_Cost",
    "amount": "Proposal_Cost",
    "principal": "Proposal_Cost",
    "duration": "Duration",
    "end date": "End_Date",
    "end_date": "End_Date",
    "profit rate": "Rate",
    "rate": "Rate",
    "rate %": "Rate",
    "final amount": "Final_Cost",
    "final_amount": "Final_Cost",
    "profit": "Profit",
    "status": "Status",
}

DEFAULT_RAW_IMPORT_PATH = Path(r"c:\Users\Smart\Desktop\dATa.xlsx")


def normalize_raw_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Map common spreadsheet headers to internal column names."""
    out = df.copy()
    rename = {}
    for col in out.columns:
        key = str(col).strip().lower()
        if key in RAW_IMPORT_ALIASES:
            rename[col] = RAW_IMPORT_ALIASES[key]
    return out.rename(columns=rename)


def prepare_raw_import_df(df: pd.DataFrame) -> pd.DataFrame:
    """Clean and validate a raw import sheet."""
    raw = normalize_raw_columns(df)
    required = ["Client_Name", "Start_Date", "End_Date", "Proposal_Cost", "Rate"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError(f"Missing columns: {', '.join(missing)}")

    raw = raw.copy()
    raw["Client_Name"] = raw["Client_Name"].fillna("").astype(str).str.strip()
    raw = raw[raw["Client_Name"] != ""]
    raw["Start_Date"] = pd.to_datetime(raw["Start_Date"], errors="coerce")
    raw["End_Date"] = pd.to_datetime(raw["End_Date"], errors="coerce")
    raw["Proposal_Cost"] = pd.to_numeric(raw["Proposal_Cost"], errors="coerce")
    raw["Rate"] = pd.to_numeric(raw["Rate"], errors="coerce")
    if "Final_Cost" not in raw.columns:
        raw["Final_Cost"] = pd.NA
    if "Profit" not in raw.columns:
        raw["Profit"] = pd.NA
    raw["Final_Cost"] = pd.to_numeric(raw["Final_Cost"], errors="coerce")
    raw["Profit"] = pd.to_numeric(raw["Profit"], errors="coerce")

    raw = raw.dropna(subset=["Client_Name", "Start_Date", "End_Date", "Proposal_Cost", "Rate"])
    if raw.empty:
        raise ValueError("No valid rows found after cleaning.")

    duration = (raw["End_Date"] - raw["Start_Date"]).dt.days
    recalc_final, recalc_profit = zip(
        *[
            calc(float(p), float(r), max(int(d), 0))
            for p, r, d in zip(raw["Proposal_Cost"], raw["Rate"], duration)
        ]
    )
    raw["Final_Cost"] = raw["Final_Cost"].fillna(pd.Series(recalc_final, index=raw.index))
    raw["Profit"] = raw["Profit"].fillna(pd.Series(recalc_profit, index=raw.index))

    if "Status" in raw.columns:
        raw["Status"] = raw["Status"].apply(normalize_status)
    else:
        today = pd.Timestamp(date.today())
        raw["Status"] = raw["End_Date"].apply(
            lambda end: "Close" if pd.notna(end) and end < today else "Open"
        )
    raw["Closing_Date"] = raw.apply(
        lambda row: row["End_Date"] if row["Status"] == "Close" else pd.NaT,
        axis=1,
    )
    return raw.reset_index(drop=True)


def upsert_clients_from_raw(raw: pd.DataFrame, clients_df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str]]:
    """Create Client_ID for each unique client name."""
    clients = clients_df.copy()
    name_to_id: dict[str, str] = {}
    for name, cid in zip(clients["Client_Name"], clients["Client_ID"]):
        key = str(name).strip()
        if key:
            name_to_id[key.lower()] = str(cid)

    new_rows = []
    for name in sorted(raw["Client_Name"].unique()):
        clean = str(name).strip()
        if not clean:
            continue
        if clean.lower() in name_to_id:
            continue
        cid = next_id(
            clients["Client_ID"] if not clients.empty else pd.Series([], dtype=str),
            "SIG-C-",
            CLIENT_ID_RE,
        )
        name_to_id[clean.lower()] = cid
        new_rows.append(
            {
                "Client_ID": cid,
                "Client_Name": clean,
                "Created_Date": pd.Timestamp.now(),
                "Is_Archived": False,
                "Notes": "Imported from Excel",
            }
        )
        clients = pd.concat([clients, pd.DataFrame([new_rows[-1]])], ignore_index=True)

    for name, cid in zip(clients["Client_Name"], clients["Client_ID"]):
        key = str(name).strip()
        if key:
            name_to_id[key.lower()] = str(cid)
    return clients, name_to_id


def build_proposals_from_raw(
    raw: pd.DataFrame,
    name_to_id: dict[str, str],
    proposals_df: pd.DataFrame,
    group_by_terms: bool = True,
) -> pd.DataFrame:
    """
    Build proposal rows with auto Proposal_ID.
    Rows with same start date, end date, and rate become one proposal (multi-client).
    """
    work = raw.copy()
    if group_by_terms:
        work["_Group_Key"] = work.apply(
            lambda row: (
                pd.Timestamp(row["Start_Date"]).normalize(),
                pd.Timestamp(row["End_Date"]).normalize(),
                round(float(row["Rate"]), 4),
            ),
            axis=1,
        )
    else:
        work["_Group_Key"] = work.index

    base = proposals_df["Proposal_ID"] if not proposals_df.empty else pd.Series([], dtype=str)
    group_to_pid: dict = {}
    rows = []
    for _, row in work.iterrows():
        gkey = row["_Group_Key"]
        if gkey not in group_to_pid:
            group_to_pid[gkey] = next_id(
                pd.concat([base, pd.Series(list(group_to_pid.values()))], ignore_index=True),
                "SIG-P-",
                PROPOSAL_ID_RE,
            )
            base = pd.concat([base, pd.Series([group_to_pid[gkey]])], ignore_index=True)

        cname = str(row["Client_Name"]).strip()
        cid = name_to_id.get(cname.lower())
        if not cid:
            raise ValueError(f"Client ID not found for: {cname}")

        rows.append(
            {
                "Proposal_ID": group_to_pid[gkey],
                "Client_ID": cid,
                "Client_Name": cname,
                "Proposal_Cost": float(row["Proposal_Cost"]),
                "Rate": float(row["Rate"]),
                "Final_Cost": float(row["Final_Cost"]),
                "Profit": float(row["Profit"]),
                "Start_Date": pd.Timestamp(row["Start_Date"]),
                "End_Date": pd.Timestamp(row["End_Date"]),
                "Status": row["Status"],
                "Closing_Date": row["Closing_Date"],
            }
        )
    return pd.DataFrame(rows, columns=PROPOSAL_COLS)


def import_raw_excel(
    raw_df: pd.DataFrame,
    clients_df: pd.DataFrame,
    proposals_df: pd.DataFrame,
    append: bool = True,
    group_by_terms: bool = True,
) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Import raw spreadsheet; auto-create Client_ID and Proposal_ID."""
    prepared = prepare_raw_import_df(raw_df)
    clients, name_to_id = upsert_clients_from_raw(prepared, clients_df)
    imported = build_proposals_from_raw(prepared, name_to_id, proposals_df, group_by_terms)

    if append:
        final_clients = clients
        final_proposals = pd.concat([proposals_df, imported], ignore_index=True)
    else:
        final_clients = clients
        final_proposals = imported

    stats = {
        "rows": len(prepared),
        "clients_new": len(clients) - len(clients_df),
        "proposals": imported["Proposal_ID"].nunique(),
        "lines": len(imported),
    }
    return final_clients, final_proposals, stats


def active_clients_df() -> pd.DataFrame:
    clients = st.session_state.clients_df
    return clients.loc[~clients["Is_Archived"]].copy()


def unique_sorted(series: pd.Series) -> list:
    return sorted(series.dropna().unique().tolist())


def ensure_logo() -> None:
    if LOGO_FILE.exists() or not LOGO_FALLBACK.exists():
        return
    try:
        import shutil

        shutil.copy2(LOGO_FALLBACK, LOGO_FILE)
    except OSError:
        pass


def resolve_logo_path() -> Path | None:
    ensure_logo()
    if LOGO_FILE.exists():
        return LOGO_FILE
    if LOGO_FALLBACK.exists():
        return LOGO_FALLBACK
    return None


def logo_image_data() -> tuple[bytes, str] | None:
    """Return logo file bytes and MIME type (no transparency processing)."""
    logo_path = resolve_logo_path()
    if not logo_path:
        return None
    suffix = logo_path.suffix.lower()
    mime = "image/jpeg" if suffix in {".jpg", ".jpeg"} else "image/png"
    return logo_path.read_bytes(), mime


def render_logo(context: str, tagline: str = "") -> None:
    """Show Sigma logo sized for sidebar, login, or welcome."""
    data = logo_image_data()
    if not data:
        return
    raw, mime = data

    encoded = base64.b64encode(raw).decode()
    img_html = (
        f'<img src="data:{mime};base64,{encoded}" '
        f'alt="Sigma Consultants" class="sigma-logo-img" />'
    )
    tagline_html = f'<p class="sidebar-brand-tagline">{tagline}</p>' if tagline else ""

    if context == "sidebar":
        st.markdown(
            f"""
<div class="sidebar-brand">
  <div class="logo-wrap logo-{context}">{img_html}</div>
  {tagline_html}
</div>
""",
            unsafe_allow_html=True,
        )
        return

    st.markdown(
        f'<div class="logo-wrap logo-{context}">{img_html}</div>',
        unsafe_allow_html=True,
    )


# =====================================================
# STYLE
# =====================================================

def inject_css() -> None:
    st.markdown(
        f"""
<style>
@import url("https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap");

:root {{
    --sigma-primary: {PRIMARY_COLOR};
    --sigma-secondary: {SECONDARY_COLOR};
    --sigma-accent: {ACCENT_COLOR};
    --sigma-bg: {BG_COLOR};
    --sigma-card: {CARD_BG};
    --sigma-text: {TEXT_COLOR};
    --sigma-muted: {MUTED_COLOR};
    --sigma-border: {BORDER_COLOR};
    --sigma-surface-alt: {SURFACE_ALT};
    --sigma-radius: 14px;
    --sigma-radius-lg: 18px;
    --sigma-shadow: 0 4px 14px rgba(15, 23, 42, 0.06);
    --sigma-shadow-md: 0 12px 32px rgba(15, 23, 42, 0.10);
    --sigma-fs-base: 16px;
    --sigma-fs-label: 15px;
    --sigma-fs-caption: 14px;
}}

html, body, [class*="css"], .stApp, button, input, textarea, select {{
    font-family: "Plus Jakarta Sans", Inter, "Segoe UI", system-ui, sans-serif !important;
    font-size: var(--sigma-fs-base);
}}

.stApp {{
    background: linear-gradient(165deg, #F8FAFC 0%, #EEF2F7 45%, #E8EDF5 100%);
    color: var(--sigma-text);
}}

[data-testid="stHeader"] {{
    background: transparent;
}}

div.block-container {{
    padding: 1.75rem 2.25rem 2.5rem;
    max-width: 1240px;
}}

.page-header {{
    margin: 0 0 1.5rem 0;
    padding: 1.25rem 1.5rem;
    background: linear-gradient(135deg, #FFFFFF 0%, #F8FAFC 100%);
    border: 1px solid var(--sigma-border);
    border-radius: var(--sigma-radius-lg);
    box-shadow: var(--sigma-shadow);
    border-left: 5px solid {SECONDARY_COLOR};
}}

.page-header h1 {{
    margin: 0;
    font-size: 1.875rem;
    line-height: 1.2;
    font-weight: 800;
    letter-spacing: -0.03em;
    color: var(--sigma-text);
}}

.page-subtitle {{
    margin: 0.5rem 0 0 0;
    font-size: 1rem;
    line-height: 1.55;
    color: var(--sigma-muted);
    font-weight: 400;
    max-width: 60ch;
}}

.section-title {{
    margin: 0 0 1rem 0;
    font-size: 1.125rem;
    font-weight: 700;
    color: var(--sigma-text);
    letter-spacing: -0.02em;
    padding-bottom: 0.35rem;
    border-bottom: 2px solid #E8EDF5;
}}

h1, h2, h3, h4 {{
    color: var(--sigma-text);
    font-weight: 700;
    letter-spacing: -0.02em;
}}

h1 {{ font-size: 1.875rem; }}
h2 {{ font-size: 1.125rem; }}
h3 {{ font-size: 1rem; }}

[data-testid="stWidgetLabel"] p,
label {{
    font-size: var(--sigma-fs-label) !important;
    font-weight: 600 !important;
    color: #334155 !important;
    margin-bottom: 0.35rem !important;
}}

.stCaption, [data-testid="stCaptionContainer"] {{
    font-size: var(--sigma-fs-caption) !important;
    color: var(--sigma-muted) !important;
    line-height: 1.5 !important;
}}

[data-testid="stVerticalBlock"] > div {{
    gap: 0.5rem;
}}

.stButton > button {{
    min-height: 46px;
    border-radius: 12px !important;
    font-size: var(--sigma-fs-label) !important;
    font-weight: 600 !important;
    border: 1px solid var(--sigma-border);
    color: var(--sigma-text);
    background: #ffffff;
    padding: 0.55rem 1.25rem !important;
    transition: all 0.18s ease;
    box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
}}

.stButton > button:hover {{
    background: var(--sigma-surface-alt);
    border-color: #B8C4D0;
    color: var(--sigma-text);
    transform: translateY(-1px);
}}

.stButton > button[kind="primary"],
button[kind="primary"],
[data-testid="stBaseButton-primary"],
[data-testid="stFormSubmitButton"] button {{
    background: linear-gradient(135deg, {SECONDARY_COLOR} 0%, {SECONDARY_DARK} 100%) !important;
    color: #ffffff !important;
    border: none !important;
    box-shadow: 0 6px 18px rgba(37, 99, 235, 0.28) !important;
}}

.stButton > button[kind="primary"]:hover,
button[kind="primary"]:hover,
[data-testid="stBaseButton-primary"]:hover,
[data-testid="stFormSubmitButton"] button:hover {{
    background: linear-gradient(135deg, #3B82F6 0%, {SECONDARY_COLOR} 100%) !important;
    box-shadow: 0 8px 22px rgba(37, 99, 235, 0.35) !important;
    color: #ffffff !important;
}}

[data-testid="stToolbar"],
[data-testid="stDecoration"],
.stDeployButton,
#MainMenu {{
    visibility: hidden;
    height: 0;
}}

[data-testid="stDataFrame"], .stDataFrame {{
    border: 1px solid var(--sigma-border);
    border-radius: var(--sigma-radius);
    overflow: hidden;
    background: #fff;
    font-size: var(--sigma-fs-label);
    box-shadow: var(--sigma-shadow);
}}

div[data-baseweb="select"] > div,
div[data-baseweb="input"] > div,
.stNumberInput input,
.stTextInput input,
.stDateInput input {{
    min-height: 46px;
    font-size: var(--sigma-fs-base) !important;
    border-radius: 12px !important;
    color: var(--sigma-text) !important;
}}

[data-testid="stForm"] {{
    background: #ffffff;
    border: 1px solid var(--sigma-border);
    border-radius: var(--sigma-radius-lg);
    padding: 1.5rem 1.6rem 1.25rem;
    box-shadow: var(--sigma-shadow-md);
}}

[data-testid="stRadio"] [role="radiogroup"] {{
    gap: 0.65rem;
    flex-wrap: wrap;
}}

[data-testid="stRadio"] label {{
    background: #ffffff;
    border: 1px solid var(--sigma-border);
    border-radius: 999px;
    padding: 0.5rem 1.1rem !important;
    font-size: var(--sigma-fs-caption) !important;
    font-weight: 600 !important;
    box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
}}

[data-testid="stRadio"] label[data-checked="true"] {{
    background: linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%) !important;
    border-color: #93C5FD !important;
    color: {SECONDARY_DARK} !important;
}}

div[data-testid="stVerticalBlockBorderWrapper"] {{
    background: linear-gradient(180deg, #FFFFFF 0%, #FAFBFD 100%);
    border: 1px solid var(--sigma-border);
    border-radius: var(--sigma-radius-lg);
    padding: 1.35rem 1.5rem 1.2rem;
    box-shadow: var(--sigma-shadow);
    margin-bottom: 1rem;
}}

.hero-card, .login-heading {{
    background: linear-gradient(180deg, #FFFFFF 0%, #F8FAFC 100%);
    border: 1px solid var(--sigma-border);
    border-radius: var(--sigma-radius-lg);
    padding: 2.25rem 2rem 2rem;
    text-align: center;
    box-shadow: var(--sigma-shadow-md);
}}

.hero-kicker, .login-heading .hero-kicker {{
    display: inline-block;
    font-size: 0.8125rem;
    font-weight: 700;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: {SECONDARY_COLOR};
    margin-bottom: 0.75rem;
    padding: 0.35rem 0.85rem;
    background: #EFF6FF;
    border-radius: 999px;
}}

.hero-card h1, .login-heading h1 {{
    margin: 0;
    font-size: 1.75rem;
    font-weight: 800;
    letter-spacing: -0.03em;
    color: var(--sigma-text);
}}

.hero-card p, .login-heading p {{
    margin: 0.75rem auto 0;
    max-width: 440px;
    font-size: 1.0625rem;
    line-height: 1.6;
    color: var(--sigma-muted);
}}

.welcome-actions [data-testid="stVerticalBlock"] {{
    gap: 0.65rem;
}}

.kpi-grid [data-testid="column"] {{
    display: flex;
}}

.kpi-grid [data-testid="column"] > div {{
    width: 100%;
}}

.kpi-card {{
    background: linear-gradient(180deg, #FFFFFF 0%, #FAFBFD 100%);
    border: 1px solid var(--sigma-border);
    border-radius: var(--sigma-radius);
    padding: 1.15rem 1.25rem;
    margin: 0;
    height: 100%;
    min-height: 108px;
    display: flex;
    flex-direction: column;
    justify-content: center;
    box-shadow: var(--sigma-shadow);
    position: relative;
    overflow: hidden;
}}

.kpi-card::before {{
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 4px;
    background: linear-gradient(90deg, {SECONDARY_COLOR}, #60A5FA);
}}

.kpi-card .kpi-label {{
    font-size: 0.8125rem;
    color: var(--sigma-muted);
    text-transform: uppercase;
    letter-spacing: 0.06em;
    font-weight: 700;
    line-height: 1.35;
}}

.kpi-card .kpi-value {{
    font-size: 1.75rem;
    font-weight: 800;
    color: var(--sigma-text);
    margin-top: 0.45rem;
    line-height: 1.15;
    letter-spacing: -0.03em;
}}

.summary-item {{
    background: #ffffff;
    border: 1px solid var(--sigma-border);
    border-radius: var(--sigma-radius);
    padding: 1.1rem 1.35rem;
    margin: 0 0 0.65rem 0;
    display: grid;
    grid-template-columns: 1fr auto;
    align-items: center;
    gap: 1.25rem;
    box-shadow: var(--sigma-shadow);
}}

.summary-item .left {{
    font-size: 1.0625rem;
    font-weight: 700;
    color: var(--sigma-text);
}}

.summary-item .meta {{
    font-size: var(--sigma-fs-caption);
    color: var(--sigma-muted);
    margin-top: 0.25rem;
    font-weight: 500;
}}

.summary-item .right {{
    text-align: right;
    font-size: var(--sigma-fs-caption);
    color: #334155;
    line-height: 1.55;
    white-space: nowrap;
    font-weight: 500;
}}

.profit-pos {{ color: {ACCENT_COLOR}; font-weight: 700; font-size: 1rem; }}
.profit-neg {{ color: #DC2626; font-weight: 700; font-size: 1rem; }}

.maturity-badge {{
    display: inline-block;
    padding: 0.2rem 0.55rem;
    border-radius: 999px;
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 0.02em;
}}
.badge-overdue {{ background: #FEE2E2; color: #B91C1C; }}
.badge-today {{ background: #FEF3C7; color: #B45309; }}
.badge-soon {{ background: #DBEAFE; color: #1D4ED8; }}
.badge-open {{ background: #D1FAE5; color: #047857; }}
.badge-closed {{ background: #E2E8F0; color: #475569; }}

/* ── Sidebar (boxed menu) ── */
section[data-testid="stSidebar"] {{
    background: #151C2B;
    border-right: 1px solid #2A3548;
}}

section[data-testid="stSidebar"] > div {{
    padding: 1.35rem 1rem 1.75rem;
    background: #151C2B;
}}

section[data-testid="stSidebar"] .sidebar-brand {{
    background: linear-gradient(145deg, #1E2A40 0%, #1A2234 100%);
    border: 1px solid #334155;
    border-radius: 14px;
    padding: 0.85rem 0.75rem 0.95rem;
    margin-bottom: 1rem;
    box-shadow: 0 4px 14px rgba(0, 0, 0, 0.18);
    text-align: center;
}}

section[data-testid="stSidebar"] .sidebar-brand-tagline {{
    margin: 0.55rem 0 0;
    color: #94A3B8 !important;
    font-size: 0.82rem;
    line-height: 1.4;
    font-weight: 500;
    text-align: center;
}}

.logo-wrap {{
    display: flex;
    justify-content: center;
    align-items: center;
    margin: 0 auto;
}}

.logo-login {{
    margin-bottom: 1rem;
    background: #ffffff;
    border-radius: 14px;
    padding: 0.75rem 1rem;
    box-shadow: 0 4px 14px rgba(15, 23, 42, 0.06);
}}

.logo-welcome {{
    margin-bottom: 1.15rem;
    background: #ffffff;
    border-radius: 14px;
    padding: 0.85rem 1.1rem;
    box-shadow: 0 4px 14px rgba(15, 23, 42, 0.06);
}}

.logo-sidebar {{
    margin-bottom: 0.35rem;
    background: #ffffff;
    border-radius: 10px;
    padding: 0.5rem 0.65rem;
}}

.sigma-logo-img {{
    display: block;
    width: 100%;
    height: auto;
    margin: 0 auto;
    object-fit: contain;
    border-radius: 6px;
}}

.logo-sidebar .sigma-logo-img {{
    max-width: 210px;
}}

.logo-login .sigma-logo-img {{
    max-width: 320px;
}}

.logo-welcome .sigma-logo-img {{
    max-width: 360px;
}}

section[data-testid="stSidebar"] .sidebar-toggle-box {{
    background: #1E2A40;
    border: 1px solid #334155;
    border-radius: 12px;
    padding: 0.65rem 0.85rem;
    margin-bottom: 1rem;
}}

section[data-testid="stSidebar"] .sidebar-section {{
    margin: 0.85rem 0 0.55rem;
    padding: 0 0.15rem;
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #64748B;
}}

section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] {{
    color: #CBD5E1 !important;
    font-size: 0.9rem !important;
    font-weight: 600 !important;
}}

section[data-testid="stSidebar"] [data-testid="stToggle"] label span {{
    color: #E2E8F0 !important;
    font-size: 0.9rem !important;
    font-weight: 600 !important;
}}

section[data-testid="stSidebar"] .sidebar-btn,
section[data-testid="stSidebar"] .sidebar-active,
section[data-testid="stSidebar"] .sidebar-util,
section[data-testid="stSidebar"] .sidebar-logout {{
    margin-bottom: 0.5rem;
}}

section[data-testid="stSidebar"] .stButton {{
    margin-bottom: 0;
}}

section[data-testid="stSidebar"] .sidebar-btn .stButton > button {{
    background: #1E2A40 !important;
    color: #E2E8F0 !important;
    border: 1px solid #3D4F68 !important;
    justify-content: flex-start;
    text-align: left;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    min-height: 50px;
    padding: 0.7rem 1rem !important;
    width: 100%;
    border-radius: 12px !important;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.12) !important;
    letter-spacing: 0.01em;
    transition: all 0.15s ease;
}}

section[data-testid="stSidebar"] .sidebar-btn .stButton > button p,
section[data-testid="stSidebar"] .sidebar-btn .stButton > button div {{
    font-size: 0.95rem !important;
    font-weight: 600 !important;
}}

section[data-testid="stSidebar"] .sidebar-btn .stButton > button:hover {{
    background: #253347 !important;
    color: #FFFFFF !important;
    border-color: #4B6A8F !important;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2) !important;
    transform: translateY(-1px);
}}

section[data-testid="stSidebar"] .sidebar-active .stButton > button {{
    background: linear-gradient(135deg, {SECONDARY_COLOR} 0%, {SECONDARY_DARK} 100%) !important;
    color: #FFFFFF !important;
    border: 1px solid #60A5FA !important;
    font-weight: 700 !important;
    min-height: 50px;
    border-radius: 12px !important;
    box-shadow: 0 6px 18px rgba(37, 99, 235, 0.4) !important;
}}

section[data-testid="stSidebar"] .sidebar-active .stButton > button p,
section[data-testid="stSidebar"] .sidebar-active .stButton > button div {{
    font-size: 0.95rem !important;
    font-weight: 700 !important;
    color: #FFFFFF !important;
}}

section[data-testid="stSidebar"] .sidebar-active .stButton > button:hover {{
    background: linear-gradient(135deg, #3B82F6 0%, {SECONDARY_COLOR} 100%) !important;
    color: #FFFFFF !important;
    transform: translateY(-1px);
}}

section[data-testid="stSidebar"] .sidebar-util .stButton > button {{
    background: #1E2A40 !important;
    color: #CBD5E1 !important;
    border: 1px dashed #475569 !important;
    font-size: 0.9rem !important;
    min-height: 46px;
    border-radius: 12px !important;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1) !important;
}}

section[data-testid="stSidebar"] .sidebar-util .stButton > button:hover {{
    background: #253347 !important;
    color: #FFFFFF !important;
    border-color: #64748B !important;
    border-style: solid !important;
}}

section[data-testid="stSidebar"] .sidebar-logout .stButton > button {{
    background: #2A1F28 !important;
    color: #FCA5A5 !important;
    border: 1px solid #7F3D45 !important;
    font-size: 0.9rem !important;
    min-height: 46px;
    border-radius: 12px !important;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.12) !important;
}}

section[data-testid="stSidebar"] .sidebar-logout .stButton > button:hover {{
    background: #3D2430 !important;
    color: #FECACA !important;
    border-color: #F87171 !important;
}}

section[data-testid="stSidebar"] hr {{
    margin: 1rem 0;
    border: none;
    border-top: 1px solid #2E3A50 !important;
}}

section[data-testid="stSidebar"] [data-testid="stAlert"] {{
    font-size: 0.875rem;
    border-radius: 10px;
}}

div[data-testid="stImage"] {{
    display: flex;
    justify-content: center;
}}

[data-testid="stHorizontalBlock"] {{
    gap: 1rem;
    align-items: flex-start;
}}

hr, [data-testid="stDivider"] {{
    margin: 1.25rem 0 !important;
    border: none;
    border-top: 1px solid var(--sigma-border) !important;
}}

[data-testid="stAlert"] {{
    font-size: var(--sigma-fs-caption);
    border-radius: 12px;
    padding: 0.85rem 1rem;
}}

@media (max-width: 768px) {{
    div.block-container {{
        padding: 1.25rem 1rem 1.75rem;
    }}
    .page-header {{ padding: 1rem 1.15rem; }}
    .page-header h1 {{ font-size: 1.5rem; }}
    .kpi-card .kpi-value {{ font-size: 1.5rem; }}
    .summary-item {{
        grid-template-columns: 1fr;
    }}
    .summary-item .right {{
        text-align: left;
        white-space: normal;
    }}
    .logo-login .sigma-logo-img {{ max-width: 260px; }}
    .logo-welcome .sigma-logo-img {{ max-width: 280px; }}
    .logo-sidebar .sigma-logo-img {{ max-width: 185px; }}
}}
</style>
""",
        unsafe_allow_html=True,
    )


def page_header(title: str, subtitle: str = "") -> None:
    subtitle_html = f'<p class="page-subtitle">{subtitle}</p>' if subtitle else ""
    st.markdown(
        f'<div class="page-header"><h1>{title}</h1>{subtitle_html}</div>',
        unsafe_allow_html=True,
    )


def section_title(title: str) -> None:
    st.markdown(f'<h2 class="section-title">{title}</h2>', unsafe_allow_html=True)


def card(title, value) -> None:
    st.markdown(
        f"""
<div class="kpi-card">
  <div class="kpi-label">{title}</div>
  <div class="kpi-value">{value}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def metrics_row(items: list[tuple[str, object]], mobile: bool) -> None:
    """Render aligned KPI cards in a responsive grid."""
    n = len(items)
    if mobile or n == 1:
        for title, value in items:
            card(title, value)
        return

    st.markdown('<div class="kpi-grid">', unsafe_allow_html=True)
    cols = st.columns(n, gap="small")
    for col, (title, value) in zip(cols, items):
        with col:
            card(title, value)
    st.markdown("</div>", unsafe_allow_html=True)


def money_cols(*names: str) -> dict:
    return {name: st.column_config.NumberColumn(format="₹ %.2f") for name in names}


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
        st.session_state.page = "Maturity"
    if "selected_proposal_id" not in st.session_state:
        st.session_state.selected_proposal_id = None
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
        render_logo("login")
        st.markdown(
            """
<div class="login-heading">
  <h1>Welcome back</h1>
  <p>Enter your password to access the CRM dashboard.</p>
</div>
""",
            unsafe_allow_html=True,
        )
        st.write("")
        with st.form("login_form"):
            pwd = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Continue", use_container_width=True, type="primary")
        if submitted:
            if pwd == expected_password():
                st.session_state.auth = True
                st.session_state.page = "Maturity"
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
        render_logo("sidebar", tagline="Client &amp; proposal management")

        st.markdown('<div class="sidebar-toggle-box">', unsafe_allow_html=True)
        st.session_state.is_mobile = st.toggle(
            "Compact card view",
            value=st.session_state.is_mobile,
            help="Stack metric cards for smaller screens.",
        )
        st.markdown("</div>", unsafe_allow_html=True)

        current = st.session_state.page

        def nav(icon: str, label: str, page_key: str) -> None:
            css = "sidebar-active" if current == page_key else "sidebar-btn"
            st.markdown(f'<div class="{css}">', unsafe_allow_html=True)
            if st.button(f"{icon}  {label}", use_container_width=True, key=f"nav_{page_key}"):
                go(page_key)
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<p class="sidebar-section">Daily use</p>', unsafe_allow_html=True)
        nav("📅", "Maturity board", "Maturity")
        nav("📁", "Proposal file", "ProposalDetail")
        nav("👤", "Client ledger", "ClientLedger")
        nav("🔍", "Search", "Search")

        st.markdown('<p class="sidebar-section">Operations</p>', unsafe_allow_html=True)
        nav("➕", "Add proposal", "AddProposal")
        nav("✏️", "Edit line", "Edit")
        nav("👥", "Clients", "Clients")

        st.markdown('<p class="sidebar-section">Reports</p>', unsafe_allow_html=True)
        nav("📊", "Summary", "Summary")
        nav("📥", "Export / backup", "Export")
        nav("📤", "Import Excel", "Import")

        st.markdown('<p class="sidebar-section">System</p>', unsafe_allow_html=True)
        st.markdown('<div class="sidebar-util">', unsafe_allow_html=True)
        if st.button("🔄  Reload Excel files", use_container_width=True, key="nav_reload"):
            st.session_state.clients_df = load_clients()
            st.session_state.proposals_df = load_proposals()
            st.success("Reloaded from disk")
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="sidebar-logout">', unsafe_allow_html=True)
        if st.button("🚪  Log out", use_container_width=True, key="nav_logout"):
            st.session_state.auth = False
            st.session_state.page = "Maturity"
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)


# =====================================================
# PAGES — PHASE 1 (maturity-centric)
# =====================================================

MATURITY_FILTERS = [
    "Due today",
    "Next 7 days",
    "Next 15 days",
    "Next 30 days",
    "Overdue",
    "All open",
]


def maturity_badge_html(bucket: str) -> str:
    css = {
        "Overdue": "badge-overdue",
        "Due today": "badge-today",
        "Due in 7 days": "badge-soon",
        "Due in 15 days": "badge-soon",
        "Due in 30 days": "badge-soon",
        "Later": "badge-open",
        "Closed": "badge-closed",
    }.get(bucket, "badge-open")
    return f'<span class="maturity-badge {css}">{bucket}</span>'


def page_maturity(is_mobile: bool) -> None:
    page_header(
        "Maturity board",
        "See what is ending today, soon, or overdue — your daily starting point.",
    )
    df = enrich_proposals(st.session_state.proposals_df)
    if df.empty:
        st.info("No proposals yet. Add your first proposal to start tracking maturities.")
        if st.button("Add proposal", type="primary"):
            go("AddProposal")
        return

    open_df = df[df["Status"] == "Open"]
    due_today = len(open_df[open_df["Days_Left"] == 0])
    due_7 = len(open_df[open_df["Days_Left"].between(0, 7)])
    overdue = len(open_df[open_df["Days_Left"] < 0])

    metrics_row(
        [
            ("Due today", due_today),
            ("Due in 7 days", due_7),
            ("Overdue", overdue),
            ("Open investment", fmt_money(open_df["Proposal_Cost"].sum())),
        ],
        is_mobile,
    )

    st.markdown("---")
    bucket = st.radio(
        "Show proposals",
        MATURITY_FILTERS,
        horizontal=True,
        key="maturity_filter",
    )
    filtered = filter_by_maturity_bucket(df, bucket)
    if filtered.empty:
        st.success(f"No proposals in “{bucket}”.")
        return

    display = proposal_display_table(filtered)
    table_cols = [
        "Proposal_ID",
        "Client_Name",
        "End_Date",
        "Days_Left",
        "Proposal_Cost",
        "Rate",
        "Profit",
        "Final_Cost",
        "Maturity_Bucket",
        "Status",
    ]
    st.dataframe(
        display[table_cols],
        use_container_width=True,
        hide_index=True,
        column_config={
            **money_cols("Proposal_Cost", "Final_Cost", "Profit"),
            "Days_Left": st.column_config.NumberColumn("Days left", format="%d"),
            "Rate": st.column_config.NumberColumn("Rate %", format="%.2f"),
        },
    )

    st.markdown("---")
    section_title("Open proposal details")
    pid = st.selectbox(
        "Select proposal ID",
        unique_sorted(filtered["Proposal_ID"]),
        key="maturity_pick_proposal",
    )
    if st.button("View full proposal", type="primary", key="maturity_view_btn"):
        view_proposal(pid)


def page_proposal_detail(is_mobile: bool) -> None:
    page_header(
        "Proposal file",
        "One proposal — all clients, totals, close or renew.",
    )
    df = enrich_proposals(st.session_state.proposals_df)
    if df.empty:
        st.info("No proposals available")
        return

    ids = unique_sorted(df["Proposal_ID"])
    default_idx = 0
    if st.session_state.selected_proposal_id in ids:
        default_idx = ids.index(st.session_state.selected_proposal_id)

    proposal_id = st.selectbox(
        "Proposal ID",
        ids,
        index=default_idx,
        key="detail_proposal_id",
    )
    st.session_state.selected_proposal_id = proposal_id
    proposal_df = df[df["Proposal_ID"] == proposal_id].copy()
    if proposal_df.empty:
        st.warning("Proposal not found")
        return

    start = proposal_df["Start_Date"].iloc[0]
    end = proposal_df["End_Date"].iloc[0]
    rate = proposal_df["Rate"].iloc[0]
    status = proposal_df["Status"].iloc[0]
    days_left = proposal_df["Days_Left"].iloc[0]
    bucket = proposal_df["Maturity_Bucket"].iloc[0]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Start", fmt_date(start))
    c2.metric("End", fmt_date(end))
    c3.metric("Rate", f"{float(rate):.2f}%")
    c4.metric("Days left", int(days_left) if pd.notna(days_left) else "—")

    st.markdown(maturity_badge_html(bucket), unsafe_allow_html=True)

    render_grand_total(
        proposal_df["Proposal_Cost"].sum(),
        proposal_df["Final_Cost"].sum(),
        proposal_df["Profit"].sum(),
        is_mobile,
    )

    section_title("Clients in this proposal")
    display = proposal_display_table(proposal_df)
    st.dataframe(
        display[
            [
                "Client_Name",
                "Proposal_Cost",
                "Rate",
                "Profit",
                "Final_Cost",
                "Status",
                "Days_Left",
            ]
        ],
        use_container_width=True,
        hide_index=True,
        column_config={
            **money_cols("Proposal_Cost", "Final_Cost", "Profit"),
            "Days_Left": st.column_config.NumberColumn("Days left", format="%d"),
        },
    )

    st.markdown("---")
    section_title("Actions")
    if status == "Open":
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Close entire proposal", type="primary", key="close_whole_proposal"):
                close_proposal(proposal_id)
                st.success(f"Closed {proposal_id}")
                st.rerun()
        with c2:
            if st.button("Go to edit line", key="detail_go_edit"):
                go("Edit")

        with st.container(border=True):
            section_title("Renew proposal")
            new_start = st.date_input("New start date", value=date.today(), key="renew_start")
            new_days = st.selectbox("Duration (days)", DURATION_DAYS, key="renew_days")
            if st.button("Create renewed proposal", key="renew_btn"):
                new_id = renew_proposal(proposal_id, new_start, int(new_days))
                st.success(f"Created {new_id} from {proposal_id}")
                view_proposal(new_id)
    else:
        st.info("This proposal is closed. Use Renew to create a new open proposal with the same clients.")


def page_client_ledger(is_mobile: bool) -> None:
    page_header(
        "Client ledger",
        "Pick a client — see every proposal, amount, rate, end date, and profit.",
    )
    clients = active_clients_df()
    if clients.empty:
        st.info("No active clients")
        return

    df = enrich_proposals(st.session_state.proposals_df)
    client = st.selectbox(
        "Client",
        unique_sorted(clients["Client_Name"]),
        key="ledger_client",
    )
    client_id = clients[clients["Client_Name"] == client]["Client_ID"].iloc[0]
    client_df = df[df["Client_ID"] == client_id].copy()
    if client_df.empty:
        st.info("No proposals for this client")
        return

    status = st.radio("Status", ["All", "Open", "Close"], horizontal=True, key="ledger_status")
    if status != "All":
        client_df = client_df[client_df["Status"] == status]
    if client_df.empty:
        st.info("No records for this filter")
        return

    open_rows = client_df[client_df["Status"] == "Open"]
    metrics_row(
        [
            ("Open investment", fmt_money(open_rows["Proposal_Cost"].sum())),
            ("Expected profit", fmt_money(open_rows["Profit"].sum())),
            ("Open proposals", len(open_rows)),
            ("Due in 7 days", int(open_rows["Days_Left"].between(0, 7).sum())),
        ],
        is_mobile,
    )

    display = proposal_display_table(client_df.sort_values(["End_Date", "Proposal_ID"]))
    st.dataframe(
        display[
            [
                "Proposal_ID",
                "Start_Date",
                "End_Date",
                "Days_Left",
                "Proposal_Cost",
                "Rate",
                "Profit",
                "Final_Cost",
                "Status",
                "Maturity_Bucket",
            ]
        ],
        use_container_width=True,
        hide_index=True,
        column_config={
            **money_cols("Proposal_Cost", "Final_Cost", "Profit"),
            "Days_Left": st.column_config.NumberColumn("Days left", format="%d"),
        },
    )

    pid = st.selectbox("Open proposal file", unique_sorted(client_df["Proposal_ID"]), key="ledger_open")
    if st.button("View proposal", key="ledger_view"):
        view_proposal(pid)


def page_search(is_mobile: bool) -> None:
    page_header("Search", "One search box — client name, proposal ID, or date (dd-mm-yyyy).")
    df = enrich_proposals(st.session_state.proposals_df)
    if df.empty:
        st.warning("No data to search")
        return

    query = st.text_input("Search", placeholder="e.g. ABC, SIG-P-001, 18-09-2026", key="global_search")
    status = st.selectbox("Status", ["All", "Open", "Close"], key="search_status")

    result = df.copy()
    if status != "All":
        result = result[result["Status"] == status]

    if query.strip():
        q = query.strip().lower()
        mask = (
            result["Client_Name"].str.lower().str.contains(q, na=False)
            | result["Proposal_ID"].str.lower().str.contains(q, na=False)
            | result["End_Date"].map(fmt_date).str.contains(q, na=False)
            | result["Start_Date"].map(fmt_date).str.contains(q, na=False)
        )
        result = result[mask]

    if result.empty:
        st.info("No matching records")
        return

    render_grand_total(
        result["Proposal_Cost"].sum(),
        result["Final_Cost"].sum(),
        result["Profit"].sum(),
        is_mobile,
    )

    display = proposal_display_table(result.sort_values(["End_Date", "Proposal_ID", "Client_Name"]))
    st.dataframe(
        display[
            [
                "Proposal_ID",
                "Client_Name",
                "Start_Date",
                "End_Date",
                "Days_Left",
                "Proposal_Cost",
                "Rate",
                "Profit",
                "Final_Cost",
                "Status",
            ]
        ],
        use_container_width=True,
        hide_index=True,
        column_config={
            **money_cols("Proposal_Cost", "Final_Cost", "Profit"),
            "Days_Left": st.column_config.NumberColumn("Days left", format="%d"),
        },
    )

    pid = st.selectbox("Open proposal", unique_sorted(result["Proposal_ID"]), key="search_open")
    if st.button("View proposal file", key="search_view"):
        view_proposal(pid)


def page_welcome() -> None:
    _, center, _ = st.columns([1, 2, 1])
    with center:
        render_logo("welcome")
        st.markdown(
            """
<div class="hero-card">
  <h1>Welcome to your CRM</h1>
  <p>Manage clients, proposals, investments, profits, and maturity dates — all in one place.</p>
</div>
""",
            unsafe_allow_html=True,
        )
        st.markdown('<div class="welcome-actions">', unsafe_allow_html=True)
        if st.button("➕  Add client", use_container_width=True, type="primary"):
            go("Clients")
        if st.button("📄  Add proposal", use_container_width=True):
            go("AddProposal")
        if st.button("🔍  Find details", use_container_width=True):
            go("Find")
        if st.button("📊  View summary", use_container_width=True):
            go("Summary")
        st.markdown("</div>", unsafe_allow_html=True)


def page_summary(is_mobile: bool) -> None:
    page_header("Summary", "Investment, profit, and date-wise proposal totals.")
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
    section_title("Date-based summary")

    with st.container(border=True):
        status_options = ["All"] + unique_sorted(df["Status"])
        c1, c2 = st.columns(2)
        with c1:
            selected_status = st.selectbox("Proposal status", status_options, key="summary_status")
        with c2:
            date_type = st.radio("Date type", ["Start Date", "End Date"], horizontal=True, key="summary_date_type")

    if selected_status != "All":
        df = df[df["Status"] == selected_status]
    if df.empty:
        st.warning("No data for selected status")
        return

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

    section_title("Filtered totals")
    render_grand_total(df["Proposal_Cost"].sum(), df["Final_Cost"].sum(), df["Profit"].sum(), is_mobile)
    st.markdown("---")
    section_title("By rate and date")

    summary_df = (
        df.groupby(["Rate_Int", "DateOnly"], as_index=False)
        .agg({"Proposal_Cost": "sum", "Final_Cost": "sum", "Profit": "sum"})
        .sort_values(["Rate_Int", "DateOnly"])
    )
    for _, row in summary_df.iterrows():
        profit_cls = "profit-pos" if row["Profit"] >= 0 else "profit-neg"
        st.markdown(
            f"""
<div class="summary-item">
  <div>
    <div class="left">{row['DateOnly'].strftime('%d-%m-%Y')}</div>
    <div class="meta">Rate {int(row['Rate_Int'])}%</div>
  </div>
  <div class="right">
    Investment {fmt_money(row['Proposal_Cost'])}<br>
    Final {fmt_money(row['Final_Cost'])}<br>
    <span class="{profit_cls}">Profit {fmt_money(row['Profit'])}</span>
  </div>
</div>
""",
            unsafe_allow_html=True,
        )


def page_add_proposal() -> None:
    page_header("Add proposal", "Add one or more clients to a single proposal, then save.")
    active = active_clients_df()
    if active.empty:
        st.warning("Please add an active client first.")
        if st.button("Go to Clients"):
            go("Clients")
        return

    with st.container(border=True):
        section_title("Proposal details")
        c1, c2, c3 = st.columns(3)
        with c1:
            start_date = st.date_input("Start date", value=date.today(), key="add_start")
        with c2:
            days = st.selectbox("Duration (days)", DURATION_DAYS, key="add_days")
        with c3:
            rate = st.number_input("Monthly rate (%)", min_value=0.0, step=0.25, key="add_rate")
        end_date = (pd.Timestamp(start_date) + pd.Timedelta(days=int(days))).date()
        st.caption(f"End date: **{end_date.strftime('%d-%m-%Y')}**")

    with st.container(border=True):
        section_title("Add client to proposal")
        c1, c2 = st.columns(2)
        names = unique_sorted(active["Client_Name"])
        with c1:
            selected_client = st.selectbox("Client name", ["Select Client"] + names, key="add_client")
        with c2:
            principal = st.number_input("Proposal amount (₹)", min_value=0.0, step=1000.0, key="add_principal")

    c1, c2 = st.columns(2)
    with c1:
        add_clicked = st.button("Add to proposal", use_container_width=True)
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
        preview = pd.DataFrame(st.session_state.proposal_clients)
        st.dataframe(
            preview,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Principal": st.column_config.NumberColumn("Principal", format="₹ %.2f"),
                "Profit": st.column_config.NumberColumn("Profit", format="₹ %.2f"),
                "Final_Amount": st.column_config.NumberColumn("Final amount", format="₹ %.2f"),
            },
        )

    if st.button("Save proposal", type="primary"):
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
        view_proposal(proposal_id)


def page_edit() -> None:
    page_header("Edit proposal", "Update amount, rate, dates, or close a proposal.")
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

    with st.container(border=True):
        c1, c2 = st.columns(2)
        with c1:
            proposal_cost = st.number_input(
                "Proposal amount", min_value=0.0, value=float(row["Proposal_Cost"]), key="edit_cost"
            )
            rate = st.number_input("Rate (%)", min_value=0.0, value=float(row["Rate"]), key="edit_rate")
        with c2:
            new_start = st.date_input("Start date", value=to_date(row["Start_Date"]), key="edit_start")
            new_end = st.date_input("End date", value=to_date(row["End_Date"], fallback=new_start), key="edit_end")
            new_status = st.selectbox(
                "Status", ["Open", "Close"], index=0 if row["Status"] == "Open" else 1, key="edit_new_status"
            )

    days = (new_end - new_start).days
    if days < 0:
        st.error("End date cannot be before start date")
        return

    final_cost, profit = calc(proposal_cost, rate, days)
    st.caption(f"Duration: {days} days · Final: {fmt_money(final_cost)} · Profit: {fmt_money(profit)}")

    if st.button("Save changes", type="primary"):
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
        view_proposal(proposal_id)


def render_by_proposal(df_master: pd.DataFrame, is_mobile: bool) -> None:
    section_title("By proposal")
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
        column_config=money_cols("Proposal_Cost", "Final_Cost", "Profit"),
    )


def render_by_client(df_master: pd.DataFrame, is_mobile: bool) -> None:
    section_title("By client name")
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
        column_config=money_cols("Proposal_Cost", "Final_Cost", "Profit"),
    )


def render_by_date(df_master: pd.DataFrame, is_mobile: bool) -> None:
    section_title("By start or end date")
    status = st.selectbox("Status", ["All", "Open", "Close"], key="date_status")
    df = df_master if status == "All" else df_master[df_master["Status"] == status]
    if df.empty:
        st.info("No proposals found")
        return

    if st.button("Add another date"):
        st.session_state.date_blocks.append(max(st.session_state.date_blocks) + 1)
        st.rerun()

    st.divider()
    for i, block_id in enumerate(list(st.session_state.date_blocks)):
        col1, col2 = st.columns([6, 1])
        with col1:
            st.markdown(f'<h2 class="section-title">Date filter {block_id}</h2>', unsafe_allow_html=True)
        with col2:
            if i != 0 and st.button("Remove", key=f"remove_{block_id}"):
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
            column_config=money_cols("Proposal_Cost", "Final_Cost", "Profit"),
        )
        st.divider()


def page_find(is_mobile: bool) -> None:
    page_header("Find details", "Look up proposals by ID, client, or date.")
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
    page_header("Clients", "Add, archive, and restore client records.")
    with st.container(border=True):
        section_title("New client")
        c1, c2 = st.columns(2)
        with c1:
            cname = st.text_input("Client name", key="new_client_name")
        with c2:
            notes = st.text_input("Notes (optional)", key="new_client_notes")

    if st.button("Add client", type="primary"):
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

    section_title("Directory")
    show_archived = st.toggle("Show archived clients", value=False)
    view = st.session_state.clients_df.copy()
    if not show_archived:
        view = view.loc[~view["Is_Archived"]]

    if view.empty:
        st.info("No clients to display")
        return

    st.dataframe(view, use_container_width=True, hide_index=True)

    with st.container(border=True):
        section_title("Archive or restore")
        target = st.selectbox("Client", unique_sorted(st.session_state.clients_df["Client_Name"]))
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("Archive", use_container_width=True):
                mask = st.session_state.clients_df["Client_Name"] == target
                st.session_state.clients_df.loc[mask, "Is_Archived"] = True
                save_clients()
                st.rerun()
        with col_b:
            if st.button("Restore", use_container_width=True):
                mask = st.session_state.clients_df["Client_Name"] == target
                st.session_state.clients_df.loc[mask, "Is_Archived"] = False
                save_clients()
                st.rerun()


def page_client_dashboard(is_mobile: bool) -> None:
    page_header("Client dashboard", "Single-client detail or all-client totals.")
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
        section_title("Proposal details")
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
            column_config=money_cols("Proposal_Cost", "Final_Cost", "Profit"),
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
    st.dataframe(
        summary_df,
        use_container_width=True,
        hide_index=True,
        column_config=money_cols(
            "Total of Proposal price",
            "Total of Profit amount",
            "Total of Proposal Price and Profit amount",
        ),
    )


def build_excel_bytes(clients: pd.DataFrame, proposals: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if not clients.empty:
            clients.to_excel(writer, index=False, sheet_name="Clients")
        if not proposals.empty:
            proposals.to_excel(writer, index=False, sheet_name="Proposals")
    return output.getvalue()


def page_import() -> None:
    page_header(
        "Import Excel",
        "Upload raw proposal data — Client ID and Proposal ID are created automatically.",
    )

    st.markdown(
        """
Your file can use columns like **Start Date**, **Client Name**, **Initial Amount**,
**Duration**, **End Date**, **Profit Rate**, **Final Amount**, **Profit** (as in `dATa.xlsx`).

**Proposal ID rule:** rows with the same **start date + end date + rate** become **one proposal**
with multiple clients. Everything else is filled in automatically.
"""
    )

    with st.container(border=True):
        section_title("Choose file")
        uploaded = st.file_uploader(
            "Upload Excel (.xlsx)",
            type=["xlsx", "xls"],
            key="import_upload",
        )
        if DEFAULT_RAW_IMPORT_PATH.exists():
            if st.button(f"Load desktop file: {DEFAULT_RAW_IMPORT_PATH.name}", key="load_desktop"):
                st.session_state.import_preview_df = pd.read_excel(DEFAULT_RAW_IMPORT_PATH)
                st.rerun()

    preview_source = None
    if uploaded is not None:
        preview_source = pd.read_excel(uploaded)
    elif "import_preview_df" in st.session_state:
        preview_source = st.session_state.import_preview_df

    if preview_source is None:
        st.info("Upload your Excel file or load the desktop `dATa.xlsx` file.")
        return

    st.caption(f"**{len(preview_source)}** rows · columns: {', '.join(map(str, preview_source.columns))}")
    st.dataframe(preview_source.head(20), use_container_width=True, hide_index=True)

    with st.container(border=True):
        section_title("Import options")
        append_mode = st.radio(
            "How to import",
            ["Add to existing data (recommended)", "Replace all existing data"],
            key="import_mode",
        )
        group_terms = st.checkbox(
            "Group rows into one proposal when start date, end date, and rate match",
            value=True,
            key="import_group",
        )
        if append_mode.startswith("Replace"):
            st.warning("Replace will remove current clients and proposals after you confirm.")

    try:
        prepared = prepare_raw_import_df(preview_source)
        st.success(
            f"Ready: **{len(prepared)}** valid rows · "
            f"**{prepared['Client_Name'].nunique()}** clients · "
            f"preview proposals: **{prepared.groupby(['Start_Date', 'End_Date', 'Rate']).ngroups if group_terms else len(prepared)}**"
        )
    except ValueError as exc:
        st.error(str(exc))
        st.caption("Rename your columns to match: Start Date, Client Name, Initial Amount, End Date, Profit Rate, etc.")
        return

    if st.button("Import into CRM", type="primary", key="run_import"):
        append = append_mode.startswith("Add")
        if not append:
            st.session_state.clients_df = pd.DataFrame(columns=CLIENT_COLS)
            st.session_state.proposals_df = pd.DataFrame(columns=PROPOSAL_COLS)

        clients_df, proposals_df, stats = import_raw_excel(
            preview_source,
            st.session_state.clients_df,
            st.session_state.proposals_df,
            append=append,
            group_by_terms=group_terms,
        )
        st.session_state.clients_df = clients_df
        st.session_state.proposals_df = proposals_df
        save_clients()
        save_proposals()
        if "import_preview_df" in st.session_state:
            del st.session_state.import_preview_df
        st.success(
            f"Imported **{stats['lines']}** lines · **{stats['proposals']}** proposals · "
            f"**{stats['clients_new']}** new clients"
        )
        go("Maturity")


def page_export() -> None:
    page_header("Export & backup", "Download data or find auto-backups saved on every change.")
    proposals_df = st.session_state.proposals_df.copy()
    clients_df = st.session_state.clients_df.copy()
    if proposals_df.empty and clients_df.empty:
        st.warning("No data available to export.")
        return

    if BACKUP_DIR.exists():
        backups = sorted(BACKUP_DIR.glob("*.xlsx"), reverse=True)
        if backups:
            st.caption(f"Auto-backup folder: `{BACKUP_DIR}` — latest: **{backups[0].name}**")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    excel_bytes = build_excel_bytes(clients_df, proposals_df)

    with st.container(border=True):
        section_title("Excel export")
        st.download_button(
            label="Download Excel (clients + proposals)",
            data=excel_bytes,
            file_name=f"sigma_consultants_data_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="export_xlsx",
        )

    with st.container(border=True):
        section_title("CSV backup")
        if not proposals_df.empty:
            st.download_button(
                label="Download proposals CSV",
                data=proposals_df.to_csv(index=False, encoding="utf-8-sig"),
                file_name=f"sigma_proposals_{timestamp}.csv",
                mime="text/csv",
                key="export_proposals_csv",
            )
        if not clients_df.empty:
            st.download_button(
                label="Download clients CSV",
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
if page == "Maturity":
    page_maturity(is_mobile)
elif page == "ProposalDetail":
    page_proposal_detail(is_mobile)
elif page == "ClientLedger":
    page_client_ledger(is_mobile)
elif page == "Search":
    page_search(is_mobile)
elif page == "Welcome":
    page_welcome()
elif page == "Summary":
    page_summary(is_mobile)
elif page == "AddProposal":
    page_add_proposal()
elif page == "Edit":
    page_edit()
elif page == "Find":
    page_search(is_mobile)
elif page == "Clients":
    page_clients()
elif page == "ClientDashboard":
    page_client_ledger(is_mobile)
elif page == "Export":
    page_export()
elif page == "Import":
    page_import()
else:
    st.session_state.page = "Maturity"
    st.rerun()
