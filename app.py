# =====================================================
# Sigma Consultants CRM — clean rebuild
# Run: streamlit run sigma_consultants_crm.py
# Data: clients.xlsx + proposals.xlsx (same folder as this file)
# =====================================================
from __future__ import annotations

import base64
import io
import json
import os
import re
import shutil
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# ── Paths (always next to this script) ──────────────────────────────────────

BASE_DIR = Path(__file__).resolve().parent
CLIENT_FILE = BASE_DIR / "clients.xlsx"
PROPOSAL_FILE = BASE_DIR / "proposals.xlsx"
BACKUP_DIR = BASE_DIR / "backups"
MANIFEST_FILE = BASE_DIR / "data_manifest.json"
LOGO_FILE = BASE_DIR / "sigma_logo.jpg"
LOGO_FALLBACK = Path(
    r"C:\Users\Smart\.cursor\projects\empty-window\assets"
    r"\c__Users_Smart_AppData_Roaming_Cursor_User_workspaceStorage_empty-window_images_"
    r"Sigma_Con_Logo-e928b1a9-6cdc-44e2-837a-3d8331acffba.jpg"
)
DEFAULT_IMPORT = Path(r"c:\Users\Smart\Desktop\dATa.xlsx")

CLIENT_ID_RE = re.compile(r"SIG-C-(\d+)", re.I)
PROPOSAL_ID_RE = re.compile(r"SIG-P-(\d+)", re.I)

CLIENT_COLS = ["Client_ID", "Client_Name", "Created_Date", "Is_Archived", "Notes"]
PROPOSAL_COLS = [
    "Proposal_ID", "Client_ID", "Client_Name", "Proposal_Cost", "Rate",
    "Final_Cost", "Profit", "Start_Date", "End_Date", "Status", "Closing_Date",
]
DATE_COLS = ["Start_Date", "End_Date", "Closing_Date"]
MONEY_COLS = ["Proposal_Cost", "Rate", "Final_Cost", "Profit"]
DURATION_DAYS = [15, 20, 30, 45, 60, 90]

C_PRIMARY = "#0F172A"
C_ACCENT = "#2563EB"
C_ACCENT_DARK = "#1D4ED8"
C_SUCCESS = "#059669"
C_BG = "#EEF2F7"
C_CARD = "#FFFFFF"
C_MUTED = "#64748B"
C_BORDER = "#E2E8F0"

PAGES = {
    "Maturity": "📅  Maturity board",
    "ProposalDetail": "📁  Proposal file",
    "ClientLedger": "👤  Client ledger",
    "Search": "🔍  Search",
    "AddProposal": "➕  Add proposal",
    "Edit": "✏️  Edit line",
    "Clients": "👥  Clients",
    "Summary": "📊  Summary",
    "Import": "📤  Import Excel",
    "Export": "📥  Export / backup",
}

NAV_GROUPS = [
    ("Daily use", ["Maturity", "ProposalDetail", "ClientLedger", "Search"]),
    ("Operations", ["AddProposal", "Edit", "Clients"]),
    ("Reports", ["Summary", "Import", "Export"]),
]

MATURITY_FILTERS = [
    "Due today", "Next 7 days", "Next 15 days", "Next 30 days",
    "Overdue", "All open", "All data (incl. closed)",
]

RAW_ALIASES = {
    "start date": "Start_Date", "start_date": "Start_Date",
    "client name": "Client_Name", "client_name": "Client_Name",
    "initial amount": "Proposal_Cost", "proposal amount": "Proposal_Cost",
    "amount": "Proposal_Cost", "principal": "Proposal_Cost",
    "duration": "Duration", "end date": "End_Date", "end_date": "End_Date",
    "profit rate": "Rate", "rate": "Rate", "rate %": "Rate",
    "final amount": "Final_Cost", "final_amount": "Final_Cost",
    "profit": "Profit", "status": "Status",
}

st.set_page_config(
    page_title="Sigma Consultants CRM",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="auto",
)


# ── Utilities ───────────────────────────────────────────────────────────────

def normalize_status(v) -> str:
    if pd.isna(v):
        return "Open"
    s = str(v).strip().lower()
    if s in {"close", "closed", "done", "complete", "completed"}:
        return "Close"
    if s in {"open", "active"}:
        return "Open"
    return str(v).strip().title()


def as_bool(v) -> bool:
    if isinstance(v, bool):
        return v
    if pd.isna(v):
        return False
    return str(v).strip().lower() in {"true", "1", "yes", "y"}


def fmt_money(v) -> str:
    return f"₹ {float(v):,.2f}"


def fmt_date(v) -> str:
    if pd.isna(v):
        return "—"
    ts = pd.to_datetime(v, errors="coerce")
    return "—" if pd.isna(ts) else ts.strftime("%d-%m-%Y")


def to_date(v, fallback: date | None = None) -> date:
    ts = pd.to_datetime(v, errors="coerce")
    return ts.date() if pd.notna(ts) else (fallback or date.today())


def calc(principal: float, rate: float, days: int) -> tuple[float, float]:
    days = max(int(days), 0)
    profit = principal * rate * (days / 30.0) / 100.0
    return principal + profit, profit


def next_id(series: pd.Series, prefix: str, pat: re.Pattern) -> str:
    nums = [int(m.group(1)) for x in series.dropna().astype(str) if (m := pat.search(x.strip()))]
    return f"{prefix}{(max(nums) + 1) if nums else 1:03d}"


def unique_sorted(s: pd.Series) -> list:
    return sorted(s.dropna().unique().tolist())


def money_cfg(*cols):
    return {c: st.column_config.NumberColumn(format="₹ %.2f") for c in cols}


def go(page: str) -> None:
    st.session_state.page = page
    st.rerun()


# ── Data I/O ────────────────────────────────────────────────────────────────

def _ensure(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df[cols].copy()


def _save_excel(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    df.to_excel(tmp, index=False, engine="openpyxl")
    if path.exists():
        path.unlink()
    os.replace(tmp, path)
    n = len(pd.read_excel(path))
    if n != len(df):
        raise OSError(f"{path.name}: saved {n} rows, expected {len(df)}")


def load_clients() -> pd.DataFrame:
    df = pd.read_excel(CLIENT_FILE) if CLIENT_FILE.exists() else pd.DataFrame(columns=CLIENT_COLS)
    if not CLIENT_FILE.exists():
        _save_excel(df, CLIENT_FILE)
    df = _ensure(df, CLIENT_COLS)
    df["Client_ID"] = df["Client_ID"].astype(str).replace({"nan": "", "<NA>": ""})
    df["Client_Name"] = df["Client_Name"].fillna("").astype(str).str.strip()
    df["Notes"] = df["Notes"].fillna("").astype(str)
    df["Is_Archived"] = df["Is_Archived"].map(as_bool)
    df["Created_Date"] = pd.to_datetime(df["Created_Date"], errors="coerce")
    return df


def load_proposals() -> pd.DataFrame:
    df = pd.read_excel(PROPOSAL_FILE) if PROPOSAL_FILE.exists() else pd.DataFrame(columns=PROPOSAL_COLS)
    if not PROPOSAL_FILE.exists():
        _save_excel(df, PROPOSAL_FILE)
    df = _ensure(df, PROPOSAL_COLS)
    for c in MONEY_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    for c in DATE_COLS:
        df[c] = pd.to_datetime(df[c], errors="coerce")
    df["Status"] = df["Status"].apply(normalize_status)
    df["Client_Name"] = df["Client_Name"].fillna("").astype(str).str.strip()
    df["Proposal_ID"] = df["Proposal_ID"].astype(str)
    df["Client_ID"] = df["Client_ID"].astype(str)
    return df


def load_all_from_disk() -> None:
    """Always read Excel files from disk — single source of truth."""
    st.session_state.clients_df = load_clients()
    st.session_state.proposals_df = load_proposals()
    st.session_state._data_mtime = _disk_mtime()
    MANIFEST_FILE.write_text(
        json.dumps({
            "clients": len(st.session_state.clients_df),
            "proposals": len(st.session_state.proposals_df),
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "folder": str(BASE_DIR),
        }, indent=2),
        encoding="utf-8",
    )


def _disk_mtime() -> float:
    mt = [p.stat().st_mtime for p in (CLIENT_FILE, PROPOSAL_FILE) if p.exists()]
    return max(mt) if mt else 0.0


def _backup() -> None:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    for src, name in ((CLIENT_FILE, "clients"), (PROPOSAL_FILE, "proposals")):
        if src.exists():
            shutil.copy2(src, BACKUP_DIR / f"{name}_{ts}.xlsx")


def save_clients() -> None:
    _save_excel(st.session_state.clients_df, CLIENT_FILE)
    st.session_state._data_mtime = _disk_mtime()
    _backup()


def save_proposals() -> None:
    _save_excel(st.session_state.proposals_df, PROPOSAL_FILE)
    st.session_state._data_mtime = _disk_mtime()
    _backup()


def new_client_id() -> str:
    return next_id(st.session_state.clients_df["Client_ID"], "SIG-C-", CLIENT_ID_RE)


def new_proposal_id() -> str:
    return next_id(st.session_state.proposals_df["Proposal_ID"], "SIG-P-", PROPOSAL_ID_RE)


def active_clients() -> pd.DataFrame:
    return st.session_state.clients_df.loc[~st.session_state.clients_df["Is_Archived"]].copy()


# ── Business logic ──────────────────────────────────────────────────────────

def enrich(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in DATE_COLS:
        out[c] = pd.to_datetime(out[c], errors="coerce")
    out["Status"] = out["Status"].apply(normalize_status)
    today = date.today()

    def days_left(end):
        return None if pd.isna(end) else (pd.Timestamp(end).date() - today).days

    def bucket(row):
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
    out["Maturity_Bucket"] = out.apply(bucket, axis=1)
    return out


def filter_maturity(df: pd.DataFrame, label: str) -> pd.DataFrame:
    open_df = df[df["Status"] == "Open"].copy()
    if label == "All open":
        return open_df.sort_values(["End_Date", "Proposal_ID", "Client_Name"])
    if label == "All data (incl. closed)":
        return df.sort_values(["End_Date", "Proposal_ID", "Client_Name"])
    if label == "Overdue":
        return open_df[open_df["Days_Left"] < 0].sort_values("End_Date")
    if label == "Due today":
        return open_df[open_df["Days_Left"] == 0].sort_values("Proposal_ID")
    if label == "Next 7 days":
        return open_df[open_df["Days_Left"].between(0, 7)].sort_values("End_Date")
    if label == "Next 15 days":
        return open_df[open_df["Days_Left"].between(0, 15)].sort_values("End_Date")
    if label == "Next 30 days":
        return open_df[open_df["Days_Left"].between(0, 30)].sort_values("End_Date")
    return open_df.sort_values("End_Date")


def display_table(df: pd.DataFrame) -> pd.DataFrame:
    show = df.copy()
    show["Start_Date"] = show["Start_Date"].map(fmt_date)
    show["End_Date"] = show["End_Date"].map(fmt_date)
    show["Rate"] = pd.to_numeric(show["Rate"], errors="coerce").round(2)
    for c in ["Proposal_Cost", "Final_Cost", "Profit"]:
        show[c] = pd.to_numeric(show[c], errors="coerce").round(2)
    return show


def close_proposal(pid: str) -> None:
    m = st.session_state.proposals_df["Proposal_ID"] == pid
    st.session_state.proposals_df.loc[m, "Status"] = "Close"
    st.session_state.proposals_df.loc[m, "Closing_Date"] = pd.Timestamp(date.today())
    save_proposals()


def renew_proposal(pid: str, start: date, days: int) -> str:
    src = st.session_state.proposals_df[st.session_state.proposals_df["Proposal_ID"] == pid]
    if src.empty:
        raise ValueError("Proposal not found")
    new_id = new_proposal_id()
    end = pd.Timestamp(start) + pd.Timedelta(days=int(days))
    rows = []
    for _, r in src.iterrows():
        final, profit = calc(float(r["Proposal_Cost"]), float(r["Rate"]), days)
        rows.append({
            "Proposal_ID": new_id, "Client_ID": r["Client_ID"], "Client_Name": r["Client_Name"],
            "Proposal_Cost": r["Proposal_Cost"], "Rate": r["Rate"], "Profit": profit,
            "Final_Cost": final, "Start_Date": pd.Timestamp(start), "End_Date": end,
            "Status": "Open", "Closing_Date": pd.NaT,
        })
    st.session_state.proposals_df = pd.concat(
        [st.session_state.proposals_df, pd.DataFrame(rows)], ignore_index=True
    )
    save_proposals()
    return new_id


# ── Import ──────────────────────────────────────────────────────────────────

def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    rename = {c: RAW_ALIASES[str(c).strip().lower()] for c in df.columns if str(c).strip().lower() in RAW_ALIASES}
    return df.rename(columns=rename)


def prepare_import(df: pd.DataFrame) -> pd.DataFrame:
    raw = _norm_cols(df)
    need = ["Client_Name", "Start_Date", "End_Date", "Proposal_Cost", "Rate"]
    miss = [c for c in need if c not in raw.columns]
    if miss:
        raise ValueError(f"Missing columns: {', '.join(miss)}")
    raw = raw.copy()
    raw["Client_Name"] = raw["Client_Name"].fillna("").astype(str).str.strip()
    raw = raw[raw["Client_Name"] != ""]
    raw["Start_Date"] = pd.to_datetime(raw["Start_Date"], errors="coerce")
    raw["End_Date"] = pd.to_datetime(raw["End_Date"], errors="coerce")
    raw["Proposal_Cost"] = pd.to_numeric(raw["Proposal_Cost"], errors="coerce")
    raw["Rate"] = pd.to_numeric(raw["Rate"], errors="coerce")
    raw["Final_Cost"] = pd.to_numeric(raw.get("Final_Cost", pd.NA), errors="coerce")
    raw["Profit"] = pd.to_numeric(raw.get("Profit", pd.NA), errors="coerce")
    raw = raw.dropna(subset=need)
    if raw.empty:
        raise ValueError("No valid rows after cleaning.")
    dur = (raw["End_Date"] - raw["Start_Date"]).dt.days
    finals, profits = zip(*[calc(float(p), float(r), max(int(d), 0)) for p, r, d in zip(raw["Proposal_Cost"], raw["Rate"], dur)])
    raw["Final_Cost"] = raw["Final_Cost"].fillna(pd.Series(finals, index=raw.index))
    raw["Profit"] = raw["Profit"].fillna(pd.Series(profits, index=raw.index))
    today = pd.Timestamp(date.today())
    raw["Status"] = raw["End_Date"].apply(lambda e: "Close" if pd.notna(e) and e < today else "Open") if "Status" not in raw.columns else raw["Status"].apply(normalize_status)
    raw["Closing_Date"] = raw.apply(lambda r: r["End_Date"] if r["Status"] == "Close" else pd.NaT, axis=1)
    return raw.reset_index(drop=True)


def import_excel(raw_df: pd.DataFrame, replace: bool) -> dict:
    prepared = prepare_import(raw_df)
    clients = pd.DataFrame(columns=CLIENT_COLS) if replace else st.session_state.clients_df.copy()
    proposals = pd.DataFrame(columns=PROPOSAL_COLS) if replace else st.session_state.proposals_df.copy()
    name_map: dict[str, str] = {}
    for n, cid in zip(clients["Client_Name"], clients["Client_ID"]):
        if str(n).strip():
            name_map[str(n).strip().lower()] = str(cid)
    for name in sorted(prepared["Client_Name"].unique()):
        clean = str(name).strip()
        if clean.lower() in name_map:
            continue
        cid = next_id(clients["Client_ID"] if not clients.empty else pd.Series([], dtype=str), "SIG-C-", CLIENT_ID_RE)
        name_map[clean.lower()] = cid
        clients = pd.concat([clients, pd.DataFrame([{
            "Client_ID": cid, "Client_Name": clean, "Created_Date": pd.Timestamp.now(),
            "Is_Archived": False, "Notes": "Imported",
        }])], ignore_index=True)
    work = prepared.copy()
    work["_gk"] = work.apply(lambda r: (pd.Timestamp(r["Start_Date"]).normalize(), pd.Timestamp(r["End_Date"]).normalize(), round(float(r["Rate"]), 4)), axis=1)
    base = proposals["Proposal_ID"] if not proposals.empty else pd.Series([], dtype=str)
    gmap: dict = {}
    rows = []
    for _, r in work.iterrows():
        gk = r["_gk"]
        if gk not in gmap:
            gmap[gk] = next_id(pd.concat([base, pd.Series(list(gmap.values()))], ignore_index=True), "SIG-P-", PROPOSAL_ID_RE)
            base = pd.concat([base, pd.Series([gmap[gk]])], ignore_index=True)
        cid = name_map[str(r["Client_Name"]).strip().lower()]
        rows.append({
            "Proposal_ID": gmap[gk], "Client_ID": cid, "Client_Name": r["Client_Name"],
            "Proposal_Cost": float(r["Proposal_Cost"]), "Rate": float(r["Rate"]),
            "Final_Cost": float(r["Final_Cost"]), "Profit": float(r["Profit"]),
            "Start_Date": pd.Timestamp(r["Start_Date"]), "End_Date": pd.Timestamp(r["End_Date"]),
            "Status": r["Status"], "Closing_Date": r["Closing_Date"],
        })
    imported = pd.DataFrame(rows, columns=PROPOSAL_COLS)
    st.session_state.clients_df = clients
    st.session_state.proposals_df = pd.concat([proposals, imported], ignore_index=True) if not replace else imported
    save_clients()
    save_proposals()
    load_all_from_disk()
    return {"lines": len(imported), "proposals": imported["Proposal_ID"].nunique(), "clients": len(clients)}


# ── Logo & style ────────────────────────────────────────────────────────────

def _logo_bytes() -> tuple[bytes, str] | None:
    path = LOGO_FILE if LOGO_FILE.exists() else (LOGO_FALLBACK if LOGO_FALLBACK.exists() else None)
    if not path:
        return None
    if not LOGO_FILE.exists() and LOGO_FALLBACK.exists():
        try:
            shutil.copy2(LOGO_FALLBACK, LOGO_FILE)
        except OSError:
            pass
    mime = "image/jpeg" if path.suffix.lower() in {".jpg", ".jpeg"} else "image/png"
    return path.read_bytes(), mime


def logo_html(width: int = 200, css_class: str = "sigma-logo") -> str:
    data = _logo_bytes()
    if not data:
        return ""
    b64 = base64.b64encode(data[0]).decode()
    return (
        f'<img src="data:{data[1]};base64,{b64}" class="{css_class}" '
        f'style="width:{width}px;max-width:100%;" alt="Sigma Consultants" />'
    )


def css() -> None:
    st.markdown(
        f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

:root {{
    --c-primary: {C_PRIMARY};
    --c-accent: {C_ACCENT};
    --c-accent-dark: {C_ACCENT_DARK};
    --c-success: {C_SUCCESS};
    --c-bg: {C_BG};
    --c-card: {C_CARD};
    --c-muted: {C_MUTED};
    --c-border: {C_BORDER};
    --radius: 14px;
    --shadow: 0 4px 18px rgba(15, 23, 42, 0.06);
    --shadow-md: 0 10px 28px rgba(15, 23, 42, 0.09);
}}

html, body, [class*="css"], .stApp, button, input, textarea, select {{
    font-family: "Plus Jakarta Sans", Inter, system-ui, sans-serif !important;
}}

.stApp {{
    background: linear-gradient(165deg, #F8FAFC 0%, {C_BG} 50%, #E8EDF5 100%);
    color: var(--c-primary);
}}

[data-testid="stHeader"] {{ background: transparent; }}

div.block-container {{
    padding: 1.5rem 2rem 2.5rem;
    max-width: 1180px;
}}

/* ── Page header card ── */
.page-header {{
    background: var(--c-card);
    border: 1px solid var(--c-border);
    border-radius: var(--radius);
    padding: 1.15rem 1.35rem;
    margin-bottom: 1.15rem;
    box-shadow: var(--shadow);
}}
.page-header h1 {{
    margin: 0;
    font-size: 1.55rem;
    font-weight: 800;
    color: var(--c-primary);
    letter-spacing: -0.02em;
}}
.page-header p {{
    margin: 0.35rem 0 0;
    color: var(--c-muted);
    font-size: 0.92rem;
    line-height: 1.45;
}}

/* ── KPI cards ── */
.kpi-card {{
    background: var(--c-card);
    border: 1px solid var(--c-border);
    border-radius: 12px;
    padding: 0.9rem 1rem;
    box-shadow: var(--shadow);
    min-height: 88px;
}}
.kpi-label {{
    font-size: 0.72rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: var(--c-muted);
    margin-bottom: 0.35rem;
}}
.kpi-value {{
    font-size: 1.45rem;
    font-weight: 800;
    color: var(--c-primary);
    line-height: 1.2;
    letter-spacing: -0.02em;
}}
.kpi-value.accent {{ color: var(--c-accent); }}
.kpi-value.success {{ color: var(--c-success); }}
.kpi-value.warn {{ color: #D97706; }}
.kpi-value.danger {{ color: #DC2626; }}

.section-label {{
    font-size: 0.72rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: var(--c-muted);
    margin: 0.25rem 0 0.65rem;
}}

/* ── Login ── */
.login-wrap {{
    max-width: 420px;
    margin: 2rem auto 0;
}}
.login-card {{
    background: var(--c-card);
    border: 1px solid var(--c-border);
    border-radius: 18px;
    padding: 2rem 1.75rem 1.75rem;
    box-shadow: var(--shadow-md);
    text-align: center;
}}
.login-card .logo-box {{
    background: #fff;
    border-radius: 12px;
    padding: 0.75rem 1rem;
    margin-bottom: 1.25rem;
    box-shadow: var(--shadow);
}}
.login-card h2 {{
    margin: 0 0 0.35rem;
    font-size: 1.45rem;
    font-weight: 800;
    color: var(--c-primary);
}}
.login-card .login-sub {{
    color: var(--c-muted);
    font-size: 0.92rem;
    margin-bottom: 1.25rem;
}}

/* ── Sidebar ── */
section[data-testid="stSidebar"] {{
    background: #151C2B !important;
    border-right: 1px solid #2A3548;
}}
section[data-testid="stSidebar"] > div {{
    padding: 1.25rem 0.9rem 1.5rem;
    background: #151C2B;
}}
.sidebar-brand {{
    background: linear-gradient(145deg, #1E2A40, #1A2234);
    border: 1px solid #334155;
    border-radius: 14px;
    padding: 0.85rem 0.65rem;
    margin-bottom: 0.85rem;
    text-align: center;
    box-shadow: 0 4px 14px rgba(0,0,0,0.18);
}}
.sidebar-brand .logo-box {{
    background: #fff;
    border-radius: 10px;
    padding: 0.45rem 0.55rem;
    margin-bottom: 0.45rem;
}}
.sidebar-brand .tagline {{
    margin: 0;
    font-size: 0.78rem;
    color: #94A3B8;
    font-weight: 500;
}}
.sidebar-stats {{
    background: #1E2A40;
    border: 1px solid #334155;
    border-radius: 12px;
    padding: 0.7rem 0.85rem;
    margin-bottom: 0.85rem;
    font-size: 0.78rem;
    color: #94A3B8;
    line-height: 1.5;
}}
.sidebar-stats strong {{ color: #F1F5F9; font-size: 0.95rem; }}
.sidebar-stats .path {{ font-size: 0.65rem; color: #64748B; word-break: break-all; }}

section[data-testid="stSidebar"] .nav-section {{
    margin: 0.75rem 0 0.35rem;
    padding: 0 0.2rem;
    font-size: 0.65rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #64748B;
}}
section[data-testid="stSidebar"] .nav-btn .stButton > button {{
    width: 100%;
    justify-content: flex-start;
    text-align: left;
    min-height: 46px;
    padding: 0.55rem 0.85rem !important;
    border-radius: 11px !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
    background: #1E2A40 !important;
    color: #E2E8F0 !important;
    border: 1px solid #3D4F68 !important;
    box-shadow: 0 2px 6px rgba(0,0,0,0.1) !important;
    margin-bottom: 0.35rem;
}}
section[data-testid="stSidebar"] .nav-btn .stButton > button:hover {{
    background: #253347 !important;
    border-color: #4B6A8F !important;
    color: #fff !important;
}}
section[data-testid="stSidebar"] .nav-active .stButton > button {{
    background: linear-gradient(135deg, {C_ACCENT}, {C_ACCENT_DARK}) !important;
    color: #fff !important;
    border-color: #60A5FA !important;
    box-shadow: 0 4px 14px rgba(37, 99, 235, 0.35) !important;
}}
section[data-testid="stSidebar"] .nav-util .stButton > button {{
    background: #1E2A40 !important;
    color: #CBD5E1 !important;
    border: 1px dashed #475569 !important;
    border-radius: 11px !important;
    min-height: 42px;
    font-size: 0.85rem !important;
}}
section[data-testid="stSidebar"] .nav-logout .stButton > button {{
    background: #2A1F28 !important;
    color: #FCA5A5 !important;
    border: 1px solid #7F3D45 !important;
    border-radius: 11px !important;
    min-height: 42px;
    font-size: 0.85rem !important;
}}
section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p {{
    color: #CBD5E1 !important;
    font-weight: 600 !important;
}}

.mobile-tip {{
    display: none;
    background: #1E3A5F;
    border: 1px solid #2563EB;
    color: #BFDBFE;
    padding: 0.5rem 0.65rem;
    border-radius: 9px;
    font-size: 0.75rem;
    line-height: 1.4;
    margin-bottom: 0.75rem;
}}

/* ── Main content polish ── */
div[data-testid="stVerticalBlockBorderWrapper"] {{
    border-color: var(--c-border) !important;
    border-radius: 12px !important;
    box-shadow: var(--shadow);
    background: var(--c-card);
}}
.stButton > button[kind="primary"] {{
    background: linear-gradient(135deg, {C_ACCENT}, {C_ACCENT_DARK}) !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    box-shadow: 0 4px 12px rgba(37, 99, 235, 0.28) !important;
}}
.stButton > button[kind="secondary"] {{
    border-radius: 10px !important;
    font-weight: 600 !important;
    border-color: var(--c-border) !important;
}}
[data-testid="stMetric"] {{
    background: var(--c-card);
    border: 1px solid var(--c-border);
    border-radius: 12px;
    padding: 0.75rem 1rem;
    box-shadow: var(--shadow);
}}
[data-testid="stMetricLabel"] {{
    font-size: 0.72rem !important;
    font-weight: 700 !important;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: var(--c-muted) !important;
}}
[data-testid="stMetricValue"] {{
    font-size: 1.35rem !important;
    font-weight: 800 !important;
    color: var(--c-primary) !important;
}}
hr, [data-testid="stDivider"] {{
    margin: 1.1rem 0 !important;
    border-color: var(--c-border) !important;
}}
[data-testid="stDataFrame"] {{
    border: 1px solid var(--c-border);
    border-radius: 12px;
    overflow: hidden;
}}

@media (max-width: 768px) {{
    div.block-container {{ padding: 0.9rem 0.9rem 1.5rem; }}
    .page-header h1 {{ font-size: 1.3rem; }}
    .kpi-value {{ font-size: 1.2rem; }}
    .mobile-tip {{ display: block; }}
    [data-testid="stSidebarCollapsedControl"] {{
        display: flex !important;
        color: #fff !important;
        background: {C_ACCENT} !important;
        border-radius: 9px !important;
    }}
}}
</style>
""",
        unsafe_allow_html=True,
    )


def header(title: str, sub: str = "") -> None:
    sub_html = f"<p>{sub}</p>" if sub else ""
    st.markdown(
        f'<div class="page-header"><h1>{title}</h1>{sub_html}</div>',
        unsafe_allow_html=True,
    )


def kpi_row(items: list) -> None:
    """Render a row of KPI cards. Optional 3rd element: accent|success|warn|danger."""
    cols = st.columns(len(items))
    for col, item in zip(cols, items):
        label, value = item[0], item[1]
        tone = item[2] if len(item) > 2 else ""
        cls = f"kpi-value {tone}" if tone else "kpi-value"
        with col:
            st.markdown(
                f'<div class="kpi-card"><div class="kpi-label">{label}</div>'
                f'<div class="{cls.strip()}">{value}</div></div>',
                unsafe_allow_html=True,
            )


def section_label(text: str) -> None:
    st.markdown(f'<p class="section-label">{text}</p>', unsafe_allow_html=True)


# ── Auth & session ──────────────────────────────────────────────────────────

def password() -> str:
    try:
        s = st.secrets.get("SIGMA_PASSWORD")
        if s:
            return s
    except Exception:
        pass
    return os.getenv("SIGMA_PASSWORD") or "sigma123"


def init_session() -> None:
    defaults = {
        "auth": False, "page": "Maturity", "selected_proposal_id": None,
        "proposal_clients": [], "_data_ready": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v
    if st.session_state.auth and not st.session_state._data_ready:
        load_all_from_disk()
        st.session_state._data_ready = True


def login_page() -> None:
    img = logo_html(300)
    logo_block = f'<div class="logo-box">{img}</div>' if img else ""
    st.markdown(
        f"""
<div class="login-wrap">
  <div class="login-card">
    {logo_block}
    <h2>Welcome back</h2>
    <p class="login-sub">Sign in to manage clients, proposals, and maturities.</p>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )
    _, col, _ = st.columns([1, 1.1, 1])
    with col:
        with st.form("login"):
            pwd = st.text_input("Password", type="password", placeholder="Enter your password")
            if st.form_submit_button("Continue to dashboard", use_container_width=True, type="primary"):
                if pwd == password():
                    st.session_state.auth = True
                    st.session_state._data_ready = False
                    st.rerun()
                else:
                    st.error("Incorrect password")


def sidebar() -> None:
    with st.sidebar:
        img = logo_html(200)
        logo_block = f'<div class="logo-box">{img}</div>' if img else ""
        st.markdown(
            f"""
<div class="sidebar-brand">
  {logo_block}
  <p class="tagline">Client &amp; proposal management</p>
</div>
""",
            unsafe_allow_html=True,
        )
        n_c = len(st.session_state.clients_df)
        n_p = len(st.session_state.proposals_df)
        st.markdown(
            f"""
<div class="sidebar-stats">
  <strong>{n_c}</strong> clients &nbsp;·&nbsp; <strong>{n_p}</strong> proposals<br>
  <span class="path">Data: {BASE_DIR}</span>
</div>
""",
            unsafe_allow_html=True,
        )
        st.markdown(
            '<p class="mobile-tip">📱 Tap <strong>☰</strong> (top-left) to open or close this menu.</p>',
            unsafe_allow_html=True,
        )
        cur = st.session_state.page
        for section_name, keys in NAV_GROUPS:
            st.markdown(f'<p class="nav-section">{section_name}</p>', unsafe_allow_html=True)
            for key in keys:
                wrap = "nav-active" if cur == key else "nav-btn"
                st.markdown(f'<div class="{wrap}">', unsafe_allow_html=True)
                if st.button(PAGES[key], key=f"nav_{key}", use_container_width=True):
                    go(key)
                st.markdown("</div>", unsafe_allow_html=True)
        st.markdown('<p class="nav-section">System</p>', unsafe_allow_html=True)
        st.markdown('<div class="nav-util">', unsafe_allow_html=True)
        if st.button("🔄  Reload from Excel", use_container_width=True, key="nav_reload"):
            load_all_from_disk()
            st.session_state._data_ready = True
            st.toast("Reloaded from disk")
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown('<div class="nav-logout">', unsafe_allow_html=True)
        if st.button("🚪  Log out", use_container_width=True, key="nav_logout"):
            st.session_state.auth = False
            st.session_state._data_ready = False
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)


# ── Pages ───────────────────────────────────────────────────────────────────

def page_maturity() -> None:
    header("Maturity board", "What is ending today, soon, or overdue.")
    df = enrich(st.session_state.proposals_df)
    if df.empty:
        st.info("No proposals yet. Import Excel or add a proposal.")
        return
    open_df = df[df["Status"] == "Open"]
    kpi_row([
        ("Due today", int((open_df["Days_Left"] == 0).sum()), "warn"),
        ("Due in 7 days", int(open_df["Days_Left"].between(0, 7).sum()), "accent"),
        ("Overdue", int((open_df["Days_Left"] < 0).sum()), "danger"),
        ("Open investment", fmt_money(open_df["Proposal_Cost"].sum()), "success"),
    ])
    with st.container(border=True):
        section_label("Filter proposals")
        bucket = st.selectbox("Show", MATURITY_FILTERS, index=len(MATURITY_FILTERS) - 1, label_visibility="collapsed")
    filtered = filter_maturity(df, bucket)
    if filtered.empty:
        st.success(f"No proposals in “{bucket}”.")
        return
    st.dataframe(
        display_table(filtered)[["Client_Name", "End_Date", "Days_Left", "Proposal_Cost", "Rate", "Profit", "Status", "Maturity_Bucket"]],
        use_container_width=True, hide_index=True,
        column_config={**money_cfg("Proposal_Cost", "Profit"), "Days_Left": st.column_config.NumberColumn(format="%d")},
    )
    with st.expander("Open proposal file"):
        pid = st.selectbox("Proposal ID", unique_sorted(filtered["Proposal_ID"]))
        if st.button("View", type="primary"):
            st.session_state.selected_proposal_id = pid
            go("ProposalDetail")


def page_proposal_detail() -> None:
    header("Proposal file", "All clients in one proposal.")
    df = enrich(st.session_state.proposals_df)
    if df.empty:
        st.info("No proposals.")
        return
    ids = unique_sorted(df["Proposal_ID"])
    idx = ids.index(st.session_state.selected_proposal_id) if st.session_state.selected_proposal_id in ids else 0
    pid = st.selectbox("Proposal ID", ids, index=idx)
    st.session_state.selected_proposal_id = pid
    pdf = df[df["Proposal_ID"] == pid]
    dl = pdf["Days_Left"].iloc[0]
    kpi_row([
        ("Start date", fmt_date(pdf["Start_Date"].iloc[0])),
        ("End date", fmt_date(pdf["End_Date"].iloc[0])),
        ("Rate", f"{float(pdf['Rate'].iloc[0]):.2f}%", "accent"),
        ("Days left", int(dl) if pd.notna(dl) else "—", "warn"),
    ])
    section_label("Clients in this proposal")
    st.dataframe(
        display_table(pdf)[["Client_Name", "Proposal_Cost", "Rate", "Profit", "Final_Cost", "Status", "Days_Left"]],
        use_container_width=True, hide_index=True, column_config=money_cfg("Proposal_Cost", "Profit", "Final_Cost"),
    )
    if pdf["Status"].iloc[0] == "Open":
        if st.button("Close entire proposal", type="primary"):
            close_proposal(pid)
            st.success("Closed")
            st.rerun()
        with st.expander("Renew proposal"):
            ns = st.date_input("New start", value=date.today())
            nd = st.selectbox("Duration (days)", DURATION_DAYS)
            if st.button("Create renewed proposal"):
                nid = renew_proposal(pid, ns, int(nd))
                st.session_state.selected_proposal_id = nid
                st.success(f"Created {nid}")
                st.rerun()
    else:
        st.info("Proposal is closed. Use Renew to create a new open proposal.")


def page_client_ledger() -> None:
    header("Client ledger", "All proposals and totals for one client.")
    clients = active_clients()
    if clients.empty:
        st.info("No active clients.")
        return
    df = enrich(st.session_state.proposals_df)
    name = st.selectbox("Client", unique_sorted(clients["Client_Name"]))
    cid = clients[clients["Client_Name"] == name]["Client_ID"].iloc[0]
    cdf = df[df["Client_ID"] == cid]
    status = st.radio("Status", ["All", "Open", "Close"], horizontal=True)
    if status != "All":
        cdf = cdf[cdf["Status"] == status]
    if cdf.empty:
        st.info("No records.")
        return
    open_r = cdf[cdf["Status"] == "Open"]
    kpi_row([
        ("Open investment", fmt_money(open_r["Proposal_Cost"].sum()), "success"),
        ("Expected profit", fmt_money(open_r["Profit"].sum()), "accent"),
        ("Open lines", len(open_r)),
    ])
    section_label("Proposal history")
    st.dataframe(
        display_table(cdf.sort_values("End_Date"))[["Proposal_ID", "Start_Date", "End_Date", "Days_Left", "Proposal_Cost", "Profit", "Status"]],
        use_container_width=True, hide_index=True, column_config=money_cfg("Proposal_Cost", "Profit"),
    )


def page_search() -> None:
    header("Search", "Client name, proposal ID, or date (dd-mm-yyyy).")
    df = enrich(st.session_state.proposals_df)
    if df.empty:
        st.warning("No data.")
        return
    q = st.text_input("Search")
    status = st.selectbox("Status", ["All", "Open", "Close"])
    res = df.copy()
    if status != "All":
        res = res[res["Status"] == status]
    if q.strip():
        ql = q.strip().lower()
        res = res[
            res["Client_Name"].str.lower().str.contains(ql, na=False)
            | res["Proposal_ID"].str.lower().str.contains(ql, na=False)
            | res["End_Date"].map(fmt_date).str.contains(ql, na=False)
            | res["Start_Date"].map(fmt_date).str.contains(ql, na=False)
        ]
    if res.empty:
        st.info("No matches.")
        return
    kpi_row([
        ("Investment", fmt_money(res["Proposal_Cost"].sum())),
        ("Final amount", fmt_money(res["Final_Cost"].sum()), "accent"),
        ("Profit", fmt_money(res["Profit"].sum()), "success"),
    ])
    section_label("Results")
    st.dataframe(
        display_table(res)[["Proposal_ID", "Client_Name", "End_Date", "Proposal_Cost", "Profit", "Status"]],
        use_container_width=True, hide_index=True, column_config=money_cfg("Proposal_Cost", "Profit"),
    )


def page_add_proposal() -> None:
    header("Add proposal", "Create one proposal with one or more clients.")
    active = active_clients()
    if active.empty:
        st.warning("Add a client first.")
        return
    with st.container(border=True):
        section_label("Proposal terms")
        c1, c2, c3 = st.columns(3)
        start = c1.date_input("Start date", value=date.today())
        days = c2.selectbox("Duration (days)", DURATION_DAYS)
        rate = c3.number_input("Rate %", min_value=0.0, step=0.25)
        end = (pd.Timestamp(start) + pd.Timedelta(days=int(days))).date()
        st.caption(f"Maturity date: **{end.strftime('%d-%m-%Y')}**")
    with st.container(border=True):
        section_label("Add client line")
        c1, c2 = st.columns(2)
        client = c1.selectbox("Client", ["—"] + unique_sorted(active["Client_Name"]))
        amount = c2.number_input("Amount (₹)", min_value=0.0, step=1000.0)
    if st.button("Add to draft"):
        if client == "—" or amount <= 0:
            st.warning("Select client and amount.")
        elif any(x["Client_Name"] == client for x in st.session_state.proposal_clients):
            st.warning("Client already on draft.")
        else:
            f, p = calc(amount, rate, int(days))
            st.session_state.proposal_clients.append({"Client_Name": client, "Principal": amount, "Profit": p, "Final": f})
    if st.session_state.proposal_clients:
        st.dataframe(pd.DataFrame(st.session_state.proposal_clients), use_container_width=True, hide_index=True)
    if st.button("Save proposal", type="primary") and st.session_state.proposal_clients:
        pid = new_proposal_id()
        rows = []
        for item in st.session_state.proposal_clients:
            m = active[active["Client_Name"] == item["Client_Name"]].iloc[0]
            f, p = calc(item["Principal"], rate, int(days))
            rows.append({
                "Proposal_ID": pid, "Client_ID": m["Client_ID"], "Client_Name": item["Client_Name"],
                "Proposal_Cost": item["Principal"], "Rate": rate, "Profit": p, "Final_Cost": f,
                "Start_Date": pd.Timestamp(start), "End_Date": pd.Timestamp(end),
                "Status": "Open", "Closing_Date": pd.NaT,
            })
        st.session_state.proposals_df = pd.concat([st.session_state.proposals_df, pd.DataFrame(rows)], ignore_index=True)
        save_proposals()
        st.session_state.proposal_clients = []
        st.session_state.selected_proposal_id = pid
        st.success(f"Saved {pid}")
        go("ProposalDetail")


def page_edit() -> None:
    header("Edit line", "Update amount, rate, dates, or close a single client line.")
    prop = st.session_state.proposals_df
    if prop.empty:
        st.info("No proposals.")
        return
    with st.container(border=True):
        section_label("Select record")
        sf = st.selectbox("Status filter", ["Open", "Close"])
        df = prop[prop["Status"] == sf]
        if df.empty:
            st.info(f"No {sf} proposals.")
            return
        pid = st.selectbox("Proposal ID", unique_sorted(df["Proposal_ID"]))
        pdf = df[df["Proposal_ID"] == pid]
        client = st.selectbox("Client", unique_sorted(pdf["Client_Name"]))
    row = pdf[pdf["Client_Name"] == client].iloc[0]
    ix = pdf[pdf["Client_Name"] == client].index[0]
    with st.container(border=True):
        section_label("Edit values")
        c1, c2 = st.columns(2)
        cost = c1.number_input("Amount", min_value=0.0, value=float(row["Proposal_Cost"]))
        rate = c2.number_input("Rate %", min_value=0.0, value=float(row["Rate"]))
        c3, c4 = st.columns(2)
        sd = c3.date_input("Start", value=to_date(row["Start_Date"]))
        ed = c4.date_input("End", value=to_date(row["End_Date"], sd))
        nst = st.selectbox("Status", ["Open", "Close"], index=0 if row["Status"] == "Open" else 1)
        d = (ed - sd).days
        if d < 0:
            st.error("End before start.")
            return
        final, profit = calc(cost, rate, d)
        st.caption(f"Final {fmt_money(final)} · Profit {fmt_money(profit)}")
        if st.button("Save changes", type="primary"):
            st.session_state.proposals_df.at[ix, "Proposal_Cost"] = cost
            st.session_state.proposals_df.at[ix, "Rate"] = rate
            st.session_state.proposals_df.at[ix, "Start_Date"] = pd.Timestamp(sd)
            st.session_state.proposals_df.at[ix, "End_Date"] = pd.Timestamp(ed)
            st.session_state.proposals_df.at[ix, "Final_Cost"] = final
            st.session_state.proposals_df.at[ix, "Profit"] = profit
            st.session_state.proposals_df.at[ix, "Status"] = nst
            st.session_state.proposals_df.at[ix, "Closing_Date"] = pd.Timestamp(date.today()) if nst == "Close" else pd.NaT
            save_proposals()
            st.success("Updated")


def page_clients() -> None:
    header("Clients", "Add, archive, and restore client records.")
    with st.container(border=True):
        section_label("New client")
        c1, c2 = st.columns(2)
        name = c1.text_input("Client name")
        notes = c2.text_input("Notes (optional)")
    if st.button("Add client", type="primary"):
        n = name.strip()
        if not n:
            st.warning("Enter a name.")
        elif n.lower() in st.session_state.clients_df["Client_Name"].str.lower().tolist():
            st.warning("Client exists.")
        else:
            st.session_state.clients_df = pd.concat([st.session_state.clients_df, pd.DataFrame([{
                "Client_ID": new_client_id(), "Client_Name": n, "Created_Date": pd.Timestamp.now(),
                "Is_Archived": False, "Notes": notes.strip(),
            }])], ignore_index=True)
            save_clients()
            st.success("Added")
            st.rerun()
    show_arch = st.toggle("Show archived", False)
    view = st.session_state.clients_df.copy()
    if not show_arch:
        view = view[~view["Is_Archived"]]
    st.dataframe(view, use_container_width=True, hide_index=True)
    if not st.session_state.clients_df.empty:
        t = st.selectbox("Archive / restore", unique_sorted(st.session_state.clients_df["Client_Name"]))
        c1, c2 = st.columns(2)
        if c1.button("Archive"):
            st.session_state.clients_df.loc[st.session_state.clients_df["Client_Name"] == t, "Is_Archived"] = True
            save_clients()
            st.rerun()
        if c2.button("Restore"):
            st.session_state.clients_df.loc[st.session_state.clients_df["Client_Name"] == t, "Is_Archived"] = False
            save_clients()
            st.rerun()


def page_summary() -> None:
    header("Summary", "Investment totals and date-wise breakdowns.")
    df = st.session_state.proposals_df.copy()
    if df.empty:
        st.info("No data.")
        return
    kpi_row([
        ("Total investment", fmt_money(df["Proposal_Cost"].sum())),
        ("Total profit", fmt_money(df["Profit"].sum()), "success"),
        ("Open proposals", int((df["Status"] == "Open").sum()), "accent"),
    ])
    with st.container(border=True):
        section_label("Breakdown by rate & date")
        stat = st.selectbox("Status", ["All", "Open", "Close"])
        dtype = st.radio("Date", ["Start Date", "End Date"], horizontal=True)
    if stat != "All":
        df = df[df["Status"] == stat]
    col = "Start_Date" if dtype == "Start Date" else "End_Date"
    df = df.dropna(subset=[col])
    if df.empty:
        return
    df["DateOnly"] = df[col].dt.date
    g = df.groupby(["Rate", "DateOnly"], as_index=False).agg({"Proposal_Cost": "sum", "Final_Cost": "sum", "Profit": "sum"})
    st.dataframe(g.round(2), use_container_width=True, hide_index=True)


def page_import() -> None:
    header("Import Excel", "Upload raw data — Client ID and Proposal ID are created automatically.")
    with st.container(border=True):
        section_label("Choose file")
        st.caption("Expected columns: Start Date, Client Name, Initial Amount, End Date, Profit Rate, etc.")
        up = st.file_uploader("Excel file", type=["xlsx", "xls"])
    src = None
    if up is not None:
        src = pd.read_excel(up)
    elif DEFAULT_IMPORT.exists() and st.button(f"Load {DEFAULT_IMPORT.name} from Desktop"):
        src = pd.read_excel(DEFAULT_IMPORT)
    if src is None:
        st.info("Upload a file or load from Desktop.")
        return
    st.caption(f"{len(src)} rows")
    st.dataframe(src.head(15), use_container_width=True, hide_index=True)
    replace = st.radio("Mode", ["Replace all data (recommended for full import)", "Add to existing"], index=0)
    if st.button("Import", type="primary"):
        try:
            stats = import_excel(src, replace=replace.startswith("Replace"))
            st.success(f"Imported {stats['lines']} lines · {stats['proposals']} proposals · {stats['clients']} clients")
            go("Maturity")
        except Exception as e:
            st.error(str(e))


def page_export() -> None:
    header("Export & backup", "Download your data or use auto-backups saved on every change.")
    c, p = st.session_state.clients_df, st.session_state.proposals_df
    with st.container(border=True):
        section_label("Auto-backup")
        if BACKUP_DIR.exists():
            backups = sorted(BACKUP_DIR.glob("*.xlsx"), reverse=True)
            if backups:
                st.caption(f"Latest backup: **{backups[0].name}** · folder: `{BACKUP_DIR}`")
            else:
                st.caption("Backups are created automatically when you save data.")
        else:
            st.caption("Backups folder will be created on first save.")
    with st.container(border=True):
        section_label("Download")
        buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        if not c.empty:
            c.to_excel(w, sheet_name="Clients", index=False)
        if not p.empty:
            p.to_excel(w, sheet_name="Proposals", index=False)
        st.download_button(
            "Download Excel (clients + proposals)",
            buf.getvalue(),
            f"sigma_crm_{datetime.now():%Y%m%d}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )


# ── Main ────────────────────────────────────────────────────────────────────

ROUTES = {
    "Maturity": page_maturity,
    "ProposalDetail": page_proposal_detail,
    "ClientLedger": page_client_ledger,
    "Search": page_search,
    "AddProposal": page_add_proposal,
    "Edit": page_edit,
    "Clients": page_clients,
    "Summary": page_summary,
    "Import": page_import,
    "Export": page_export,
}

css()
init_session()

if not st.session_state.auth:
    login_page()
    st.stop()

sidebar()
page = st.session_state.page
ROUTES.get(page, page_maturity)()
