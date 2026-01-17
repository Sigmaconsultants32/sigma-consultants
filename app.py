# =====================================================
# Sigma Consultants – Full CRM Production Build
# =====================================================

import streamlit as st
import pandas as pd
from datetime import datetime
import os, io

# ---------------- PAGE CONFIG ----------------
st.set_page_config(page_title="Sigma Consultants", layout="wide")

# ---------------- MOBILE TOGGLE ----------------
is_mobile = st.toggle("📱 Mobile View", value=True)

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
        df.to_excel(CLIENT_FILE,index=False)

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

def save_clients(): clients_df.to_excel(CLIENT_FILE,index=False)
def save_proposals(): proposals_df.to_excel(PROPOSAL_FILE,index=False)

# ---------------- UTIL ----------------
def new_client_id(): return f"SIG-C-{len(clients_df)+1:03d}"
def new_proposal_id(): return f"SIG-P-{len(proposals_df)+1:03d}"

# ---------------- PAGE STATE ----------------
if "page" not in st.session_state:
    st.session_state.page = "Summary"

# ---------------- SIDEBAR ----------------
with st.sidebar:
    st.markdown("## 📂 Sigma Consultants")

    if st.button("🏠 Summary", use_container_width=True):
        st.session_state.page = "Summary"; st.rerun()

    if st.button("➕ Add Proposal", use_container_width=True):
        st.session_state.page = "AddProposal"; st.rerun()

    if st.button("🔍 Find Details", use_container_width=True):
        st.session_state.page = "Find"; st.rerun()

    if st.button("✏️ Edit Proposal", use_container_width=True):
        st.session_state.page = "Edit"; st.rerun()

    if st.button("👤 Clients", use_container_width=True):
        st.session_state.page = "Clients"; st.rerun()

    if st.button("📊 Client Dashboard", use_container_width=True):
        st.session_state.page = "ClientDashboard"; st.rerun()

    st.markdown("---")
    if st.button("📥 Export Data", use_container_width=True):
        st.session_state.page = "Export"

    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.auth = False; st.rerun()

# =====================================================
# ================= WELCOME SCREEN ====================
# =====================================================
if st.session_state.page == "Welcome":

    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:

        logo_path = "sigma_logo.png"

        if os.path.exists(logo_path):
            st.image(logo_path, use_container_width=True)
        else:
            st.markdown(
                "<h2 style='text-align:center;'>Sigma Consultants</h2>",
                unsafe_allow_html=True
            )

        st.markdown(
            """
            <div style="
            text-align:center;
            padding:20px;
            background:#ffffff;
            border-radius:16px;
            box-shadow:0 4px 12px rgba(0,0,0,0.08)">
            
            <h3>Welcome</h3>
            <p style="font-size:16px;color:#555">
            Manage clients, proposals, profits, and follow-ups  
            in one secure platform.
            </p>

            <p style="font-size:14px;color:#777">
            Please choose an option from the menu to get started.
            </p>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("🚀 Go to Dashboard", use_container_width=True):
            st.session_state.page = "Summary"
            st.rerun()
            
# =====================================================
# ================= SUMMARY ===========================
# =====================================================
if st.session_state.page == "Summary":
    st.header("📊 Summary")

    # ---------- SAFETY CHECK ----------
    if proposals_df.empty:
        st.info("No proposals available")
        st.stop()

    # ---------- BASIC SUMMARY ----------
    total_inv = proposals_df["Proposal_Cost"].sum()
    total_profit = proposals_df["Profit"].sum()
    open_cnt = len(proposals_df[proposals_df["Status"] == "Open"])

    if is_mobile:
        card("Total Investment", f"₹ {total_inv:,.2f}")
        card("Total Profit", f"₹ {total_profit:,.2f}")
        card("Open Proposals", open_cnt)
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Investment", f"₹ {total_inv:,.2f}")
        c2.metric("Total Profit", f"₹ {total_profit:,.2f}")
        c3.metric("Open Proposals", open_cnt)

    # =====================================================
    # ============ END DATE BASED SUMMARY =================
    # =====================================================
    st.markdown("---")
    st.subheader("📅 End Date Based Summary")
    st.caption("ℹ️ All amounts are shown in actual ₹ (rupees)")

    # ---------- DATE CONVERSION ----------
    proposals_df["Start_Date"] = pd.to_datetime(
        proposals_df["Start_Date"], errors="coerce"
    )
    proposals_df["End_Date"] = pd.to_datetime(
        proposals_df["End_Date"], errors="coerce"
    )

    # =====================================================
    # ================= STATUS FILTER =====================
    # =====================================================
    st.subheader("📌 Status Filter")

    status_options = ["All"] + sorted(
        proposals_df["Status"].dropna().unique().tolist()
    )

    selected_status = st.selectbox(
        "Select Proposal Status",
        status_options
    )

    if selected_status != "All":
        status_filtered_df = proposals_df[
            proposals_df["Status"] == selected_status
        ]
    else:
        status_filtered_df = proposals_df.copy()

    if status_filtered_df.empty:
        st.warning("No data available for selected status")
        st.stop()

    # =====================================================
    # ================= END DATE FILTER ===================
    # =====================================================
    end_dates = sorted(status_filtered_df["End_Date"].dropna().unique())

    if not end_dates:
        st.info("No End Dates available for selected status")
        st.stop()

    selected_end_date = st.selectbox(
        "Select End Date",
        end_dates,
        format_func=lambda x: x.strftime("%d-%m-%Y")
    )

    # ---------- FILTER DATA ----------
    filtered_df = status_filtered_df[
        status_filtered_df["End_Date"] == selected_end_date
    ]

    if filtered_df.empty:
        st.warning("No data for selected End Date")
        st.stop()

    # =====================================================
    # ========== GROUP BY RATE + START DATE ================
    # =====================================================
    summary_df = (
        filtered_df
        .groupby(["Rate", "Start_Date"], as_index=False)
        .agg({
            "Proposal_Cost": "sum",
            "Final_Cost": "sum",
            "Profit": "sum"
        })
        .sort_values(["Rate", "Start_Date"])
    )

    # =====================================================
    # ============ DISPLAY SUMMARY BOXES ===================
    # =====================================================
    for _, row in summary_df.iterrows():

        rate_display = round(row["Rate"], 2)
        inv = round(row["Proposal_Cost"], 2)
        final_amt = round(row["Final_Cost"], 2)
        profit = round(row["Profit"], 2)

        st.markdown(
            f"""
            <div style="
                border:1px solid #d0d0d0;
                border-radius:12px;
                padding:16px;
                margin-bottom:14px;
                background-color:#fafafa;">
                
                <b>Start Date</b> : {row['Start_Date'].strftime('%d-%m-%Y')}<br>
                <b>Total Investment</b> : ₹ {inv:,.2f}<br>
                <b>Rate</b> : {rate_display:.2f} %<br>
                <b>Total Final Amount</b> : ₹ {final_amt:,.2f}<br>
                <b>Total Profit</b> :
                <span style="color:green;"><b>₹ {profit:,.2f}</b></span>
            </div>
            """,
            unsafe_allow_html=True
        )

# =====================================================
# ================= ADD PROPOSAL ======================
# =====================================================
if st.session_state.page == "AddProposal":

    st.header("➕ Add New Proposal")

    active_clients = clients_df[clients_df["Is_Archived"]==False]
    if active_clients.empty:
        st.warning("Add a client first")
        st.stop()

    cname = st.selectbox("Client Name", active_clients["Client_Name"])
    start = st.date_input("Start Date", datetime.today())
    end = st.date_input("End Date", datetime.today())
    cost = st.number_input("Proposal Amount (₹)", step=1000.0, format="%.2f")
    rate = st.number_input("Monthly Rate (%)", step=0.1, format="%.2f")

    days = (end - start).days
    months = days / 30 if days > 0 else 0
    profit = cost * (rate/100) * months
    final = cost + profit

    st.markdown("### 💰 Auto Calculation")
    st.write(f"Duration: {months:.2f} months")
    st.write(f"Profit: ₹ {profit:,.2f}")
    st.write(f"Final Amount: ₹ {final:,.2f}")

    if st.button("💾 Save Proposal", use_container_width=True):
        cid = active_clients.loc[
            active_clients["Client_Name"]==cname,"Client_ID"
        ].values[0]

        proposals_df.loc[len(proposals_df)] = [
            new_proposal_id(), cid, cname,
            round(cost,2), round(rate,2),
            round(final,2), round(profit,2),
            start, end, "Open", ""
        ]
        save_proposals()
        st.success("Proposal added successfully")
        st.session_state.page = "Summary"
        st.rerun()

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
    pid = st.selectbox("Select Proposal", proposals_df["Proposal_ID"])
    row = proposals_df[proposals_df["Proposal_ID"]==pid].iloc[0]

    cost = st.number_input("Proposal Amount", value=float(row["Proposal_Cost"]))
    rate = st.number_input("Rate (%)", value=float(row["Rate"]))
    status = st.selectbox("Status", ["Open","Closed"], index=0 if row["Status"]=="Open" else 1)

    if st.button("💾 Update", use_container_width=True):
        proposals_df.loc[proposals_df["Proposal_ID"]==pid,
            ["Proposal_Cost","Rate","Status"]] = [cost,rate,status]
        save_proposals()
        st.success("Updated successfully")

# =====================================================
# ================= CLIENTS ===========================
# =====================================================
if st.session_state.page == "Clients":

    st.header("👤 Clients")

    with st.expander("➕ Add Client"):
        cname = st.text_input("Client Name")
        if st.button("Add Client"):
            clients_df.loc[len(clients_df)] = [
                new_client_id(), cname, datetime.now(), False, ""
            ]
            save_clients()
            st.success("Client added")
            st.rerun()

    for _,c in clients_df[clients_df["Is_Archived"]==False].iterrows():
        st.subheader(c["Client_Name"])
        st.write(f"Notes: {c['Notes']}")

        note = st.text_area("Update Notes", value=c["Notes"], key=c["Client_ID"])
        if st.button("Save Notes", key=f"note_{c['Client_ID']}"):
            clients_df.loc[clients_df["Client_ID"]==c["Client_ID"],"Notes"]=note
            save_clients()
            st.success("Notes saved")

        if st.button("📦 Archive Client", key=f"arc_{c['Client_ID']}"):
            clients_df.loc[clients_df["Client_ID"]==c["Client_ID"],"Is_Archived"]=True
            save_clients()
            st.rerun()

# =====================================================
# ================= CLIENT DASHBOARD ==================
# =====================================================
if st.session_state.page == "ClientDashboard":

    st.header("📊 Client Dashboard")
    cname = st.selectbox("Client", clients_df["Client_Name"])
    cid = clients_df.loc[clients_df["Client_Name"]==cname,"Client_ID"].values[0]
    data = proposals_df[proposals_df["Client_ID"]==cid]

    card("Total Investment", f"₹ {data['Proposal_Cost'].sum():,.2f}")
    card("Total Profit", f"₹ {data['Profit'].sum():,.2f}")
    card("Open Proposals", len(data[data["Status"]=="Open"]))

# =====================================================
# ================= EXPORT ============================
# =====================================================
if st.session_state.page == "Export":

    def export_excel():
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as w:
            clients_df.to_excel(w, sheet_name="Clients", index=False)
            proposals_df.to_excel(w, sheet_name="Proposals", index=False)
        out.seek(0)
        return out

    st.download_button(
        "📥 Download Excel Backup",
        data=export_excel(),
        file_name="Sigma_Consultants_Data.xlsx"
    )












