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

# ---------------- STYLES ----------------
st.markdown("""
<style>
.header {border-bottom:3px solid #1f77ff;margin-bottom:20px;padding-bottom:6px;}
.btn-blue button {background:#1f77ff !important;color:white;}
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
def new_client_id():
    return f"SIG-C-{len(clients_df)+1:03d}"

def new_proposal_id():
    return f"SIG-P-{len(proposals_df)+1:03d}"

def calc(cost, rate, days):
    months = days / 30
    profit = cost * (rate / 100) * months
    return round(cost + profit, 2), round(profit, 2)

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
        card("Investment", f"₹ {total_inv:,.2f}")
        card("Profit", f"₹ {total_profit:,.2f}")
        card("Open", open_cnt)
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
    st.caption("ℹ️ Amounts shown in actual ₹ (rupees)")

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
    status_options = ["All"] + sorted(
        proposals_df["Status"].dropna().unique().tolist()
    )

    selected_status = st.selectbox("Select Proposal Status", status_options)

    if selected_status != "All":
        status_filtered_df = proposals_df[
            proposals_df["Status"] == selected_status
        ].copy()
    else:
        status_filtered_df = proposals_df.copy()

    if status_filtered_df.empty:
        st.warning("No data for selected status")
        st.stop()

    # =====================================================
    # ================= END DATE FILTER ===================
    # =====================================================
    end_dates = sorted(status_filtered_df["End_Date"].dropna().unique())

    if not end_dates:
        st.info("No End Dates available")
        st.stop()

    selected_end_date = st.selectbox(
        "Select End Date",
        end_dates,
        format_func=lambda x: x.strftime("%d-%m-%Y")
    )

    filtered_df = status_filtered_df[
        status_filtered_df["End_Date"] == selected_end_date
    ].copy()

    if filtered_df.empty:
        st.warning("No data for selected End Date")
        st.stop()

    # =====================================================
    # ===== NORMALIZE RATE (INTEGER BASED GROUPING) =======
    # =====================================================
    filtered_df["Rate_Int"] = filtered_df["Rate"].round(0).astype(int)

    # =====================================================
    # ========== END DATE GRAND TOTALS ====================
    # =====================================================
    grand_inv = filtered_df["Proposal_Cost"].sum()
    grand_final = filtered_df["Final_Cost"].sum()
    grand_profit = filtered_df["Profit"].sum()

    st.markdown("### 🔢 End Date Grand Totals")

    if is_mobile:
        card("Investment", f"₹ {grand_inv:,.2f}")
        card("Final Amt", f"₹ {grand_final:,.2f}")
        card("Profit", f"₹ {grand_profit:,.2f}")
    else:
        g1, g2, g3 = st.columns(3)
        g1.metric("Total Investment", f"₹ {grand_inv:,.2f}")
        g2.metric("Total Final Amount", f"₹ {grand_final:,.2f}")
        g3.metric("Total Profit", f"₹ {grand_profit:,.2f}")

    st.markdown("---")

    # =====================================================
    # ========== GROUP BY RATE + START DATE ================
    # =====================================================
    summary_df = (
        filtered_df
        .groupby(["Rate_Int", "Start_Date"], as_index=False)
        .agg({
            "Proposal_Cost": "sum",
            "Final_Cost": "sum",
            "Profit": "sum"
        })
        .sort_values(["Rate_Int", "Start_Date"])
    )

    # =====================================================
    # ============ COMPACT SUMMARY DISPLAY ================
    # =====================================================
    for _, row in summary_df.iterrows():

        profit_color = "🟢" if row["Profit"] >= 0 else "🔴"

        if is_mobile:
            st.markdown(
                f"""
**📅 {row['Start_Date'].strftime('%d-%m-%Y')} | {row['Rate_Int']} %**  
💰 Invested : ₹ {row['Proposal_Cost']:,.2f}  
📈 Final : ₹ {row['Final_Cost']:,.2f}  
{profit_color} Profit : ₹ {row['Profit']:,.2f}
"""
            )
        else:
            st.markdown(
                f"""
**Start Date** : {row['Start_Date'].strftime('%d-%m-%Y')}  
**Rate** : {row['Rate_Int']} %  
**Investment** : ₹ {row['Proposal_Cost']:,.2f}  
**Final Amount** : ₹ {row['Final_Cost']:,.2f}  
**Profit** : ₹ {row['Profit']:,.2f}
"""
            )

        st.markdown("---")

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
    st.subheader("💰 Auto Calculated")

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

# =====================================================
# ================= FIND DETAILS ======================
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

    if is_mobile:
        cost = st.number_input("Proposal Cost (₹)", value=int(row["Proposal_Cost"]), step=1000)
        rate = st.number_input("Rate (%)", value=float(row["Rate"]), step=0.1)
        status = st.selectbox("Status", ["Active","Closed"], index=0 if row["Status"]=="Active" else 1)
    else:
        c1, c2 = st.columns(2)
        cost = c1.number_input("Proposal Cost (₹)", value=int(row["Proposal_Cost"]), step=1000)
        rate = c2.number_input("Rate (%)", value=float(row["Rate"]), step=0.1)
        status = st.selectbox("Status", ["Active","Closed"], index=0 if row["Status"]=="Active" else 1)

    if st.button("💾 Save Changes", use_container_width=True):
        proposals_df.loc[proposals_df["Proposal_ID"] == pid, "Proposal_Cost"] = cost
        proposals_df.loc[proposals_df["Proposal_ID"] == pid, "Rate"] = rate
        proposals_df.loc[proposals_df["Proposal_ID"] == pid, "Status"] = status
        save_proposals()
        st.success("Proposal updated successfully")

# =====================================================
# ================= CLIENTS ===========================
# =====================================================
if st.session_state.page == "Clients":

    st.header("👤 Clients Management")

    # -------------------------------------------------
    # ADD NEW CLIENT
    # -------------------------------------------------
    with st.expander("➕ Add New Client"):
        cname = st.text_input("Client Name", placeholder="Enter client name")
        if st.button("Add Client", use_container_width=True):
            if cname.strip():
                clients_df.loc[len(clients_df)] = [
                    new_client_id(),
                    cname.strip(),
                    datetime.now()
                ]
                save_clients()
                st.success("Client added successfully")
                st.rerun()
            else:
                st.error("Client name cannot be empty")

    # -------------------------------------------------
    # CLIENT SEARCH
    # -------------------------------------------------
    search = st.text_input("🔍 Search Client", placeholder="Type client name")

    filtered_clients = clients_df.copy()
    if search.strip():
        filtered_clients = filtered_clients[
            filtered_clients["Client_Name"].str.contains(search, case=False, na=False)
        ]

    # -------------------------------------------------
    # CLIENT LIST WITH ACTIONS
    # -------------------------------------------------
    with st.expander("📋 Client List", expanded=True):

        if filtered_clients.empty:
            st.info("No clients found")
            st.stop()

        for _, client in filtered_clients.iterrows():

            client_id = client["Client_ID"]
            client_name = client["Client_Name"]

            # Proposal count
            proposal_count = len(
                proposals_df[proposals_df["Client_ID"] == client_id]
            )

            st.markdown(
                f"""
                <div style="
                border:1px solid #ddd;
                border-radius:12px;
                padding:12px;
                margin-bottom:10px;
                background:#fafafa">
                <b>{client_name}</b><br>
                <span style="color:#555">Client ID:</span> {client_id}<br>
                <span style="color:#555">Proposals:</span> {proposal_count}
                </div>
                """,
                unsafe_allow_html=True
            )

            col1, col2, col3 = st.columns([3, 2, 2])

            # -------- EDIT CLIENT --------
            with col1:
                new_name = st.text_input(
                    "Edit Name",
                    value=client_name,
                    key=f"edit_{client_id}"
                )
                if new_name.strip() and new_name != client_name:
                    clients_df.loc[
                        clients_df["Client_ID"] == client_id,
                        "Client_Name"
                    ] = new_name.strip()

                    proposals_df.loc[
                        proposals_df["Client_ID"] == client_id,
                        "Client_Name"
                    ] = new_name.strip()

                    save_clients()
                    save_proposals()
                    st.success("Client name updated")
                    st.rerun()

            # -------- VIEW PROPOSALS --------
            with col2:
                if st.button("📄 View Proposals", key=f"view_{client_id}"):
                    st.session_state.page = "Find"
                    st.rerun()

            # -------- DELETE CLIENT --------
            with col3:
                if st.button("📦 Archive Client", key=f"arc_{client_id}"):
                    clients_df.loc[
                    clients_df["Client_ID"] == client_id,
                    "Is_Archived"
                    ] = True
                    save_clients()
                    st.success("Client archived (data preserved)")
                    st.rerun()


# =====================================================
# ================= CLIENT DASHBOARD ==================
# =====================================================
if st.session_state.page == "ClientDashboard":

    st.header("📊 Client Summary Dashboard")

    active_clients = clients_df[clients_df["Is_Archived"] == False]

    if active_clients.empty:
        st.info("No active clients")
        st.stop()

    client = st.selectbox(
        "Select Client",
        active_clients["Client_Name"]
    )

    client_id = active_clients.loc[
        active_clients["Client_Name"] == client,
        "Client_ID"
    ].values[0]

    data = proposals_df[proposals_df["Client_ID"] == client_id]

    total_invest = data["Proposal_Cost"].sum()
    total_profit = data["Profit"].sum()
    open_props = len(data[data["Status"] == "Open"])
    closed_props = len(data[data["Status"] == "Closed"])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Investment", f"₹ {total_invest:,.2f}")
    c2.metric("Total Profit", f"₹ {total_profit:,.2f}")
    c3.metric("Open Proposals", open_props)
    c4.metric("Closed Proposals", closed_props)

    st.markdown("### 📄 Proposal Details")
    st.dataframe(data, use_container_width=True)

# =====================================================
# 📤 EXPORT DATA SECTION (COMPLETE & FINAL)
# =====================================================

st.markdown("---")
st.subheader("📤 Export Data")

# ---- SAFETY CHECK ----
if "clients_df" not in st.session_state or st.session_state.clients_df.empty:
    st.warning("⚠️ No data available to export")
else:
    df = st.session_state.clients_df.copy()

    # ---------------- FILTER CONTROLS ----------------
    col1, col2, col3 = st.columns(3)

    with col1:
        client_filter = st.selectbox(
            "👤 Client",
            ["All"] + sorted(df["Client Name"].dropna().unique().tolist())
        )

    with col2:
        status_filter = st.selectbox(
            "📌 Status",
            ["All"] + sorted(df["Status"].dropna().unique().tolist())
        )

    with col3:
        date_range = st.date_input(
            "📅 Proposal Date Range",
            []
        )

    # ---------------- APPLY FILTERS ----------------
    filtered_df = df.copy()

    if client_filter != "All":
        filtered_df = filtered_df[
            filtered_df["Client Name"] == client_filter
        ]

    if status_filter != "All":
        filtered_df = filtered_df[
            filtered_df["Status"] == status_filter
        ]

    if len(date_range) == 2:
        start_date, end_date = date_range
        filtered_df["Proposal Date"] = pd.to_datetime(
            filtered_df["Proposal Date"],
            errors="coerce"
        )
        filtered_df = filtered_df[
            (filtered_df["Proposal Date"] >= pd.to_datetime(start_date)) &
            (filtered_df["Proposal Date"] <= pd.to_datetime(end_date))
        ]

    # ---------------- EXPORT FUNCTION ----------------
    def export_filtered_excel(dataframe):
        output = io.BytesIO()
        try:
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

                # Main data
                dataframe.to_excel(
                    writer,
                    sheet_name="Filtered_Data",
                    index=False
                )

                # Summary
                summary = {
                    "Export Date": [datetime.now().strftime("%d-%m-%Y %H:%M")],
                    "Total Records": [len(dataframe)],
                    "Client Filter": [client_filter],
                    "Status Filter": [status_filter]
                }

                pd.DataFrame(summary).to_excel(
                    writer,
                    sheet_name="Summary",
                    index=False
                )

            output.seek(0)
            return output.getvalue()

        except Exception as e:
            st.error(f"❌ Export failed: {e}")
            return None

    # ---------------- DOWNLOAD BUTTON ----------------
    excel_bytes = export_filtered_excel(filtered_df)

    if excel_bytes and not filtered_df.empty:
        st.download_button(
            "📥 Download Excel",
            excel_bytes,
            file_name=f"Sigma_Export_{datetime.now().strftime('%d%m%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("ℹ️ No records match selected filters")
