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
    # ============ DATE BASED SUMMARY =====================
    # =====================================================
    st.markdown("---")
    st.subheader("📅 Date Based Summary")
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
    # ============ START / END DATE SELECTOR ===============
    # =====================================================
    date_type = st.radio(
        "Select Date Type",
        ["Start Date", "End Date"],
        horizontal=True
    )

    date_col = "Start_Date" if date_type == "Start Date" else "End_Date"

    available_dates = sorted(
        status_filtered_df[date_col].dropna().dt.date.unique()
    )

    if not available_dates:
        st.info("No dates available")
        st.stop()

    date_options = ["All"] + available_dates

    selected_dates = st.multiselect(
        f"Select {date_type}(s)",
        date_options,
        default="All",
        format_func=lambda x: x if x == "All" else x.strftime("%d-%m-%Y")
    )

    filtered_df = status_filtered_df.copy()

    if "All" not in selected_dates:
        filtered_df = filtered_df[
            filtered_df[date_col].dt.date.isin(selected_dates)
        ]

    if filtered_df.empty:
        st.warning("No data for selected date(s)")
        st.stop()

    # =====================================================
    # ================ NORMALIZE RATE =====================
    # =====================================================
    filtered_df["Rate_Int"] = filtered_df["Rate"].round(0).astype(int)

    # =====================================================
    # ================= PROPOSAL STATS ====================
    # =====================================================
    total_inv = filtered_df["Proposal_Cost"].sum()
    total_final = filtered_df["Final_Cost"].sum()
    total_profit = filtered_df["Profit"].sum()

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
    # ========= GROUP BY RATE + START DATE ================
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
# ================= EDIT PROPOSAL =====================
# =====================================================
if st.session_state.page == "Edit":

    st.header("✏️ Edit Proposal")

    if proposals_df.empty:
        st.warning("No proposals available")
        st.stop()

    df = proposals_df.copy()

    # =================================================
    # STEP 1: STATUS FILTER
    # =================================================
    status_filter = st.selectbox(
        "Select Proposal Status",
        ["Open", "Closed"]
    )

    df = df[df["Status"] == status_filter]

    if df.empty:
        st.info(f"No {status_filter} proposals found")
        st.stop()

    # =================================================
    # STEP 2: PROPOSAL ID FILTER
    # =================================================
    proposal_id = st.selectbox(
        "Select Proposal ID",
        sorted(df["Proposal_ID"].unique())
    )

    proposal_df = df[df["Proposal_ID"] == proposal_id]

    # =================================================
    # STEP 3: AUTO START / END DATE (READ ONLY)
    # =================================================
    start_date = proposal_df["Start_Date"].iloc[0]
    end_date = proposal_df["End_Date"].iloc[0]
    rate_master = proposal_df["Rate"].iloc[0]

    c1, c2, c3 = st.columns(3)
    c1.text_input("Start Date", start_date.strftime("%d-%m-%Y"), disabled=True)
    c2.text_input("End Date", end_date.strftime("%d-%m-%Y"), disabled=True)
    c3.text_input("Rate (%)", int(round(rate_master, 0)), disabled=True)

    # =================================================
    # STEP 4: CLIENT LIST UNDER PROPOSAL
    # =================================================
    client_name = st.selectbox(
        "Select Client Included in Proposal",
        sorted(proposal_df["Client_Name"].unique())
    )

    row = proposal_df[
        proposal_df["Client_Name"] == client_name
    ].iloc[0]

    st.markdown("---")

    # =================================================
    # STEP 5: DISPLAY CURRENT DETAILS
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
        # STEP 7: AUTO CALCULATION
        # =================================================
        days = (pd.to_datetime(new_end) - pd.to_datetime(new_start)).days
        months = days / 30 if days > 0 else 0

        profit = proposal_cost * (rate / 100) * months
        final_cost = proposal_cost + profit

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

            mask = (
                (proposals_df["Proposal_ID"] == proposal_id) &
                (proposals_df["Client_Name"] == client_name)
            )

            proposals_df.loc[mask, [
                "Proposal_Cost",
                "Rate",
                "Start_Date",
                "End_Date",
                "Final_Cost",
                "Profit",
                "Status"
            ]] = [
                round(proposal_cost, 2),
                round(rate, 2),
                pd.to_datetime(new_start),
                pd.to_datetime(new_end),
                round(final_cost, 2),
                round(profit, 2),
                new_status
            ]

            save_proposals()

            st.success("✅ Proposal updated successfully")
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

    # -------------------------------------------------
    # MODE SELECTION (⚠️ MUST EXIST BEFORE ANY IF)
    # -------------------------------------------------
    find_mode = st.radio(
        "Find Details Mode",
        [
            "By Proposal",
            "By Client Name",
            "By Start / End Date"
        ],
        horizontal=True
    )

    # =================================================
    # =============== BY PROPOSAL MODE =================
    # =================================================
    if find_mode == "By Proposal":

        st.subheader("📄 Find Details By Proposal")

        df = proposals_df.copy()

        # -------------------------------------------------
        # STEP 1: CLIENT + STATUS FILTER
        # -------------------------------------------------
        col1, col2 = st.columns(2)

        selected_client = col1.selectbox(
            "Client Name",
            ["All"] + sorted(df["Client_Name"].dropna().unique())
        )

        selected_status = col2.selectbox(
            "Status",
            ["All", "Open", "Closed"]
        )

        if selected_client != "All":
            df = df[df["Client_Name"] == selected_client]

        if selected_status != "All":
            df = df[df["Status"] == selected_status]

        if df.empty:
            st.info("No proposals found for selected filters")
            st.stop()

        # -------------------------------------------------
        # STEP 2: PROPOSAL ID FILTER
        # -------------------------------------------------
        proposal_ids = sorted(df["Proposal_ID"].unique())

        selected_proposal = st.selectbox(
            "Select Proposal ID",
            proposal_ids
        )

        proposal_df = df[df["Proposal_ID"] == selected_proposal]

        # -------------------------------------------------
        # STEP 3: AUTO START / END / RATE (READ ONLY)
        # -------------------------------------------------
        p_start = proposal_df["Start_Date"].iloc[0].date()
        p_end = proposal_df["End_Date"].iloc[0].date()
        p_rate = int(round(proposal_df["Rate"].iloc[0], 0))

        c1, c2, c3 = st.columns(3)
        c1.text_input("Start Date", p_start.strftime("%d-%m-%Y"), disabled=True)
        c2.text_input("End Date", p_end.strftime("%d-%m-%Y"), disabled=True)
        c3.text_input("Rate (%)", p_rate, disabled=True)

        # -------------------------------------------------
        # STEP 4: CLIENTS UNDER PROPOSAL
        # -------------------------------------------------
        proposal_clients = sorted(proposal_df["Client_Name"].unique())

        selected_client_final = st.selectbox(
            "Select Client Included in Proposal",
            proposal_clients
        )

        final_df = proposal_df[
            proposal_df["Client_Name"] == selected_client_final
        ]

        record = final_df.iloc[0]

        st.markdown("---")

        # -------------------------------------------------
        # STEP 5: DISPLAY
        # -------------------------------------------------
        if is_mobile:
            st.markdown(
                f"""
                <div style="border:1px solid #ddd;
                border-radius:12px;
                padding:12px;
                background:#fafafa">

                <b>Client:</b> {record['Client_Name']}<br>
                <b>Status:</b> {record['Status']}<br>
                <b>Start:</b> {record['Start_Date'].date()}<br>
                <b>End:</b> {record['End_Date'].date()}<br>
                <b>Amount:</b> ₹ {record['Proposal_Cost']:,.2f}<br>
                <b>Rate:</b> {int(round(record['Rate'],0))} %<br>
                <b>Final:</b> ₹ {record['Final_Cost']:,.2f}<br>
                <b>Profit:</b> ₹ {record['Profit']:,.2f}
                </div>
                """,
                unsafe_allow_html=True
            )
        else:
            st.dataframe(
                final_df[[
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

        # -------------------------------------------------
        # STEP 6: EDIT
        # -------------------------------------------------
        st.markdown("---")
        edit_mode = st.checkbox("✏️ Edit Proposal")

        if edit_mode:
            new_cost = st.number_input(
                "Edit Proposal Amount",
                value=float(record["Proposal_Cost"]),
                step=1000.0
            )

            new_rate = st.number_input(
                "Edit Rate (%)",
                value=float(record["Rate"]),
                step=0.5
            )

            if st.button("💾 Save Changes"):
                mask = (
                    (proposals_df["Proposal_ID"] == selected_proposal) &
                    (proposals_df["Client_Name"] == selected_client_final)
                )

                proposals_df.loc[mask, "Proposal_Cost"] = new_cost
                proposals_df.loc[mask, "Rate"] = new_rate

                st.success("Proposal updated successfully")
                st.rerun()

    # =================================================
    # ============ BY CLIENT NAME MODE ================
    # =================================================
    if find_mode == "By Client Name":

        st.subheader("🔎 Find Details Using Client Name")

        status = st.selectbox("Status", ["All", "Open", "Closed"])
        client = st.selectbox(
            "Client Name",
            sorted(proposals_df["Client_Name"].dropna().unique())
        )

        date_type = st.radio(
            "Date Type",
            ["Start Date", "End Date"],
            horizontal=True
        )

        df = proposals_df.copy()

        if status != "All":
            df = df[df["Status"] == status]

        df = df[df["Client_Name"] == client]

        date_col = "Start_Date" if date_type == "Start Date" else "End_Date"
        available_dates = sorted(df[date_col].dropna().unique())

        if not available_dates:
            st.info("No data found")
            st.stop()

        selected_date = st.selectbox(
            f"Select {date_type}",
            available_dates,
            format_func=lambda x: x.strftime("%d-%m-%Y")
        )

        result = df[df[date_col] == selected_date]

        st.markdown("### 📋 Proposal Details")

        if is_mobile:
            for _, r in result.iterrows():
                st.markdown(
                    f"""
                    <div style="border:1px solid #ddd;border-radius:12px;
                    padding:12px;margin-bottom:10px;background:#f9f9f9">

                    <b>Proposal Amount:</b> ₹ {r['Proposal_Cost']:,.2f}<br>
                    <b>Rate:</b> {int(round(r['Rate'], 0))} %<br>
                    <b>Final Amount:</b> ₹ {r['Final_Cost']:,.2f}<br>
                    <b>Profit:</b> ₹ {r['Profit']:,.2f}

                    </div>
                    """,
                    unsafe_allow_html=True
                )
        else:
            table = result[[
                "Proposal_Cost","Rate","Final_Cost","Profit"
            ]].copy()

            table["Rate"] = table["Rate"].round(0).astype(int)
            table = table.round(2)

    # =================================================
    # ========= BY START / END DATE MODE ===============
    # =================================================
    if find_mode == "By Start / End Date":

        st.subheader("📅 Find Details Using Date & Clients")

        col1, col2 = st.columns(2)
        status = col1.selectbox("Status", ["All", "Open", "Closed"])
        date_type = col2.radio(
            "Date Type",
            ["Start Date", "End Date"],
            horizontal=True
        )

        clients = sorted(proposals_df["Client_Name"].dropna().unique())
        selected_clients = st.multiselect(
            "Select Clients (leave blank for All)",
            clients
        )

        df = proposals_df.copy()

        if status != "All":
            df = df[df["Status"] == status]

        if selected_clients:
            df = df[df["Client_Name"].isin(selected_clients)]

        date_col = "Start_Date" if date_type == "Start Date" else "End_Date"

        summary = (
            df.groupby([date_col, "Client_Name"], as_index=False)
            .agg({
                "Proposal_Cost": "sum",
                "Final_Cost": "sum",
                "Profit": "sum"
            })
            .round(2)
        )

        if summary.empty:
            st.info("No records found")
            st.stop()

        summary[date_col] = summary[date_col].dt.date

        st.markdown("### 📊 Client-wise Summary")

        if is_mobile:
            for _, r in summary.iterrows():
                st.markdown(
                    f"""
                    <div style="border:1px solid #ddd;border-radius:12px;
                    padding:12px;margin-bottom:10px;background:#fafafa">

                    <b>Date:</b> {r[date_col]}<br>
                    <b>Client:</b> {r['Client_Name']}<br>
                    <b>Investment:</b> ₹ {r['Proposal_Cost']:,.2f}<br>
                    <b>Final:</b> ₹ {r['Final_Cost']:,.2f}<br>
                    <b>Profit:</b> ₹ {r['Profit']:,.2f}

                    </div>
                    """,
                    unsafe_allow_html=True
                )

            st.markdown("### 🔢 Grand Total")
            st.metric("Investment", f"₹ {summary['Proposal_Cost'].sum():,.2f}")
            st.metric("Final Amount", f"₹ {summary['Final_Cost'].sum():,.2f}")
            st.metric("Profit", f"₹ {summary['Profit'].sum():,.2f}")

        else:
            st.dataframe(summary, use_container_width=True)

            st.markdown("### 🔢 Grand Total")
            c1, c2, c3 = st.columns(3)
            c1.metric("Investment", f"₹ {summary['Proposal_Cost'].sum():,.2f}")
            c2.metric("Final Amount", f"₹ {summary['Final_Cost'].sum():,.2f}")
            c3.metric("Profit", f"₹ {summary['Profit'].sum():,.2f}")

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

    # =================================================
    # ============== CLIENT FILTER ====================
    # =================================================
    client_options = ["Select"] + sorted(active_clients["Client_Name"].unique())

    client = st.selectbox(
        "Select Client",
        client_options
    )

    if client == "Select":
        st.info("Please select a client to continue")
        st.stop()

    client_id = active_clients.loc[
        active_clients["Client_Name"] == client,
        "Client_ID"
    ].values[0]

    client_data = proposals_df[
        proposals_df["Client_ID"] == client_id
    ].copy()

    if client_data.empty:
        st.info("No proposals available for this client")
        st.stop()

    # =================================================
    # ============== STATUS FILTER ====================
    # =================================================
    status_options = ["Select", "All", "Open", "Closed"]

    status = st.selectbox(
        "Select Status",
        status_options
    )

    if status == "Select":
        st.info("Please select status to view data")
        st.stop()

    if status != "All":
        client_data = client_data[
            client_data["Status"] == status
        ]

    if client_data.empty:
        st.info("No records found for selected filters")
        st.stop()

    # =================================================
    # ============== GRAND TOTALS =====================
    # =================================================
    total_invest = client_data["Proposal_Cost"].sum()
    total_final = client_data["Final_Cost"].sum()
    total_profit = client_data["Profit"].sum()
    open_props = len(client_data[client_data["Status"] == "Open"])
    closed_props = len(client_data[client_data["Status"] == "Closed"])

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
    client_data = client_data.sort_values("Start_Date")

    client_data["Start_Date"] = pd.to_datetime(
        client_data["Start_Date"]
    ).dt.strftime("%d/%m/%Y")

    client_data["End_Date"] = pd.to_datetime(
        client_data["End_Date"]
    ).dt.strftime("%d/%m/%Y")

    client_data["Rate"] = client_data["Rate"].round(0).astype(int)

    # ---------------- MOBILE VIEW -------------------
    if is_mobile:
        for _, r in client_data.iterrows():
            st.markdown(
                f"""
                <div style="
                border:1px solid #ddd;
                border-radius:12px;
                padding:14px;
                margin-bottom:12px;
                background:#fafafa">

                <b>Client:</b> {r['Client_Name']}<br>
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
        display_df = client_data[[
            "Proposal_ID",
            "Start_Date",
            "End_Date",
            "Proposal_Cost",
            "Final_Cost",
            "Profit",
            "Status"
        ]]

        st.dataframe(display_df, use_container_width=True)

# =====================================================
# ================= EXPORT DATA =======================
# =====================================================

if st.session_state.page == "Export Data":
    st.header("📤 Export Data")

    # ---------- SAFETY CHECK ----------
    if proposals_df.empty:
        st.warning("No data available to export.")
        st.stop()

    # ---------- EXPORT FUNCTION ----------
    def export_excel():
        import io
        import pandas as pd

        output = io.BytesIO()

        # Streamlit Cloud safe engine
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
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
        label="⬇️ Download Excel",
        data=excel_file,
        file_name="sigma_consultants_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ---------- OPTIONAL CSV BACKUP ----------
    st.markdown("### 🔁 Alternative Format")
    st.download_button(
        label="⬇️ Download CSV",
        data=proposals_df.to_csv(index=False),
        file_name="sigma_consultants_data.csv",
        mime="text/csv"
    )







