# db.py
from sqlalchemy import create_engine
import pandas as pd

# ================= DATABASE =================
DB_URL = "sqlite:///sigma_consultants.db"
engine = create_engine(DB_URL, echo=False)

# ================= LOADERS =================
def load_clients():
    try:
        return pd.read_sql("SELECT * FROM clients", engine)
    except Exception:
        return pd.DataFrame(columns=[
            "Client_ID","Client_Name","Created_Date",
            "Is_Archived","Notes"
        ])

def load_proposals():
    try:
        df = pd.read_sql("SELECT * FROM proposals", engine)
        for c in ["Start_Date","End_Date","Closing_Date"]:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], errors="coerce")
        return df
    except Exception:
        return pd.DataFrame(columns=[
            "Proposal_ID","Client_ID","Client_Name",
            "Proposal_Cost","Rate","Final_Cost","Profit",
            "Start_Date","End_Date","Status","Closing_Date"
        ])

# ================= SAVERS =================
def save_clients(df):
    df.to_sql("clients", engine, if_exists="replace", index=False)

def save_proposals(df):
    df.to_sql("proposals", engine, if_exists="replace", index=False)


