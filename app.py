import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

# 1. Premium Executive Theme Engine Setup
st.set_page_config(
    page_title="ZV STEELS – Executive Dashboard",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    .stApp { background-color: #0d1b2e; color: #f5f5f0; }
    h1, h2, h3 { color: #c9a227 !important; font-family: 'DM Sans', sans-serif; font-weight: 700; }
    div[data-testid="stMetricValue"] { color: #f5f5f0 !important; font-size: 26px !important; font-weight: bold; }
    div[data-testid="stMetricLabel"] { color: #8899aa !important; text-transform: uppercase; font-size: 11px; letter-spacing: 1px; }
    .dataframe { background-color: #162440 !important; border: 1px solid rgba(255,255,255,0.08) !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] { background-color: #162440; border-radius: 4px 4px 0px 0px; color: #8899aa; padding: 10px 20px; }
    .stTabs [aria-selected="true"] { background-color: #c9a227 !important; color: #0d1b2e !important; font-weight: bold; }
    </style>
""", unsafe_allow_index=True)

# 2. Live Dynamic Sheet Pipeline (Using Direct Excel Stream Format)
RAW_URL = "https://docs.google.com/spreadsheets/d/1IHtomumrXZrNv9nrn3q7wjT7M2Nj0Jj8TkasAH6kwrw/edit?usp=sharing"
EXPORT_URL = RAW_URL.split("/edit")[0] + "/export?format=xlsx"

@st.cache_data(ttl=15)
def load_all_sheets(url):
    try:
        return pd.ExcelFile(url)
    except Exception as e:
        return None

excel_data = load_all_sheets(EXPORT_URL)

def get_clean_df(excel_obj, sheet_name):
    if excel_obj is None or sheet_name not in excel_obj.sheet_names:
        return pd.DataFrame()
    try:
        df_raw = pd.read_excel(excel_obj, sheet_name=sheet_name, header=None)
        header_idx = 0
        for i in range(min(15, len(df_raw))):
            row_str = " ".join(df_raw.iloc[i].astype(str).tolist())
            if any(kw in row_str for kw in ["Week", "Entity", "Customer", "Salesperson", "Category"]):
                header_idx = i
                break
        
        df = pd.read_excel(excel_obj, sheet_name=sheet_name, skiprows=header_idx)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        df.columns = df.columns.str.strip()
        return df.dropna(how='all')
    except:
        return pd.DataFrame()

# Extract tables
df_summary = get_clean_df(excel_data, "Weekly Summary")
df_bifurcation = get_clean_df(excel_data, "Sales Bifurcation")
df_crr_team = get_clean_df(excel_data, "CRR Consolidated")
df_stock_interest = get_clean_df(excel_data, "Stock Interest")

# Dashboard Banner Header
st.title("ZV STEELS")
st.caption("Performance Activity Dashboard • Live Sync Engine")

if df_summary.empty:
    st.warning("⚠️ Access stream initializing... Please verify your Google Sheet is set to 'Anyone with the link can view'.")
else:
    first_col = df_summary.columns[0]
    df_w1_search = df_summary[df_summary[first_col].astype(str).str.contains('Week 1', na=False, case=False)]
    
    def get_col_val(df_source, keywords, default=0.0):
        if df_source.empty:
            return default
        for col in df_source.columns:
            if all(kw.lower() in col.lower() for kw in keywords):
                try:
                    val_raw = str(df_source[col].iloc[0]).replace('₹','').replace(',','').strip()
                    return float(val_raw)
                except:
                    return default
        return default

    # Metrics
    val_crr = get_col_val(df_w1_search, ["CRR", "Actual"])
    diff_crr = get_col_val(df_w1_search, ["CRR", "Diff"])
    val_nbd = get_col_val(df_w1_search, ["NBD", "Actual"])
    diff_nbd = get_col_val(df_w1_search, ["NBD", "Diff"])
    val_coll = get_col_val(df_w1_search, ["Collection", "Actual"])
    diff_coll = get_col_val(df_w1_search, ["Collection", "Diff"])
    
    raw_otd = get_col_val(df_w1_search, ["OTD"])
    otd_pct = raw_otd * 100 if raw_otd <= 1.0 else raw_otd

    # Render Executive Ledger
    st.markdown("### ◆ Executive KPI Ledger")
    m1, m2, m3 = st.columns(3)
    
    with m1:
        st.metric(label="CRR Sales (Actual)", value=f"₹ {val_crr:.2f} Cr", 
                  delta=f"{diff_crr:.2f} Cr Shortfall" if diff_crr < 0 else f"▲ {diff_crr:.2f} Cr",
                  delta_color="normal" if diff_crr >= 0 else "inverse")
    with m2:
        st.metric(label="NBD Sales (Actual)", value=f"₹ {val_nbd:.2f} Cr", 
                  delta=f"{diff_nbd:.2f} Cr Shortfall" if diff_nbd < 0 else f"▲ {diff_nbd:.2f} Cr",
                  delta_color="normal" if diff_nbd >= 0 else "inverse")
    with m3:
        st.metric(label="Financial Collections", value=f"₹ {val_coll:.2f} Cr", 
                  delta=f"{diff_coll:.2f} Cr Shortfall" if diff_coll < 0 else f"▲ {diff_coll:.2f} Cr",
                  delta_color="normal" if diff_coll >= 0 else "inverse")

    # Metrics Row 2
    st.markdown("### ◆ Logistics & Penalties Risk Metrics")
    m4, m5, m6 = st.columns(3)
    
    with m4:
        st.metric(label="On-Time Delivery Rate", value=f"{otd_pct:.1f}%", delta="Target threshold: 80%")
    with m5:
        interest_sum = 0.0
        if not df_stock_interest.empty:
            for col in df_stock_interest.columns:
                if "Interest" in col or "Loss" in col:
                    interest_sum = pd.to_numeric(df_stock_interest[col], errors='coerce').sum()
                    break
        st.metric(label="Overdue Interest Loss Leakage", value=f"₹ {interest_sum:,.2f}", delta="Action Required", delta_color="inverse")
    with m6:
        st.metric(label="Total Inventory Ageing Status", value="9,549.64 MT", delta="Raw Material + FG Stock Status", delta_color="off")

    # Chart
    st.markdown("### ◆ Weekly Performance Realization Graph")
    df_weeks_only = df_summary[df_summary[first_col].astype(str).str.contains('Week', na=False, case=False)]
    
    fig = go.Figure()
    if not df_weeks_only.empty:
        col_plan = [c for c in df_weeks_only.columns if "CRR" in c and "Plan" in c]
        col_act = [c for c in df_weeks_only.columns if "CRR" in c and "Actual" in c]
        if col_plan:
            fig.add_trace(go.Bar(x=df_weeks_only[first_col], y=df_weeks_only[col_plan[0]], name="Plan Target Profile", marker_color='#4a8fd4'))
        if col_act:
            fig.add_trace(go.Bar(x=df_weeks_only[first_col], y=df_weeks_only[col_act[0]], name="Actual Performance Realized", marker_color='#c9a227'))
    
    fig.update_layout(barmode='group', margin=dict(l=10, r=10, t=15, b=10), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="#f5f5f0")
    st.plotly_chart(fig, use_container_width=True)

# Tabs
st.markdown("### ◆ Granular Operational Ledgers")
tab_corp, tab_team, tab_stock = st.tabs(["🏢 Corporate Entity Bifurcation", "👥 Sales Team Performance", "📋 Penalty Accounts"])

with tab_corp:
    st.markdown("#### ZVS & ZVC Sales Breakdown Metrics")
    if not df_bifurcation.empty: st.dataframe(df_bifurcation, use_container_width=True, hide_index=True)
    else: st.info("Awaiting live corporate entity data synchronization...")
with tab_team:
    st.markdown("#### Salesperson Performance Tracking Profile")
    if not df_crr_team.empty: st.dataframe(df_crr_team, use_container_width=True, hide_index=True)
    else: st.info("Awaiting live team sales log data synchronization...")
with tab_stock:
    st.markdown("#### Overdue Balances Accounts Audit Ledger")
    if not df_stock_interest.empty: st.dataframe(df_stock_interest, use_container_width=True, hide_index=True)
    else: st.info("Awaiting live risk ledger penalty tracking records...")
