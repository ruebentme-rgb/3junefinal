import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import re
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="APRIL Group – Bilateral Banking Facilities",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Custom CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a3a5c 0%, #2e6da4 100%);
        padding: 18px 24px; border-radius: 10px; margin-bottom: 20px;
        color: white;
    }
    .main-header h1 { margin: 0; font-size: 1.6rem; font-weight: 700; }
    .main-header p  { margin: 4px 0 0; font-size: 0.85rem; opacity: 0.85; }
    .metric-card {
        background: #f8fafc; border: 1px solid #e2e8f0;
        border-radius: 8px; padding: 16px; text-align: center;
    }
    .metric-card .value { font-size: 2rem; font-weight: 700; color: #1a3a5c; }
    .metric-card .label { font-size: 0.8rem; color: #64748b; margin-top: 4px; }
    .expired-row { background-color: #fff1f0 !important; }
    .renewal-badge {
        background: #ff4d4f; color: white; padding: 2px 8px;
        border-radius: 12px; font-size: 0.75rem; font-weight: 600;
    }
    .current-badge {
        background: #52c41a; color: white; padding: 2px 8px;
        border-radius: 12px; font-size: 0.75rem; font-weight: 600;
    }
    .negotiating-badge {
        background: #fa8c16; color: white; padding: 2px 8px;
        border-radius: 12px; font-size: 0.75rem; font-weight: 600;
    }
    .stDataFrame { font-size: 0.85rem; }
    div[data-testid="stSidebarNav"] { display: none; }
    .sidebar-section { font-size: 0.78rem; color: #94a3b8;
        text-transform: uppercase; letter-spacing: 0.08em;
        margin: 16px 0 6px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# ─── Data Loading ─────────────────────────────────────────────────────────────
TODAY = date.today()

@st.cache_data
def load_data(filepath: str) -> pd.DataFrame:
    raw = pd.read_excel(filepath, sheet_name="Sheet1", header=None)
    raw.columns = ["Bank", "Office", "Borrower", "Date_of_Facility",
                   "Currency", "Quantum", "Facility_Type", "Tenure",
                   "Status", "Security_Package"]
    df = raw.iloc[1:].copy().reset_index(drop=True)

    # ── Office mapping ────────────────────────────────────────────────────────
    df["Office"] = df["Office"].fillna("").str.strip()
    office_map = {"HK": "Hong Kong", "SG": "Singapore", "UAE": "UAE"}
    df["Office"] = df["Office"].map(office_map).fillna(df["Office"])

    # ── Clean Borrower (multi-line → list-like string) ────────────────────────
    df["Borrower"] = df["Borrower"].fillna("").astype(str).str.replace("\n", ", ")

    # ── Clean Security / Guarantees ───────────────────────────────────────────
    df["Security_Package"] = (df["Security_Package"]
        .fillna("Nil").astype(str).str.replace("\n", " | "))

    # ── Parse Tenure for expiry detection ────────────────────────────────────
    def parse_tenure(val):
        if pd.isna(val):
            return pd.NaT
        s = str(val).strip()
        try:
            return pd.to_datetime(s)
        except Exception:
            return pd.NaT

    df["Tenure_Date"] = df["Tenure"].apply(parse_tenure)

    # ── Expired flag ──────────────────────────────────────────────────────────
    df["Expired"] = df["Tenure_Date"].apply(
        lambda x: (x.date() < TODAY) if pd.notna(x) else False
    )

    # ── Parse Quantum ─────────────────────────────────────────────────────────
    def parse_quantum(val):
        if pd.isna(val) or str(val).strip() in ("", "NaN", "Negotiating"):
            return np.nan
        s = str(val).split("\n")[0].replace(",", "").strip()
        try:
            return float(s)
        except Exception:
            return np.nan

    df["Quantum_Num"] = df["Quantum"].apply(parse_quantum)

    # ── USD-equivalent quantum (rough: treat all as USD for analytics) ────────
    df["USD_Quantum"] = df["Quantum_Num"]   # currency-mixed; flagged in UI

    # ── Facility Type clean ───────────────────────────────────────────────────
    df["Facility_Type"] = df["Facility_Type"].fillna("Unknown").astype(str).str.strip()

    # ── Status badge ─────────────────────────────────────────────────────────
    df["Status"] = df["Status"].fillna("Unknown").astype(str).str.strip()

    # ── Extract guarantors from Security_Package ──────────────────────────────
    def extract_guarantors(s):
        tokens = re.findall(r"(APRIL(?:HL)?|APRIHL|SGF|DFF|MFF|PVHL|AFPT|Clean)\s+(?:Guarantee|dd|Security)",
                            str(s), flags=re.IGNORECASE)
        guarantors = set()
        for t in re.findall(r"(APRIL(?:HL)?|APRIHL|SGF|DFF|MFF|PVHL|AFPT)\s+(?:Guarantee|dd)",
                            str(s), flags=re.IGNORECASE):
            guarantors.add(t.strip().upper())
        if "Clean" in str(s):
            guarantors.add("Clean (No Guarantee)")
        if not guarantors and str(s).strip().upper() in ("NIL", ""):
            guarantors.add("Nil")
        return sorted(guarantors)

    df["Guarantors"] = df["Security_Package"].apply(extract_guarantors)

    return df


df = load_data("/home/claude/APRIL_Bilateral_Facilities_dd_02_06_26.xlsx")

# ─── Sidebar – Filters ────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🔍 Filters")

    st.markdown('<div class="sidebar-section">Office</div>', unsafe_allow_html=True)
    all_offices = sorted(df["Office"].dropna().unique())
    sel_office = st.multiselect("Office", all_offices, default=all_offices, label_visibility="collapsed")

    st.markdown('<div class="sidebar-section">Borrower</div>', unsafe_allow_html=True)
    all_borrowers = sorted(set(
        b.strip()
        for cell in df["Borrower"]
        for b in str(cell).split(",")
        if b.strip()
    ))
    sel_borrower = st.multiselect("Borrower", all_borrowers, label_visibility="collapsed")

    st.markdown('<div class="sidebar-section">Bank</div>', unsafe_allow_html=True)
    all_banks = sorted(df["Bank"].dropna().unique())
    sel_bank = st.multiselect("Bank", all_banks, label_visibility="collapsed")

    st.markdown('<div class="sidebar-section">Facility Type</div>', unsafe_allow_html=True)
    all_types = sorted(df["Facility_Type"].dropna().unique())
    sel_type = st.multiselect("Facility Type", all_types, label_visibility="collapsed")

    st.markdown('<div class="sidebar-section">Guarantor</div>', unsafe_allow_html=True)
    all_guarantors = sorted(set(g for gs in df["Guarantors"] for g in gs if g))
    sel_guarantor = st.multiselect("Guarantor", all_guarantors, label_visibility="collapsed")

    st.markdown('<div class="sidebar-section">Status</div>', unsafe_allow_html=True)
    all_statuses = sorted(df["Status"].dropna().unique())
    sel_status = st.multiselect("Status", all_statuses, default=all_statuses, label_visibility="collapsed")

    st.markdown("---")
    show_expired_only = st.checkbox("⚠️ Show Expired Facilities Only", value=False)

# ─── Apply Filters ────────────────────────────────────────────────────────────
fdf = df.copy()

if sel_office:
    fdf = fdf[fdf["Office"].isin(sel_office)]
if sel_bank:
    fdf = fdf[fdf["Bank"].isin(sel_bank)]
if sel_type:
    fdf = fdf[fdf["Facility_Type"].isin(sel_type)]
if sel_status:
    fdf = fdf[fdf["Status"].isin(sel_status)]
if sel_borrower:
    fdf = fdf[fdf["Borrower"].apply(
        lambda x: any(b in x for b in sel_borrower)
    )]
if sel_guarantor:
    fdf = fdf[fdf["Guarantors"].apply(
        lambda gs: any(g in gs for g in sel_guarantor)
    )]
if show_expired_only:
    fdf = fdf[fdf["Expired"] == True]

# ─── Navigation Tabs ─────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>🏦 APRIL Group – Bilateral Banking Facilities</h1>
  <p>As at 2 June 2026 &nbsp;|&nbsp; Strategic Finance & Legal</p>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["📋 Facilities Register", "⚠️ Renewal Actions", "📊 Analytics"])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 – FACILITIES REGISTER
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:
    # KPI row
    col1, col2, col3, col4, col5 = st.columns(5)
    total = len(fdf)
    current_n = len(fdf[fdf["Status"] == "Current"])
    negotiating_n = len(fdf[fdf["Status"] == "Negotiating"])
    expired_n = int(fdf["Expired"].sum())
    total_usd = fdf["USD_Quantum"].sum()

    for col, val, lbl in zip(
        [col1, col2, col3, col4, col5],
        [total, current_n, negotiating_n, expired_n,
         f"USD {total_usd/1e6:.0f}M*" if not np.isnan(total_usd) else "—"],
        ["Total Facilities", "Current", "Negotiating",
         "⚠️ Expired Tenure", "Total Quantum"]
    ):
        with col:
            st.markdown(f"""
            <div class="metric-card">
                <div class="value">{val}</div>
                <div class="label">{lbl}</div>
            </div>""", unsafe_allow_html=True)

    st.caption("*Quantum shown in stated currency; not FX-converted. Mix of USD, EUR, INR.")
    st.markdown("---")

    # Build display dataframe
    def status_badge(s, expired):
        if expired:
            return "⚠️ RENEW"
        elif s == "Current":
            return "✅ Current"
        elif s == "Negotiating":
            return "🔄 Negotiating"
        return s

    display = fdf[[
        "Bank", "Office", "Borrower", "Date_of_Facility",
        "Currency", "Quantum", "Facility_Type",
        "Tenure", "Status", "Security_Package", "Expired"
    ]].copy()

    display["Status_Display"] = display.apply(
        lambda r: status_badge(r["Status"], r["Expired"]), axis=1)

    display["Date_of_Facility"] = display["Date_of_Facility"].astype(str).str.replace("00:00:00", "").str.strip()
    display["Tenure_Display"] = fdf["Tenure_Date"].apply(
        lambda x: x.strftime("%d %b %Y") if pd.notna(x) else
        str(fdf.loc[fdf["Tenure_Date"].isna(), "Tenure"].iloc[0]
            if False else "Undefined")
    )

    # Re-do tenure display properly
    def tenure_display(row):
        if pd.notna(row["Tenure_Date"]):
            label = row["Tenure_Date"].strftime("%d %b %Y")
            if row["Expired"]:
                return f"🔴 {label} (EXPIRED)"
            return label
        return str(row["Tenure"]) if str(row["Tenure"]) not in ("nan", "NaT") else "Undefined"

    display["Tenure_Show"] = fdf.apply(tenure_display, axis=1)

    show_cols = {
        "Bank": "Bank",
        "Office": "Office",
        "Borrower": "Borrower",
        "Date_of_Facility": "Facility Date",
        "Currency": "Ccy",
        "Quantum": "Quantum",
        "Facility_Type": "Type",
        "Tenure_Show": "Tenure / Expiry",
        "Status_Display": "Status",
        "Security_Package": "Security / Guarantee",
    }

    out = display[list(show_cols.keys())].rename(columns=show_cols)

    # Highlight expired rows
    def highlight_expired(row):
        color = "#fff1f0" if "EXPIRED" in str(row.get("Tenure / Expiry", "")) else ""
        return [f"background-color: {color}" if color else "" for _ in row]

    st.dataframe(
        out.style.apply(highlight_expired, axis=1),
        use_container_width=True,
        height=600,
        column_config={
            "Status": st.column_config.TextColumn(width="small"),
            "Ccy": st.column_config.TextColumn(width="small"),
            "Office": st.column_config.TextColumn(width="small"),
        }
    )

    st.caption(f"Showing {len(out)} of {len(df)} facilities after filters applied.")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 – RENEWAL ACTIONS
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    expired_df = df[df["Expired"] == True].copy()
    expiring_soon = df[
        (df["Tenure_Date"].notna()) &
        (~df["Expired"]) &
        (df["Tenure_Date"].apply(
            lambda x: (x.date() - TODAY).days <= 90 if pd.notna(x) else False))
    ].copy()

    st.subheader(f"🔴 Expired Facilities Requiring Renewal Action — {len(expired_df)} facility")

    if len(expired_df) > 0:
        for _, row in expired_df.iterrows():
            days_ago = (TODAY - row["Tenure_Date"].date()).days
            with st.expander(
                f"🔴 {row['Bank']} | {row['Borrower']} | {row['Facility_Type']} — "
                f"expired {days_ago} days ago ({row['Tenure_Date'].strftime('%d %b %Y')})"
            ):
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(f"**Bank:** {row['Bank']}")
                    st.markdown(f"**Office:** {row['Office']}")
                    st.markdown(f"**Borrower(s):** {row['Borrower']}")
                    st.markdown(f"**Facility Type:** {row['Facility_Type']}")
                with c2:
                    st.markdown(f"**Currency:** {row['Currency']}")
                    st.markdown(f"**Quantum:** {row['Quantum']}")
                    st.markdown(f"**Tenure Expired:** :red[{row['Tenure_Date'].strftime('%d %b %Y')}]")
                    st.markdown(f"**Status:** {row['Status']}")
                st.markdown(f"**Security / Guarantee:** {row['Security_Package']}")
                st.markdown(f"**Facility Date:** {str(row['Date_of_Facility']).replace('00:00:00','').strip()}")
    else:
        st.success("✅ No expired facilities in the current dataset.")

    st.markdown("---")
    st.subheader(f"🟡 Facilities Expiring Within 90 Days — {len(expiring_soon)} facilities")

    if len(expiring_soon) > 0:
        for _, row in expiring_soon.iterrows():
            days_left = (row["Tenure_Date"].date() - TODAY).days
            with st.expander(
                f"🟡 {row['Bank']} | {row['Borrower']} | expires in {days_left} days "
                f"({row['Tenure_Date'].strftime('%d %b %Y')})"
            ):
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(f"**Bank:** {row['Bank']}")
                    st.markdown(f"**Office:** {row['Office']}")
                    st.markdown(f"**Borrower(s):** {row['Borrower']}")
                    st.markdown(f"**Facility Type:** {row['Facility_Type']}")
                with c2:
                    st.markdown(f"**Currency:** {row['Currency']}")
                    st.markdown(f"**Quantum:** {row['Quantum']}")
                    st.markdown(f"**Tenure Expiry:** :orange[{row['Tenure_Date'].strftime('%d %b %Y')}]")
                    st.markdown(f"**Status:** {row['Status']}")
                st.markdown(f"**Security / Guarantee:** {row['Security_Package']}")
    else:
        st.success("✅ No facilities expiring within the next 90 days.")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 – ANALYTICS
# ═══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("📊 Portfolio Analytics")
    st.caption("Analytics based on full dataset (not filtered by sidebar selections).")

    adf = df.copy()

    # ── Row 1: Facilities by Office & Facility Type breakdown ────────────────
    r1c1, r1c2 = st.columns(2)

    with r1c1:
        st.markdown("**Facility Count by Office**")
        off_count = adf.groupby("Office").size().reset_index(name="Count")
        fig1 = px.pie(off_count, names="Office", values="Count",
                      color_discrete_sequence=px.colors.sequential.Blues_r,
                      hole=0.4)
        fig1.update_traces(textposition="outside", textinfo="percent+label")
        fig1.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=320,
                           showlegend=True)
        st.plotly_chart(fig1, use_container_width=True)

    with r1c2:
        st.markdown("**Facility Count by Type**")
        type_count = adf.groupby("Facility_Type").size().reset_index(name="Count").sort_values("Count")
        fig2 = px.bar(type_count, x="Count", y="Facility_Type", orientation="h",
                      color="Count", color_continuous_scale="Blues",
                      labels={"Facility_Type": ""})
        fig2.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=320,
                           coloraxis_showscale=False)
        st.plotly_chart(fig2, use_container_width=True)

    # ── Row 2: USD Quantum by Office ─────────────────────────────────────────
    st.markdown("---")
    st.markdown("**USD Quantum by Office** *(stated currency, not FX-adjusted)*")

    q_by_office = (adf[adf["Currency"] == "USD"]
                   .groupby("Office")["USD_Quantum"]
                   .sum()
                   .reset_index()
                   .sort_values("USD_Quantum", ascending=False))
    q_by_office["USD_Quantum_M"] = q_by_office["USD_Quantum"] / 1e6

    fig3 = px.bar(q_by_office, x="Office", y="USD_Quantum_M",
                  text=q_by_office["USD_Quantum_M"].apply(lambda x: f"USD {x:.0f}M"),
                  color="Office",
                  color_discrete_sequence=["#1a3a5c", "#2e6da4", "#5b9bd5"])
    fig3.update_traces(textposition="outside")
    fig3.update_layout(yaxis_title="USD (Millions)", xaxis_title="",
                       showlegend=False, height=350,
                       margin=dict(t=30, b=20))
    st.plotly_chart(fig3, use_container_width=True)

    # ── Row 3: Top 20 Banks by Facility Count ────────────────────────────────
    st.markdown("---")
    st.markdown("**Top 20 Banks by Facility Count**")

    bank_count = (adf.groupby(["Bank", "Office"])
                  .size()
                  .reset_index(name="Count")
                  .sort_values("Count", ascending=False)
                  .head(20))

    fig4 = px.bar(
        bank_count.sort_values("Count"),
        x="Count", y="Bank", orientation="h",
        color="Office",
        color_discrete_map={"Hong Kong": "#1a3a5c", "Singapore": "#2e6da4", "UAE": "#5b9bd5"},
        text="Count",
        labels={"Bank": ""}
    )
    fig4.update_traces(textposition="outside")
    fig4.update_layout(height=600, margin=dict(t=10, b=10, l=10, r=60),
                       legend_title_text="Office")
    st.plotly_chart(fig4, use_container_width=True)

    # ── Row 4: Top 20 Banks by USD Quantum ───────────────────────────────────
    st.markdown("---")
    st.markdown("**Top 20 Banks by USD Quantum** *(USD-denominated facilities only)*")

    bank_usd = (adf[adf["Currency"] == "USD"]
                .groupby(["Bank", "Office"])["USD_Quantum"]
                .sum()
                .reset_index()
                .sort_values("USD_Quantum", ascending=False)
                .head(20))
    bank_usd["USD_M"] = bank_usd["USD_Quantum"] / 1e6

    fig5 = px.bar(
        bank_usd.sort_values("USD_M"),
        x="USD_M", y="Bank", orientation="h",
        color="Office",
        color_discrete_map={"Hong Kong": "#1a3a5c", "Singapore": "#2e6da4", "UAE": "#5b9bd5"},
        text=bank_usd.sort_values("USD_M")["USD_M"].apply(lambda x: f"{x:.0f}M"),
        labels={"Bank": "", "USD_M": "USD Millions"}
    )
    fig5.update_traces(textposition="outside")
    fig5.update_layout(height=560, margin=dict(t=10, b=10, l=10, r=80),
                       legend_title_text="Office")
    st.plotly_chart(fig5, use_container_width=True)

    # ── Row 5: Status breakdown ───────────────────────────────────────────────
    st.markdown("---")
    r5c1, r5c2 = st.columns(2)

    with r5c1:
        st.markdown("**Facility Status Distribution**")
        status_ct = adf.groupby("Status").size().reset_index(name="Count")
        color_map = {"Current": "#52c41a", "Negotiating": "#fa8c16", "Unknown": "#8c8c8c"}
        fig6 = px.pie(status_ct, names="Status", values="Count",
                      color="Status", color_discrete_map=color_map, hole=0.4)
        fig6.update_traces(textposition="outside", textinfo="percent+label")
        fig6.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=300)
        st.plotly_chart(fig6, use_container_width=True)

    with r5c2:
        st.markdown("**Guarantor Frequency**")
        all_g = [g for gs in adf["Guarantors"] for g in gs if g not in ("Nil", "")]
        g_series = pd.Series(all_g).value_counts().reset_index()
        g_series.columns = ["Guarantor", "Count"]
        fig7 = px.bar(g_series, x="Count", y="Guarantor", orientation="h",
                      color="Count", color_continuous_scale="Blues",
                      labels={"Guarantor": ""})
        fig7.update_layout(height=300, margin=dict(t=10, b=10, l=10, r=10),
                           coloraxis_showscale=False)
        st.plotly_chart(fig7, use_container_width=True)

    # ── Summary table ─────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("**Summary Statistics by Office**")

    summary = adf.groupby("Office").agg(
        Facilities=("Bank", "count"),
        Banks=("Bank", "nunique"),
        USD_Facilities=("USD_Quantum", lambda x: x.notna().sum()),
        Total_USD_M=("USD_Quantum", lambda x: x.sum() / 1e6),
        Current=("Status", lambda x: (x == "Current").sum()),
        Negotiating=("Status", lambda x: (x == "Negotiating").sum()),
    ).reset_index()
    summary["Total_USD_M"] = summary["Total_USD_M"].apply(
        lambda x: f"USD {x:,.0f}M" if x > 0 else "—")
    summary.columns = ["Office", "Total Facilities", "Unique Banks",
                       "USD-Denominated", "Total USD Quantum",
                       "Current", "Negotiating"]
    st.dataframe(summary, use_container_width=True, hide_index=True)
