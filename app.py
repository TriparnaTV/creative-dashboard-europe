import pandas as pd
import numpy as np
import streamlit as st
from pathlib import Path

# ============================================================
# CONFIG
# ============================================================
st.set_page_config(page_title="Europe Creative Dashboard", layout="wide")

# If Excel is in the same folder as app.py, keep this:
FILE_PATH = "Europe_dashboard.xlsx"

# If needed, replace with full path:
# FILE_PATH = r"D:\OneDrive - Transformative Ventures Pvt. Ltd\Desktop\Creative_dashboard\Europe_dashboard.xlsx"

EURO = "€"

# ============================================================
# STYLING
# ============================================================
st.markdown(
    """
    <style>
    .main {
        background-color: #0f1117;
        color: white;
    }
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 1rem;
    }
    .kpi-card {
        background: #171923;
        border: 1px solid #2a2f3a;
        border-radius: 12px;
        padding: 14px 18px;
        margin-bottom: 10px;
    }
    .kpi-title {
        font-size: 13px;
        color: #b9c0cc;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 22px;
        font-weight: 700;
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# HELPERS
# ============================================================
def excel_col_to_idx(col_letters: str) -> int:
    """Convert Excel column letters to 0-based index. Example: A->0, Z->25, AM->38"""
    col_letters = col_letters.upper().strip()
    idx = 0
    for char in col_letters:
        idx = idx * 26 + (ord(char) - ord("A") + 1)
    return idx - 1


def safe_divide(a, b):
    return np.where((b == 0) | pd.isna(b), 0, a / b)


def format_int(x):
    return f"{int(round(x)):,}" if pd.notna(x) else "0"


def format_currency(x):
    return f"{EURO}{x:,.2f}" if pd.notna(x) else f"{EURO}0.00"


def format_percent(x):
    return f"{x:.2%}" if pd.notna(x) else "0.00%"


def format_ratio(x):
    return f"{x:.2f}x" if pd.notna(x) else "0.00x"


def add_period_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["month_start"] = out["date"].dt.to_period("M").dt.start_time
    out["month_label"] = out["month_start"].dt.strftime("%Y-%m")

    iso = out["date"].dt.isocalendar()
    out["iso_year"] = iso["year"].astype(int)
    out["iso_week"] = iso["week"].astype(int)
    out["week_label"] = out["iso_year"].astype(str) + "-W" + out["iso_week"].astype(str).str.zfill(2)

    # Monday of ISO week
    out["week_start"] = pd.to_datetime(
        out["iso_year"].astype(str) + "-W" + out["iso_week"].astype(str).str.zfill(2) + "-1",
        format="%G-W%V-%u",
        errors="coerce"
    )
    return out


@st.cache_data
def load_data(file_path: str) -> pd.DataFrame:
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    raw = pd.read_excel(path, engine="openpyxl")

    col_map_by_letter = {
        "AM": "date",
        "K": "store",
        "D": "product",
        "AA": "audience_type",
        "Z": "creative_format",
        "G": "ad_name",
        "E": "campaign",
        "AE": "campaign_type",
        "L": "spend",
        "M": "revenue",
        "N": "orders",
        "O": "clicks",
        "P": "impressions",
        "U": "nb_revenue",
        "V": "b_revenue",
        "W": "nb_spend",
        "X": "b_spend",
    }

    rename_map = {}
    for excel_letter, clean_name in col_map_by_letter.items():
        idx = excel_col_to_idx(excel_letter)
        if idx >= len(raw.columns):
            raise ValueError(
                f"Expected Excel column {excel_letter} at position {idx}, "
                f"but sheet only has {len(raw.columns)} columns."
            )
        actual_col_name = raw.columns[idx]
        rename_map[actual_col_name] = clean_name

    df = raw.rename(columns=rename_map)[list(rename_map.values())].copy()

    # Clean strings
    str_cols = [
        "store", "product", "audience_type", "creative_format",
        "ad_name", "campaign", "campaign_type"
    ]
    for c in str_cols:
        df[c] = df[c].astype(str).str.strip()
        df[c] = df[c].replace({"nan": np.nan, "None": np.nan, "": np.nan})

    # Dates
    df["date"] = pd.to_datetime(df["date"], errors="coerce")

    # Numerics
    numeric_cols = [
        "spend", "revenue", "orders", "clicks", "impressions",
        "nb_revenue", "b_revenue", "nb_spend", "b_spend"
    ]
    for c in numeric_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # Standardize text
    df["store"] = df["store"].str.upper()
    df["creative_format"] = df["creative_format"].str.title()
    df["campaign_type"] = df["campaign_type"].str.title()
    df["audience_type"] = df["audience_type"].str.replace("_", " ", regex=False).str.title()

    # Remove bad rows
    df = df.dropna(subset=["date", "creative_format", "ad_name"]).copy()

    # Add month/week fields
    df = add_period_columns(df)

    return df


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Filters")

    min_date = df["date"].min().date()
    max_date = df["date"].max().date()

    date_range = st.sidebar.date_input(
        "Date Range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )

    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date, end_date = min_date, max_date

    stores = sorted(x for x in df["store"].dropna().unique())
    products = sorted(x for x in df["product"].dropna().unique())
    campaign_types = sorted(x for x in df["campaign_type"].dropna().unique())
    audience_types = sorted(x for x in df["audience_type"].dropna().unique())

    selected_stores = st.sidebar.multiselect("Store / Country", stores, default=stores)
    selected_products = st.sidebar.multiselect("Product", products, default=products)
    selected_campaign_types = st.sidebar.multiselect("Campaign Type", campaign_types, default=campaign_types)
    selected_audience_types = st.sidebar.multiselect("Buyer / Non-Buyer", audience_types, default=audience_types)

    filtered = df[
        (df["date"].dt.date >= start_date) &
        (df["date"].dt.date <= end_date) &
        (df["store"].isin(selected_stores)) &
        (df["product"].isin(selected_products)) &
        (df["campaign_type"].isin(selected_campaign_types)) &
        (df["audience_type"].isin(selected_audience_types))
    ].copy()

    return filtered


def build_summary(filtered: pd.DataFrame) -> pd.DataFrame:
    if filtered.empty:
        return pd.DataFrame()

    summary = (
        filtered.groupby("creative_format", dropna=False)
        .agg(
            Ads=("ad_name", "nunique"),
            Spend=("spend", "sum"),
            Revenue=("revenue", "sum"),
            Orders=("orders", "sum"),
            Clicks=("clicks", "sum"),
            Impressions=("impressions", "sum"),
            NB_Spend=("nb_spend", "sum"),
            B_Spend=("b_spend", "sum"),
            NB_Revenue=("nb_revenue", "sum"),
            B_Revenue=("b_revenue", "sum"),
        )
        .reset_index()
        .rename(columns={"creative_format": "Creative"})
    )

    total_ads = summary["Ads"].sum()
    total_spend = summary["Spend"].sum()
    total_b_spend = summary["B_Spend"].sum()
    total_nb_spend = summary["NB_Spend"].sum()

    summary["Ads %"] = safe_divide(summary["Ads"], total_ads)
    summary["Spend %"] = safe_divide(summary["Spend"], total_spend)
    summary["B-Spend %"] = safe_divide(summary["B_Spend"], total_b_spend)
    summary["NB-Spend %"] = safe_divide(summary["NB_Spend"], total_nb_spend)

    summary["B-ROAS"] = safe_divide(summary["B_Revenue"], summary["B_Spend"])
    summary["NB-ROAS"] = safe_divide(summary["NB_Revenue"], summary["NB_Spend"])
    summary["CTR"] = safe_divide(summary["Clicks"], summary["Impressions"])
    summary["CPM"] = safe_divide(summary["Spend"] * 1000, summary["Impressions"])
    summary["CPC"] = safe_divide(summary["Spend"], summary["Clicks"])
    summary["CVR"] = safe_divide(summary["Orders"], summary["Clicks"])
    summary["AOV"] = safe_divide(summary["Revenue"], summary["Orders"])
    summary["EPC"] = safe_divide(summary["Revenue"], summary["Clicks"])

    summary = summary.sort_values("Spend", ascending=False).reset_index(drop=True)

    total_row = pd.DataFrame([{
        "Creative": "Total",
        "Ads": summary["Ads"].sum(),
        "Spend": summary["Spend"].sum(),
        "Revenue": summary["Revenue"].sum(),
        "Orders": summary["Orders"].sum(),
        "Clicks": summary["Clicks"].sum(),
        "Impressions": summary["Impressions"].sum(),
        "NB_Spend": summary["NB_Spend"].sum(),
        "B_Spend": summary["B_Spend"].sum(),
        "NB_Revenue": summary["NB_Revenue"].sum(),
        "B_Revenue": summary["B_Revenue"].sum(),
    }])

    total_row["Ads %"] = safe_divide(total_row["Ads"], total_row["Ads"])
    total_row["Spend %"] = safe_divide(total_row["Spend"], total_row["Spend"])
    total_row["B-Spend %"] = safe_divide(total_row["B_Spend"], total_row["B_Spend"])
    total_row["NB-Spend %"] = safe_divide(total_row["NB_Spend"], total_row["NB_Spend"])
    total_row["B-ROAS"] = safe_divide(total_row["B_Revenue"], total_row["B_Spend"])
    total_row["NB-ROAS"] = safe_divide(total_row["NB_Revenue"], total_row["NB_Spend"])
    total_row["CTR"] = safe_divide(total_row["Clicks"], total_row["Impressions"])
    total_row["CPM"] = safe_divide(total_row["Spend"] * 1000, total_row["Impressions"])
    total_row["CPC"] = safe_divide(total_row["Spend"], total_row["Clicks"])
    total_row["CVR"] = safe_divide(total_row["Orders"], total_row["Clicks"])
    total_row["AOV"] = safe_divide(total_row["Revenue"], total_row["Orders"])
    total_row["EPC"] = safe_divide(total_row["Revenue"], total_row["Clicks"])

    summary = pd.concat([summary, total_row], ignore_index=True)
    return summary


def build_time_summary(filtered: pd.DataFrame, period_type: str) -> pd.DataFrame:
    if filtered.empty:
        return pd.DataFrame()

    if period_type == "monthly":
        period_col = "month_label"
        sort_col = "month_start"
    elif period_type == "weekly":
        period_col = "week_label"
        sort_col = "week_start"
    else:
        raise ValueError("period_type must be 'monthly' or 'weekly'")

    grouped = (
        filtered.groupby([period_col, "creative_format", sort_col], dropna=False)
        .agg(
            Ads=("ad_name", "nunique"),
            Spend=("spend", "sum"),
            Revenue=("revenue", "sum"),
            Orders=("orders", "sum"),
            Clicks=("clicks", "sum"),
            Impressions=("impressions", "sum"),
            NB_Spend=("nb_spend", "sum"),
            B_Spend=("b_spend", "sum"),
            NB_Revenue=("nb_revenue", "sum"),
            B_Revenue=("b_revenue", "sum"),
        )
        .reset_index()
        .rename(columns={period_col: "Period", "creative_format": "Creative"})
    )

    grouped["B-ROAS"] = safe_divide(grouped["B_Revenue"], grouped["B_Spend"])
    grouped["NB-ROAS"] = safe_divide(grouped["NB_Revenue"], grouped["NB_Spend"])
    grouped["CTR"] = safe_divide(grouped["Clicks"], grouped["Impressions"])
    grouped["CPM"] = safe_divide(grouped["Spend"] * 1000, grouped["Impressions"])
    grouped["CPC"] = safe_divide(grouped["Spend"], grouped["Clicks"])
    grouped["CVR"] = safe_divide(grouped["Orders"], grouped["Clicks"])
    grouped["AOV"] = safe_divide(grouped["Revenue"], grouped["Orders"])
    grouped["EPC"] = safe_divide(grouped["Revenue"], grouped["Clicks"])

    grouped = grouped.sort_values([sort_col, "Spend"], ascending=[True, False]).reset_index(drop=True)
    return grouped.drop(columns=[sort_col])


def build_display_table(summary: pd.DataFrame) -> pd.DataFrame:
    if summary.empty:
        return summary

    display = summary.copy()

    display["Ads"] = display["Ads"].apply(format_int)
    display["Ads %"] = display["Ads %"].apply(format_percent)
    display["Spend"] = display["Spend"].apply(format_currency)
    display["Spend %"] = display["Spend %"].apply(format_percent)
    display["B_Spend"] = display["B_Spend"].apply(format_currency)
    display["B-Spend %"] = display["B-Spend %"].apply(format_percent)
    display["B-ROAS"] = display["B-ROAS"].apply(format_ratio)
    display["NB_Spend"] = display["NB_Spend"].apply(format_currency)
    display["NB-Spend %"] = display["NB-Spend %"].apply(format_percent)
    display["NB-ROAS"] = display["NB-ROAS"].apply(format_ratio)
    display["CTR"] = display["CTR"].apply(format_percent)
    display["CPM"] = display["CPM"].apply(format_currency)
    display["CPC"] = display["CPC"].apply(format_currency)
    display["CVR"] = display["CVR"].apply(format_percent)
    display["AOV"] = display["AOV"].apply(format_currency)
    display["EPC"] = display["EPC"].apply(format_currency)

    display = display.rename(columns={
        "B_Spend": "B-Spends",
        "NB_Spend": "NB-Spends"
    })

    final_cols = [
        "Creative", "Ads", "Ads %", "Spend", "Spend %",
        "B-Spends", "B-Spend %", "B-ROAS",
        "NB-Spends", "NB-Spend %", "NB-ROAS",
        "CTR", "CPM", "CPC", "CVR", "AOV", "EPC"
    ]
    return display[final_cols]


def build_display_time_table(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    display = df.copy()

    display["Ads"] = display["Ads"].apply(format_int)
    display["Spend"] = display["Spend"].apply(format_currency)
    display["Revenue"] = display["Revenue"].apply(format_currency)
    display["Orders"] = display["Orders"].apply(format_int)
    display["Clicks"] = display["Clicks"].apply(format_int)
    display["Impressions"] = display["Impressions"].apply(format_int)
    display["B_Spend"] = display["B_Spend"].apply(format_currency)
    display["NB_Spend"] = display["NB_Spend"].apply(format_currency)
    display["B-ROAS"] = display["B-ROAS"].apply(format_ratio)
    display["NB-ROAS"] = display["NB-ROAS"].apply(format_ratio)
    display["CTR"] = display["CTR"].apply(format_percent)
    display["CPM"] = display["CPM"].apply(format_currency)
    display["CPC"] = display["CPC"].apply(format_currency)
    display["CVR"] = display["CVR"].apply(format_percent)
    display["AOV"] = display["AOV"].apply(format_currency)
    display["EPC"] = display["EPC"].apply(format_currency)

    display = display.rename(columns={
        "B_Spend": "B-Spends",
        "NB_Spend": "NB-Spends"
    })

    final_cols = [
        "Period", "Creative", "Ads", "Spend", "Revenue", "Orders",
        "Clicks", "Impressions", "B-Spends", "NB-Spends",
        "B-ROAS", "NB-ROAS", "CTR", "CPM", "CPC", "CVR", "AOV", "EPC"
    ]
    return display[final_cols]


def build_metric_pivot(df: pd.DataFrame, metric: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    pivot = df.pivot_table(
        index="Period",
        columns="Creative",
        values=metric,
        aggfunc="sum"
    ).reset_index()

    return pivot


def render_kpis(filtered: pd.DataFrame):
    total_ads = filtered["ad_name"].nunique()
    total_spend = filtered["spend"].sum()
    total_revenue = filtered["revenue"].sum()
    total_orders = filtered["orders"].sum()
    total_clicks = filtered["clicks"].sum()
    total_impressions = filtered["impressions"].sum()

    ctr = (total_clicks / total_impressions) if total_impressions else 0
    roas = (total_revenue / total_spend) if total_spend else 0

    c1, c2, c3, c4, c5, c6 = st.columns(6)

    with c1:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-title">Distinct Ads</div><div class="kpi-value">{format_int(total_ads)}</div></div>',
            unsafe_allow_html=True
        )
    with c2:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-title">Spend</div><div class="kpi-value">{format_currency(total_spend)}</div></div>',
            unsafe_allow_html=True
        )
    with c3:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-title">Revenue</div><div class="kpi-value">{format_currency(total_revenue)}</div></div>',
            unsafe_allow_html=True
        )
    with c4:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-title">Orders</div><div class="kpi-value">{format_int(total_orders)}</div></div>',
            unsafe_allow_html=True
        )
    with c5:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-title">CTR</div><div class="kpi-value">{format_percent(ctr)}</div></div>',
            unsafe_allow_html=True
        )
    with c6:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-title">ROAS</div><div class="kpi-value">{format_ratio(roas)}</div></div>',
            unsafe_allow_html=True
        )


def highlight_total(row):
    if row["Creative"] == "Total":
        return ["background-color: #262b36; font-weight: bold; color: white;" for _ in row]
    return [""] * len(row)


# ============================================================
# APP
# ============================================================
st.title("Europe Creative Format Dashboard")

try:
    df = load_data(FILE_PATH)
except Exception as e:
    st.error(f"Failed to load Excel file.\n\n{e}")
    st.stop()

filtered_df = apply_filters(df)

if filtered_df.empty:
    st.warning("No data found for the selected filters.")
    st.stop()

with st.expander("Data Validation", expanded=False):
    spend_gap = (filtered_df["spend"] - (filtered_df["b_spend"] + filtered_df["nb_spend"])).abs().sum()
    revenue_gap = (filtered_df["revenue"] - (filtered_df["b_revenue"] + filtered_df["nb_revenue"])).abs().sum()

    c1, c2, c3 = st.columns(3)
    c1.metric("Rows", f"{len(filtered_df):,}")
    c2.metric("Spend Mismatch Sum", f"{spend_gap:,.2f}")
    c3.metric("Revenue Mismatch Sum", f"{revenue_gap:,.2f}")

    st.write("Distinct Ads:", f"{filtered_df['ad_name'].nunique():,}")
    st.write("Spend:", format_currency(filtered_df["spend"].sum()))
    st.write("Revenue:", format_currency(filtered_df["revenue"].sum()))
    st.write("Orders:", f"{int(filtered_df['orders'].sum()):,}")
    st.write("Clicks:", f"{int(filtered_df['clicks'].sum()):,}")
    st.write("Impressions:", f"{int(filtered_df['impressions'].sum()):,}")
    st.write("Date Range:", f"{filtered_df['date'].min().date()} to {filtered_df['date'].max().date()}")

render_kpis(filtered_df)

tab1, tab2, tab3 = st.tabs(["Overview", "Monthly by Ad Format", "Weekly by Ad Format"])

with tab1:
    summary = build_summary(filtered_df)
    display_table = build_display_table(summary)

    st.subheader("Creative Performance Summary")
    styled = display_table.style.apply(highlight_total, axis=1)
    st.dataframe(styled, use_container_width=True, height=600)

    csv_data = summary.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download Summary CSV",
        data=csv_data,
        file_name="creative_dashboard_summary.csv",
        mime="text/csv",
    )

with tab2:
    st.subheader("Monthly Performance by Ad Format")

    monthly_summary = build_time_summary(filtered_df, "monthly")
    monthly_display = build_display_time_table(monthly_summary)
    st.dataframe(monthly_display, use_container_width=True, height=600)

    st.markdown("### Monthly Spend Pivot")
    monthly_spend_pivot = build_metric_pivot(monthly_summary, "Spend")
    if not monthly_spend_pivot.empty:
        monthly_spend_pivot_fmt = monthly_spend_pivot.copy()
        for col in monthly_spend_pivot_fmt.columns:
            if col != "Period":
                monthly_spend_pivot_fmt[col] = monthly_spend_pivot_fmt[col].apply(format_currency)
        st.dataframe(monthly_spend_pivot_fmt, use_container_width=True)

    st.markdown("### Monthly Revenue Pivot")
    monthly_revenue_pivot = build_metric_pivot(monthly_summary, "Revenue")
    if not monthly_revenue_pivot.empty:
        monthly_revenue_pivot_fmt = monthly_revenue_pivot.copy()
        for col in monthly_revenue_pivot_fmt.columns:
            if col != "Period":
                monthly_revenue_pivot_fmt[col] = monthly_revenue_pivot_fmt[col].apply(format_currency)
        st.dataframe(monthly_revenue_pivot_fmt, use_container_width=True)

    monthly_csv = monthly_summary.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download Monthly CSV",
        data=monthly_csv,
        file_name="creative_dashboard_monthly.csv",
        mime="text/csv",
        key="monthly_csv_btn",
    )

with tab3:
    st.subheader("Weekly Performance by Ad Format")

    weekly_summary = build_time_summary(filtered_df, "weekly")
    weekly_display = build_display_time_table(weekly_summary)
    st.dataframe(weekly_display, use_container_width=True, height=600)

    st.markdown("### Weekly Spend Pivot")
    weekly_spend_pivot = build_metric_pivot(weekly_summary, "Spend")
    if not weekly_spend_pivot.empty:
        weekly_spend_pivot_fmt = weekly_spend_pivot.copy()
        for col in weekly_spend_pivot_fmt.columns:
            if col != "Period":
                weekly_spend_pivot_fmt[col] = weekly_spend_pivot_fmt[col].apply(format_currency)
        st.dataframe(weekly_spend_pivot_fmt, use_container_width=True)

    st.markdown("### Weekly Revenue Pivot")
    weekly_revenue_pivot = build_metric_pivot(weekly_summary, "Revenue")
    if not weekly_revenue_pivot.empty:
        weekly_revenue_pivot_fmt = weekly_revenue_pivot.copy()
        for col in weekly_revenue_pivot_fmt.columns:
            if col != "Period":
                weekly_revenue_pivot_fmt[col] = weekly_revenue_pivot_fmt[col].apply(format_currency)
        st.dataframe(weekly_revenue_pivot_fmt, use_container_width=True)

    weekly_csv = weekly_summary.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download Weekly CSV",
        data=weekly_csv,
        file_name="creative_dashboard_weekly.csv",
        mime="text/csv",
        key="weekly_csv_btn",
    )