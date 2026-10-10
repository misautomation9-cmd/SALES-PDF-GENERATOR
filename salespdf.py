import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="Ironmart Sales Analytical Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------- CUSTOM CSS ----------------
st.markdown("""
<style>
    .main-header {
        font-size: 2rem;
        font-weight: 700;
        color: #1f4e79;
        text-align: center;
        padding: 10px;
        background: linear-gradient(90deg, #e8f0fe, #ffffff);
        border-radius: 10px;
        margin-bottom: 15px;
    }
    .kpi-card {
        background: linear-gradient(135deg, #667eea, #764ba2);
        color: white;
        padding: 15px;
        border-radius: 12px;
        text-align: center;
        box-shadow: 2px 4px 10px rgba(0,0,0,0.15);
        margin-bottom: 8px;
    }
    .kpi-card h4 { font-size: 0.85rem; margin: 0; opacity: 0.9; }
    .kpi-card h2 { font-size: 1.4rem; margin: 6px 0 0 0; }
    .kpi-card-orange { background: linear-gradient(135deg, #f7971e, #ffd200); }
    .kpi-card-red    { background: linear-gradient(135deg, #eb3349, #f45c43); }
    .kpi-card-green  { background: linear-gradient(135deg, #11998e, #38ef7d); }
    .kpi-card-blue   { background: linear-gradient(135deg, #2193b0, #6dd5ed); }
    .kpi-card-purple { background: linear-gradient(135deg, #8e2de2, #4a00e0); }
    .kpi-card-pink   { background: linear-gradient(135deg, #ec008c, #fc6767); }
</style>
""", unsafe_allow_html=True)


# ---------------- HELPER FUNCTIONS ----------------
@st.cache_data(show_spinner=False)
def load_all_sheets(file_bytes: bytes) -> dict:
    """Read every sheet of the Excel file into a dict of DataFrames (pickle-safe)."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = {}
    for name in xls.sheet_names:
        try:
            sheets[name] = pd.read_excel(xls, sheet_name=name)
        except Exception:
            sheets[name] = pd.DataFrame()
    return sheets


def normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Clean column names and standardise key ones."""
    df = df.copy()
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]

    mapping = {
        "S_NO": "S_NO", "S NO": "S_NO", "SNO": "S_NO",
        "PO_NO": "PO_NO", "PO NO": "PO_NO",
        "DO_NO": "DO_NO", "DO NO": "DO_NO",
        "SO_DATE": "SO_DATE", "SO DATE": "SO_DATE",
        "CUSTOMER_NAME": "CUSTOMER_NAME", "CUSTOMER NAME": "CUSTOMER_NAME",
        "BROKER": "BROKER",
        "SECTOR": "SECTOR",
        "PLACE": "PLACE",
        "SALES_EXECUTIVE": "SALES_EXECUTIVE", "SALES EXECUTIVE": "SALES_EXECUTIVE",
        "ITEM": "ITEM", "ITEM_NAME": "ITEM", "ITEM NAME": "ITEM",
        "THIKNESS": "THIKNESS", "THICKNESS": "THIKNESS",
        "SIZE_(MM)": "SIZE_(MM)", "SIZE (MM)": "SIZE_(MM)", "SIZE_MM": "SIZE_(MM)",
        "GRADE": "GRADE",
        "SO_QTY_(MT)": "SO_QTY_(MT)", "SO QTY (MT)": "SO_QTY_(MT)",
        "SO_QTY": "SO_QTY_(MT)", "SO QTY": "SO_QTY_(MT)",
        "PER_TON": "PER_TON", "PER TON": "PER_TON",
        "INVOICE_NO": "INVOICE_NO", "INVOICE NO": "INVOICE_NO",
        "INVOICE_DATE": "INVOICE_DATE", "INVOICE DATE": "INVOICE_DATE",
        "DISPATCH_QTY": "DISPATCH_QTY", "DISPATCH QTY": "DISPATCH_QTY",
        "PENDING": "PENDING", "PENDING_QTY": "PENDING",
        "PAYMENT": "PAYMENT",
        "DISPATCH_THROUGH": "DISPATCH_THROUGH", "DISPATCH THROUGH": "DISPATCH_THROUGH",
        "STATUS": "STATUS",
        "REMARK": "REMARK", "REMARKS": "REMARK",
        "MOBILE NO": "MOBILE_NO", "MOBILE_NO": "MOBILE_NO",
    }
    df.rename(columns={c: mapping.get(c.upper(), c) for c in df.columns}, inplace=True)

    # Numeric conversions
    for col in ["SO_QTY_(MT)", "PER_TON", "DISPATCH_QTY", "PENDING"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0.0

    # Date conversion
    if "SO_DATE" in df.columns:
        df["SO_DATE"] = pd.to_datetime(df["SO_DATE"], errors="coerce", dayfirst=True)
    else:
        df["SO_DATE"] = pd.NaT

    if "INVOICE_DATE" in df.columns:
        df["INVOICE_DATE"] = pd.to_datetime(df["INVOICE_DATE"], errors="coerce", dayfirst=True)
    else:
        df["INVOICE_DATE"] = pd.NaT

    # Total amount
    df["TOTAL_AMOUNT"] = df["SO_QTY_(MT)"] * df["PER_TON"]

    # Uppercase string columns
    for col in ["STATUS", "REMARK", "CUSTOMER_NAME", "SALES_EXECUTIVE", "ITEM",
                "GRADE", "SECTOR", "PLACE", "BROKER", "DISPATCH_THROUGH"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()
            df[col] = df[col].replace({"NAN": "", "NONE": ""})
        else:
            df[col] = ""

    # Extract width from SIZE_(MM) — "1250X5635" → 1250
    if "SIZE_(MM)" in df.columns:
        df["WIDTH"] = (
            df["SIZE_(MM)"].astype(str)
            .str.upper()
            .str.extract(r"(\d+)\s*X", expand=False)
        )
        df["WIDTH"] = pd.to_numeric(df["WIDTH"], errors="coerce")
    else:
        df["WIDTH"] = np.nan
        df["SIZE_(MM)"] = ""

    # Flags
    df["IS_CANCEL"] = df["STATUS"].str.contains("CANCEL", na=False)
    df["IS_SC"] = (
        df["STATUS"].str.contains(r"\bSC\b", na=False, regex=True) |
        df["REMARK"].str.contains(r"\bSC\b", na=False, regex=True)
    )
    df["IS_NEW"] = df["REMARK"].str.contains("NEW", na=False)
    df["IS_PENDING"] = df["STATUS"].str.contains("PENDING", na=False)

    return df


def kpi_metrics(df: pd.DataFrame) -> dict:
    """Return standard KPI dictionary for a dataframe."""
    return {
        "Unique Customers": df["CUSTOMER_NAME"].nunique(),
        "SO Qty (MT)": round(df["SO_QTY_(MT)"].sum(), 2),
        "Total Amount": round(df["TOTAL_AMOUNT"].sum(), 2),
        "Dispatch Qty": round(df["DISPATCH_QTY"].sum(), 2),
        "Cancel Qty": round(df.loc[df["IS_CANCEL"], "SO_QTY_(MT)"].sum(), 2),
        "Active Pending": round(df.loc[df["IS_PENDING"], "PENDING"].sum(), 2),
        "SC Qty": round(df.loc[df["IS_SC"], "PENDING"].sum(), 2),
        "New Parties": df.loc[df["IS_NEW"], "CUSTOMER_NAME"].nunique(),
    }


def render_kpi_row(metrics: dict, key_prefix: str = ""):
    """Render 8 KPI cards with a 'View Details' button each. Returns selected KPI label."""
    cols = st.columns(8)
    colors = ["kpi-card", "kpi-card-blue", "kpi-card-green", "kpi-card-orange",
              "kpi-card-red", "kpi-card-purple", "kpi-card-pink", "kpi-card-blue"]
    selected = None
    for i, (label, val) in enumerate(metrics.items()):
        with cols[i]:
            st.markdown(
                f"""<div class="kpi-card {colors[i % len(colors)]}">
                        <h4>{label}</h4>
                        <h2>{val:,.2f}</h2>
                    </div>""",
                unsafe_allow_html=True
            )
            if st.button("View Details", key=f"{key_prefix}_{label}", use_container_width=True):
                selected = label
    return selected


def drilldown_filter(df: pd.DataFrame, kpi: str) -> pd.DataFrame:
    """Return dataframe filtered for a selected KPI drill-down."""
    if kpi == "Unique Customers":
        return df.drop_duplicates(subset=["CUSTOMER_NAME"])
    if kpi == "SO Qty (MT)":
        return df[df["SO_QTY_(MT)"] > 0]
    if kpi == "Total Amount":
        return df[df["TOTAL_AMOUNT"] > 0]
    if kpi == "Dispatch Qty":
        return df[df["DISPATCH_QTY"] > 0]
    if kpi == "Cancel Qty":
        return df[df["IS_CANCEL"]]
    if kpi == "Active Pending":
        return df[df["IS_PENDING"]]
    if kpi == "SC Qty":
        return df[df["IS_SC"]]
    if kpi == "New Parties":
        return df[df["IS_NEW"]]
    return df


def apply_multiselect_filters(df: pd.DataFrame, key_prefix: str):
    """Common 7-column filter row."""
    f1, f2, f3, f4, f5, f6, f7 = st.columns(7)

    def _sel(col, colname):
        opts = sorted([x for x in df[colname].dropna().unique() if str(x).strip() != ""])
        return col.multiselect(colname.replace("_", " ").title(), opts,
                               key=f"{key_prefix}_{colname}")

    cust = _sel(f1, "CUSTOMER_NAME")
    se = _sel(f2, "SALES_EXECUTIVE")
    item = _sel(f3, "ITEM")
    thick = _sel(f4, "THIKNESS")
    width = _sel(f5, "WIDTH")
    grade = _sel(f6, "GRADE")
    status = _sel(f7, "STATUS")

    out = df.copy()
    if cust:   out = out[out["CUSTOMER_NAME"].isin(cust)]
    if se:     out = out[out["SALES_EXECUTIVE"].isin(se)]
    if item:   out = out[out["ITEM"].isin(item)]
    if thick:  out = out[out["THIKNESS"].isin(thick)]
    if width:  out = out[out["WIDTH"].isin(width)]
    if grade:  out = out[out["GRADE"].isin(grade)]
    if status: out = out[out["STATUS"].isin(status)]
    return out


# ---------------- SIDEBAR ----------------
st.sidebar.markdown("## 📊 Ironmart Dashboard")
st.sidebar.markdown("### 📁 Upload Excel File")
uploaded = st.sidebar.file_uploader("Upload Sales Excel (.xlsx / .xls)", type=["xlsx", "xls"])

st.sidebar.markdown("---")
st.sidebar.markdown("### 📌 Sections")
section = st.sidebar.radio(
    "Navigate",
    ["Sales Analysis", "Date/Month/Year Comparison", "Dispatch Pending"],
    label_visibility="collapsed"
)

st.markdown('<div class="main-header">Ironmart Sales Analytical Dashboard</div>',
            unsafe_allow_html=True)

# ---------------- NO FILE ----------------
if uploaded is None:
    st.info("⬅️  Please upload an Excel file from the left sidebar to begin.")
    st.stop()

# ---------------- LOAD DATA ----------------
try:
    file_bytes = uploaded.getvalue()
    all_sheets = load_all_sheets(file_bytes)
    sheet_names = list(all_sheets.keys())
    st.sidebar.success(f"Sheets found: {', '.join(sheet_names)}")
except Exception as e:
    st.error(f"Error reading Excel file: {e}")
    st.stop()


def read_sheet(name):
    if name not in all_sheets:
        return pd.DataFrame()
    return normalise_columns(all_sheets[name].copy())


# Auto-detect main sales sheet
main_sheet = None
for s in sheet_names:
    if "dispatch" not in s.lower() and "pending" not in s.lower():
        main_sheet = s
        break
if main_sheet is None and sheet_names:
    main_sheet = sheet_names[0]

df_main = read_sheet(main_sheet) if main_sheet else pd.DataFrame()

# Auto-detect Dispatch Pending sheet
dispatch_sheet = next((s for s in sheet_names
                       if "dispatch" in s.lower() and "pending" in s.lower()), None)
if dispatch_sheet is None:
    dispatch_sheet = next((s for s in sheet_names if "dispatch" in s.lower()), None)

df_dispatch = read_sheet(dispatch_sheet) if dispatch_sheet else pd.DataFrame()


# ================= SECTION 1: SALES ANALYSIS =================
if section == "Sales Analysis":
    st.subheader("📈 Sales Analysis")

    if df_main.empty:
        st.warning("Main sales sheet could not be loaded or has no data.")
        st.stop()

    metrics = kpi_metrics(df_main)
    selected_kpi = render_kpi_row(metrics, key_prefix="sa")

    if selected_kpi:
        st.markdown(f"### 🔍 Drill-down: **{selected_kpi}**")
        drill_df = drilldown_filter(df_main, selected_kpi)
        st.dataframe(drill_df, use_container_width=True, height=300)
        st.caption(f"Rows: {len(drill_df)}")
        st.markdown("---")

    # Donut chart
    st.markdown("#### 🍩 Ordered / Pending / Cancel / SC / Dispatch")
    donut_data = {
        "Ordered": df_main["SO_QTY_(MT)"].sum(),
        "Pending": df_main["PENDING"].sum(),
        "Cancel": df_main.loc[df_main["IS_CANCEL"], "SO_QTY_(MT)"].sum(),
        "SC": df_main.loc[df_main["IS_SC"], "PENDING"].sum(),
        "Dispatch": df_main["DISPATCH_QTY"].sum(),
    }
    fig = px.pie(
        names=list(donut_data.keys()),
        values=list(donut_data.values()),
        hole=0.55,
        color_discrete_sequence=px.colors.qualitative.Set2
    )
    fig.update_traces(textinfo="percent+label")
    st.plotly_chart(fig, use_container_width=True)

    # Filters
    st.markdown("#### 🔎 Filters")
    df_filtered = apply_multiselect_filters(df_main, "sa")

    # Salesperson-wise charts
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 📊 Sales Person Wise Ordered Qty")
        se_ordered = (df_filtered.groupby("SALES_EXECUTIVE")["SO_QTY_(MT)"]
                      .sum().reset_index().sort_values("SO_QTY_(MT)", ascending=False))
        fig = px.bar(se_ordered, x="SALES_EXECUTIVE", y="SO_QTY_(MT)",
                     text="SO_QTY_(MT)", color="SALES_EXECUTIVE")
        fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown("#### 📊 Sales Person Wise Dispatched Qty")
        se_disp = (df_filtered.groupby("SALES_EXECUTIVE")["DISPATCH_QTY"]
                   .sum().reset_index().sort_values("DISPATCH_QTY", ascending=True))
        fig = px.bar(se_disp, x="DISPATCH_QTY", y="SALES_EXECUTIVE",
                     orientation="h", text="DISPATCH_QTY", color="SALES_EXECUTIVE")
        fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

    # Summary table
    st.markdown("#### 📋 Sales Executive Summary Table")
    if not df_filtered.empty:
        grouped = df_filtered.groupby(["SALES_EXECUTIVE", "ITEM"])
        summary = grouped.agg(
            OrderedQty=("SO_QTY_(MT)", "sum"),
            DispatchedQty=("DISPATCH_QTY", "sum"),
            PendingQty=("PENDING", "sum"),
        ).reset_index()

        cancel_map = (df_filtered[df_filtered["IS_CANCEL"]]
                      .groupby(["SALES_EXECUTIVE", "ITEM"])["SO_QTY_(MT)"].sum()
                      .rename("CancelQty"))
        sc_map = (df_filtered[df_filtered["IS_SC"]]
                  .groupby(["SALES_EXECUTIVE", "ITEM"])["PENDING"].sum()
                  .rename("SCQty"))

        summary = summary.merge(cancel_map, on=["SALES_EXECUTIVE", "ITEM"], how="left")
        summary = summary.merge(sc_map, on=["SALES_EXECUTIVE", "ITEM"], how="left")
        summary[["CancelQty", "SCQty"]] = summary[["CancelQty", "SCQty"]].fillna(0)

        st.dataframe(summary, use_container_width=True)

    # Top parties
    st.markdown("#### 🏆 Top Parties - Ordered / Pending / Cancel / Dispatch / SC")
    if not df_filtered.empty:
        top_parties = (df_filtered.groupby("CUSTOMER_NAME")
                       .agg(Ordered=("SO_QTY_(MT)", "sum"),
                            Pending=("PENDING", "sum"),
                            Dispatch=("DISPATCH_QTY", "sum"))
                       .reset_index()
                       .sort_values("Ordered", ascending=False).head(15))

        cancel_by_cust = (df_filtered[df_filtered["IS_CANCEL"]]
                          .groupby("CUSTOMER_NAME")["SO_QTY_(MT)"].sum())
        sc_by_cust = (df_filtered[df_filtered["IS_SC"]]
                      .groupby("CUSTOMER_NAME")["PENDING"].sum())

        top_parties["Cancel"] = top_parties["CUSTOMER_NAME"].map(cancel_by_cust).fillna(0)
        top_parties["SC"] = top_parties["CUSTOMER_NAME"].map(sc_by_cust).fillna(0)

        fig = go.Figure()
        for col in ["Ordered", "Pending", "Cancel", "Dispatch", "SC"]:
            fig.add_trace(go.Bar(name=col, x=top_parties["CUSTOMER_NAME"],
                                 y=top_parties[col], text=top_parties[col],
                                 textposition="outside"))
        fig.update_layout(barmode="group", xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)

    # Item-wise analysis
    st.markdown("#### 🔩 Item Wise Analysis")
    ci1, ci2, ci3 = st.columns(3)
    item_opts = sorted([x for x in df_main["ITEM"].unique() if x])
    thick_opts = sorted([x for x in df_main["THIKNESS"].unique() if x])
    width_opts = sorted([x for x in df_main["WIDTH"].dropna().unique()])

    sel_item = ci1.multiselect("Item Name", item_opts, key="item_ms")
    sel_thick = ci2.multiselect("Thickness", thick_opts, key="thick_ms")
    sel_width = ci3.multiselect("Width", width_opts, key="width_ms")

    item_df = df_filtered.copy()
    if sel_item:  item_df = item_df[item_df["ITEM"].isin(sel_item)]
    if sel_thick: item_df = item_df[item_df["THIKNESS"].isin(sel_thick)]
    if sel_width: item_df = item_df[item_df["WIDTH"].isin(sel_width)]

    if not item_df.empty:
        item_grouped = item_df.groupby(["ITEM", "THIKNESS", "WIDTH"])
        item_summary = item_grouped.agg(
            Ordered=("SO_QTY_(MT)", "sum"),
            Dispatch=("DISPATCH_QTY", "sum"),
            Pending=("PENDING", "sum"),
        ).reset_index()

        itm_cancel = (item_df[item_df["IS_CANCEL"]]
                      .groupby(["ITEM", "THIKNESS", "WIDTH"])["SO_QTY_(MT)"].sum()
                      .rename("Cancel"))
        itm_sc = (item_df[item_df["IS_SC"]]
                  .groupby(["ITEM", "THIKNESS", "WIDTH"])["PENDING"].sum()
                  .rename("SC"))

        item_summary = item_summary.merge(itm_cancel, on=["ITEM", "THIKNESS", "WIDTH"], how="left")
        item_summary = item_summary.merge(itm_sc, on=["ITEM", "THIKNESS", "WIDTH"], how="left")
        item_summary[["Cancel", "SC"]] = item_summary[["Cancel", "SC"]].fillna(0)

        st.dataframe(item_summary, use_container_width=True)

        fig = px.bar(
            item_summary.melt(id_vars=["ITEM", "THIKNESS", "WIDTH"],
                              value_vars=["Ordered", "Dispatch", "Cancel", "Pending", "SC"],
                              var_name="Metric", value_name="Qty"),
            x="ITEM", y="Qty", color="Metric", barmode="group", text="Qty"
        )
        fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        st.plotly_chart(fig, use_container_width=True)


# ================= SECTION 2: DATE/MONTH/YEAR COMPARISON =================
elif section == "Date/Month/Year Comparison":
    st.subheader("📅 Date / Month / Year Comparison")

    if df_main.empty:
        st.warning("Main sales sheet could not be loaded.")
        st.stop()

    df_valid = df_main.dropna(subset=["SO_DATE"]).copy()
    if df_valid.empty:
        st.warning("No valid SO_DATE entries found.")
        st.stop()

    df_valid["YEAR"] = df_valid["SO_DATE"].dt.year
    df_valid["MONTH"] = df_valid["SO_DATE"].dt.month
    df_valid["DATE"] = df_valid["SO_DATE"].dt.date

    mode = st.radio(
        "Select Comparison Mode",
        ["Compare Two Dates", "Compare Two Months", "Compare Two Years"],
        horizontal=True
    )

    def kpi_table(dfA, dfB, labelA, labelB):
        kA = kpi_metrics(dfA)
        kB = kpi_metrics(dfB)
        rows = []
        for k in kA:
            a, b = kA[k], kB[k]
            diff = b - a
            pct = (diff / a * 100) if a not in (0, 0.0) else 0
            rows.append([k, a, b, diff, f"{pct:+.2f}%"])
        return pd.DataFrame(rows, columns=["KPI", labelA, labelB, "Difference", "% Change"]), kA, kB

    col1, col2 = st.columns(2)

    if mode == "Compare Two Dates":
        min_d = df_valid["SO_DATE"].min().date()
        max_d = df_valid["SO_DATE"].max().date()
        d1 = col1.date_input("Date A", value=min_d, min_value=min_d, max_value=max_d)
        d2 = col2.date_input("Date B", value=max_d, min_value=min_d, max_value=max_d)
        dfA = df_valid[df_valid["DATE"] == d1]
        dfB = df_valid[df_valid["DATE"] == d2]
        labelA, labelB = str(d1), str(d2)

    elif mode == "Compare Two Months":
        years = sorted(df_valid["YEAR"].unique())
        c1, c2 = col1.columns(2)
        y1 = c1.selectbox("Year A", years, index=0)
        m1 = c2.selectbox("Month A", range(1, 13),
                          format_func=lambda m: datetime(2000, m, 1).strftime("%B"))
        c3, c4 = col2.columns(2)
        y2 = c3.selectbox("Year B", years, index=len(years) - 1)
        m2 = c4.selectbox("Month B", range(1, 13), index=0,
                          format_func=lambda m: datetime(2000, m, 1).strftime("%B"))
        dfA = df_valid[(df_valid["YEAR"] == y1) & (df_valid["MONTH"] == m1)]
        dfB = df_valid[(df_valid["YEAR"] == y2) & (df_valid["MONTH"] == m2)]
        labelA = f"{datetime(2000, m1, 1).strftime('%b')}-{y1}"
        labelB = f"{datetime(2000, m2, 1).strftime('%b')}-{y2}"

    else:  # Two Years
        years = sorted(df_valid["YEAR"].unique())
        y1 = col1.selectbox("Year A", years, index=0)
        y2 = col2.selectbox("Year B", years, index=len(years) - 1)
        dfA = df_valid[df_valid["YEAR"] == y1]
        dfB = df_valid[df_valid["YEAR"] == y2]
        labelA, labelB = str(y1), str(y2)

    st.markdown("---")
    st.markdown("#### 📊 KPI Comparison")
    kpi_df, kA, kB = kpi_table(dfA, dfB, labelA, labelB)

    cols = st.columns(len(kA))
    for i, k in enumerate(kA):
        cols[i].metric(k, f"{kB[k]:,.2f}", delta=f"{kB[k] - kA[k]:,.2f}")

    st.dataframe(kpi_df, use_container_width=True)

    kpi_names = list(kA.keys())
    fig = go.Figure()
    fig.add_trace(go.Bar(name=labelA, x=kpi_names, y=[kA[k] for k in kpi_names],
                         text=[kA[k] for k in kpi_names], textposition="outside"))
    fig.add_trace(go.Bar(name=labelB, x=kpi_names, y=[kB[k] for k in kpi_names],
                         text=[kB[k] for k in kpi_names], textposition="outside"))
    fig.update_layout(barmode="group", xaxis_tickangle=-30)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 📈 Order Qty Trend")
    trend = df_valid.groupby("SO_DATE")["SO_QTY_(MT)"].sum().reset_index()
    fig = px.line(trend, x="SO_DATE", y="SO_QTY_(MT)", markers=True, text="SO_QTY_(MT)")
    fig.update_traces(textposition="top center")
    st.plotly_chart(fig, use_container_width=True)


# ================= SECTION 3: DISPATCH PENDING =================
elif section == "Dispatch Pending":
    st.subheader("🚚 Dispatch Pending Analysis")

    if df_dispatch.empty:
        st.warning("No 'Dispatch Pending' sheet found in the uploaded file.")
        st.stop()

    qty_col = "PENDING" if "PENDING" in df_dispatch.columns else (
        "SO_QTY_(MT)" if "SO_QTY_(MT)" in df_dispatch.columns else None
    )

    m = {
        "Unique Customers": df_dispatch["CUSTOMER_NAME"].nunique(),
        "Total Pending": round(df_dispatch[qty_col].sum(), 2) if qty_col else 0,
        "Total Ordered": round(df_dispatch["SO_QTY_(MT)"].sum(), 2) if "SO_QTY_(MT)" in df_dispatch else 0,
        "Dispatched": round(df_dispatch["DISPATCH_QTY"].sum(), 2) if "DISPATCH_QTY" in df_dispatch else 0,
        "Cancel": round(df_dispatch.loc[df_dispatch["IS_CANCEL"], "SO_QTY_(MT)"].sum(), 2)
            if "SO_QTY_(MT)" in df_dispatch else 0,
        "Active Pending": round(df_dispatch.loc[df_dispatch["IS_PENDING"], "PENDING"].sum(), 2)
            if "PENDING" in df_dispatch else 0,
        "SC": round(df_dispatch.loc[df_dispatch["IS_SC"], "PENDING"].sum(), 2)
            if "PENDING" in df_dispatch else 0,
        "New Parties": df_dispatch.loc[df_dispatch["IS_NEW"], "CUSTOMER_NAME"].nunique(),
    }

    selected = render_kpi_row(m, key_prefix="dp")
    if selected:
        st.markdown(f"### 🔍 Drill-down: **{selected}**")
        st.dataframe(drilldown_filter(df_dispatch, selected),
                     use_container_width=True, height=300)
        st.markdown("---")

    colA, colB = st.columns(2)

    with colA:
        st.markdown("#### 🍩 Dispatch Status Distribution")
        donut_data = {
            "Pending": df_dispatch["PENDING"].sum() if "PENDING" in df_dispatch else 0,
            "Dispatched": df_dispatch["DISPATCH_QTY"].sum() if "DISPATCH_QTY" in df_dispatch else 0,
            "Cancel": df_dispatch.loc[df_dispatch["IS_CANCEL"], "SO_QTY_(MT)"].sum()
                if "SO_QTY_(MT)" in df_dispatch else 0,
            "SC": df_dispatch.loc[df_dispatch["IS_SC"], "PENDING"].sum()
                if "PENDING" in df_dispatch else 0,
        }
        fig = px.pie(names=list(donut_data.keys()), values=list(donut_data.values()),
                     hole=0.55, color_discrete_sequence=px.colors.qualitative.Bold)
        fig.update_traces(textinfo="percent+label")
        st.plotly_chart(fig, use_container_width=True)

    with colB:
        st.markdown("#### 📊 Sales Executive Wise Pending Qty")
        if "SALES_EXECUTIVE" in df_dispatch and qty_col:
            se_pending = (df_dispatch.groupby("SALES_EXECUTIVE")[qty_col].sum()
                          .reset_index().sort_values(qty_col, ascending=True))
            fig = px.bar(se_pending, x=qty_col, y="SALES_EXECUTIVE", orientation="h",
                         text=qty_col, color="SALES_EXECUTIVE")
            fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 🏆 Top Customers with Pending Qty")
    if "CUSTOMER_NAME" in df_dispatch and qty_col:
        cust_pending = (df_dispatch.groupby("CUSTOMER_NAME")[qty_col].sum()
                        .reset_index().sort_values(qty_col, ascending=False).head(15))
        fig = px.bar(cust_pending, x="CUSTOMER_NAME", y=qty_col,
                     text=qty_col, color="CUSTOMER_NAME")
        fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        fig.update_layout(xaxis_tickangle=-45, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 📋 Dispatch Pending Detail")
    st.dataframe(df_dispatch, use_container_width=True, height=400)

    csv = df_dispatch.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download Dispatch Pending CSV", csv,
                       "dispatch_pending.csv", "text/csv")
