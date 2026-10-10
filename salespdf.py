import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import re
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="Ironmart Sales & Dispatch Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Ironmart Sales & Dispatch Analytics Dashboard")

# ---------------------------------------------------------
# Flexible Header Mapping & Normalization
# ---------------------------------------------------------
EXPECTED_COLUMNS = {
    'S_NO': ['s_no', 's.no', 'sr no', 'serial no', 'sno', 's_no.', 'sr no.'],
    'PO_NO': ['po_no', 'po no', 'po.no', 'po number', 'pono'],
    'DO_NO': ['do_no', 'do no', 'do.no', 'dono', 'do .no.'],
    'SO_DATE': ['so_date', 'po date', 'podate', 'po_date', 'order date', 'so date'],
    'CUSTOMER_NAME': ['customer_name', 'party name', 'party_name', 'customer name', 'party'],
    'BROKER': ['broker'],
    'SECTOR': ['sector'],
    'PLACE': ['place', 'city', 'location'],
    'SALES_EXECUTIVE': ['sales_executive', 'seller name', 'seller_name', 'sales person', 'salesperson', 'sales executive'],
    'ITEM': ['item', 'item name', 'item_name', 'product', 'discription'],
    'THIKNESS': ['thikness', 'thickness', 'thk', 'thk.'],
    'SIZE_(MM)': ['size_(mm)', 'size', 'size (mm)', 'size'],
    'GRADE': ['grade'],
    'SO_QTY_(MT)': ['so_qty_(mt)', 'so qty (mt)', 'po qty (mt)', 'po qty', 'po_qty', 'ordered qty', 'order qty'],
    'PER_TON': ['per_ton', 'per ton', 'rate', 'price/ton', 'rate per ton'],
    'INVOICE_NO': ['invoice_no', 'inv no.', 'inv no', 'invoice no', 'inv_no'],
    'INVOICE_DATE': ['invoice_date', 'date', 'disp date', 'dispatch date', 'inv date'],
    'DISP.QTY': ['dispatch_qty', 'disp.qty', 'disp qty', 'dispatched qty', 'disp_qty'],
    'PENDING': ['pending', 'pending qty', 'bal qty'],
    'PAYMENT': ['payment', 'payment mode', 'terms'],
    'DISPATCH_THROUGH': ['dispatch_through', 'disp. th.', 'disp th', 'dispatch through', 'disp_th'],
    'STATUS': ['status', 'order status'],
    'REMARK': ['remark', 'remarks'],
    'MOBILE NO': ['mobile no', 'mobile', 'phone']
}

@st.cache_data
def load_and_clean_sheet(file_bytes, sheet_name):
    df_raw = pd.read_excel(file_bytes, sheet_name=sheet_name, header=None)
    
    header_row_idx = 0
    max_matches = 0
    
    for row_idx in range(min(10, len(df_raw))):
        row_values = df_raw.iloc[row_idx].astype(str).str.strip().str.lower().tolist()
        matches = 0
        for std_col, aliases in EXPECTED_COLUMNS.items():
            if any(alias in row_values for alias in aliases):
                matches += 1
        if matches > max_matches:
            max_matches = matches
            header_row_idx = row_idx
            
    df = pd.read_excel(file_bytes, sheet_name=sheet_name, header=header_row_idx)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Remove duplicate columns safely to prevent DataFrame indexing errors
    df = df.loc[:, ~df.columns.duplicated()].copy()
    
    renamed_cols = {}
    for col in df.columns:
        col_lower = str(col).strip().lower()
        matched = False
        for std_col, aliases in EXPECTED_COLUMNS.items():
            if col_lower in aliases or any(a in col_lower for a in aliases):
                renamed_cols[col] = std_col
                matched = True
                break
        if not matched:
            renamed_cols[col] = col

    df.rename(columns=renamed_cols, inplace=True)
    
    for std_col in EXPECTED_COLUMNS.keys():
        if std_col not in df.columns:
            df[std_col] = np.nan

    if 'SO_DATE' in df.columns:
        df['SO_DATE'] = df['SO_DATE'].ffill()

    numeric_cols = ['SO_QTY_(MT)', 'PER_TON', 'DISP.QTY', 'PENDING']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    if 'CUSTOMER_NAME' in df.columns:
        is_total_word = df['CUSTOMER_NAME'].astype(str).str.lower().str.contains('total|sum', na=False)
        is_empty_cust = df['CUSTOMER_NAME'].isna() | (df['CUSTOMER_NAME'].astype(str).str.strip() == '') | (df['CUSTOMER_NAME'].astype(str).str.lower() == 'nan')
        is_empty_do = df['DO_NO'].isna() | (df['DO_NO'].astype(str).str.strip() == '') | (df['DO_NO'].astype(str).str.lower() == 'nan')
        df = df[~(is_total_word | (is_empty_cust & is_empty_do))].copy()

    if 'SO_DATE' in df.columns:
        so_date_clean = df['SO_DATE'].astype(str).str.replace('.', '/', regex=False)
        df['SO_DATE'] = pd.to_datetime(so_date_clean, dayfirst=True, errors='coerce')
    
    df['SO_DATE_STR'] = df['SO_DATE'].dt.strftime('%d/%m/%Y').fillna('N/A')
    df['TOTAL_AMOUNT'] = df['SO_QTY_(MT)'] * df['PER_TON']
    
    str_cols = ['CUSTOMER_NAME', 'SALES_EXECUTIVE', 'ITEM', 'STATUS', 'REMARK', 'BROKER', 'PLACE', 'PO_NO', 'DO_NO', 'THIKNESS']
    for c in str_cols:
        if c in df.columns and isinstance(df[c], pd.Series):
            df[c] = df[c].fillna('N/A').astype(str).str.strip()

    df['THICKNESS_MM'] = df['THIKNESS'].astype(str).str.replace(r'(?i)\s*mm', '', regex=True).str.strip()

    status_lower = df['STATUS'].str.lower()
    remark_lower = df['REMARK'].str.lower()
    
    df['IS_CANCELLED'] = status_lower.str.contains('cancel') | remark_lower.str.contains('cancel')
    df['IS_SHORT_CLOSE'] = status_lower.str.contains(r'\bsc\b|short close') | remark_lower.str.contains(r'\bsc\b|short close')
    df['IS_NEW_CUSTOMER'] = remark_lower.str.contains('new')

    def determine_status(row):
        if row['IS_CANCELLED']:
            return 'CANCEL'
        elif row['IS_SHORT_CLOSE']:
            return 'SC'
        elif row['DISP.QTY'] >= row['SO_QTY_(MT)'] - 0.01:
            return 'OK'
        elif row['PENDING'] <= 0.01:
            return 'OK'
        else:
            return 'PENDING'

    df['ORDER_STATUS'] = df.apply(determine_status, axis=1)

    def parse_width(size_val):
        if pd.isna(size_val) or str(size_val).strip() in ['', 'nan', 'N/A']:
            return "N/A"
        match = re.search(r'(\d+)\s*[xX*]\s*(\d+)', str(size_val))
        if match:
            return match.group(1)
        num = re.findall(r'\d+', str(size_val))
        return num[0] if num else str(size_val)

    df['WIDTH_MM'] = df['SIZE_(MM)'].apply(parse_width)

    return df

# ---------------------------------------------------------
# Sidebar Navigation Controls
# ---------------------------------------------------------
with st.sidebar:
    st.subheader("📌 Navigation & Controls")
    uploaded_file = st.file_uploader("Upload Sales Excel Workbook", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        
        st.markdown("---")
        section = st.radio("Select Section", [
            "1) Dispatch Analysis",
            "2) Date, Month & Year Comparison",
            "3) Pending Dispatch Report"
        ])
    else:
        st.info("👈 Please upload your sales Excel workbook to begin analysis.")
        st.stop()

# ---------------------------------------------------------
# Helper Chart & PDF Functions
# ---------------------------------------------------------
def make_pie_chart_bytes(labels, values, title):
    fig, ax = plt.subplots(figsize=(6, 3.2), dpi=200)
    valid_vals = [v if v > 0 else 0 for v in values]
    if sum(valid_vals) == 0:
        plt.close(fig)
        return None
    wedges, texts, autotexts = ax.pie(
        valid_vals, labels=labels, autopct="%1.1f%%", startangle=90, 
        colors=['#10B981', '#F59E0B', '#EF4444', '#6366F1']
    )
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_weight('bold')
        autotext.set_fontsize(8)
    ax.set_title(title, fontsize=10, fontweight='bold', color='#1E3A8A', pad=10)
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf

def make_bar_chart_bytes(df_data, x_col, y_col, title, color='#2563EB'):
    if df_data is None or df_data.empty:
        return None
    fig, ax = plt.subplots(figsize=(7, 3.5), dpi=200)
    rects = ax.bar(df_data[x_col].astype(str), df_data[y_col], color=color, width=0.45)
    ax.bar_label(rects, fmt='%.1f', padding=2, fontsize=6.5, fontweight='bold')
    plt.xticks(rotation=35, ha='right', fontsize=6.5)
    ax.set_title(title, fontsize=10, fontweight='bold', color='#1E3A8A', pad=10)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf

def generate_dashboard_pdf(sheet_name, kpis, chart_buffers, tables_dict):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=15, textColor=colors.HexColor('#1E3A8A'), spaceAfter=8)
    section_style = ParagraphStyle('DocSection', parent=styles['Heading2'], fontSize=11, textColor=colors.HexColor('#1E40AF'), spaceBefore=12, spaceAfter=6)
    cell_style = ParagraphStyle('TableCell', parent=styles['Normal'], fontSize=7, leading=8.5)
    cell_header = ParagraphStyle('HeaderCell', parent=styles['Normal'], fontSize=7.5, leading=9, textColor=colors.whitesmoke, fontName="Helvetica-Bold")
    
    story.append(Paragraph(f"<b>📊 Dispatch Report — {sheet_name}</b>", title_style))
    story.append(Spacer(1, 4))
    
    if kpis:
        story.append(Paragraph("<b>1. Key Performance Indicators (KPIs)</b>", section_style))
        kpi_items = [(k, f"{v:,.2f}" if isinstance(v, float) else f"{v:,}") for k, v in kpis.items()]
        kpi_matrix = []
        for i in range(0, len(kpi_items), 4):
            chunk = kpi_items[i:i+4]
            row_titles = [Paragraph(f"<b>{item[0]}</b>", cell_header) for item in chunk]
            row_vals = [Paragraph(f"<b>{item[1]}</b>", cell_style) for item in chunk]
            while len(row_titles) < 4:
                row_titles.append(Paragraph("", cell_header))
                row_vals.append(Paragraph("", cell_style))
            kpi_matrix.append(row_titles)
            kpi_matrix.append(row_vals)
            
        t_kpi = Table(kpi_matrix, colWidths=[135]*4)
        t_kpi.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2563EB')),
            ('BACKGROUND', (0,2), (-1,2), colors.HexColor('#2563EB')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t_kpi)
        story.append(Spacer(1, 8))

    if chart_buffers:
        story.append(Paragraph("<b>2. Visual Analytics & Charts</b>", section_style))
        for fig_title, buf in chart_buffers.items():
            if buf is not None:
                story.append(KeepTogether([
                    Paragraph(f"<b>{fig_title}</b>", ParagraphStyle('SubHead', parent=styles['Normal'], fontSize=9, fontName="Helvetica-Bold", textColor=colors.HexColor('#1E40AF'))),
                    Spacer(1, 2),
                    Image(buf, width=520, height=220),
                    Spacer(1, 6)
                ]))

    story.append(Paragraph("<b>3. Detailed Data Tables</b>", section_style))
    for title, df_table in tables_dict.items():
        if df_table is not None and not df_table.empty:
            story.append(Paragraph(f"<b>{title}</b>", ParagraphStyle('THead', parent=styles['Normal'], fontSize=9, fontName="Helvetica-Bold", textColor=colors.HexColor('#1E3A8A'), spaceBefore=8, spaceAfter=4)))
            sub_df = df_table.copy().reset_index(drop=True)
            cols = sub_df.columns.tolist()
            table_data = [[Paragraph(f"<b>{col}</b>", cell_header) for col in cols]]
            for row in sub_df.values.tolist():
                formatted_row = [Paragraph(f"{val:,.2f}" if isinstance(val, (float, np.floating)) else str(val), cell_style) for val in row]
                table_data.append(formatted_row)
            col_width = 540 / max(len(cols), 1)
            t_data = Table(table_data, colWidths=[col_width]*len(cols), repeatRows=1)
            t_data.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F2937')),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F9FAFB')]),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
                ('TOPPADDING', (0,0), (-1,-1), 3),
                ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ]))
            story.append(t_data)
            story.append(Spacer(1, 8))
            
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ---------------------------------------------------------
# SECTION 1: DISPATCH ANALYSIS (WITH CUSTOMER & ITEM WISE)
# ---------------------------------------------------------
if section == "1) Dispatch Analysis":
    st.header("📦 Dispatch Analysis & Monthly Performance")
    
    selected_month = st.selectbox("Select Month / Sheet", sheet_names)
    df = load_and_clean_sheet(uploaded_file, selected_month)
    
    po_count = df['PO_NO'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique()
    do_count = df['DO_NO'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique()
    unique_parties = df['CUSTOMER_NAME'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique()
    sum_so_qty = df['SO_QTY_(MT)'].sum()
    total_amount = df['TOTAL_AMOUNT'].sum()
    sum_disp_qty = df['DISP.QTY'].sum()
    sum_pending = df[df['ORDER_STATUS'] == 'PENDING']['PENDING'].sum()
    sum_cancelled = df[df['ORDER_STATUS'] == 'CANCEL']['SO_QTY_(MT)'].sum()
    sum_sc = df[df['ORDER_STATUS'] == 'SC']['PENDING'].sum()
    new_customers_df = df[df['IS_NEW_CUSTOMER']]
    new_customers_count = new_customers_df['CUSTOMER_NAME'].nunique()

    kpi_dict = {
        'PO Count': po_count,
        'DO Count': do_count,
        'Unique Parties': unique_parties,
        'Total SO Qty (MT)': sum_so_qty,
        'Total Amount (₹)': total_amount,
        'Dispatched Qty (MT)': sum_disp_qty,
        'Active Pending (MT)': sum_pending,
        'Cancelled Qty (MT)': sum_cancelled,
        'Short Close (MT)': sum_sc,
        'New Parties Added': new_customers_count
    }

    st.subheader("📌 Key Performance Indicators (Click any metric card below to inspect details)")
    
    if 'active_kpi_drill' not in st.session_state:
        st.session_state.active_kpi_drill = None

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        if st.button(f"📌 PO Count\n\n{po_count:,}", use_container_width=True): st.session_state.active_kpi_drill = "PO"
    with c2:
        if st.button(f"📌 DO Count\n\n{do_count:,}", use_container_width=True): st.session_state.active_kpi_drill = "DO"
    with c3:
        if st.button(f"📌 Unique Parties\n\n{unique_parties:,}", use_container_width=True): st.session_state.active_kpi_drill = "PARTIES"
    with c4:
        if st.button(f"📌 Total SO Qty\n\n{sum_so_qty:,.2f} MT", use_container_width=True): st.session_state.active_kpi_drill = "SO_QTY"
    with c5:
        if st.button(f"📌 Total Amount\n\n₹{total_amount:,.0f}", use_container_width=True): st.session_state.active_kpi_drill = "AMOUNT"

    c6, c7, c8, c9, c10 = st.columns(5)
    with c6:
        if st.button(f"📌 Dispatched Qty\n\n{sum_disp_qty:,.2f} MT", use_container_width=True): st.session_state.active_kpi_drill = "DISP"
    with c7:
        if st.button(f"📌 Active Pending\n\n{sum_pending:,.2f} MT", use_container_width=True): st.session_state.active_kpi_drill = "PENDING"
    with c8:
        if st.button(f"📌 Cancelled Qty\n\n{sum_cancelled:,.2f} MT", use_container_width=True): st.session_state.active_kpi_drill = "CANCEL"
    with c9:
        if st.button(f"📌 Short Close Qty\n\n{sum_sc:,.2f} MT", use_container_width=True): st.session_state.active_kpi_drill = "SC"
    with c10:
        if st.button(f"📌 New Parties\n\n{new_customers_count}", use_container_width=True): st.session_state.active_kpi_drill = "NEW_PARTIES"

    if st.session_state.active_kpi_drill:
        st.markdown(f"### 🔍 Detailed View: {st.session_state.active_kpi_drill}")
        drill_df = df[df['ORDER_STATUS'] == 'PENDING'] if st.session_state.active_kpi_drill == "PENDING" else (df[df['ORDER_STATUS'] == 'CANCEL'] if st.session_state.active_kpi_drill == "CANCEL" else (df[df['ORDER_STATUS'] == 'SC'] if st.session_state.active_kpi_drill == "SC" else (new_customers_df if st.session_state.active_kpi_drill == "NEW_PARTIES" else df)))
        display_cols = [c for c in ['PO_NO', 'DO_NO', 'SO_DATE_STR', 'CUSTOMER_NAME', 'SALES_EXECUTIVE', 'ITEM', 'SO_QTY_(MT)', 'DISP.QTY', 'PENDING', 'ORDER_STATUS', 'REMARK'] if c in drill_df.columns]
        st.dataframe(drill_df[display_cols], use_container_width=True)
        if st.button("Close Drill-down"):
            st.session_state.active_kpi_drill = None
            st.rerun()

    st.markdown("---")

    # Donut Chart for Status
    st.subheader("🍩 Order Status Distribution")
    status_counts = df['ORDER_STATUS'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Count']
    fig_donut = px.pie(
        status_counts, names='Status', values='Count', hole=0.4,
        color='Status', color_discrete_map={'OK': '#10B981', 'PENDING': '#F59E0B', 'CANCEL': '#EF4444', 'SC': '#6366F1'}
    )
    fig_donut.update_traces(textinfo='label+value+percent', textfont_size=12)
    st.plotly_chart(fig_donut, use_container_width=True)

    st.markdown("---")
    st.subheader("📋 Sales Executive Item-Wise Breakdown")
    sp_item_grp = df.groupby(['SALES_EXECUTIVE', 'ITEM']).agg(
        OrderedQty=('SO_QTY_(MT)', 'sum'),
        DispatchedQty=('DISP.QTY', 'sum'),
        CancelledQty=('SO_QTY_(MT)', lambda x: x[df['ORDER_STATUS'] == 'CANCEL'].sum()),
        ShortCloseQty=('PENDING', lambda x: x[df['ORDER_STATUS'] == 'SC'].sum()),
        Status=('ORDER_STATUS', lambda x: ', '.join(x.unique()))
    ).reset_index()
    st.dataframe(sp_item_grp, use_container_width=True)

    # Customer & Item Wise Section (Integrated into Dispatch)
    st.markdown("---")
    st.subheader("🏢 Customer Wise Detailed Table")
    cust_grp = df.groupby(['CUSTOMER_NAME', 'SALES_EXECUTIVE']).agg(
        Ordered=('SO_QTY_(MT)', 'sum'),
        Dispatched=('DISP.QTY', 'sum'),
        Cancelled=('SO_QTY_(MT)', lambda x: x[df['ORDER_STATUS'] == 'CANCEL'].sum()),
        Pending=('PENDING', lambda x: x[df['ORDER_STATUS'] == 'PENDING'].sum()),
        ShortClose=('PENDING', lambda x: x[df['ORDER_STATUS'] == 'SC'].sum())
    ).reset_index()
    st.dataframe(cust_grp, use_container_width=True)
    
    st.markdown("---")
    st.subheader("📊 Top 15 Customers by Order Volume")
    top_cust = cust_grp.sort_values(by='Ordered', ascending=False).head(15)
    fig_topcust = px.bar(
        top_cust, x='CUSTOMER_NAME', y=['Ordered', 'Dispatched', 'Cancelled', 'Pending', 'ShortClose'],
        barmode='group', text_auto=',.1f'
    )
    fig_topcust.update_traces(textposition='outside')
    st.plotly_chart(fig_topcust, use_container_width=True)

    st.markdown("---")
    st.subheader("🔍 Detailed Item, Thickness & Width Table with Filters")
    item_spec_grp = df.groupby(['ITEM', 'THICKNESS_MM', 'WIDTH_MM']).agg(
        OrderedQty=('SO_QTY_(MT)', 'sum'),
        DispatchedQty=('DISP.QTY', 'sum'),
        Cancelled=('SO_QTY_(MT)', lambda x: x[df['ORDER_STATUS'] == 'CANCEL'].sum()),
        Pending=('PENDING', lambda x: x[df['ORDER_STATUS'] == 'PENDING'].sum()),
        ShortClose=('PENDING', lambda x: x[df['ORDER_STATUS'] == 'SC'].sum())
    ).reset_index()
    item_spec_grp.rename(columns={'ITEM': 'ItemName', 'THICKNESS_MM': 'Thickness', 'WIDTH_MM': 'Width'}, inplace=True)

    c1, c2, c3 = st.columns(3)
    with c1: selected_items = st.multiselect("Filter Item Name(s)", sorted(item_spec_grp['ItemName'].unique().tolist()))
    with c2: selected_thickness = st.multiselect("Filter Thickness", sorted(item_spec_grp['Thickness'].unique().tolist(), key=str))
    with c3: selected_width = st.multiselect("Filter Width", sorted(item_spec_grp['Width'].unique().tolist(), key=str))

    filtered_spec = item_spec_grp.copy()
    if selected_items: filtered_spec = filtered_spec[filtered_spec['ItemName'].isin(selected_items)]
    if selected_thickness: filtered_spec = filtered_spec[filtered_spec['Thickness'].isin(selected_thickness)]
    if selected_width: filtered_spec = filtered_spec[filtered_spec['Width'].isin(selected_width)]

    st.dataframe(filtered_spec, use_container_width=True)

    st.markdown("---")
    st.subheader("📥 Download Dispatch PDF Report")
    pdf_data = generate_dashboard_pdf(selected_month, kpi_dict, {}, {"Customer Wise Summary": cust_grp})
    st.download_button("📥 Download PDF Report", data=pdf_data, file_name=f"Dispatch_Report_{selected_month}.pdf", mime="application/pdf")

# ---------------------------------------------------------
# SECTION 2: DATE, MONTH & YEAR COMPARISON
# ---------------------------------------------------------
elif section == "2) Date, Month & Year Comparison":
    st.header("⚖️ Date, Month & Year Wise Comparison")
    
    comp_mode = st.radio("Select Comparison Type", ["Date Wise", "Month Wise", "Year Wise"], horizontal=True)
    
    all_dfs = []
    for s_name in sheet_names:
        if 'pending' not in s_name.lower():
            temp_df = load_and_clean_sheet(uploaded_file, s_name)
            temp_df['SHEET_NAME'] = s_name
            all_dfs.append(temp_df)
            
    master_df = pd.concat(all_dfs, ignore_index=True)
    master_df['YEAR'] = master_df['SO_DATE'].dt.year.fillna(0).astype(int)
    master_df['DATE_STR'] = master_df['SO_DATE'].dt.strftime('%d/%m/%Y')
    
    if comp_mode == "Date Wise":
        valid_dates = sorted(master_df.dropna(subset=['SO_DATE'])['DATE_STR'].unique())
        if not valid_dates:
            st.warning("No valid dates found for comparison.")
            st.stop()
        c1, c2 = st.columns(2)
        with c1: d1 = st.selectbox("Select Period 1 (Date)", valid_dates, index=0)
        with c2: d2 = st.selectbox("Select Period 2 (Date)", valid_dates, index=min(1, len(valid_dates)-1))
        
        df1 = master_df[master_df['DATE_STR'] == d1]
        df2 = master_df[master_df['DATE_STR'] == d2]
        label1, label2 = d1, d2
        
    elif comp_mode == "Month Wise":
        valid_months = [s for s in sheet_names if 'pending' not in s.lower()]
        c1, c2 = st.columns(2)
        with c1: m1 = st.selectbox("Select Period 1 (Month/Sheet)", valid_months, index=0)
        with c2: m2 = st.selectbox("Select Period 2 (Month/Sheet)", valid_months, index=min(1, len(valid_months)-1))
        
        df1 = load_and_clean_sheet(uploaded_file, m1)
        df2 = load_and_clean_sheet(uploaded_file, m2)
        label1, label2 = m1, m2
        
    else: # Year Wise
        valid_years = sorted([y for y in master_df['YEAR'].unique() if y > 2000])
        if not valid_years:
            st.warning("No valid years found for comparison.")
            st.stop()
        c1, c2 = st.columns(2)
        with c1: y1 = st.selectbox("Select Period 1 (Year)", valid_years, index=0)
        with c2: y2 = st.selectbox("Select Period 2 (Year)", valid_years, index=min(1, len(valid_years)-1))
        
        df1 = master_df[master_df['YEAR'] == y1]
        df2 = master_df[master_df['YEAR'] == y2]
        label1, label2 = str(y1), str(y2)

    def calc_comp_kpis(d):
        return {
            'PO Count': d['PO_NO'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique(),
            'SO Qty (MT)': d['SO_QTY_(MT)'].sum(),
            'Dispatched Qty (MT)': d['DISP.QTY'].sum(),
            'Total Amount (₹)': d['TOTAL_AMOUNT'].sum(),
            'Pending Qty (MT)': d[d['ORDER_STATUS'] == 'PENDING']['PENDING'].sum()
        }

    kpi1 = calc_comp_kpis(df1)
    kpi2 = calc_comp_kpis(df2)
    
    st.subheader(f"📊 Comparison Summary: {label1} vs {label2}")
    
    comp_table_data = []
    for k in kpi1.keys():
        val1 = kpi1[k]
        val2 = kpi2[k]
        diff = val2 - val1
        pct_var = (diff / val1 * 100) if val1 > 0 else 0.0
        comp_table_data.append({
            "Metric": k,
            f"{label1}": val1,
            f"{label2}": val2,
            "Variance (Diff)": diff,
            "Variance (%)": f"{pct_var:+.2f}%"
        })
        
    comp_df = pd.DataFrame(comp_table_data)
    st.dataframe(comp_df, use_container_width=True)
    
    fig_comp = go.Figure(data=[
        go.Bar(name=str(label1), x=list(kpi1.keys()), y=list(kpi1.values()), text=[f"{v:,.1f}" for v in kpi1.values()], textposition='outside'),
        go.Bar(name=str(label2), x=list(kpi2.keys()), y=list(kpi2.values()), text=[f"{v:,.1f}" for v in kpi2.values()], textposition='outside')
    ])
    fig_comp.update_layout(barmode='group', title=f"Performance Comparison: {label1} vs {label2}")
    st.plotly_chart(fig_comp, use_container_width=True)

# ---------------------------------------------------------
# SECTION 3: PENDING DISPATCH REPORT
# ---------------------------------------------------------
elif section == "3) Pending Dispatch Report":
    st.header("🚚 Pending Dispatch Report")
    
    pd_sheet_candidates = [s for s in sheet_names if 'pending' in s.lower() and 'dispatch' in s.lower()]
    target_pd_sheet = pd_sheet_candidates[0] if pd_sheet_candidates else (sheet_names[-1] if sheet_names else None)
    
    pd_sheet = st.selectbox("Select Pending Dispatch Sheet", sheet_names, index=sheet_names.index(target_pd_sheet) if target_pd_sheet in sheet_names else 0)
    
    df_pd = load_and_clean_sheet(uploaded_file, pd_sheet)
    active_pd = df_pd[(df_pd['ORDER_STATUS'] == 'PENDING') & (~df_pd['IS_CANCELLED'])].copy()
    
    pd_kpis = {
        'Delivery Orders (DOs)': active_pd['DO_NO'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique(),
        'Purchase Orders (POs)': active_pd['PO_NO'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique(),
        'Parties Impacted': active_pd['CUSTOMER_NAME'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique(),
        'Pending Line Items': len(active_pd),
        'Total Pending Qty (MT)': active_pd['PENDING'].sum(),
        'Est. Pending Value (₹)': (active_pd['PENDING'] * active_pd['PER_TON']).sum()
    }
    
    st.subheader("📌 Pending Dispatch KPIs")
    c1, c2, c3 = st.columns(3)
    c1.metric("Pending DOs", f"{pd_kpis['Delivery Orders (DOs)']:,}")
    c2.metric("Pending POs", f"{pd_kpis['Purchase Orders (POs)']:,}")
    c3.metric("Parties Impacted", f"{pd_kpis['Parties Impacted']:,}")

    c4, c5, c6 = st.columns(3)
    c4.metric("Pending Items", f"{pd_kpis['Pending Line Items']:,}")
    c5.metric("Total Pending Qty", f"{pd_kpis['Total Pending Qty (MT)']:,.2f} MT")
    c6.metric("Est. Pending Value", f"₹{pd_kpis['Est. Pending Value (₹)']:,.2f}")
    
    st.markdown("---")
    st.subheader("📊 Pending Quantity by Sales Person")
    sp_pd_df = active_pd.groupby('SALES_EXECUTIVE')['PENDING'].sum().reset_index()
    fig_pd = px.bar(sp_pd_df, x='SALES_EXECUTIVE', y='PENDING', title="Pending Dispatch Qty (MT) by Sales Person", text_auto=',.2f', color_discrete_sequence=['#EF4444'])
    fig_pd.update_traces(textposition='outside')
    st.plotly_chart(fig_pd, use_container_width=True)
    
    st.markdown("---")
    st.subheader("📋 Pending Dispatch Detailed Data Table")
    pending_details_df = active_pd[['PO_NO', 'DO_NO', 'SO_DATE_STR', 'CUSTOMER_NAME', 'SALES_EXECUTIVE', 'ITEM', 'THICKNESS_MM', 'SIZE_(MM)', 'SO_QTY_(MT)', 'DISP.QTY', 'PENDING', 'REMARK']].copy()
    st.dataframe(pending_details_df, use_container_width=True)
