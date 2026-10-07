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
    page_title="Sales & Dispatch Analytics Dashboard",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for UI polish
st.markdown("""
<style>
    div.stButton > button {
        width: 100%;
        height: auto;
        padding: 10px 8px;
        background-color: #F8FAFC;
        border: 1px solid #CBD5E1;
        border-radius: 8px;
        color: #1E293B;
        font-weight: 600;
        font-size: 14px;
        text-align: center;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        transition: all 0.2s ease-in-out;
    }
    div.stButton > button:hover {
        background-color: #EFF6FF;
        border-color: #3B82F6;
        color: #1D4ED8;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 Sales & Dispatch Analytics Dashboard")

# ---------------------------------------------------------
# Flexible Excel Header & Column Normalization
# ---------------------------------------------------------
EXPECTED_COLUMNS = {
    'S_NO': ['s_no', 's.no', 'sr no', 'serial no', 'sno', 's_no.'],
    'PO NO': ['po no', 'po.no', 'po number', 'pono', 'po_no'],
    'DO NO': ['do .no.', 'do no', 'do.no', 'do_no', 'dono'],
    'PO DATE': ['po date', 'podate', 'po_date', 'order date'],
    'PARTY NAME': ['party name', 'party_name', 'customer name', 'party'],
    'BROKER': ['broker'],
    'SECTOR': ['sector'],
    'PLACE': ['place', 'city', 'location'],
    'SELLER NAME': ['seller name', 'seller_name', 'sales person', 'salesperson', 'sales executive'],
    'ITEM': ['item', 'item name', 'item_name', 'product', 'discription'],
    'THICKNESS': ['thikness', 'thickness', 'thk', 'thk.'],
    'SIZE': ['size', 'size (mm)'],
    'GRADE': ['grade'],
    'PO QTY (MT)': ['po qty (mt)', 'po qty', 'po_qty', 'ordered qty', 'order qty'],
    'PER TON': ['per ton', 'rate', 'price/ton', 'rate per ton'],
    'INV NO': ['inv no.', 'inv no', 'invoice no', 'inv_no'],
    'DATE': ['date', 'disp date', 'dispatch date', 'inv date'],
    'DISP.QTY': ['disp.qty', 'disp qty', 'dispatched qty', 'disp_qty'],
    'PENDING': ['pending', 'pending qty', 'bal qty'],
    'PAYMENT': ['payment', 'payment mode', 'terms'],
    'DISP. TH.': ['disp. th.', 'disp th', 'dispatch through', 'disp_th'],
    'STATUS': ['status', 'order status'],
    'REMARK': ['remark', 'remarks'],
    'MOBILE NO': ['mobile no', 'mobile', 'phone']
}

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

    numeric_cols = ['PO QTY (MT)', 'PER TON', 'DISP.QTY', 'PENDING']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    if 'PARTY NAME' in df.columns:
        is_total_word = df['PARTY NAME'].astype(str).str.lower().str.contains('total|sum', na=False)
        is_empty_party = df['PARTY NAME'].isna() | (df['PARTY NAME'].astype(str).str.strip() == '') | (df['PARTY NAME'].astype(str).str.lower() == 'nan')
        is_empty_do = df['DO NO'].isna() | (df['DO NO'].astype(str).str.strip() == '') | (df['DO NO'].astype(str).str.lower() == 'nan')
        df = df[~(is_total_word | (is_empty_party & is_empty_do))].copy()

    if 'PO DATE' in df.columns:
        po_date_clean = df['PO DATE'].astype(str).str.replace('.', '/', regex=False)
        df['PO DATE'] = pd.to_datetime(po_date_clean, dayfirst=True, errors='coerce')

    if 'DATE' in df.columns:
        disp_date_clean = df['DATE'].astype(str).str.replace('.', '/', regex=False)
        df['DATE'] = pd.to_datetime(disp_date_clean, dayfirst=True, errors='coerce')
    
    df['PO_DATE_STR'] = df['PO DATE'].dt.strftime('%d/%m/%Y').fillna('N/A')
    df['DISP_DATE_STR'] = df['DATE'].dt.strftime('%d/%m/%Y').fillna('N/A')
    df['AMOUNT'] = df['PO QTY (MT)'] * df['PER TON']
    
    str_cols = ['PARTY NAME', 'SELLER NAME', 'ITEM', 'STATUS', 'REMARK', 'BROKER', 'SECTOR', 'PLACE', 'PO NO', 'DO NO', 'THICKNESS']
    for c in str_cols:
        df[c] = df[c].fillna('N/A').astype(str).str.strip()

    df['THICKNESS_MM'] = df['THICKNESS'].astype(str).str.replace(r'(?i)\s*mm', '', regex=True).str.strip()
    df['IS_CANCELLED'] = df['STATUS'].str.lower().str.contains('cancel') | df['REMARK'].str.lower().str.contains('cancel')
    df['CANCELLED_QTY'] = np.where(df['IS_CANCELLED'], df['PO QTY (MT)'], 0.0)
    df['ACTIVE_PENDING_QTY'] = np.where(df['IS_CANCELLED'], 0.0, df['PENDING'])

    def parse_width(size_val):
        if pd.isna(size_val) or str(size_val).strip() in ['', 'nan', 'N/A']:
            return "N/A"
        match = re.search(r'(\d+)\s*[xX*]\s*(\d+)', str(size_val))
        if match:
            return match.group(1)
        num = re.findall(r'\d+', str(size_val))
        return num[0] if num else str(size_val)

    df['WIDTH_MM'] = df['SIZE'].apply(parse_width)
    return df

# ---------------------------------------------------------
# Matplotlib High-Quality Chart Helpers for PDF Export
# ---------------------------------------------------------
def make_pie_chart_bytes(labels, values, title):
    fig, ax = plt.subplots(figsize=(6, 3.2), dpi=200)
    valid_vals = [v if v > 0 else 0 for v in values]
    if sum(valid_vals) == 0:
        plt.close(fig)
        return None
    wedges, texts, autotexts = ax.pie(
        valid_vals, labels=labels, autopct="%1.1f%%", startangle=90, 
        colors=['#10B981', '#F59E0B', '#EF4444', '#3B82F6']
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

def make_bar_chart_bytes(df_data, x_col, y_cols, title, color='#2563EB'):
    if df_data is None or df_data.empty:
        return None
    fig, ax = plt.subplots(figsize=(7, 3.5), dpi=200)
    if isinstance(y_cols, list):
        x = np.arange(len(df_data[x_col]))
        width = 0.8 / len(y_cols)
        for idx, col in enumerate(y_cols):
            rects = ax.bar(x + idx*width, df_data[col], width, label=col)
            ax.bar_label(rects, fmt='%.1f', padding=2, fontsize=5)
        ax.set_xticks(x + width*(len(y_cols)-1)/2)
        ax.set_xticklabels(df_data[x_col].astype(str), rotation=35, ha='right', fontsize=6)
        ax.legend(fontsize=6, loc='upper right')
    else:
        rects = ax.bar(df_data[x_col].astype(str), df_data[y_cols], color=color, width=0.45)
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

# ---------------------------------------------------------
# Dynamic PDF Generator
# ---------------------------------------------------------
def generate_exact_screen_pdf(sheet_name, kpis, chart_buffers, tables_dict):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#1E3A8A'), spaceAfter=8)
    section_style = ParagraphStyle('DocSection', parent=styles['Heading2'], fontSize=11, textColor=colors.HexColor('#1E40AF'), spaceBefore=14, spaceAfter=6)
    cell_style = ParagraphStyle('TableCell', parent=styles['Normal'], fontSize=7, leading=8.5, alignment=0)
    cell_header = ParagraphStyle('HeaderCell', parent=styles['Normal'], fontSize=7.5, leading=9, textColor=colors.whitesmoke, fontName="Helvetica-Bold", alignment=0)
    
    story.append(Paragraph(f"<b>📊 Sales & Dispatch Report — {sheet_name}</b>", title_style))
    story.append(Spacer(1, 6))
    
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
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
        ]))
        story.append(t_kpi)
        story.append(Spacer(1, 10))

    if chart_buffers:
        story.append(Paragraph("<b>2. Visual Analytics</b>", section_style))
        for fig_title, buf in chart_buffers.items():
            if buf is not None:
                story.append(KeepTogether([
                    Paragraph(f"<b>{fig_title}</b>", ParagraphStyle('SubHead', parent=styles['Normal'], fontSize=9, fontName="Helvetica-Bold", textColor=colors.HexColor('#1E40AF'))),
                    Spacer(1, 3),
                    Image(buf, width=520, height=230),
                    Spacer(1, 8)
                ]))

    for title, df_table in tables_dict.items():
        if df_table is not None and not df_table.empty:
            story.append(Paragraph(f"<b>{title}</b>", section_style))
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
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('TOPPADDING', (0,0), (-1,-1), 3),
                ('BOTTOMPADDING', (0,0), (-1,-1), 3),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F9FAFB')]),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
            ]))
            story.append(t_data)
            story.append(Spacer(1, 10))
            
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ---------------------------------------------------------
# Sidebar Controls & Navigation
# ---------------------------------------------------------
st.sidebar.title("📌 Navigation & Controls")
uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("👈 Please upload an Excel workbook from the left sidebar to view dashboard metrics.")
    st.stop()

xl = pd.ExcelFile(uploaded_file)
sheet_names = xl.sheet_names

section = st.sidebar.radio("Go to Section", [
    "📊 All Sales & Dispatch Analysis",
    "📅 Date, Month & Year Comparison",
    "🚚 Pending Dispatch"
])

def calculate_kpis(data):
    return {
        'Overall PO Count': data['PO NO'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique(),
        'Overall DO Count': data['DO NO'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique(),
        'Number of Parties': data['PARTY NAME'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique(),
        'Total PO Quantity (MT)': data['PO QTY (MT)'].sum(),
        'Total PO Amount': data['AMOUNT'].sum(),
        'Dispatched Qty (MT)': data['DISP.QTY'].sum(),
        'Cancelled Qty (MT)': data['CANCELLED_QTY'].sum(),
        'Pending Qty (MT)': data['ACTIVE_PENDING_QTY'].sum()
    }

# =========================================================
# SECTION 1: ALL SALES & DISPATCH ANALYSIS (SELECTED MONTH)
# =========================================================
if section == "📊 All Sales & Dispatch Analysis":
    selected_sheet = st.sidebar.selectbox("Select Month / Sheet", sheet_names)
    df = load_and_clean_sheet(uploaded_file, selected_sheet)
    
    st.header(f"📊 All Sales & Dispatch Analysis — {selected_sheet}")
    
    kpis = calculate_kpis(df)
    
    if 'active_kpi_drill' not in st.session_state:
        st.session_state.active_kpi_drill = None

    st.markdown("👉 **Click any KPI card below to instantly open its detailed records:**")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if st.button(f"📌 PO Count: {kpis['Overall PO Count']:,}", use_container_width=True):
            st.session_state.active_kpi_drill = "Overall PO Count"
    with col2:
        if st.button(f"📌 DO Count: {kpis['Overall DO Count']:,}", use_container_width=True):
            st.session_state.active_kpi_drill = "Overall DO Count"
    with col3:
        if st.button(f"📌 Parties: {kpis['Number of Parties']:,}", use_container_width=True):
            st.session_state.active_kpi_drill = "Number of Parties"
    with col4:
        if st.button(f"📌 PO Qty: {kpis['Total PO Quantity (MT)']:,.1f} MT", use_container_width=True):
            st.session_state.active_kpi_drill = "Total PO Qty (MT)"

    col5, col6, col7, col8 = st.columns(4)
    with col5:
        if st.button(f"📌 Amount: ₹{kpis['Total PO Amount']:,.0f}", use_container_width=True):
            st.session_state.active_kpi_drill = "Total Amount"
    with col6:
        if st.button(f"📌 Disp. Qty: {kpis['Dispatched Qty (MT)']:,.1f} MT", use_container_width=True):
            st.session_state.active_kpi_drill = "Dispatched Qty"
    with col7:
        if st.button(f"📌 Canc. Qty: {kpis['Cancelled Qty (MT)']:,.1f} MT", use_container_width=True):
            st.session_state.active_kpi_drill = "Cancelled Qty"
    with col8:
        if st.button(f"📌 Pend. Qty: {kpis['Pending Qty (MT)']:,.1f} MT", use_container_width=True):
            st.session_state.active_kpi_drill = "Pending Qty"

    # Drill-down viewer (Excluded from PDF report)
    if st.session_state.active_kpi_drill:
        st.markdown("---")
        c_title, c_close = st.columns([8, 1])
        with c_title:
            st.subheader(f"🔍 Drill-Down Details: {st.session_state.active_kpi_drill}")
        with c_close:
            if st.button("❌ Close", use_container_width=True):
                st.session_state.active_kpi_drill = None
                st.rerun()

        drill = st.session_state.active_kpi_drill
        if drill == "Overall PO Count":
            st.dataframe(df.groupby('PO NO').agg(PO_Date=('PO_DATE_STR', 'first'), Party=('PARTY NAME', 'first'), Seller=('SELLER NAME', 'first'), Total_Ordered=('PO QTY (MT)', 'sum'), Status=('STATUS', 'first')).reset_index(), use_container_width=True)
        elif drill == "Overall DO Count":
            st.dataframe(df[df['DO NO'] != 'N/A'].groupby('DO NO').agg(PO_NO=('PO NO', 'first'), Party=('PARTY NAME', 'first'), Item=('ITEM', 'first'), Ordered=('PO QTY (MT)', 'sum'), Dispatched=('DISP.QTY', 'sum')).reset_index(), use_container_width=True)
        elif drill == "Number of Parties":
            st.dataframe(df.groupby('PARTY NAME').agg(Total_POs=('PO NO', 'nunique'), OrderedQty=('PO QTY (MT)', 'sum'), DispatchedQty=('DISP.QTY', 'sum')).reset_index(), use_container_width=True)
        elif drill in ["Total PO Qty (MT)", "Total Amount"]:
            st.dataframe(df, use_container_width=True)
        elif drill == "Dispatched Qty":
            st.dataframe(df[df['DISP.QTY'] > 0][['PO NO', 'DO NO', 'PARTY NAME', 'SELLER NAME', 'ITEM', 'PO QTY (MT)', 'DISP.QTY']], use_container_width=True)
        elif drill == "Cancelled Qty":
            st.dataframe(df[df['IS_CANCELLED']][['PO NO', 'DO NO', 'PARTY NAME', 'SELLER NAME', 'ITEM', 'PO QTY (MT)', 'REMARK']], use_container_width=True)
        elif drill == "Pending Qty":
            st.dataframe(df[df['ACTIVE_PENDING_QTY'] > 0][['PO NO', 'DO NO', 'PARTY NAME', 'SELLER NAME', 'ITEM', 'PO QTY (MT)', 'ACTIVE_PENDING_QTY']], use_container_width=True)

    st.markdown("---")
    
    # Donut Chart for Status Breakdown
    fig_donut = go.Figure(data=[go.Pie(
        labels=['Dispatched Qty', 'Pending Qty', 'Cancelled Qty'],
        values=[kpis['Dispatched Qty (MT)'], kpis['Pending Qty (MT)'], kpis['Cancelled Qty (MT)']],
        hole=.4,
        textinfo='label+value+percent',
        texttemplate='%{label}<br>%{value:,.2f} MT (%{percent})',
        marker_colors=['#10B981', '#3B82F6', '#EF4444']
    )])
    fig_donut.update_layout(title="Overall Status Breakdown (Ordered, Dispatched, Pending, Cancelled)")
    st.plotly_chart(fig_donut, use_container_width=True)
    
    st.markdown("---")
    st.subheader("Sales Executive Analysis Item-Wise Breakdown")
    sp_item_grp = df.groupby(['SELLER NAME', 'ITEM']).agg(
        OrderedQty=('PO QTY (MT)', 'sum'),
        DispatchedQty=('DISP.QTY', 'sum'),
        CancelledQty=('CANCELLED_QTY', 'sum'),
        PendingQty=('ACTIVE_PENDING_QTY', 'sum')
    ).reset_index()
    st.dataframe(sp_item_grp, use_container_width=True)
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("Cancelled Qty per Sales Person")
        sp_cancelled = df[df['IS_CANCELLED']].groupby('SELLER NAME').agg(CancelledOrders=('PO NO', 'nunique'), CancelledQty=('CANCELLED_QTY', 'sum')).reset_index()
        st.dataframe(sp_cancelled, use_container_width=True)
    with col_b:
        st.subheader("Pending Order Qty per Sales Person")
        sp_pending = df[df['ACTIVE_PENDING_QTY'] > 0].groupby('SELLER NAME').agg(PendingOrders=('PO NO', 'nunique'), PendingQty=('ACTIVE_PENDING_QTY', 'sum')).reset_index()
        st.dataframe(sp_pending, use_container_width=True)
        
    col_c, col_d = st.columns(2)
    with col_c:
        sp_rec = df.groupby('SELLER NAME')['PO QTY (MT)'].sum().reset_index()
        fig_rec = px.bar(sp_rec, x='SELLER NAME', y='PO QTY (MT)', title="Ordered Qty Sales Person-Wise", text_auto=',.1f', color_discrete_sequence=['#3B82F6'])
        fig_rec.update_traces(textposition='outside')
        st.plotly_chart(fig_rec, use_container_width=True)
    with col_d:
        sp_disp = df.groupby('SELLER NAME')['DISP.QTY'].sum().reset_index()
        fig_disp = px.bar(sp_disp, x='SELLER NAME', y='DISP.QTY', title="Dispatched Qty Sales Person-Wise", text_auto=',.1f', color_discrete_sequence=['#10B981'])
        fig_disp.update_traces(textposition='outside')
        st.plotly_chart(fig_disp, use_container_width=True)

    st.subheader("Short Closed Orders (Remarks with 'SC')")
    sc_df = df[df['REMARK'].str.lower().str.contains(r'\bsc\b|short close', na=False)]
    if not sc_df.empty:
        st.dataframe(sc_df.groupby(['SELLER NAME', 'PARTY NAME', 'PO NO', 'REMARK']).agg(ShortClosedQty=('PENDING', 'sum')).reset_index(), use_container_width=True)
    else:
        st.info("No Short Closed ('SC') orders found.")

    st.markdown("---")
    st.subheader("Party-Wise Analytics")
    party_grp = df.groupby(['PARTY NAME', 'SELLER NAME']).agg(
        OrderedQty=('PO QTY (MT)', 'sum'),
        DispatchedQty=('DISP.QTY', 'sum'),
        CancelledQty=('CANCELLED_QTY', 'sum'),
        PendingQty=('ACTIVE_PENDING_QTY', 'sum')
    ).reset_index()
    st.dataframe(party_grp, use_container_width=True)
    
    st.subheader("Top Parties Graphical Representation")
    top_parties = party_grp.sort_values(by='OrderedQty', ascending=False).head(15)
    fig_party = px.bar(top_parties, x='PARTY NAME', y=['OrderedQty', 'DispatchedQty', 'CancelledQty', 'PendingQty'], title="Top Parties - Ordered, Dispatched, Cancelled, Pending", barmode='group', text_auto=',.1f')
    fig_party.update_traces(textposition='outside')
    st.plotly_chart(fig_party, use_container_width=True)

    st.markdown("---")
    st.subheader("Detailed Item, Thickness & Width Summary")
    item_spec_grp = df.groupby(['ITEM', 'THICKNESS_MM', 'WIDTH_MM']).agg(
        OrderedQty=('PO QTY (MT)', 'sum'),
        DispatchedQty=('DISP.QTY', 'sum'),
        CancelledQty=('CANCELLED_QTY', 'sum'),
        PendingQty=('ACTIVE_PENDING_QTY', 'sum')
    ).reset_index()
    item_spec_grp.rename(columns={'ITEM': 'Item Name', 'THICKNESS_MM': 'Thickness (mm)', 'WIDTH_MM': 'Width (mm)'}, inplace=True)

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        sel_items = st.multiselect("Filter Item Name(s)", sorted(item_spec_grp['Item Name'].unique().tolist()))
    with col_f2:
        sel_thk = st.multiselect("Filter Thickness (mm)", sorted(item_spec_grp['Thickness (mm)'].unique().tolist(), key=lambda x: str(x)))
    with col_f3:
        sel_width = st.multiselect("Filter Width (mm)", sorted(item_spec_grp['Width (mm)'].unique().tolist(), key=lambda x: str(x)))

    display_spec = item_spec_grp.copy()
    if sel_items: display_spec = display_spec[display_spec['Item Name'].isin(sel_items)]
    if sel_thk: display_spec = display_spec[display_spec['Thickness (mm)'].isin(sel_thk)]
    if sel_width: display_spec = display_spec[display_spec['Width (mm)'].isin(sel_width)]
    st.dataframe(display_spec, use_container_width=True)

    st.markdown("---")
    st.subheader("📥 Export Clean PDF Report")
    chart_bufs = {
        "Status Breakdown Donut": make_pie_chart_bytes(['Dispatched', 'Pending', 'Cancelled'], [kpis['Dispatched Qty (MT)'], kpis['Pending Qty (MT)'], kpis['Cancelled Qty (MT)']], "Overall Status Breakdown"),
        "Ordered Qty by Sales Person": make_bar_chart_bytes(sp_rec, 'SELLER NAME', 'PO QTY (MT)', "Ordered Qty Sales Person-Wise", color='#3B82F6'),
        "Dispatched Qty by Sales Person": make_bar_chart_bytes(sp_disp, 'SELLER NAME', 'DISP.QTY', "Dispatched Qty Sales Person-Wise", color='#10B981')
    }
    tables_to_pdf = {
        "Sales Executive Item-Wise Breakdown": sp_item_grp,
        "Party-Wise Analytics Summary": party_grp,
        "Detailed Item, Thickness & Width Summary": display_spec
    }
    pdf_bytes = generate_exact_screen_pdf(selected_sheet, kpis, chart_bufs, tables_to_pdf)
    st.download_button("📥 Download Clean PDF Report", data=pdf_bytes, file_name=f"Sales_Report_{selected_sheet}.pdf", mime="application/pdf")

# =========================================================
# SECTION 2: DATE, MONTH & YEAR COMPARISON ANALYSIS
# =========================================================
elif section == "📅 Date, Month & Year Comparison":
    st.header("📅 Date, Month & Year Comparison Analysis")
    
    comp_mode = st.radio("Select Comparison Mode", ["Date-Wise", "Month-Wise", "Year-Wise"], horizontal=True)
    
    df_valid = df.dropna(subset=['PO DATE']).copy()
    df_valid['YEAR'] = df_valid['PO DATE'].dt.year
    df_valid['MONTH'] = df_valid['PO DATE'].dt.month
    df_valid['YEAR_MONTH'] = df_valid['PO DATE'].dt.strftime('%B %Y')
    
    if comp_mode == "Date-Wise":
        dates = sorted(df_valid['PO DATE'].dt.date.unique())
        date_strs = [d.strftime('%d/%m/%Y') for d in dates]
        c1, c2 = st.columns(2)
        with c1: d1_str = st.selectbox("Primary Date", date_strs)
        with c2: d2_str = st.selectbox("Comparison Date", date_strs, index=min(1, len(date_strs)-1))
        
        df1 = df_valid[df_valid['PO DATE'].dt.date == pd.to_datetime(d1_str, format='%d/%m/%Y').date()]
        df2 = df_valid[df_valid['PO DATE'].dt.date == pd.to_datetime(d2_str, format='%d/%m/%Y').date()]
        lbl1, lbl2 = d1_str, d2_str
        
    elif comp_mode == "Month-Wise":
        months = sorted(df_valid['YEAR_MONTH'].unique(), key=lambda x: pd.to_datetime(x, format='%B %Y'))
        c1, c2 = st.columns(2)
        with c1: m1 = st.selectbox("Primary Month", months, index=len(months)-1 if len(months)>0 else 0)
        with c2: m2 = st.selectbox("Comparison Month", months, index=0)
        
        df1 = df_valid[df_valid['YEAR_MONTH'] == m1]
        df2 = df_valid[df_valid['YEAR_MONTH'] == m2]
        lbl1, lbl2 = m1, m2
        
    else:
        years = sorted(df_valid['YEAR'].unique())
        c1, c2 = st.columns(2)
        with c1: y1 = st.selectbox("Primary Year", years, index=len(years)-1 if len(years)>0 else 0)
        with c2: y2 = st.selectbox("Comparison Year", years, index=0)
        
        df1 = df_valid[df_valid['YEAR'] == y1]
        df2 = df_valid[df_valid['YEAR'] == y2]
        lbl1, lbl2 = str(y1), str(y2)

    kpi1 = calculate_kpis(df1)
    kpi2 = calculate_kpis(df2)
    
    st.subheader(f"Key Metrics Comparison ({lbl1} vs {lbl2})")
    comp_summary_df = pd.DataFrame({
        "Metric": ["PO Count", "DO Count", "Parties Count", "Ordered Qty (MT)", "Dispatched Qty (MT)", "Cancelled Qty (MT)", "Pending Qty (MT)", "Total Amount (₹)"],
        lbl1: [kpi1['Overall PO Count'], kpi1['Overall DO Count'], kpi1['Number of Parties'], kpi1['Total PO Quantity (MT)'], kpi1['Dispatched Qty (MT)'], kpi1['Cancelled Qty (MT)'], kpi1['Pending Qty (MT)'], kpi1['Total PO Amount']],
        lbl2: [kpi2['Overall PO Count'], kpi2['Overall DO Count'], kpi2['Number of Parties'], kpi2['Total PO Quantity (MT)'], kpi2['Dispatched Qty (MT)'], kpi2['Cancelled Qty (MT)'], kpi2['Pending Qty (MT)'], kpi2['Total PO Amount']],
        "Variance (Diff)": [
            kpi2['Overall PO Count'] - kpi1['Overall PO Count'],
            kpi2['Overall DO Count'] - kpi1['Overall DO Count'],
            kpi2['Number of Parties'] - kpi1['Number of Parties'],
            kpi2['Total PO Quantity (MT)'] - kpi1['Total PO Quantity (MT)'],
            kpi2['Dispatched Qty (MT)'] - kpi1['Dispatched Qty (MT)'],
            kpi2['Cancelled Qty (MT)'] - kpi1['Cancelled Qty (MT)'],
            kpi2['Pending Qty (MT)'] - kpi1['Pending Qty (MT)'],
            kpi2['Total PO Amount'] - kpi1['Total PO Amount']
        ]
    })
    st.table(comp_summary_df)
    
    st.subheader("Graphical Comparison Representation")
    comp_chart_data = pd.DataFrame({
        "Metric": ["Ordered Qty", "Dispatched Qty", "Cancelled Qty", "Pending Qty"],
        lbl1: [kpi1['Total PO Quantity (MT)'], kpi1['Dispatched Qty (MT)'], kpi1['Cancelled Qty (MT)'], kpi1['Pending Qty (MT)']],
        lbl2: [kpi2['Total PO Quantity (MT)'], kpi2['Dispatched Qty (MT)'], kpi2['Cancelled Qty (MT)'], kpi2['Pending Qty (MT)']]
    }).melt(id_vars="Metric", var_name="Period", value_name="Quantity")
    
    fig_comp = px.bar(comp_chart_data, x="Metric", y="Quantity", color="Period", barmode="group", title=f"Comparison: {lbl1} vs {lbl2}", text_auto=".1f")
    st.plotly_chart(fig_comp, use_container_width=True)

# =========================================================
# SECTION 3: PENDING DISPATCH
# =========================================================
elif section == "🚚 Pending Dispatch":
    st.header("🚚 Pending Dispatch Standalone Report")
    
    pd_sheet_candidates = [s for s in sheet_names if 'pending' in s.lower() and 'dispatch' in s.lower()]
    target_pd_sheet = pd_sheet_candidates[0] if pd_sheet_candidates else sheet_names[0]
    pd_sheet = st.sidebar.selectbox("Select Pending Dispatch Sheet", sheet_names, index=sheet_names.index(target_pd_sheet) if target_pd_sheet in sheet_names else 0)
    
    df_pd = load_and_clean_sheet(uploaded_file, pd_sheet)
    active_pd = df_pd[(df_pd['ACTIVE_PENDING_QTY'] > 0) & (~df_pd['IS_CANCELLED'])].copy()
    
    pd_kpis = {
        'Pending DOs': active_pd['DO NO'].replace(['N/A', 'nan'], np.nan).dropna().nunique(),
        'Pending POs': active_pd['PO NO'].replace(['N/A', 'nan'], np.nan).dropna().nunique(),
        'Parties Impacted': active_pd['PARTY NAME'].replace(['N/A', 'nan'], np.nan).dropna().nunique(),
        'Total Pending Qty (MT)': active_pd['ACTIVE_PENDING_QTY'].sum(),
        'Est. Pending Value (₹)': (active_pd['ACTIVE_PENDING_QTY'] * active_pd['PER TON']).sum()
    }
    
    if 'active_pd_drill' not in st.session_state:
        st.session_state.active_pd_drill = None

    st.markdown("👉 **Click any Pending KPI button to view specific records:**")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if st.button(f"📌 Pending DOs: {pd_kpis['Pending DOs']:,}", use_container_width=True): st.session_state.active_pd_drill = "Pending DOs"
    with c2:
        if st.button(f"📌 Pending POs: {pd_kpis['Pending POs']:,}", use_container_width=True): st.session_state.active_pd_drill = "Pending POs"
    with c3:
        if st.button(f"📌 Parties Impacted: {pd_kpis['Parties Impacted']:,}", use_container_width=True): st.session_state.active_pd_drill = "Parties Impacted"
    with c4:
        if st.button(f"📌 Total Pending: {pd_kpis['Total Pending Qty (MT)']:,.1f} MT", use_container_width=True): st.session_state.active_pd_drill = "Total Pending"

    if st.session_state.active_pd_drill:
        st.markdown("---")
        st.subheader(f"🔍 Pending Drill-Down: {st.session_state.active_pd_drill}")
        st.dataframe(active_pd[['PO NO', 'DO NO', 'PARTY NAME', 'SELLER NAME', 'ITEM', 'PO QTY (MT)', 'ACTIVE_PENDING_QTY']], use_container_width=True)

    st.markdown("---")
    st.subheader("Pending Quantity Breakdown by Sales Person")
    sp_pd = active_pd.groupby('SELLER NAME')['ACTIVE_PENDING_QTY'].sum().reset_index()
    fig_pd = px.bar(sp_pd, x='SELLER NAME', y='ACTIVE_PENDING_QTY', title="Pending Qty by Sales Person", text_auto=',.1f', color_discrete_sequence=['#EF4444'])
    st.plotly_chart(fig_pd, use_container_width=True)

    st.subheader("📋 Detailed Pending Dispatch Summary Table")
    pending_details = active_pd[['PO NO', 'DO NO', 'PO_DATE_STR', 'PARTY NAME', 'SELLER NAME', 'ITEM', 'THICKNESS_MM', 'SIZE', 'PO QTY (MT)', 'DISP.QTY', 'ACTIVE_PENDING_QTY', 'REMARK']].copy()
    st.dataframe(pending_details, use_container_width=True)
