import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import re

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
    'DISP.QTY': ['disp.qty', 'disp qty', 'dispatched qty', 'disp_qty'],
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

    # Forward fill SO_DATE so it continues until the next date appears
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
        if c in df.columns:
            df[c] = df[c].fillna('N/A').astype(str).str.strip()

    df['THICKNESS_MM'] = df['THIKNESS'].astype(str).str.replace(r'(?i)\s*mm', '', regex=True).str.strip()

    # Status classifications
    status_lower = df['STATUS'].str.lower()
    remark_lower = df['REMARK'].str.lower()
    
    df['IS_CANCELLED'] = status_lower.str.contains('cancel') | remark_lower.str.contains('cancel')
    df['IS_SHORT_CLOSE'] = status_lower.str.contains(r'\bsc\b|short close') | remark_lower.str.contains(r'\bsc\b|short close')
    df['IS_NEW_CUSTOMER'] = remark_lower.str.contains('new')

    # Status category for donut chart
    def determine_status(row):
        if row['IS_CANCELLED']:
            return 'CANCEL'
        elif row['IS_SHORT_CLOSE']:
            return 'SC'
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
# Sidebar Controls
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
            "2) Sales Executive Performance",
            "3) Customer & Item Deep-Dive"
        ])
    else:
        st.info("👈 Please upload your sales Excel workbook to begin analysis.")
        st.stop()

# ---------------------------------------------------------
# SECTION 1: DISPATCH ANALYSIS
# ---------------------------------------------------------
if section == "1) Dispatch Analysis":
    st.header("📦 Dispatch Analysis & Monthly Overview")
    
    selected_month = st.selectbox("Select Month / Sheet", sheet_names)
    df = load_and_clean_sheet(uploaded_file, selected_month)
    
    # Calculate KPIs
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

    st.subheader("📌 Key Performance Indicators (Click a metric button to view details)")
    
    # Initialize session state for drill-down toggles
    if 'active_drilldown' not in st.session_state:
        st.session_state.active_drilldown = None

    col1, col2, col3, col4, col5 = st.columns(5)
    if col1.button(f"PO Count: {po_count:,}"):
        st.session_state.active_drilldown = "PO"
    if col2.button(f"DO Count: {do_count:,}"):
        st.session_state.active_drilldown = "DO"
    if col3.button(f"Unique Parties: {unique_parties:,}"):
        st.session_state.active_drilldown = "PARTIES"
    if col4.button(f"Total SO Qty: {sum_so_qty:,.2f} MT"):
        st.session_state.active_drilldown = "SO_QTY"
    if col5.button(f"Total Amount: ₹{total_amount:,.0f}"):
        st.session_state.active_drilldown = "AMOUNT"

    col6, col7, col8, col9, col10 = st.columns(5)
    if col6.button(f"Dispatched Qty: {sum_disp_qty:,.2f} MT"):
        st.session_state.active_drilldown = "DISP"
    if col7.button(f"Active Pending: {sum_pending:,.2f} MT"):
        st.session_state.active_drilldown = "PENDING"
    if col8.button(f"Cancelled Qty: {sum_cancelled:,.2f} MT"):
        st.session_state.active_drilldown = "CANCEL"
    if col9.button(f"Short Close Qty: {sum_sc:,.2f} MT"):
        st.session_state.active_drilldown = "SC"
    if col10.button(f"New Parties Added: {new_customers_count}"):
        st.session_state.active_drilldown = "NEW_PARTIES"

    # Display Drill-down Table based on clicked KPI button
    if st.session_state.active_drilldown:
        st.markdown(f"### 🔍 Detailed View: {st.session_state.active_drilldown}")
        if st.session_state.active_drilldown == "PENDING":
            drill_df = df[df['ORDER_STATUS'] == 'PENDING']
        elif st.session_state.active_drilldown == "CANCEL":
            drill_df = df[df['ORDER_STATUS'] == 'CANCEL']
        elif st.session_state.active_drilldown == "SC":
            drill_df = df[df['ORDER_STATUS'] == 'SC']
        elif st.session_state.active_drilldown == "NEW_PARTIES":
            drill_df = new_customers_df
        else:
            drill_df = df
            
        display_cols = [c for c in ['PO_NO', 'DO_NO', 'SO_DATE_STR', 'CUSTOMER_NAME', 'SALES_EXECUTIVE', 'ITEM', 'SO_QTY_(MT)', 'DISP.QTY', 'PENDING', 'ORDER_STATUS', 'REMARK'] if c in drill_df.columns]
        st.dataframe(drill_df[display_cols], use_container_width=True)
        if st.button("Close Drill-down"):
            st.session_state.active_drilldown = None
            st.rerun()

    st.markdown("---")

    # Donut Chart for Order Status
    st.subheader("🍩 Order Status Distribution (Pending, Cancel, OK, SC)")
    status_counts = df['ORDER_STATUS'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Count']
    
    fig_donut = px.pie(
        status_counts, 
        names='Status', 
        values='Count', 
        hole=0.4,
        color='Status',
        color_discrete_map={'OK': '#10B981', 'PENDING': '#F59E0B', 'CANCEL': '#EF4444', 'SC': '#6366F1'}
    )
    fig_donut.update_traces(textinfo='label+value+percent', textfont_size=12)
    st.plotly_chart(fig_donut, use_container_width=True)

# ---------------------------------------------------------
# SECTION 2: SALES EXECUTIVE PERFORMANCE
# ---------------------------------------------------------
elif section == "2) Sales Executive Performance":
    st.header("👨‍💼 Sales Executive & Bar Chart Analytics")
    
    selected_month = st.selectbox("Select Month / Sheet", sheet_names, key="sec2_sheet")
    df = load_and_clean_sheet(uploaded_file, selected_month)
    
    st.subheader("📋 Sales Executive Item-Wise Breakdown")
    sp_item_grp = df.groupby(['SALES_EXECUTIVE', 'ITEM']).agg(
        OrderedQty=('SO_QTY_(MT)', 'sum'),
        DispatchedQty=('DISP.QTY', 'sum'),
        CancelledQty=('SO_QTY_(MT)', lambda x: x[df['ORDER_STATUS'] == 'CANCEL'].sum()),
        ShortCloseQty=('PENDING', lambda x: x[df['ORDER_STATUS'] == 'SC'].sum()),
        Status=('ORDER_STATUS', lambda x: ', '.join(x.unique()))
    ).reset_index()
    st.dataframe(sp_item_grp, use_container_width=True)
    
    st.markdown("---")
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Cancelled Orders Per Sales Person")
        sp_cancelled = df[df['ORDER_STATUS'] == 'CANCEL'].groupby('SALES_EXECUTIVE')['SO_QTY_(MT)'].sum().reset_index()
        sp_cancelled.columns = ['Sales Executive', 'Cancelled Qty']
        fig_cancel = px.bar(sp_cancelled, x='Sales Executive', y='Cancelled Qty', text_auto=',.2f', color_discrete_sequence=['#EF4444'])
        fig_cancel.update_traces(textposition='outside')
        st.plotly_chart(fig_cancel, use_container_width=True)
        
    with col2:
        st.subheader("🥧 Pending Order Share Per Sales Person")
        sp_pending = df[df['ORDER_STATUS'] == 'PENDING'].groupby('SALES_EXECUTIVE')['PENDING'].sum().reset_index()
        sp_pending.columns = ['Sales Executive', 'Pending Qty']
        fig_pending = px.pie(sp_pending, names='Sales Executive', values='Pending Qty', hole=0.3)
        fig_pending.update_traces(textinfo='label+value+percent')
        st.plotly_chart(fig_pending, use_container_width=True)

    st.markdown("---")
    col3, col4 = st.columns(2)
    
    with col3:
        st.subheader("📊 Sum of SO Qty Received — Sales Person Wise")
        sp_so = df.groupby('SALES_EXECUTIVE')['SO_QTY_(MT)'].sum().reset_index()
        sp_so.columns = ['Sales Executive', 'SO Qty (MT)']
        fig_so = px.bar(sp_so, x='Sales Executive', y='SO Qty (MT)', text_auto=',.2f', color_discrete_sequence=['#3B82F6'])
        fig_so.update_traces(textposition='outside')
        st.plotly_chart(fig_so, use_container_width=True)
        
    with col4:
        st.subheader("📊 Sum of Dispatched Qty — Sales Person Wise")
        sp_disp = df.groupby('SALES_EXECUTIVE')['DISP.QTY'].sum().reset_index()
        sp_disp.columns = ['Sales Executive', 'Dispatched Qty']
        fig_disp = px.bar(sp_disp, x='Sales Executive', y='Dispatched Qty', text_auto=',.2f', color_discrete_sequence=['#10B981'])
        fig_disp.update_traces(textposition='outside')
        st.plotly_chart(fig_disp, use_container_width=True)

# ---------------------------------------------------------
# SECTION 3: CUSTOMER & ITEM DEEP-DIVE
# ---------------------------------------------------------
elif section == "3) Customer & Item Deep-Dive":
    st.header("🏢 Customer & Item Detailed Analytics")
    
    selected_month = st.selectbox("Select Month / Sheet", sheet_names, key="sec3_sheet")
    df = load_and_clean_sheet(uploaded_file, selected_month)
    
    st.subheader("📋 Customer Wise Detailed Table")
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
        top_cust, 
        x='CUSTOMER_NAME', 
        y=['Ordered', 'Dispatched', 'Cancelled', 'Pending', 'ShortClose'],
        barmode='group',
        text_auto=',.1f'
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
    
    item_spec_grp.rename(columns={
        'ITEM': 'ItemName',
        'THICKNESS_MM': 'Thickness',
        'WIDTH_MM': 'Width'
    }, inplace=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        selected_items = st.multiselect("Filter Item Name(s)", sorted(item_spec_grp['ItemName'].unique().tolist()))
    with c2:
        selected_thickness = st.multiselect("Filter Thickness", sorted(item_spec_grp['Thickness'].unique().tolist(), key=str))
    with c3:
        selected_width = st.multiselect("Filter Width", sorted(item_spec_grp['Width'].unique().tolist(), key=str))

    filtered_spec = item_spec_grp.copy()
    if selected_items:
        filtered_spec = filtered_spec[filtered_spec['ItemName'].isin(selected_items)]
    if selected_thickness:
        filtered_spec = filtered_spec[filtered_spec['Thickness'].isin(selected_thickness)]
    if selected_width:
        filtered_spec = filtered_spec[filtered_spec['Width'].isin(selected_width)]

    st.dataframe(filtered_spec, use_container_width=True)
