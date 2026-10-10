import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import re

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="Sales & Dispatch Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Sales & Dispatch Analytics Dashboard")

# ---------------------------------------------------------
# Flexible Excel Header & Column Normalization
# ---------------------------------------------------------
EXPECTED_COLUMNS = {
    'S_NO': ['s_no', 's.no', 'sr no', 'serial no', 'sno', 's_no.'],
    'PO NO': ['po no', 'po_no', 'pono', 'order no', 'orderno'],
    'PO DATE': ['po date', 'podate', 'po_date', 'order date', 'so_date'],
    'DO NO': ['do no', 'do_no', 'dono', 'dispatch no', 'dispatch_no'],
    'DATE': ['date', 'disp date', 'dispatch date', 'inv date'],
    'PARTY NAME': ['party name', 'customer name', 'customer_name', 'client', 'buyer', 'party'],
    'SELLER NAME': ['seller name', 'seller', 'sales executive', 'sales_executive', 'executive', 'broker'],
    'PLACE': ['place', 'location', 'city'],
    'ITEM': ['item', 'material', 'product', 'item name'],
    'THICKNESS': ['thickness', 'thikness', 'thick', 'thk'],
    'SIZE': ['size', 'size_(mm)', 'width', 'dimension'],
    'PO QTY (MT)': ['po qty (mt)', 'po qty', 'ordered qty', 'qty', 'order qty', 'quantity', 'so qty'],
    'PER TON': ['per ton', 'rate', 'price', 'unit rate'],
    'DISP.QTY': ['disp.qty', 'disp qty', 'dispatched qty', 'dispatched', 'dispatch_qty'],
    'PENDING': ['pending', 'pending qty', 'balance', 'bal qty'],
    'STATUS': ['status', 'order status'],
    'REMARK': ['remark', 'remarks', 'notes', 'comment']
}

def normalize_columns(df):
    df.columns = df.columns.astype(str).str.strip()
    rename_map = {}
    for col in df.columns:
        col_lower = col.lower()
        matched = False
        for standard_name, variants in EXPECTED_COLUMNS.items():
            if col_lower in variants or any(v in col_lower for v in variants):
                rename_map[col] = standard_name
                matched = True
                break
        if not matched:
            rename_map[col] = col.upper()
    df = df.rename(columns=rename_map)
    return df

@st.cache_data
def load_data(file):
    xls = pd.ExcelFile(file)
    sheets_data = {}
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet=sheet)
            if df.empty or len(df.columns) < 2:
                continue
            df = normalize_columns(df)
            
            # Clean and parse dates
            for d_col in ['PO DATE', 'DATE']:
                if d_col in df.columns:
                    d_clean = df[d_col].astype(str).str.replace('.', '/', regex=False)
                    df[d_col] = pd.to_datetime(d_clean, dayfirst=True, errors='coerce')
            
            df['PO_DATE_STR'] = df['PO DATE'].dt.strftime('%d/%m/%Y').fillna('N/A') if 'PO DATE' in df.columns else 'N/A'
            df['MONTH_YEAR'] = df['PO DATE'].dt.strftime('%B %Y').fillna('Unknown') if 'PO DATE' in df.columns else sheet
            
            # Numeric conversion
            for num_col in ['PO QTY (MT)', 'DISP.QTY', 'PENDING', 'PER TON', 'THICKNESS']:
                if num_col in df.columns:
                    df[num_col] = pd.to_numeric(df[num_col], errors='coerce').fillna(0)
            
            # Calculate Total Amount
            if 'PO QTY (MT)' in df.columns and 'PER TON' in df.columns:
                df['TOTAL_AMOUNT'] = df['PO QTY (MT)'] * df['PER TON']
            else:
                df['TOTAL_AMOUNT'] = 0.0
                
            # Status and Category Derived Columns
            if 'STATUS' not in df.columns:
                df['STATUS'] = 'PENDING'
            if 'REMARK' not in df.columns:
                df['REMARK'] = ''
                
            df['STATUS_CLEAN'] = df['STATUS'].astype(str).str.upper().str.strip()
            df['REMARK_CLEAN'] = df['REMARK'].astype(str).str.upper().str.strip()
            
            # Categorize Status / Short Close / Cancel / OK
            def determine_order_state(row):
                status = row['STATUS_CLEAN']
                remark = row['REMARK_CLEAN']
                disp = row.get('DISP.QTY', 0)
                po = row.get('PO QTY (MT)', 0)
                
                if 'CANCEL' in status or 'CANCEL' in remark:
                    return 'CANCEL'
                elif 'SC' in status or 'SHORT' in status or 'SC' in remark.split() or 'SHORT' in remark:
                    return 'SC'
                elif status == 'PENDING' and disp < po:
                    return 'PENDING'
                else:
                    return 'OK'
                    
            df['DERIVED_STATUS'] = df.apply(determine_order_state, axis=1)
            sheets_data[sheet] = df
        except Exception as e:
            st.warning(f"Could not process sheet '{sheet}': {e}")
    return sheets_data

# ---------------------------------------------------------
# Sidebar Collapsible File Upload & Navigation
# ---------------------------------------------------------
with st.sidebar:
    st.header("📁 Upload & Navigation")
    uploaded_file = st.file_uploader("Upload Sales Order Excel File", type=["xlsx", "xls"])
    
    st.markdown("---")
    section = st.radio(
        "Navigation Sections",
        [
            "1) Dispatch Analysis",
            "2) Date Wise / Month Wise Comparison",
            "3) Pending Dispatch"
        ]
    )

if uploaded_file is not None:
    data_dict = load_data(uploaded_file)
    sheet_names = list(data_dict.keys())
    
    if not sheet_names:
        st.error("No valid data sheets found in the uploaded workbook.")
    else:
        # ---------------------------------------------------------
        # SECTION 1: DISPATCH ANALYSIS
        # ---------------------------------------------------------
        if section == "1) Dispatch Analysis":
            st.header("📊 Dispatch Analysis & Performance KPIs")
            
            # Select Month / Sheet
            selected_month = st.selectbox("Select Month / Sheet", sheet_names)
            df = data_dict[selected_month].copy()
            
            # KPI Calculations
            po_count = df['PO NO'].nunique() if 'PO NO' in df.columns else len(df)
            do_count = df['DO NO'].nunique() if 'DO NO' in df.columns else 0
            unique_parties = df['PARTY NAME'].nunique() if 'PARTY NAME' in df.columns else 0
            sum_so_qty = df['PO QTY (MT)'].sum() if 'PO QTY (MT)' in df.columns else 0
            total_amount = df['TOTAL_AMOUNT'].sum() if 'TOTAL_AMOUNT' in df.columns else 0
            sum_disp_qty = df['DISP.QTY'].sum() if 'DISP.QTY' in df.columns else 0
            
            pending_df = df[df['DERIVED_STATUS'] == 'PENDING']
            sum_pending = pending_df['PO QTY (MT)'].sum() if 'PO QTY (MT)' in pending_df.columns else 0
            
            cancel_df = df[df['DERIVED_STATUS'] == 'CANCEL']
            sum_cancel = cancel_df['PO QTY (MT)'].sum() if 'PO QTY (MT)' in cancel_df.columns else 0
            
            sc_df = df[df['DERIVED_STATUS'] == 'SC']
            sum_sc = sc_df['PO QTY (MT)'].sum() if 'PO QTY (MT)' in sc_df.columns else 0
            
            new_parties_df = df[df['REMARK_CLEAN'].str.contains('NEW', na=False)]
            new_parties_count = new_parties_df['PARTY NAME'].nunique() if 'PARTY NAME' in new_parties_df.columns else 0

            st.subheader(f"Key Performance Metrics ({selected_month})")
            
            # KPI Button / Expandable Grid
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                if st.button(f"📌 PO Count: {po_count:,}", use_container_width=True):
                    st.session_state['kpi_view'] = 'po'
            with col2:
                if st.button(f"📦 DO Count: {do_count:,}", use_container_width=True):
                    st.session_state['kpi_view'] = 'do'
            with col3:
                if st.button(f"🏢 Unique Parties: {unique_parties:,}", use_container_width=True):
                    st.session_state['kpi_view'] = 'parties'
            with col4:
                if st.button(f"⚖️ Sum SO Qty: {sum_so_qty:,.2f} MT", use_container_width=True):
                    st.session_state['kpi_view'] = 'so_qty'

            col5, col6, col7, col8 = st.columns(4)
            with col5:
                if st.button(f"💰 Total Amount: ₹{total_amount:,.2f}", use_container_width=True):
                    st.session_state['kpi_view'] = 'amount'
            with col6:
                if st.button(f"🚚 Dispatched Qty: {sum_disp_qty:,.2f} MT", use_container_width=True):
                    st.session_state['kpi_view'] = 'disp'
            with col7:
                if st.button(f"⏳ Active Pending: {sum_pending:,.2f} MT", use_container_width=True):
                    st.session_state['kpi_view'] = 'pending'
            with col8:
                if st.button(f"❌ Cancelled Qty: {sum_cancel:,.2f} MT", use_container_width=True):
                    st.session_state['kpi_view'] = 'cancel'

            col9, col10 = st.columns(2)
            with col9:
                if st.button(f"⚠️ Short Close Qty: {sum_sc:,.2f} MT", use_container_width=True):
                    st.session_state['kpi_view'] = 'sc'
            with col10:
                if st.button(f"✨ New Parties Added: {new_parties_count}", use_container_width=True):
                    st.session_state['kpi_view'] = 'new_parties'

            # Render Detailed View if KPI Button Clicked
            if 'kpi_view' in st.session_state:
                st.markdown("---")
                kv = st.session_state['kpi_view']
                if kv == 'po':
                    st.markdown("### 📋 Detailed PO Records")
                    st.dataframe(df, use_container_width=True)
                elif kv == 'do':
                    st.markdown("### 📦 Detailed DO Records")
                    st.dataframe(df[df['DO NO'].notna()], use_container_width=True)
                elif kv == 'parties':
                    st.markdown("### 🏢 Unique Parties List")
                    st.dataframe(df[['PARTY NAME', 'PLACE', 'SELLER NAME']].drop_duplicates(), use_container_width=True)
                elif kv == 'so_qty':
                    st.markdown("### ⚖️ SO Quantity Details")
                    st.dataframe(df[['PO NO', 'PARTY NAME', 'ITEM', 'PO QTY (MT)']], use_container_width=True)
                elif kv == 'amount':
                    st.markdown("### 💰 Total Amount Breakdown")
                    st.dataframe(df[['PO NO', 'PARTY NAME', 'ITEM', 'PO QTY (MT)', 'PER TON', 'TOTAL_AMOUNT']], use_container_width=True)
                elif kv == 'disp':
                    st.markdown("### 🚚 Dispatched Quantity Breakdown")
                    st.dataframe(df[['PO NO', 'DO NO', 'PARTY NAME', 'ITEM', 'DISP.QTY']], use_container_width=True)
                elif kv == 'pending':
                    st.markdown("### ⏳ Active Pending Orders")
                    st.dataframe(pending_df, use_container_width=True)
                elif kv == 'cancel':
                    st.markdown("### ❌ Cancelled Orders")
                    st.dataframe(cancel_df, use_container_width=True)
                elif kv == 'sc':
                    st.markdown("### ⚠️ Short Closed Orders")
                    st.dataframe(sc_df, use_container_width=True)
                elif kv == 'new_parties':
                    st.markdown("### ✨ Newly Added Parties")
                    st.dataframe(new_parties_df[['PARTY NAME', 'REMARK', 'PO DATE', 'SELLER NAME']].drop_duplicates(), use_container_width=True)
                if st.button("Close Detailed View"):
                    del st.session_state['kpi_view']
                    st.rerun()

            st.markdown("---")
            
            # Donut Chart for Status Breakdown (Pending, Cancel, OK, SC)
            status_counts = df['DERIVED_STATUS'].value_counts().reset_index()
            status_counts.columns = ['Status', 'Count']
            fig_donut = px.pie(status_counts, names='Status', values='Count', hole=0.4, title="Order Status Distribution (Pending, Cancel, OK, SC)")
            fig_donut.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_donut, use_container_width=True)

            st.markdown("---")
            
            # 1. Sales Executive Item Wise BreakDown Table
            st.subheader("📋 Sales Executive & Item-wise Breakdown")
            if 'SELLER NAME' in df.columns and 'ITEM' in df.columns:
                exec_item_grp = df.groupby(['SELLER NAME', 'ITEM', 'STATUS']).agg({
                    'PO QTY (MT)': 'sum',
                    'DISP.QTY': 'sum'
                }).reset_index()
                
                # Calculate specific status totals per group
                exec_item_summary = df.groupby(['SELLER NAME', 'ITEM']).agg({
                    'PO QTY (MT)': 'sum',
                    'DISP.QTY': 'sum'
                }).reset_index()
                
                exec_item_summary['CancelledQty'] = df[df['DERIVED_STATUS'] == 'CANCEL'].groupby(['SELLER NAME', 'ITEM'])['PO QTY (MT)'].sum().reset_index(drop=True)
                exec_item_summary['CancelledQty'] = exec_item_summary['CancelledQty'].fillna(0)
                
                exec_item_summary['ShortCloseQty'] = df[df['DERIVED_STATUS'] == 'SC'].groupby(['SELLER NAME', 'ITEM'])['PO QTY (MT)'].sum().reset_index(drop=True)
                exec_item_summary['ShortCloseQty'] = exec_item_summary['ShortCloseQty'].fillna(0)
                
                exec_item_summary.rename(columns={'PO QTY (MT)': 'OrderedQty', 'DISP.QTY': 'DispatchedQty'}, inplace=True)
                exec_item_summary['Status'] = 'Active'
                st.dataframe(exec_item_summary, use_container_width=True)

            st.markdown("---")
            
            # 2. Cancelled Orders Per Sales Person Bar Graph
            st.subheader("❌ Cancelled Orders Per Sales Person")
            cancel_sales = cancel_df.groupby('SELLER NAME')['PO QTY (MT)'].sum().reset_index()
            if not cancel_sales.empty:
                fig_cancel = px.bar(cancel_sales, x='SELLER NAME', y='PO QTY (MT)', text='PO QTY (MT)', title="Cancelled Orders by Sales Person")
                fig_cancel.update_traces(texttemplate='%{text:.2f}', textposition='outside')
                st.plotly_chart(fig_cancel, use_container_width=True)
            else:
                st.info("No cancelled orders found in this selection.")

            st.markdown("---")
            
            # 3. Pending Order Per Sales Person Pie Chart
            st.subheader("⏳ Pending Orders Per Sales Person")
            pending_sales = pending_df.groupby('SELLER NAME')['PO QTY (MT)'].sum().reset_index()
            if not pending_sales.empty:
                fig_pend_pie = px.pie(pending_sales, names='SELLER NAME', values='PO QTY (MT)', title="Pending Orders Share by Sales Person")
                fig_pend_pie.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_pend_pie, use_container_width=True)
            else:
                st.info("No pending orders found in this selection.")

            st.markdown("---")
            
            # 4. Sum of SO Qty Received Sales Person Wise Bar Graph
            st.subheader("📊 Total SO Quantity Received Sales Person Wise")
            so_sales = df.groupby('SELLER NAME')['PO QTY (MT)'].sum().reset_index()
            fig_so_sales = px.bar(so_sales, x='SELLER NAME', y='PO QTY (MT)', text='PO QTY (MT)', title="SO Qty Received by Sales Executive")
            fig_so_sales.update_traces(texttemplate='%{text:.2f}', textposition='outside')
            st.plotly_chart(fig_so_sales, use_container_width=True)

            st.markdown("---")
            
            # 5. Sum of Dispatched Qty Bar Graph
            st.subheader("🚚 Dispatched Quantity Sales Person Wise")
            disp_sales = df.groupby('SELLER NAME')['DISP.QTY'].sum().reset_index()
            fig_disp_sales = px.bar(disp_sales, x='SELLER NAME', y='DISP.QTY', text='DISP.QTY', title="Dispatched Qty by Sales Executive")
            fig_disp_sales.update_traces(texttemplate='%{text:.2f}', textposition='outside')
            st.plotly_chart(fig_disp_sales, use_container_width=True)

            st.markdown("---")
            
            # 6. Customer Wise Detailed Table
            st.subheader("🏢 Customer-wise Detailed Table")
            cust_summary = df.groupby(['PARTY NAME', 'SELLER NAME']).agg({
                'PO QTY (MT)': 'sum',
                'DISP.QTY': 'sum'
            }).reset_index()
            
            cust_summary['Cancelled'] = df[df['DERIVED_STATUS'] == 'CANCEL'].groupby(['PARTY NAME', 'SELLER NAME'])['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
            cust_summary['Pending'] = df[df['DERIVED_STATUS'] == 'PENDING'].groupby(['PARTY NAME', 'SELLER NAME'])['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
            cust_summary['Short Close'] = df[df['DERIVED_STATUS'] == 'SC'].groupby(['PARTY NAME', 'SELLER NAME'])['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
            
            cust_summary.rename(columns={'PO QTY (MT)': 'Ordered', 'DISP.QTY': 'Dispatched'}, inplace=True)
            st.dataframe(cust_summary, use_container_width=True)

            st.markdown("---")
            
            # 7. Top 15 Customers Graph
            st.subheader("🏆 Top 15 Customers Breakdown")
            top_15 = cust_summary.sort_values(by='Ordered', ascending=False).head(15)
            fig_top15 = px.bar(top_15, x='PARTY NAME', y=['Ordered', 'Dispatched', 'Cancelled', 'Pending', 'Short Close'], barmode='group', title="Top 15 Customers Performance")
            fig_top15.update_traces(textposition='outside')
            st.plotly_chart(fig_top15, use_container_width=True)

            st.markdown("---")
            
            # 8. Detailed Item, Thickness, Width Table with Multi-Select Filters
            st.subheader("📦 Detailed Item, Thickness & Width Analysis")
            all_items = sorted(df['ITEM'].dropna().unique().tolist()) if 'ITEM' in df.columns else []
            all_thickness = sorted(df['THICKNESS'].dropna().unique().tolist()) if 'THICKNESS' in df.columns else []
            all_width = sorted(df['SIZE'].dropna().unique().tolist()) if 'SIZE' in df.columns else []
            
            f_item = st.multiselect("Filter Item Name(s)", all_items)
            f_thick = st.multiselect("Filter Thickness", all_thickness)
            f_width = st.multiselect("Filter Width / Size", all_width)
            
            filtered_spec_df = df.copy()
            if f_item:
                filtered_spec_df = filtered_spec_df[filtered_spec_df['ITEM'].isin(f_item)]
            if f_thick:
                filtered_spec_df = filtered_spec_df[filtered_spec_df['THICKNESS'].isin(f_thick)]
            if f_width:
                filtered_spec_df = filtered_spec_df[filtered_spec_df['SIZE'].isin(f_width)]
                
            spec_summary = filtered_spec_df.groupby(['ITEM', 'THICKNESS', 'SIZE']).agg({
                'PO QTY (MT)': 'sum',
                'DISP.QTY': 'sum'
            }).reset_index()
            
            spec_summary['Cancelled'] = filtered_spec_df[filtered_spec_df['DERIVED_STATUS'] == 'CANCEL'].groupby(['ITEM', 'THICKNESS', 'SIZE'])['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
            spec_summary['Pending'] = filtered_spec_df[filtered_spec_df['DERIVED_STATUS'] == 'PENDING'].groupby(['ITEM', 'THICKNESS', 'SIZE'])['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
            spec_summary['ShortClose'] = filtered_spec_df[filtered_spec_df['DERIVED_STATUS'] == 'SC'].groupby(['ITEM', 'THICKNESS', 'SIZE'])['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
            
            spec_summary.rename(columns={'ITEM': 'ItemName', 'PO QTY (MT)': 'OrderedQty', 'DISP.QTY': 'DispatchedQty'}, inplace=True)
            st.dataframe(spec_summary, use_container_width=True)

        # ---------------------------------------------------------
        # SECTION 2: DATE WISE / MONTH WISE COMPARISON
        # ---------------------------------------------------------
        elif section == "2) Date Wise / Month Wise Comparison":
            st.header("📅 Date-Wise, Month-Wise & Year-Wise Comparison")
            
            comp_mode = st.radio("Select Comparison Level", ["Date-Wise", "Month-Wise", "Year-Wise"], horizontal=True)
            
            # Combine all sheets for multi-month/year comparison
            combined_df = pd.concat(data_dict.values(), ignore_index=True)
            
            if comp_mode == "Date-Wise" and 'PO DATE' in combined_df.columns:
                valid_dates_df = combined_df.dropna(subset=['PO DATE']).copy()
                valid_dates_df['DATE_ONLY'] = valid_dates_df['PO DATE'].dt.date
                date_options = sorted(valid_dates_df['DATE_ONLY'].unique())
                
                if len(date_options) >= 2:
                    d1 = st.selectbox("Select First Date", date_options, index=0)
                    d2 = st.selectbox("Select Second Date", date_options, index=min(1, len(date_options)-1))
                    
                    df_d1 = valid_dates_df[valid_dates_df['DATE_ONLY'] == d1]
                    df_d2 = valid_dates_df[valid_dates_df['DATE_ONLY'] == d2]
                    
                    po1, disp1 = df_d1['PO QTY (MT)'].sum(), df_d1['DISP.QTY'].sum()
                    po2, disp2 = df_d2['PO QTY (MT)'].sum(), df_d2['DISP.QTY'].sum()
                    
                    st.subheader(f"Comparison: {d1} vs {d2}")
                    col1, col2 = st.columns(2)
                    col1.metric(f"Ordered Qty ({d1})", f"{po1:,.2f} MT", delta=f"{po2 - po1:,.2f} MT vs {d2}")
                    col2.metric(f"Dispatched Qty ({d1})", f"{disp1:,.2f} MT", delta=f"{disp2 - disp1:,.2f} MT vs {d2}")
                    
                    comp_table = pd.DataFrame({
                        'Metric': ['Ordered Qty (MT)', 'Dispatched Qty (MT)', 'Pending Qty (MT)'],
                        str(d1): [po1, disp1, po1 - disp1],
                        str(d2): [po2, disp2, po2 - disp2],
                        'Variance (MT)': [po2 - po1, disp2 - disp1, (po2 - disp2) - (po1 - disp1)]
                    })
                    st.dataframe(comp_table, use_container_width=True)
                    
                    fig_comp = px.bar(comp_table, x='Metric', y=[str(d1), str(d2)], barmode='group', title=f"Date Comparison Chart")
                    st.plotly_chart(fig_comp, use_container_width=True)
                else:
                    st.warning("Insufficient dates available for comparison.")
                    
            elif comp_mode == "Month-Wise":
                month_options = list(data_dict.keys())
                if len(month_options) >= 2:
                    m1 = st.selectbox("Select First Month / Sheet", month_options, index=0)
                    m2 = st.selectbox("Select Second Month / Sheet", month_options, index=min(1, len(month_options)-1))
                    
                    df_m1 = data_dict[m1]
                    df_m2 = data_dict[m2]
                    
                    po1, disp1 = df_m1['PO QTY (MT)'].sum(), df_m1['DISP.QTY'].sum()
                    po2, disp2 = df_m2['PO QTY (MT)'].sum(), df_m2['DISP.QTY'].sum()
                    
                    st.subheader(f"Comparison: {m1} vs {m2}")
                    comp_table = pd.DataFrame({
                        'Metric': ['Ordered Qty (MT)', 'Dispatched Qty (MT)', 'Total Amount (₹)'],
                        m1: [po1, disp1, df_m1['TOTAL_AMOUNT'].sum()],
                        m2: [po2, disp2, df_m2['TOTAL_AMOUNT'].sum()],
                        'Variance': [po2 - po1, disp2 - disp1, df_m2['TOTAL_AMOUNT'].sum() - df_m1['TOTAL_AMOUNT'].sum()]
                    })
                    st.dataframe(comp_table, use_container_width=True)
                    
                    fig_mcomp = px.bar(comp_table, x='Metric', y=[m1, m2], barmode='group', title="Month Comparison Chart")
                    st.plotly_chart(fig_mcomp, use_container_width=True)
                else:
                    st.warning("Need at least two sheets for month-wise comparison.")
            else:
                st.info("Year-wise analysis aggregates data across all loaded sheets categorized by year.")
                if 'PO DATE' in combined_df.columns:
                    combined_df['YEAR'] = combined_df['PO DATE'].dt.year.fillna(2026).astype(int)
                    year_summary = combined_df.groupby('YEAR').agg({
                        'PO QTY (MT)': 'sum',
                        'DISP.QTY': 'sum',
                        'TOTAL_AMOUNT': 'sum'
                    }).reset_index()
                    st.dataframe(year_summary, use_container_width=True)
                    fig_yr = px.bar(year_summary, x='YEAR', y='PO QTY (MT)', text='PO QTY (MT)', title="Year-Wise Ordered Quantity")
                    st.plotly_chart(fig_yr, use_container_width=True)

        # ---------------------------------------------------------
        # SECTION 3: PENDING DISPATCH
        # ---------------------------------------------------------
        elif section == "3) Pending Dispatch":
            st.header("⏳ Pending Dispatch Report & Summary")
            
            # Look for pending dispatch sheet or filter active pending
            pending_candidates = [s for s in sheet_names if 'pending' in s.lower() and 'dispatch' in s.lower()]
            target_sheet = pending_candidates[0] if pending_candidates else sheet_names[0]
            
            pending_report_df = data_dict[target_sheet]
            active_pending_df = pending_report_df[pending_report_df['DERIVED_STATUS'] == 'PENDING']
            
            total_pend_qty = active_pending_df['PO QTY (MT)'].sum() if 'PO QTY (MT)' in active_pending_df.columns else 0
            total_pend_count = len(active_pending_df)
            
            col1, col2 = st.columns(2)
            col1.metric("Total Pending Orders Count", f"{total_pend_count:,}")
            col2.metric("Total Active Pending Quantity (MT)", f"{total_pend_qty:,.2f}")
            
            st.markdown("---")
            st.subheader(f"Detailed Pending Dispatch Report ({target_sheet})")
            display_cols = [c for c in ['PO NO', 'DO NO', 'PO_DATE_STR', 'PARTY NAME', 'SELLER NAME', 'ITEM', 'THICKNESS', 'SIZE', 'PO QTY (MT)', 'DISP.QTY', 'PENDING', 'REMARK'] if c in active_pending_df.columns]
            st.dataframe(active_pending_df[display_cols], use_container_width=True)

else:
    st.info("👈 Please upload your Sales Order Excel file using the sidebar to begin.")
