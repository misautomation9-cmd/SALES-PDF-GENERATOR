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
    page_title="Ironmart Sales Analytical Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Ironmart Sales Analytical Dashboard")

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
            df = pd.read_excel(xls, sheet_name=sheet)
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
            
            # Dimension parsing for width (1250X5635 -> Width = 1250)
            if 'SIZE' in df.columns:
                def extract_width(val):
                    v_str = str(val).upper().strip()
                    if 'X' in v_str:
                        parts = v_str.split('X')
                        try:
                            return float(parts[0].strip())
                        except:
                            return 0.0
                    try:
                        return float(v_str)
                    except:
                        return 0.0
                df['WIDTH_PARSED'] = df['SIZE'].apply(extract_width)
            else:
                df['WIDTH_PARSED'] = 0.0

            # Calculate Total Amount
            if 'PO QTY (MT)' in df.columns and 'PER TON' in df.columns:
                df['TOTAL_AMOUNT'] = df['PO QTY (MT)'] * df['PER TON']
            else:
                df['TOTAL_AMOUNT'] = 0.0
                
            # Status and Remarks normalization
            if 'STATUS' not in df.columns:
                df['STATUS'] = 'PENDING'
            if 'REMARK' not in df.columns:
                df['REMARK'] = ''
                
            df['STATUS_CLEAN'] = df['STATUS'].astype(str).str.upper().str.strip()
            df['REMARK_CLEAN'] = df['REMARK'].astype(str).str.upper().str.strip()
            
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
                    return 'DISPATCH'
                    
            df['DERIVED_STATUS'] = df.apply(determine_order_state, axis=1)
            sheets_data[sheet] = df
        except Exception as e:
            st.warning(f"Could not process sheet '{sheet}': {e}")
    return sheets_data

# ---------------------------------------------------------
# Sidebar Collapsible Upload & Navigation
# ---------------------------------------------------------
with st.sidebar:
    st.header("📁 Upload & Navigation")
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
    
    st.markdown("---")
    section = st.radio(
        "Navigation Sections",
        [
            "1) Sales Analysis",
            "2) Date/Month/Year Comparison",
            "3) Dispatch Pending"
        ]
    )

if uploaded_file is not None:
    data_dict = load_data(uploaded_file)
    sheet_names = list(data_dict.keys())
    
    if not sheet_names:
        st.error("No valid data sheets found in the uploaded workbook.")
    else:
        # ---------------------------------------------------------
        # SECTION 1: SALES ANALYSIS
        # ---------------------------------------------------------
        if section == "1) Sales Analysis":
            st.header("📈 Sales Analysis & Interactive KPIs")
            
            selected_month = st.selectbox("Select Month / Sheet", sheet_names)
            df = data_dict[selected_month].copy()
            
            # Global Filters for Section 1
            st.sidebar.markdown("---")
            st.sidebar.subheader("🔎 Section Filters")
            
            all_cust = ["All"] + sorted(df['PARTY NAME'].dropna().unique().tolist()) if 'PARTY NAME' in df.columns else ["All"]
            all_sellers = ["All"] + sorted(df['SELLER NAME'].dropna().unique().tolist()) if 'SELLER NAME' in df.columns else ["All"]
            all_items = ["All"] + sorted(df['ITEM'].dropna().unique().tolist()) if 'ITEM' in df.columns else ["All"]
            
            f_cust = st.sidebar.selectbox("Filter Customer Name", all_cust)
            f_seller = st.sidebar.selectbox("Filter Sales Executive", all_sellers)
            f_item = st.sidebar.selectbox("Filter Item Name", all_items)
            
            # Apply filters
            filtered_df = df.copy()
            if f_cust != "All":
                filtered_df = filtered_df[filtered_df['PARTY NAME'] == f_cust]
            if f_seller != "All":
                filtered_df = filtered_df[filtered_df['SELLER NAME'] == f_seller]
            if f_item != "All":
                filtered_df = filtered_df[filtered_df['ITEM'] == f_item]
            
            # KPI Computations
            unique_cust_count = filtered_df['PARTY NAME'].nunique() if 'PARTY NAME' in filtered_df.columns else 0
            sum_so_qty = filtered_df['PO QTY (MT)'].sum() if 'PO QTY (MT)' in filtered_df.columns else 0
            total_amount = filtered_df['TOTAL_AMOUNT'].sum() if 'TOTAL_AMOUNT' in filtered_df.columns else 0
            sum_disp_qty = filtered_df['DISP.QTY'].sum() if 'DISP.QTY' in filtered_df.columns else 0
            
            cancel_df = filtered_df[filtered_df['DERIVED_STATUS'] == 'CANCEL']
            sum_cancel_qty = cancel_df['PO QTY (MT)'].sum() if 'PO QTY (MT)' in cancel_df.columns else 0
            
            active_pend_df = filtered_df[filtered_df['DERIVED_STATUS'] == 'PENDING']
            sum_active_pend = active_pend_df['PO QTY (MT)'].sum() if 'PO QTY (MT)' in active_pend_df.columns else 0
            
            sc_df = filtered_df[filtered_df['DERIVED_STATUS'] == 'SC']
            sum_sc_qty = sc_df['PO QTY (MT)'].sum() if 'PO QTY (MT)' in sc_df.columns else 0
            
            new_parties_df = filtered_df[filtered_df['REMARK_CLEAN'].str.contains('NEW', na=False)]
            new_parties_count = new_parties_df['PARTY NAME'].nunique() if 'PARTY NAME' in new_parties_df.columns else 0

            st.subheader(f"Performance KPIs — {selected_month}")
            
            # Interactive KPI Drilldown Buttons
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                if st.button(f"👥 Unique Customers: {unique_cust_count:,}", use_container_width=True):
                    st.session_state['drill'] = 'customers'
            with c2:
                if st.button(f"📦 Sum SO Qty: {sum_so_qty:,.2f} MT", use_container_width=True):
                    st.session_state['drill'] = 'so_qty'
            with c3:
                if st.button(f"💰 Total Amount: ₹{total_amount:,.2f}", use_container_width=True):
                    st.session_state['drill'] = 'amount'
            with c4:
                if st.button(f"🚚 Dispatch Qty: {sum_disp_qty:,.2f} MT", use_container_width=True):
                    st.session_state['drill'] = 'disp_qty'

            c5, c6, c7, c8 = st.columns(4)
            with c5:
                if st.button(f"❌ Cancel Qty: {sum_cancel_qty:,.2f} MT", use_container_width=True):
                    st.session_state['drill'] = 'cancel_qty'
            with c6:
                if st.button(f"⏳ Active Pending: {sum_active_pend:,.2f} MT", use_container_width=True):
                    st.session_state['drill'] = 'active_pend'
            with c7:
                if st.button(f"⚠️ SC Qty: {sum_sc_qty:,.2f} MT", use_container_width=True):
                    st.session_state['drill'] = 'sc_qty'
            with c8:
                if st.button(f"✨ New Parties: {new_parties_count}", use_container_width=True):
                    st.session_state['drill'] = 'new_parties'

            # Drilldown Report Display
            if 'drill' in st.session_state:
                st.markdown("---")
                d_type = st.session_state['drill']
                if d_type == 'customers':
                    st.markdown("### 👥 Drilldown: Unique Customers")
                    st.dataframe(filtered_df[['PARTY NAME', 'PLACE', 'SELLER NAME']].drop_duplicates(), use_container_width=True)
                elif d_type == 'so_qty':
                    st.markdown("### 📦 Drilldown: SO Quantity Orders")
                    st.dataframe(filtered_df[['PO NO', 'PARTY NAME', 'ITEM', 'PO QTY (MT)']], use_container_width=True)
                elif d_type == 'amount':
                    st.markdown("### 💰 Drilldown: Total Amount Breakdown")
                    st.dataframe(filtered_df[['PO NO', 'PARTY NAME', 'ITEM', 'PO QTY (MT)', 'PER TON', 'TOTAL_AMOUNT']], use_container_width=True)
                elif d_type == 'disp_qty':
                    st.markdown("### 🚚 Drilldown: Dispatched Quantities")
                    st.dataframe(filtered_df[['PO NO', 'DO NO', 'PARTY NAME', 'ITEM', 'DISP.QTY']], use_container_width=True)
                elif d_type == 'cancel_qty':
                    st.markdown("### ❌ Drilldown: Cancelled Orders")
                    st.dataframe(cancel_df, use_container_width=True)
                elif d_type == 'active_pend':
                    st.markdown("### ⏳ Drilldown: Active Pending Orders")
                    st.dataframe(active_pend_df, use_container_width=True)
                elif d_type == 'sc_qty':
                    st.markdown("### ⚠️ Drilldown: Short Closed (SC) Orders")
                    st.dataframe(sc_df, use_container_width=True)
                elif d_type == 'new_parties':
                    st.markdown("### ✨ Drilldown: New Parties Added")
                    st.dataframe(new_parties_df[['PARTY NAME', 'REMARK', 'PO DATE', 'SELLER NAME']].drop_duplicates(), use_container_width=True)
                
                if st.button("Close Drilldown View"):
                    del st.session_state['drill']
                    st.rerun()

            st.markdown("---")
            
            # Donut Chart (Ordered, Pending, Cancel, SC, Dispatch)
            status_summary = filtered_df['DERIVED_STATUS'].value_counts().reset_index()
            status_summary.columns = ['Status', 'Count']
            fig_donut = px.pie(status_summary, names='Status', values='Count', hole=0.4, title="Order Status Distribution (Ordered, Pending, Cancel, SC, Dispatch)")
            fig_donut.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_donut, use_container_width=True)

            st.markdown("---")
            
            # Sales Person Wise Ordered Qty (Bar Graph with data labels)
            st.subheader("📊 Sales Person Wise Ordered Quantity")
            seller_ordered = filtered_df.groupby('SELLER NAME')['PO QTY (MT)'].sum().reset_index()
            fig_seller_ord = px.bar(seller_ordered, x='SELLER NAME', y='PO QTY (MT)', text='PO QTY (MT)', title="Ordered Qty by Sales Person")
            fig_seller_ord.update_traces(texttemplate='%{text:.2f}', textposition='outside')
            st.plotly_chart(fig_seller_ord, use_container_width=True)

            # Sales Person Wise Dispatched Qty (Horizontal Bar Graph)
            st.subheader("🚚 Sales Person Wise Dispatched Quantity")
            seller_disp = filtered_df.groupby('SELLER NAME')['DISP.QTY'].sum().reset_index()
            fig_seller_disp = px.bar(seller_disp, y='SELLER NAME', x='DISP.QTY', text='DISP.QTY', orientation='h', title="Dispatched Qty by Sales Person")
            fig_seller_disp.update_traces(texttemplate='%{text:.2f}', textposition='outside')
            st.plotly_chart(fig_seller_disp, use_container_width=True)

            st.markdown("---")
            
            # Table: SalesExecutive, Item, OrderedQty, DispatchedQty, PendingQty, CancelQty, SCQty
            st.subheader("📋 Sales Executive & Item-wise Summary Table")
            if 'SELLER NAME' in filtered_df.columns and 'ITEM' in filtered_df.columns:
                exec_item_tbl = filtered_df.groupby(['SELLER NAME', 'ITEM']).agg({
                    'PO QTY (MT)': 'sum',
                    'DISP.QTY': 'sum'
                }).reset_index()
                
                exec_item_tbl['PendingQty'] = filtered_df[filtered_df['DERIVED_STATUS'] == 'PENDING'].groupby(['SELLER NAME', 'ITEM'])['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
                exec_item_tbl['CancelQty'] = filtered_df[filtered_df['DERIVED_STATUS'] == 'CANCEL'].groupby(['SELLER NAME', 'ITEM'])['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
                exec_item_tbl['SCQty'] = filtered_df[filtered_df['DERIVED_STATUS'] == 'SC'].groupby(['SELLER NAME', 'ITEM'])['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
                
                exec_item_tbl.rename(columns={'PO QTY (MT)': 'OrderedQty', 'DISP.QTY': 'DispatchedQty'}, inplace=True)
                st.dataframe(exec_item_tbl, use_container_width=True)

            st.markdown("---")
            
            # Top Parties Chart (Ordered, Pending, Cancel, Dispatch, SC)
            st.subheader("🏆 Top Parties Performance Breakdown")
            party_tbl = filtered_df.groupby('PARTY NAME').agg({
                'PO QTY (MT)': 'sum',
                'DISP.QTY': 'sum'
            }).reset_index()
            party_tbl['Pending'] = filtered_df[filtered_df['DERIVED_STATUS'] == 'PENDING'].groupby('PARTY NAME')['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
            party_tbl['Cancel'] = filtered_df[filtered_df['DERIVED_STATUS'] == 'CANCEL'].groupby('PARTY NAME')['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
            party_tbl['SC'] = filtered_df[filtered_df['DERIVED_STATUS'] == 'SC'].groupby('PARTY NAME')['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
            party_tbl.rename(columns={'PO QTY (MT)': 'Ordered', 'DISP.QTY': 'Dispatch'}, inplace=True)
            
            top_parties = party_tbl.sort_values(by='Ordered', ascending=False).head(15)
            fig_top_parties = px.bar(top_parties, x='PARTY NAME', y=['Ordered', 'Dispatch', 'Pending', 'Cancel', 'SC'], barmode='group', title="Top Parties Breakdown")
            fig_top_parties.update_traces(textposition='outside')
            st.plotly_chart(fig_top_parties, use_container_width=True)

            st.markdown("---")
            
            # Item Wise Analysis with multi-select filters
            st.subheader("📦 Item-wise Analysis & Specifications")
            all_itms = sorted(filtered_df['ITEM'].dropna().unique().tolist()) if 'ITEM' in filtered_df.columns else []
            all_thick = sorted(filtered_df['THICKNESS'].dropna().unique().tolist()) if 'THICKNESS' in filtered_df.columns else []
            all_widths = sorted(filtered_df['WIDTH_PARSED'].dropna().unique().tolist()) if 'WIDTH_PARSED' in filtered_df.columns else []
            
            sel_items = st.multiselect("Filter Item Name(s)", all_itms)
            sel_thick = st.multiselect("Filter Thickness", all_thick)
            sel_width = st.multiselect("Filter Width (e.g. 1250)", all_widths)
            
            item_filtered_df = filtered_df.copy()
            if sel_items:
                item_filtered_df = item_filtered_df[item_filtered_df['ITEM'].isin(sel_items)]
            if sel_thick:
                item_filtered_df = item_filtered_df[item_filtered_df['THICKNESS'].isin(sel_thick)]
            if sel_width:
                item_filtered_df = item_filtered_df[item_filtered_df['WIDTH_PARSED'].isin(sel_width)]
                
            item_spec_tbl = item_filtered_df.groupby(['ITEM', 'THICKNESS', 'WIDTH_PARSED']).agg({
                'PO QTY (MT)': 'sum',
                'DISP.QTY': 'sum'
            }).reset_index()
            
            item_spec_tbl['Cancel'] = item_filtered_df[item_filtered_df['DERIVED_STATUS'] == 'CANCEL'].groupby(['ITEM', 'THICKNESS', 'WIDTH_PARSED'])['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
            item_spec_tbl['Pending'] = item_filtered_df[item_filtered_df['DERIVED_STATUS'] == 'PENDING'].groupby(['ITEM', 'THICKNESS', 'WIDTH_PARSED'])['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
            item_spec_tbl['SC'] = item_filtered_df[item_filtered_df['DERIVED_STATUS'] == 'SC'].groupby(['ITEM', 'THICKNESS', 'WIDTH_PARSED'])['PO QTY (MT)'].sum().reset_index(drop=True).fillna(0)
            
            item_spec_tbl.rename(columns={'ITEM': 'Itemname', 'THICKNESS': 'Thickness', 'WIDTH_PARSED': 'Width', 'PO QTY (MT)': 'Ordered', 'DISP.QTY': 'Dispatch'}, inplace=True)
            st.dataframe(item_spec_tbl, use_container_width=True)

        # ---------------------------------------------------------
        # SECTION 2: DATE / MONTH / YEAR COMPARISON
        # ---------------------------------------------------------
        elif section == "2) Date/Month/Year Comparison":
            st.header("📅 Date, Month & Year Comparison")
            
            comp_type = st.radio("Select Comparison Mode", ["Date-Wise", "Month-Wise", "Year-Wise"], horizontal=True)
            combined_df = pd.concat(data_dict.values(), ignore_index=True)
            
            if comp_type == "Date-Wise" and 'PO DATE' in combined_df.columns:
                valid_d_df = combined_df.dropna(subset=['PO DATE']).copy()
                valid_d_df['DATE_ONLY'] = valid_d_df['PO DATE'].dt.date
                dates_list = sorted(valid_d_df['DATE_ONLY'].unique())
                
                if len(dates_list) >= 2:
                    dt1 = st.selectbox("First Date", dates_list, index=0)
                    dt2 = st.selectbox("Second Date", dates_list, index=min(1, len(dates_list)-1))
                    
                    df1 = valid_d_df[valid_d_df['DATE_ONLY'] == dt1]
                    df2 = valid_d_df[valid_d_df['DATE_ONLY'] == dt2]
                    
                    def get_metrics(dframe):
                        cust = dframe['PARTY NAME'].nunique()
                        so = dframe['PO QTY (MT)'].sum()
                        amt = dframe['TOTAL_AMOUNT'].sum()
                        disp = dframe['DISP.QTY'].sum()
                        cancel = dframe[dframe['DERIVED_STATUS'] == 'CANCEL']['PO QTY (MT)'].sum()
                        pend = dframe[dframe['DERIVED_STATUS'] == 'PENDING']['PO QTY (MT)'].sum()
                        sc = dframe[dframe['DERIVED_STATUS'] == 'SC']['PO QTY (MT)'].sum()
                        new_p = dframe[dframe['REMARK_CLEAN'].str.contains('NEW', na=False)]['PARTY NAME'].nunique()
                        return [cust, so, amt, disp, cancel, pend, sc, new_p]
                        
                    m1 = get_metrics(df1)
                    m2 = get_metrics(df2)
                    
                    summary_df = pd.DataFrame({
                        'KPI Metric': ['Unique Customers', 'Sum SO Qty (MT)', 'Total Amount (₹)', 'Dispatch Qty (MT)', 'Cancel Qty (MT)', 'Active Pending (MT)', 'SC Qty (MT)', 'New Parties'],
                        str(dt1): m1,
                        str(dt2): m2,
                        'Variance': [m2[i] - m1[i] for i in range(8)]
                    })
                    st.dataframe(summary_df, use_container_width=True)
                    
                    fig_comp = px.bar(summary_df, x='KPI Metric', y=[str(dt1), str(dt2)], barmode='group', title=f"Comparison: {dt1} vs {dt2}")
                    fig_comp.update_traces(textposition='outside')
                    st.plotly_chart(fig_comp, use_container_width=True)
                else:
                    st.warning("Insufficient dates found for comparison.")
                    
            elif comp_type == "Month-Wise":
                m_list = list(data_dict.keys())
                if len(m_list) >= 2:
                    mo1 = st.selectbox("First Month / Sheet", m_list, index=0)
                    mo2 = st.selectbox("Second Month / Sheet", m_list, index=min(1, len(m_list)-1))
                    
                    df1 = data_dict[mo1]
                    df2 = data_dict[mo2]
                    
                    def get_metrics(dframe):
                        cust = dframe['PARTY NAME'].nunique()
                        so = dframe['PO QTY (MT)'].sum()
                        amt = dframe['TOTAL_AMOUNT'].sum()
                        disp = dframe['DISP.QTY'].sum()
                        cancel = dframe[dframe['DERIVED_STATUS'] == 'CANCEL']['PO QTY (MT)'].sum()
                        pend = dframe[dframe['DERIVED_STATUS'] == 'PENDING']['PO QTY (MT)'].sum()
                        sc = dframe[dframe['DERIVED_STATUS'] == 'SC']['PO QTY (MT)'].sum()
                        new_p = dframe[dframe['REMARK_CLEAN'].str.contains('NEW', na=False)]['PARTY NAME'].nunique()
                        return [cust, so, amt, disp, cancel, pend, sc, new_p]
                        
                    m1 = get_metrics(df1)
                    m2 = get_metrics(df2)
                    
                    summary_df = pd.DataFrame({
                        'KPI Metric': ['Unique Customers', 'Sum SO Qty (MT)', 'Total Amount (₹)', 'Dispatch Qty (MT)', 'Cancel Qty (MT)', 'Active Pending (MT)', 'SC Qty (MT)', 'New Parties'],
                        mo1: m1,
                        mo2: m2,
                        'Variance': [m2[i] - m1[i] for i in range(8)]
                    })
                    st.dataframe(summary_df, use_container_width=True)
                    
                    fig_mcomp = px.bar(summary_df, x='KPI Metric', y=[mo1, mo2], barmode='group', title=f"Comparison: {mo1} vs {mo2}")
                    fig_mcomp.update_traces(textposition='outside')
                    st.plotly_chart(fig_mcomp, use_container_width=True)
                else:
                    st.warning("Need at least two sheets for month-wise comparison.")
                    
            else:
                st.info("Year-wise comparative aggregation across loaded sheets:")
                if 'PO DATE' in combined_df.columns:
                    combined_df['YEAR'] = combined_df['PO DATE'].dt.year.fillna(2026).astype(int)
                    yr_summary = combined_df.groupby('YEAR').agg({
                        'PO QTY (MT)': 'sum',
                        'DISP.QTY': 'sum',
                        'TOTAL_AMOUNT': 'sum'
                    }).reset_index()
                    st.dataframe(yr_summary, use_container_width=True)
                    fig_yr = px.bar(yr_summary, x='YEAR', y='PO QTY (MT)', text='PO QTY (MT)', title="Year-Wise Ordered Quantity")
                    st.plotly_chart(fig_yr, use_container_width=True)

        # ---------------------------------------------------------
        # SECTION 3: DISPATCH PENDING
        # ---------------------------------------------------------
        elif section == "3) Dispatch Pending":
            st.header("⏳ Dispatch Pending Report & KPIs")
            
            pend_sheet_candidates = [s for s in sheet_names if 'pending' in s.lower() and 'dispatch' in s.lower()]
            target_pend_sheet = pend_sheet_candidates[0] if pend_sheet_candidates else sheet_names[0]
            
            pend_df = data_dict[target_pend_sheet]
            active_pend_only = pend_df[pend_df['DERIVED_STATUS'] == 'PENDING']
            
            tot_pend_count = len(active_pend_only)
            tot_pend_qty = active_pend_only['PO QTY (MT)'].sum() if 'PO QTY (MT)' in active_pend_only.columns else 0
            
            col1, col2 = st.columns(2)
            col1.metric("Pending Orders Count", f"{tot_pend_count:,}")
            col2.metric("Total Active Pending Qty (MT)", f"{tot_pend_qty:,.2f}")
            
            st.markdown("---")
            st.subheader(f"Detailed Pending Report ({target_pend_sheet})")
            disp_cols = [c for c in ['PO NO', 'DO NO', 'PO_DATE_STR', 'PARTY NAME', 'SELLER NAME', 'ITEM', 'THICKNESS', 'SIZE', 'PO QTY (MT)', 'DISP.QTY', 'PENDING', 'REMARK'] if c in active_pend_only.columns]
            st.dataframe(active_pend_only[disp_cols], use_container_width=True)

else:
    st.info("👈 Please upload your Excel file using the sidebar to begin.")
