import streamlit as st
import pandas as pd
import datetime

# Page Configuration
st.set_page_config(
    page_title="Jindal Saw - Dispatch & Logistics Tracking System",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Management Dispatch & Logistics Control System")
st.caption("Jindal Saw Limited - Live Dynamic Inward, Outward & Inventory Balance System")

# Load Raw Base Data
@st.cache_data
def load_raw_data():
    file_path = "Inward_Outward_FIXED_and_FAST.xlsx"
    df_raw = pd.read_excel(file_path, sheet_name="Inward-Outward Data")
    df_raw['Pstng Date'] = pd.to_datetime(df_raw['Pstng Date'], errors='coerce').dt.date
    return df_raw

try:
    df_raw = load_raw_data()
    
    # Date Setup
    valid_dates = df_raw['Pstng Date'].dropna()
    min_date = valid_dates.min() if not valid_dates.empty else datetime.date(2025, 1, 1)
    max_date = valid_dates.max() if not valid_dates.empty else datetime.date.today()

    # Sidebar Controls
    st.sidebar.header("⚙️ Report Control Panel")
    selected_date = st.sidebar.date_input(
        "📅 Select Cut-off Date:",
        value=max_date,
        min_value=min_date,
        max_value=max_date,
        help="Selected date tak ka exact cumulative status aur daily movement show karega."
    )
    
    report_type = st.sidebar.radio(
        "Select Excel Format View:",
        [
            "1. Daily Dispatch Report (Full)",
            "2. Print Preview (Active & Positive Balance Only)",
            "3. Plant-wise Summary",
            "4. Product-wise Summary",
            "5. Inward Register (Daily Movements)",
            "6. Outward Register (Daily Shipments)",
            "7. Balance Pipe Report"
        ]
    )

    # Clean numeric fields
    num_cols = ['No of PCS', 'Total Length (Mtr )', 'Bare Pipe Weight(TO)', 'Coated Pipe Weight(TO)']
    for nc in num_cols:
        if nc in df_raw.columns:
            df_raw[nc] = pd.to_numeric(df_raw[nc], errors='coerce').fillna(0)

    # Filter up to cut-off date
    df_upto = df_raw[df_raw['Pstng Date'] <= selected_date].copy()
    df_today = df_raw[df_raw['Pstng Date'] == selected_date].copy()

    group_cols = ['Plnt', 'Customer Name', 'TO Sales Ord.', 'TO SO item', 'Valuation Class', 'OD', 'Thickness (mm)']
    for c in group_cols:
        df_upto[c] = df_upto[c].fillna("").astype(str)
        df_today[c] = df_today[c].fillna("").astype(str)

    # Helper: Build 4-row DDR Pivot
    def build_ddr_pivot(base_df):
        metrics = [
            ('No of PCS', 'Sum of No of PCS'),
            ('Total Length (Mtr )', 'Sum of Total Length (Mtr )'),
            ('Bare Pipe Weight(TO)', 'Sum of Bare Pipe Weight(TO)'),
            ('Coated Pipe Weight(TO)', 'Sum of Coated Pipe Weight(TO)')
        ]
        list_dfs = []
        for col_name, label in metrics:
            if col_name in base_df.columns:
                t = base_df.groupby(group_cols, as_index=False)[col_name].sum()
                t['Values'] = label
                t['Val'] = t[col_name]
                list_dfs.append(t[group_cols + ['Values', 'Val']])
        
        if not list_dfs:
            return pd.DataFrame()
        return pd.concat(list_dfs, ignore_index=True)

    # Build Master DDR Sheet
    in_cum = build_ddr_pivot(df_upto[df_upto['Movement Direction'] == 'Inward']).rename(columns={'Val': 'Total Inward Till Date'})
    out_cum = build_ddr_pivot(df_upto[df_upto['Movement Direction'] == 'Outward']).rename(columns={'Val': 'Total Outward Till Date'})
    in_today = build_ddr_pivot(df_today[df_today['Movement Direction'] == 'Inward']).rename(columns={'Val': 'Today IN'})
    out_today = build_ddr_pivot(df_today[df_today['Movement Direction'] == 'Outward']).rename(columns={'Val': 'Today OUT'})

    m_keys = group_cols + ['Values']
    m_ddr = in_cum.merge(out_cum, on=m_keys, how='outer')
    m_ddr = m_ddr.merge(in_today, on=m_keys, how='left')
    m_ddr = m_ddr.merge(out_today, on=m_keys, how='left')
    m_ddr.fillna(0, inplace=True)

    # Balance calculation: Outward is negative in source, so Balance = Inward + Outward
    m_ddr['Balance'] = m_ddr['Total Inward Till Date'] + m_ddr['Total Outward Till Date']

    # -------------------------------------------------------------
    # VIEW 1 & 2: DAILY DISPATCH REPORT & PRINT PREVIEW
    # -------------------------------------------------------------
    if "Daily Dispatch Report" in report_type or "Print Preview" in report_type:
        st.subheader(f"📋 {report_type} — As of {selected_date.strftime('%d-%b-%Y')}")
        
        if "Print Preview" in report_type:
            pcs_m = m_ddr[m_ddr['Values'] == 'Sum of No of PCS']
            active_keys = pcs_m[(pcs_m['Balance'] > 0) | (pcs_m['Today IN'] != 0) | (pcs_m['Today OUT'] != 0)][group_cols]
            res_df = m_ddr.merge(active_keys, on=group_cols) if not active_keys.empty else pd.DataFrame(columns=m_ddr.columns)
        else:
            res_df = m_ddr.copy()

        res_df.sort_values(by=group_cols + ['Values'], inplace=True)

        disp_cols = group_cols + ['Values', 'Today IN', 'Total Inward Till Date', 'Today OUT', 'Total Outward Till Date', 'Balance']
        
        # Formatting zero values to "-" for clean Excel appearance
        formatted_df = res_df[disp_cols].copy()
        num_fields = ['Today IN', 'Total Inward Till Date', 'Today OUT', 'Total Outward Till Date', 'Balance']
        for col in num_fields:
            formatted_df[col] = formatted_df[col].apply(lambda x: "-" if abs(x) < 1e-4 else f"{x:,.2f}")

        st.dataframe(formatted_df, use_container_width=True, hide_index=True)

        csv_bytes = res_df[disp_cols].to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download Report (CSV)", data=csv_bytes, file_name=f"DDR_Report_{selected_date}.csv", mime="text/csv")

    # -------------------------------------------------------------
    # VIEW 3: PLANT-WISE SUMMARY
    # -------------------------------------------------------------
    elif "Plant-wise Summary" in report_type:
        st.subheader(f"🏭 Plant-wise Summary — As of {selected_date.strftime('%d-%b-%Y')}")
        
        plants = ['SGH1', 'SGH2', 'SGJ2', 'SGSS']
        rows = []
        for p in plants:
            p_df = df_upto[df_upto['Plnt'] == p]
            pcs = p_df['No of PCS'].sum()
            length = p_df['Total Length (Mtr )'].sum()
            coated = p_df['Coated Pipe Weight(TO)'].sum()
            bare = p_df['Bare Pipe Weight(TO)'].sum()
            net = coated if coated > 0 else bare
            rows.append({
                'Plant Name': p,
                'No. of Pcs': pcs,
                'Length (Mtr)': length,
                'Coated Weight (Mt)': coated,
                'Bare Weight (Mt)': bare,
                'Net Weight (Mt)': net
            })
            
        p_summary = pd.DataFrame(rows)
        
        # Grand Total Row
        gt_row = {
            'Plant Name': 'GRAND TOTAL',
            'No. of Pcs': p_summary['No. of Pcs'].sum(),
            'Length (Mtr)': p_summary['Length (Mtr)'].sum(),
            'Coated Weight (Mt)': p_summary['Coated Weight (Mt)'].sum(),
            'Bare Weight (Mt)': p_summary['Bare Weight (Mt)'].sum(),
            'Net Weight (Mt)': p_summary['Net Weight (Mt)'].sum()
        }
        p_summary = pd.concat([p_summary, pd.DataFrame([gt_row])], ignore_index=True)
        st.dataframe(p_summary, use_container_width=True, hide_index=True)

    # -------------------------------------------------------------
    # VIEW 4: PRODUCT-WISE SUMMARY
    # -------------------------------------------------------------
    elif "Product-wise Summary" in report_type:
        st.subheader(f"📦 Product-wise Summary — As of {selected_date.strftime('%d-%b-%Y')}")
        
        buckets = ['Carbon Steel', 'SEAM PIPE', 'SS WELDED', 'ss cutting plate', 'SS RESALE', 'Coating Jobwork', 'J_BARE_PIPE_MTR']
        
        rows = []
        for b in buckets:
            if b == 'J_BARE_PIPE_MTR':
                sub = df_upto[df_upto['Material'] == 'J_BARE_PIPE_MTR']
            elif b == 'Coating Jobwork':
                sub = df_upto[df_upto['Valuation Class'].str.contains('Jobwork', case=False, na=False)]
            elif b == 'SEAM PIPE':
                sub = df_upto[df_upto['Material'].str.contains('SEAM', case=False, na=False)]
            elif b == 'SS WELDED':
                sub = df_upto[df_upto['Material'].str.contains('WELD', case=False, na=False)]
            elif b == 'ss cutting plate':
                sub = df_upto[df_upto['Material'].str.contains('PLATE', case=False, na=False)]
            elif b == 'SS RESALE':
                sub = df_upto[df_upto['Material'].str.contains('RESALE', case=False, na=False)]
            else: # Carbon Steel as default / catch-all
                sub = df_upto[~df_upto['Valuation Class'].str.contains('Jobwork', case=False, na=False)]
                
            pcs = sub['No of PCS'].sum()
            length = sub['Total Length (Mtr )'].sum()
            coated = sub['Coated Pipe Weight(TO)'].sum()
            bare = sub['Bare Pipe Weight(TO)'].sum()
            net = coated if coated > 0 else bare
            rows.append({
                'Product Name': b,
                'No. of Pcs': pcs,
                'Length (Mtr)': length,
                'Coated Weight (Mt)': coated,
                'Bare Weight (Mt)': bare,
                'Net Weight (Mt)': net
            })
            
        prod_summary = pd.DataFrame(rows)
        
        # Grand Total
        gt_row = {
            'Product Name': 'GRAND TOTAL',
            'No. of Pcs': prod_summary['No. of Pcs'].sum(),
            'Length (Mtr)': prod_summary['Length (Mtr)'].sum(),
            'Coated Weight (Mt)': prod_summary['Coated Weight (Mt)'].sum(),
            'Bare Weight (Mt)': prod_summary['Bare Weight (Mt)'].sum(),
            'Net Weight (Mt)': prod_summary['Net Weight (Mt)'].sum()
        }
        prod_summary = pd.concat([prod_summary, pd.DataFrame([gt_row])], ignore_index=True)
        st.dataframe(prod_summary, use_container_width=True, hide_index=True)

    # -------------------------------------------------------------
    # VIEW 5 & 6: INWARD & OUTWARD REGISTERS
    # -------------------------------------------------------------
    elif "Inward Register" in report_type or "Outward Register" in report_type:
        direction = "Inward" if "Inward" in report_type else "Outward"
        st.subheader(f"{'📥 Inward' if direction=='Inward' else '📤 Outward'} Register — As of
