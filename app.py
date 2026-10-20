import streamlit as st
import pandas as pd
import datetime

# Page Configuration
st.set_page_config(
    page_title="Jindal Saw Dispatch & Inventory Management",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Jindal Saw - Dynamic Dispatch & Inventory Report")
st.caption("Management Control System (DDR, Plant/Product Summaries & Registers)")

# Load Raw Data
@st.cache_data
def load_data():
    file_path = "Inward_Outward_FIXED_and_FAST.xlsx"
    df_raw = pd.read_excel(file_path, sheet_name="Inward-Outward Data")
    df_raw['Pstng Date'] = pd.to_datetime(df_raw['Pstng Date'], errors='coerce').dt.date
    return df_raw

try:
    df_raw = load_data()
    
    # Date range setup
    valid_dates = df_raw['Pstng Date'].dropna()
    min_date = valid_dates.min() if not valid_dates.empty else datetime.date(2025, 1, 1)
    max_date = valid_dates.max() if not valid_dates.empty else datetime.date.today()

    # Sidebar Controls
    st.sidebar.header("⚙️ Report Control Panel")
    selected_date = st.sidebar.date_input(
        "📅 Select Report Cut-off Date:",
        value=max_date,
        min_value=min_date,
        max_value=max_date,
        help="Yeh date select karne par wahan tak ka live accurate report generate hoga."
    )
    
    report_view = st.sidebar.radio(
        "Select Report Format:",
        [
            "Daily Dispatch Report (Print Preview - Active & Positive Only)",
            "Plant-wise Summary",
            "Product-wise Summary",
            "Inward Register",
            "Outward Register"
        ]
    )

    # Data cleaning for grouping
    df_raw['Thickness (mm)'] = pd.to_numeric(df_raw['Thickness (mm)'], errors='coerce').fillna(0)
    df_raw['OD'] = pd.to_numeric(df_raw['OD'], errors='coerce').fillna(0)
    df_raw['TO Sales Ord.'] = pd.to_numeric(df_raw['TO Sales Ord.'], errors='coerce').fillna(0).astype(int)
    df_raw['TO SO item'] = pd.to_numeric(df_raw['TO SO item'], errors='coerce').fillna(0).astype(int)

    # Filter data up to selected date
    df_upto_date = df_raw[df_raw['Pstng Date'] <= selected_date].copy()
    df_today = df_raw[df_raw['Pstng Date'] == selected_date].copy()

    group_cols = ['Plnt', 'Customer Name', 'TO Sales Ord.', 'TO SO item', 'Valuation Class', 'OD', 'Thickness (mm)']
    for col in group_cols:
        df_upto_date[col] = df_upto_date[col].astype(str)
        df_today[col] = df_today[col].astype(str)

    # Helper function to generate DDR 4-row format
    def create_ddr_format(base_df):
        metrics = [
            ('No of PCS', 'Sum of No of PCS'),
            ('Total Length (Mtr )', 'Sum of Total Length (Mtr )'),
            ('Bare Pipe Weight(TO)', 'Sum of Bare Pipe Weight(TO)'),
            ('Coated Pipe Weight(TO)', 'Sum of Coated Pipe Weight(TO)')
        ]
        frames = []
        for col_name, label in metrics:
            if col_name in base_df.columns:
                temp = base_df.groupby(group_cols, as_index=False)[col_name].sum()
                temp['Values'] = label
                temp['Metric_Val'] = temp[col_name]
                frames.append(temp[['Plnt', 'Customer Name', 'TO Sales Ord.', 'TO SO item', 'Valuation Class', 'OD', 'Thickness (mm)', 'Values', 'Metric_Val']])
        
        if not frames:
            return pd.DataFrame()
            
        combined = pd.concat(frames, ignore_index=True)
        return combined

    # 1. Daily Dispatch Report View (DDR Print Preview Format)
    if "Daily Dispatch Report" in report_view:
        st.subheader(f"📋 Daily Dispatch Report — As of {selected_date.strftime('%d-%b-%Y')}")
        st.info("Note: Zero balance and inactive rows are automatically hidden as per management print preview standard.")

        inward_today_df = create_ddr_format(df_today[df_today['Movement Direction'] == 'Inward'])
        inward_today_df.rename(columns={'Metric_Val': 'Today IN'}, inplace=True)

        inward_cum_df = create_ddr_format(df_upto_date[df_upto_date['Movement Direction'] == 'Inward'])
        inward_cum_df.rename(columns={'Metric_Val': 'Total Inward Till Date'}, inplace=True)

        outward_today_df = create_ddr_format(df_today[df_today['Movement Direction'] == 'Outward'])
        outward_today_df.rename(columns={'Metric_Val': 'Today OUT'}, inplace=True)

        outward_cum_df = create_ddr_format(df_upto_date[df_upto_date['Movement Direction'] == 'Outward'])
        outward_cum_df.rename(columns={'Metric_Val': 'Total Outward Till Date'}, inplace=True)

        master_keys = group_cols + ['Values']
        merged = inward_cum_df.merge(outward_cum_df, on=master_keys, how='outer')
        merged = merged.merge(inward_today_df, on=master_keys, how='left')
        merged = merged.merge(outward_today_df, on=master_keys, how='left')
        merged.fillna(0, inplace=True)
        
        merged['Balance'] = merged['Total Inward Till Date'] - merged['Total Outward Till Date']

        # Filter out rows where Balance <= 0 and Today movements are 0
        # To handle multi-metric rows properly per group, check group balance
        balance_summary = merged[merged['Values'] == 'Sum of No of PCS'][['Plnt', 'Customer Name', 'TO Sales Ord.', 'TO SO item', 'Balance']]
        active_groups = balance_summary[balance_summary['Balance'] > 0]
        
        # Also include groups with active today movement
        today_summary = merged[(merged['Today IN'] != 0) | (merged['Today OUT'] != 0)][group_cols].drop_duplicates()
        
        valid_keys = pd.concat([active_groups[group_cols[:-1] if 'group_cols' in locals() else group_cols], today_summary]).drop_duplicates()
        
        filtered_merged = merged.merge(valid_keys, on=['Plnt', 'Customer Name', 'TO Sales Ord.', 'TO SO item', 'Valuation Class', 'OD', 'Thickness (mm)'])
        filtered_merged.sort_values(by=['Plnt', 'Customer Name', 'TO Sales Ord.', 'TO SO item', 'Values'], inplace=True)

        st.dataframe(
            filtered_merged[['Plnt', 'Customer Name', 'TO Sales Ord.', 'TO SO item', 'Valuation Class', 'OD', 'Thickness (mm)', 'Values', 'Today IN', 'Total Inward Till Date', 'Today OUT', 'Total Outward Till Date', 'Balance']],
            use_container_width=True,
            hide_index=True
        )

        csv_data = filtered_merged.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download Daily Dispatch Print Preview (CSV)", data=csv_data, file_name=f"DDR_Active_Balance_{selected_date}.csv", mime="text/csv")

    # 2. Plant-wise Summary View
    elif report_view == "Plant-wise Summary":
        st.subheader(f"🏭 Plant-wise Summary — As of {selected_date.strftime('%d-%b-%Y')}")
        plant_df = df_upto_date.groupby('Plnt', as_index=False).agg(
            No_of_Pcs=('No of PCS', 'sum'),
            Length_Mtr=('Total Length (Mtr )', 'sum'),
            Coated_Weight_Mt=('Coated Pipe Weight(TO)', 'sum'),
            Bare_Weight_Mt=('Bare Pipe Weight(TO)', 'sum')
        )
        plant_df['Net Weight (Mt)'] = plant_df['Coated_Weight_Mt'].where(plant_df['Coated_Weight_Mt'] > 0, plant_df['Bare_Weight_Mt'])
        st.dataframe(plant_df, use_container_width=True, hide_index=True)

    # 3. Product-wise Summary View
    elif report_view == "Product-wise Summary":
        st.subheader(f"📦 Product-wise Summary — As of {selected_date.strftime('%d-%b-%Y')}")
        prod_col = 'Material' if 'Material' in df_upto_date.columns else 'Valuation Class'
        prod_df = df_upto_date.groupby(prod_col, as_index=False).agg(
            No_of_Pcs=('No of PCS', 'sum'),
            Length_Mtr=('Total Length (Mtr )', 'sum'),
            Coated_Weight_Mt=('Coated Pipe Weight(TO)', 'sum'),
            Bare_Weight_Mt=('Bare Pipe Weight(TO)', 'sum')
        )
        prod_df['Net Weight (Mt)'] = prod_df['Coated_Weight_Mt'].where(prod_df['Coated_Weight_Mt'] > 0, prod_df['Bare_Weight_Mt'])
        st.rename(columns={prod_col: 'Product Name'}, inplace=True)
        st.dataframe(prod_df, use_container_width=True, hide_index=True)

    # 4. Inward Register View
    elif report_view == "Inward Register":
        st.subheader(f"📥 Inward Register (Movements up to {selected_date.strftime('%d-%b-%Y')})")
        inward_reg = df_upto_date[df_upto_date['Movement Direction'] == 'Inward']
        st.dataframe(inward_reg[['Pstng Date', 'Batch', 'Customer Name', 'Plnt', 'No of PCS', 'Total Length (Mtr )', 'Bare Pipe Weight(TO)', 'MvT', 'User Name']], use_container_width=True, hide_index=True)

    # 5. Outward Register View
    elif report_view == "Outward Register":
        st.subheader(f"📤 Outward Register (Shipments up to {selected_date.strftime('%d-%b-%Y')})")
        outward_reg = df_upto_date[df_upto_date['Movement Direction'] == 'Outward']
        st.dataframe(outward_reg[['Pstng Date', 'Batch', 'Customer Name', 'Plnt', 'No of PCS', 'Total Length (Mtr )', 'Bare Pipe Weight(TO)', 'MvT', 'User Name']], use_container_width=True, hide_index=True)

except Exception as e:
    st.error(f"Error executing report: {e}")
    st.info("Kripya verify karein ki Excel file GitHub repository me properly uploaded hai.")
