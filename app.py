import streamlit as st
import pandas as pd
import datetime

# Page Layout & Config
st.set_page_config(
    page_title="Daily Dispatch & Logistics Report",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Daily Dispatch & Inventory Report (DDR Format)")
st.caption("Jindal Saw Limited - Management Dispatch Control System")

# Load Raw Data
@st.cache_data
def load_raw_data():
    file_path = "Inward_Outward_FIXED_and_FAST.xlsx"
    df_raw = pd.read_excel(file_path, sheet_name="Inward-Outward Data")
    df_raw['Pstng Date'] = pd.to_datetime(df_raw['Pstng Date'], errors='coerce').dt.date
    return df_raw

try:
    df_raw = load_raw_data()
    
    # Get available date range
    valid_dates = df_raw['Pstng Date'].dropna()
    min_date = valid_dates.min() if not valid_dates.empty else datetime.date(2025, 1, 1)
    max_date = valid_dates.max() if not valid_dates.empty else datetime.date.today()

    # Sidebar Date Selector & Controls
    st.sidebar.header("⚙️ Report Control Panel")
    
    selected_date = st.sidebar.date_input(
        "📅 Select Report Date:",
        value=max_date,
        min_value=min_date,
        max_value=max_date,
        help="Iss date tak ka accurate cumulative aur daily movement calculate hoga."
    )
    
    view_mode = st.sidebar.radio(
        "Select Report View:",
        ["Daily Dispatch Report (Full)", "Print Preview (Active & Positive Balance Only)"]
    )

    # Clean data types to avoid merge errors
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

    # Helper function to melt into DDR 4-row format
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
        # Sort to keep the 4 metric rows together per group
        combined['TO SO item'] = pd.to_numeric(combined['TO SO item'], errors='coerce')
        combined['OD'] = pd.to_numeric(combined['OD'], errors='coerce')
        combined['Thickness (mm)'] = pd.to_numeric(combined['Thickness (mm'], errors='coerce') if 'Thickness (mm' in combined.columns else combined['Thickness (mm)']
        
        return combined

    # Calculate Inward & Outward for Cumulative and Today
    inward_today_df = create_ddr_format(df_today[df_today['Movement Direction'] == 'Inward'])
    inward_today_df.rename(columns={'Metric_Val': 'Today IN'}, inplace=True)

    inward_cum_df = create_ddr_format(df_upto_date[df_upto_date['Movement Direction'] == 'Inward'])
    inward_cum_df.rename(columns={'Metric_Val': 'Total Inward Till Date'}, inplace=True)

    outward_today_df = create_ddr_format(df_today[df_today['Movement Direction'] == 'Outward'])
    outward_today_df.rename(columns={'Metric_Val': 'Today OUT'}, inplace=True)

    outward_cum_df = create_ddr_format(df_upto_date[df_upto_date['Movement Direction'] == 'Outward'])
    outward_cum_df.rename(columns={'Metric_Val': 'Total Outward Till Date'}, inplace=True)

    # Base skeleton from cumulative inward/outward
    master_keys = group_cols + ['Values']
    
    # Merge all datasets safely using inner/outer joins on keys + Values
    merged = inward_cum_df.merge(outward_cum_df, on=master_keys, how='outer')
    merged = merged.merge(inward_today_df, on=master_keys, how='left')
    merged = merged.merge(outward_today_df, on=master_keys, how='left')
    
    merged.fillna(0, inplace=True)
    
    # Calculate Balance
    merged['Balance'] = merged['Total Inward Till Date'] - merged['Total Outward Till Date']

    # Sort properly to match Excel DDR view
    merged.sort_values(by=['Plnt', 'Customer Name', 'TO Sales Ord.', 'TO SO item', 'Values'], inplace=True)

    if view_mode == "Print Preview (Active & Positive Balance Only)":
        # Filter groups where Today IN != 0 or Today OUT != 0 or Balance > 0
        active_groups = merged[
            (merged['Today IN'] != 0) | 
            (merged['Today OUT'] != 0) | 
            (merged['Balance'] > 0)
        ]
        display_df = active_groups
    else:
        display_df = merged

    # Clean display formatting
    st.subheader(f"📋 Daily Dispatch Report — As of {selected_date.strftime('%d-%b-%Y')}")
    st.write(f"Total Rows Displayed: {len(display_df)}")
    
    st.dataframe(
        display_df[['Plnt', 'Customer Name', 'TO Sales Ord.', 'TO SO item', 'Valuation Class', 'OD', 'Thickness (mm)', 'Values', 'Today IN', 'Total Inward Till Date', 'Today OUT', 'Total Outward Till Date', 'Balance']],
        use_container_width=True,
        hide_index=True
    )

    # Export Button for Management
    csv_bytes = display_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download DDR Report (CSV)",
        data=csv_bytes,
        file_name=f"DDR_Report_{selected_date}.csv",
        mime="text/csv"
    )

except Exception as e:
    st.error(f"Error in execution: {e}")
    st.info("Kripya check karein ki 'Inward_Outward_FIXED_and_FAST.xlsx' file GitHub repo me properly uploaded hai.")
