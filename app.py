import streamlit as st
import pandas as pd
import datetime

# Page Layout & Config
st.set_page_config(
    page_title="Executive Dispatch & Logistics Report",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Management Dispatch & Logistics Control System")
st.caption("Jindal Saw Limited - Daily Inward, Outward & Inventory Balance Tracker")

# 1. Load Raw Data
@st.cache_data
def load_raw_data():
    file_path = "Inward_Outward_FIXED_and_FAST.xlsx"
    
    # Read raw data sheet
    df_raw = pd.read_excel(file_path, sheet_name="Inward-Outward Data")
    
    # Ensure Date Column is proper datetime
    df_raw['Pstng Date'] = pd.to_datetime(df_raw['Pstng Date'], errors='coerce').dt.date
    
    # Load order qty sheet for reference if needed
    try:
        so_data = pd.read_excel(file_path, sheet_name="oredr qty so line item wise")
    except:
        so_data = pd.DataFrame()
        
    return df_raw, so_data

try:
    df_raw, so_data = load_raw_data()
    
    # Get available date range
    valid_dates = df_raw['Pstng Date'].dropna()
    min_date = valid_dates.min() if not valid_dates.empty else datetime.date(2025, 1, 1)
    max_date = valid_dates.max() if not valid_dates.empty else datetime.date.today()

    # 2. Sidebar Date Selector & Controls
    st.sidebar.header("⚙️ Report Control Panel")
    
    selected_date = st.sidebar.date_input(
        "📅 Select Report Cut-off Date:",
        value=max_date,
        min_value=min_date,
        max_value=max_date,
        help="Report iss selected date tak ka accurate status aur movement dikhayega."
    )
    
    hide_zero_bal = st.sidebar.checkbox("Hide Zero Balance & Inactive Rows (Print Preview View)", value=True)

    # 3. Dynamic Calculation Engine
    # Filter data up to selected report date
    df_upto_date = df_raw[df_raw['Pstng Date'] <= selected_date].copy()
    
    # Identify Today's specific transactions
    df_today = df_raw[df_raw['Pstng Date'] == selected_date].copy()

    # Grouping key for grouping (Plant, Customer, SO, Item, Valuation Class, OD, Thickness)
    group_cols = ['Plnt', 'Customer Name', 'TO Sales Ord.', 'TO SO item', 'Valuation Class', 'OD', 'Thickness (mm)']
    
    # Clean text columns for aggregation
    for col in group_cols:
        if col in df_upto_date.columns:
            df_upto_date[col] = df_upto_date[col].fillna("N/A")

    # Aggregate Total Inward & Outward Till Selected Date
    # Note: Weight calculation uses Coated Pipe Weight if available, else Bare Pipe Weight
    df_upto_date['Effective_WT'] = df_upto_date['Coated Pipe Weight(TO)'].fillna(0)
    df_upto_date.loc[df_upto_date['Effective_WT'] == 0, 'Effective_WT'] = df_upto_date['Bare Pipe Weight(TO)'].fillna(0)
    
    # Cumulative Movement
    inward_cum = df_upto_date[df_upto_date['Movement Direction'] == 'Inward'].groupby(group_cols)['Effective_WT'].sum().reset_index(name='Total Inward Till Date')
    outward_cum = df_upto_date[df_upto_date['Movement Direction'] == 'Outward'].groupby(group_cols)['Effective_WT'].sum().reset_index(name='Total Outward Till Date')
    
    # Today Movement
    df_today['Effective_WT'] = df_today['Coated Pipe Weight(TO)'].fillna(0)
    df_today.loc[df_today['Effective_WT'] == 0, 'Effective_WT'] = df_today['Bare Pipe Weight(TO)'].fillna(0)
    
    inward_today = df_today[df_today['Movement Direction'] == 'Inward'].groupby(group_cols)['Effective_WT'].sum().reset_index(name='Today IN')
    outward_today = df_today[df_today['Movement Direction'] == 'Outward'].groupby(group_cols)['Effective_WT'].sum().reset_index(name='Today OUT')

    # Merge into Master Summary Table
    master_df = df_upto_date[group_cols].drop_duplicates().copy()
    master_df = master_df.merge(inward_today, on=group_cols, how='left').fillna({'Today IN': 0})
    master_df = master_df.merge(inward_cum, on=group_cols, how='left').fillna({'Total Inward Till Date': 0})
    master_df = master_df.merge(outward_today, on=group_cols, how='left').fillna({'Today OUT': 0})
    master_df = master_df.merge(outward_cum, on=group_cols, how='left').fillna({'Total Outward Till Date': 0})

    # Calculate Net Stock Balance
    master_df['Balance (MT)'] = master_df['Total Inward Till Date'] - master_df['Total Outward Till Date']

    # Filter Active Movements / Positive Balance if checkbox selected
    if hide_zero_bal:
        master_df = master_df[
            (master_df['Today IN'] > 0) | 
            (master_df['Today OUT'] > 0) | 
            (master_df['Balance (MT)'] > 0)
        ]

    # 4. Management Executive Summary KPI Cards
    st.subheader(f"📈 Executive Overview - As of {selected_date.strftime('%d-%b-%Y')}")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Today IN (MT)", f"{master_df['Today IN'].sum():,.2f}")
    with col2:
        st.metric("Today OUT (MT)", f"{master_df['Today OUT'].sum():,.2f}")
    with col3:
        st.metric("Cumulative Inward (MT)", f"{master_df['Total Inward Till Date'].sum():,.2f}")
    with col4:
        st.metric("Current Balance Stock (MT)", f"{master_df['Balance (MT)'].sum():,.2f}")

    st.markdown("---")

    # 5. Main Tabbed Views
    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 Daily Dispatch Summary", 
        "🏭 Plant-wise Stock", 
        "📥 Daily Movement Log",
        "📥 Export Reports"
    ])

    with tab1:
        st.subheader("Daily Dispatch & Inventory Balance Table")
        st.dataframe(master_df, use_container_width=True, hide_index=True)

    with tab2:
        st.subheader("Plant-wise Summary Breakdown")
        plant_summary = master_df.groupby('Plnt').agg({
            'Today IN': 'sum',
            'Today OUT': 'sum',
            'Total Inward Till Date': 'sum',
            'Total Outward Till Date': 'sum',
            'Balance (MT)': 'sum'
        }).reset_index()
        st.dataframe(plant_summary, use_container_width=True, hide_index=True)

    with tab3:
        st.subheader(f"Transactions Recorded On {selected_date.strftime('%d-%b-%Y')}")
        st.dataframe(df_today[['Batch', 'Customer Name', 'Plnt', 'Movement Direction', 'Effective_WT', 'MvT', 'User Name']], use_container_width=True)

    with tab4:
        st.subheader("📥 Download Clean Report for Management")
        csv_data = master_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Download Daily Dispatch Report (CSV)",
            data=csv_data,
            file_name=f"Dispatch_Report_As_Of_{selected_date}.csv",
            mime="text/csv"
        )

except Exception as e:
    st.error(f"Error executing dynamic calculation: {e}")
    st.info("Ensure that 'Inward_Outward_FIXED_and_FAST.xlsx' is committed to the GitHub repository.")
