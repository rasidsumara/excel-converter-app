import streamlit as st
import pandas as pd
import datetime
import os
import io
import requests
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Page Configuration
st.set_page_config(
    page_title="Jindal Saw - Dynamic Logistics & Inventory Dashboard",
    page_icon="📊",
    layout="wide"
)

# -------------------------------------------------------------
# PASSWORD PROTECTION LOGIC
# -------------------------------------------------------------
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    st.title("🔒 Management Report - Secure Access")
    st.subheader("Kripya App me access karne ke liye Password daalein:")
    
    user_password = st.text_input("Password:", type="password", key="password_input")
    
    if st.button("Login"):
        if user_password == "Rashid@5972":
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("❌ Galat Password! Kripya sahi password daalein.")
            
    return False

if not check_password():
    st.stop()

if st.sidebar.button("🔒 Logout"):
    st.session_state["password_correct"] = False
    st.rerun()

# -------------------------------------------------------------
# GOOGLE APPS SCRIPT SYNC HELPER
# -------------------------------------------------------------
def sync_to_google_sheet_via_script(df, web_app_url):
    try:
        data_to_send = [df.columns.values.tolist()] + df.values.tolist()
        payload = json.dumps(data_to_send)
        
        headers = {'Content-Type': 'application/json'}
        response = requests.post(web_app_url, data=payload, headers=headers)
        
        res_json = response.json()
        if res_json.get("status") == "success":
            return True, "Synced Successfully!"
        else:
            return False, res_json.get("message", "Unknown error")
    except Exception as e:
        return False, str(e)

# -------------------------------------------------------------
# EXCEL GENERATION HELPER
# -------------------------------------------------------------
def generate_excel_report(df_to_excel, title_text, selected_date):
    wb = Workbook()
    ws = wb.active
    ws.title = "Dispatch Report"

    title_font = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
    title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='0B2545', end_color='0B2545', fill_type='solid')
    data_font = Font(name='Calibri', size=10)
    bold_data_font = Font(name='Calibri', size=10, bold=True)
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    ws.merge_cells('A1:M1')
    ws['A1'] = f"JINDAL SAW LIMITED — {title_text.upper()}"
    ws['A1'].font = title_font
    ws['A1'].fill = title_fill
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

    ws['A2'] = f"Report Date: {selected_date.strftime('%d-%b-%Y')}"
    ws['A2'].font = Font(name='Calibri', size=11, italic=True)

    headers = df_to_excel.columns.tolist()
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for row_idx, row_data in enumerate(df_to_excel.values.tolist(), 5):
        is_total_row = (row_idx - 5 >= len(df_to_excel) - 4)
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = bold_data_font if is_total_row else data_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='left' if col_idx <= 8 else 'right', vertical='center')

    for col in ws.columns:
        max_len = 0
        col_letter = None
        for cell in col:
            if not col_letter and hasattr(cell, 'column_letter'):
                col_letter = cell.column_letter
            if cell.row >= 4:
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = len(val_str)
        if col_letter:
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

# -------------------------------------------------------------
# PDF GENERATION HELPER
# -------------------------------------------------------------
def generate_pdf_report(df_to_pdf, title_text, selected_date):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=15, leftMargin=15, topMargin=15, bottomMargin=15)
    elements = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=12, leading=14, textColor=colors.HexColor('#0B2545'))
    
    elements.append(Paragraph(f"<b>JINDAL SAW LIMITED</b> - {title_text}", title_style))
    elements.append(Paragraph(f"Report Date: {selected_date.strftime('%d-%b-%Y')}", styles['Normal']))
    elements.append(Spacer(1, 8))
    
    cell_style = ParagraphStyle('TableCell', fontName='Helvetica', fontSize=6, leading=7, alignment=1)
    header_style = ParagraphStyle('TableHeader', fontName='Helvetica-Bold', fontSize=7, leading=8, textColor=colors.whitesmoke, alignment=1)
    
    table_data = []
    header_row = [Paragraph(str(col), header_style) for col in df_to_pdf.columns.tolist()]
    table_data.append(header_row)
    
    for row in df_to_pdf.values.tolist():
        clean_row = [Paragraph(str(cell) if cell is not None else "", cell_style) for cell in row]
        table_data.append(clean_row)
        
    pdf_table = Table(table_data, repeatRows=1)
    pdf_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0B2545')),
        ('BOTTOMPADDING', (0,0), (-1,0), 4),
        ('TOPPADDING', (0,0), (-1,0), 4),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BOTTOMPADDING', (0,1), (-1,-1), 2),
        ('TOPPADDING', (0,1), (-1,-1), 2),
    ]))
    
    elements.append(pdf_table)
    doc.build(elements)
    buffer.seek(0)
    return buffer

# -------------------------------------------------------------
# SMART AUTOMATIC FILE DETECTION & DATA LOADER
# -------------------------------------------------------------
@st.cache_data
def load_raw_data():
    possible_files = [
        "PROJECT 11 25.10 (version 2).xlsb.xlsx",
        "Inward_Outward_FIXED_and_FAST.xlsx",
        "data.xlsx"
    ]
    
    file_path = None
    for f in possible_files:
        if os.path.exists(f):
            file_path = f
            break
            
    if not file_path:
        for file in os.listdir('.'):
            if file.endswith('.xlsx') or file.endswith('.xls'):
                file_path = file
                break
                
    if not file_path:
        raise FileNotFoundError("GitHub repository me koi bhi Excel file nahi mili. Kripya file upload karein.")
        
    # Sheet name detection
    if "PROJECT 11" in file_path:
        sheet_name = "ZMM_MB51"
    else:
        sheet_name = "Inward-Outward Data"
        
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name)
    df_raw['Pstng Date'] = pd.to_datetime(df_raw['Pstng Date'], errors='coerce')
    return df_raw

# -------------------------------------------------------------
# MAIN APPLICATION CODE
# -------------------------------------------------------------
st.title("📊 Jindal Saw - Dynamic Logistics & Inventory Control Dashboard")
st.caption("Advanced SAP MB51 Analytics, Dispatch Reconciliation & Stock Control System")

try:
    df_raw = load_raw_data()
    
    valid_dates = df_raw['Pstng Date'].dropna()
    min_date = valid_dates.min().date() if not valid_dates.empty else datetime.date(2025, 1, 1)
    max_date = valid_dates.max().date() if not valid_dates.empty else datetime.date.today()

    # Sidebar Controls & Global Filters
    st.sidebar.header("⚙️ Dashboard Controls & Filters")
    
    selected_date = st.sidebar.date_input(
        "📅 Select Cut-off Date:",
        value=max_date,
        min_value=min_date,
        max_value=max_date,
        help="Selected date tak ka exact cumulative status aur movement calculate hoga."
    )

    plants_list = sorted(df_raw['Plnt'].dropna().unique().tolist()) if 'Plnt' in df_raw.columns else []
    selected_plants = st.sidebar.multiselect("🏭 Filter Plants:", plants_list, default=plants_list)

    customers_list = sorted(df_raw['Customer Name'].dropna().unique().tolist()) if 'Customer Name' in df_raw.columns else []
    selected_customers = st.sidebar.multiselect("👥 Filter Customers:", customers_list, default=[])

    df_filtered = df_raw.copy()
    if selected_plants:
        df_filtered = df_filtered[df_filtered['Plnt'].isin(selected_plants)]
    if selected_customers:
        df_filtered = df_filtered[df_filtered['Customer Name'].isin(selected_customers)]

    num_cols = ['No of PCS', 'Total Length (Mtr )', 'Bare Pipe Weight(TO)', 'Coated Pipe Weight(TO)']
    for nc in num_cols:
        if nc in df_filtered.columns:
            df_filtered[nc] = pd.to_numeric(df_filtered[nc], errors='coerce').fillna(0)

    df_filtered['Pstng_Date_Only'] = df_filtered['Pstng Date'].dt.date
    df_upto = df_filtered[df_filtered['Pstng_Date_Only'] <= selected_date].copy()
    df_today = df_filtered[df_filtered['Pstng_Date_Only'] == selected_date].copy()

    # Top KPI Metrics Banner
    st.markdown("---")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    
    total_in_pcs = df_upto[df_upto['IN_OUT'] == 'IN']['No of PCS'].sum() if 'IN_OUT' in df_upto.columns else 0
    total_out_pcs = abs(df_upto[df_upto['IN_OUT'] == 'OUT']['No of PCS'].sum()) if 'IN_OUT' in df_upto.columns else 0
    net_pcs = total_in_pcs - total_out_pcs
    total_tonnage = df_upto['Coated Pipe Weight(TO)'].where(df_upto['Coated Pipe Weight(TO)'] > 0, df_upto['Bare Pipe Weight(TO)']).sum()

    kpi1.metric("📥 Total Inward Pcs", f"{total_in_pcs:,.0f}")
    kpi2.metric("📤 Total Outward Pcs", f"{total_out_pcs:,.0f}")
    kpi3.metric("📦 Net Active Inventory (Pcs)", f"{net_pcs:,.0f}")
    kpi4.metric("⚖️ Total Handled Tonnage (MT)", f"{total_tonnage:,.2f} MT")
    st.markdown("---")

    # Navigation Tabs
    tab_dash, tab_bal, tab_inv, tab_sap = st.tabs([
        "📈 Dashboard Overview", 
        "⚖️ 1. Balance Report", 
        "📄 2. Invoice Report", 
        "📊 3. SAP Report & DDR"
    ])

    with tab_dash:
        st.subheader("📈 Executive Logistics Analytics & Visual Breakdown")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### 🏭 Plant-wise Volume Distribution (Pcs)")
            if 'Plnt' in df_upto.columns:
                plant_grp = df_upto.groupby('Plnt')['No of PCS'].sum().reset_index()
                st.bar_chart(plant_grp.set_index('Plnt'))
        with c2:
            st.markdown("#### 📊 Movement Type Split (IN vs OUT)")
            if 'IN_OUT' in df_upto.columns:
                io_grp = df_upto.groupby('IN_OUT')['No of PCS'].apply(lambda x: abs(x.sum())).reset_index()
                st.bar_chart(io_grp.set_index('IN_OUT'))

    with tab_bal:
        st.subheader("⚖️ Comprehensive Balance Report Overview")
        bal_summary = df_upto.groupby(['Customer Name', 'Plnt', 'TO Sales Ord.'], as_index=False).agg(
            Total_PCS=('No of PCS', 'sum'),
            Total_Length=('Total Length (Mtr )', 'sum'),
            Total_Weight=('Bare Pipe Weight(TO)', 'sum')
        )
        bal_summary = bal_summary[bal_summary['Total_PCS'] > 0].copy()
        st.dataframe(bal_summary, use_container_width=True, hide_index=True)
        excel_buf = generate_excel_report(bal_summary, "Balance Report", selected_date)
        st.download_button("📥 Download Balance Report (.xlsx)", data=excel_buf, file_name=f"Balance_Report_{selected_date}.xlsx")

    with tab_inv:
        st.subheader("📄 Invoice & Billing Document Register")
        inv_col = 'Accouting Document' if 'Accouting Document' in df_upto.columns else 'Reference'
        if inv_col in df_upto.columns:
            inv_df = df_upto[df_upto[inv_col].notna()].copy()
            st.dataframe(inv_df[['Pstng Date', 'Customer Name', 'TO Sales Ord.', inv_col, 'Material', 'No of PCS']].head(200), use_container_width=True, hide_index=True)

    with tab_sap:
        st.subheader("📊 SAP Detailed Reports & Daily Dispatch System")
        sub_report = st.selectbox(
            "Select SAP Report View:",
            [
                "Daily Dispatch Report (Print Preview - Positive Balance Only)",
                "Monthly Stock Reconciliation",
                "Plant-wise Summary",
                "Product-wise Summary",
                "Inward Register",
                "Outward Register",
                "Carbon Steel Pipe Balance"
            ]
        )
        st.info(f"Selected View: {sub_report}")

except Exception as e:
    st.error(f"Error executing application: {e}")
    st.info("Kripya ensure karein ki aapki Excel file GitHub repository ke root folder me uploaded hai.")
