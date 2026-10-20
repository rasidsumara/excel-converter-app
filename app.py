import streamlit as st
import pandas as pd
import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Page Configuration
st.set_page_config(
    page_title="Jindal Saw - Dynamic Dispatch & Inventory System",
    page_icon="🔒",
    layout="wide"
)

# -------------------------------------------------------------
# PASSWORD PROTECTION LOGIC
# -------------------------------------------------------------
def check_password():
    """Returns True if the user enters the correct password."""
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    # Password Login UI
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

# Stop execution if password is not correct
if not check_password():
    st.stop()

# Logout Button in Sidebar
if st.sidebar.button("🔒 Logout"):
    st.session_state["password_correct"] = False
    st.rerun()

# -------------------------------------------------------------
# MAIN APPLICATION CODE (RUNS AFTER SUCCESSFUL LOGIN)
# -------------------------------------------------------------
st.title("📊 Management Dispatch & Logistics Control System")
st.caption("Jindal Saw Limited - Live Inward, Outward & Inventory Management System")

# Load Raw Base Data
@st.cache_data
def load_raw_data():
    file_path = "Inward_Outward_FIXED_and_FAST.xlsx"
    df_raw = pd.read_excel(file_path, sheet_name="Inward-Outward Data")
    df_raw['Pstng Date'] = pd.to_datetime(df_raw['Pstng Date'], errors='coerce')
    return df_raw

try:
    df_raw = load_raw_data()
    
    # Date Setup
    valid_dates = df_raw['Pstng Date'].dropna()
    min_date = valid_dates.min().date() if not valid_dates.empty else datetime.date(2025, 1, 1)
    max_date = valid_dates.max().date() if not valid_dates.empty else datetime.date.today()

    # Sidebar Controls
    st.sidebar.header("⚙️ Report Control Panel")
    selected_date = st.sidebar.date_input(
        "📅 Select Cut-off Date:",
        value=max_date,
        min_value=min_date,
        max_value=max_date,
        help="Selected date tak ka exact cumulative status aur movement calculate hoga."
    )
    
    report_type = st.sidebar.radio(
        "Select Excel Format View:",
        [
            "1. Daily Dispatch Report (Print Preview - Positive Balance Only)",
            "2. Monthly Stock Reconciliation (Opening / In / Out / Closing)",
            "3. Plant-wise Summary",
            "4. Product-wise Summary",
            "5. Inward Register (SO & Item Wise Sum)",
            "6. Outward Register (SO & Item Wise Sum)",
            "7. Carbon Steel Pipe Balance"
        ]
    )

    # Clean numeric fields
    num_cols = ['No of PCS', 'Total Length (Mtr )', 'Bare Pipe Weight(TO)', 'Coated Pipe Weight(TO)']
    for nc in num_cols:
        if nc in df_raw.columns:
            df_raw[nc] = pd.to_numeric(df_raw[nc], errors='coerce').fillna(0)

    # Convert date column to python date object
    df_raw['Pstng_Date_Only'] = df_raw['Pstng Date'].dt.date
    df_upto = df_raw[df_raw['Pstng_Date_Only'] <= selected_date].copy()
    df_today = df_raw[df_raw['Pstng_Date_Only'] == selected_date].copy()

    group_cols = ['Plnt', 'Customer Name', 'TO Sales Ord.', 'TO SO item', 'Valuation Class', 'OD', 'Thickness (mm)']
    for c in group_cols:
        df_upto[c] = df_upto[c].fillna("").astype(str)
        df_today[c] = df_today[c].fillna("").astype(str)

    metric_order = [
        'Sum of No of PCS',
        'Sum of Total Length (Mtr )',
        'Sum of Bare Pipe Weight(TO)',
        'Sum of Coated Pipe Weight(TO)'
    ]

    # Helper: Build DDR Pivot
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

    # Helper: Generate OpenPyXL Formatted Excel (Fixed Column Width Bug)
    def generate_excel_report(df_to_excel, title_text):
        wb = Workbook()
        ws = wb.active
        ws.title = "Dispatch Report"

        title_font = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        
        header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='0B2545', end_color='0B2545', fill_type='solid')
        
        data_font = Font(name='Calibri', size=10)
        bold_data_font = Font(name='Calibri', size=10, bold=True)
        
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        ws.merge_cells('A1:M1')
        ws['A1'] = f"DAILY DISPATCH REPORT — {title_text.upper()}"
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

        # Safe Column Width Auto-fit (Avoiding MergedCell issues)
        for col in ws.columns:
            max_len = 0
            col_letter = None
            for cell in col:
                if not col_letter and hasattr(cell, 'column_letter'):
                    col_letter = cell.column_letter
                if cell.row >= 4:  # Check length from headers onwards
                    val_str = str(cell.value or '')
                    if len(val_str) > max_len:
                        max_len = len(val_str)
            if col_letter:
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        return excel_buffer

    # Helper: PDF Generation Engine
    def generate_pdf_report(df_to_pdf, title_text):
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        elements = []
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=14, leading=16, textColor=colors.HexColor('#0B2545'))
        
        elements.append(Paragraph(f"<b>JINDAL SAW LIMITED</b> - {title_text}", title_style))
        elements.append(Paragraph(f"Report Date: {selected_date.strftime('%d-%b-%Y')}", styles['Normal']))
        elements.append(Spacer(1, 10))
        
        table_data = [df_to_pdf.columns.tolist()]
        for row in df_to_pdf.values.tolist():
            clean_row = [str(cell) if cell is not None else "" for cell in row]
            table_data.append(clean_row)
            
        pdf_table = Table(table_data, repeatRows=1)
        pdf_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0B2545')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 8),
            ('BOTTOMPADDING', (0,0), (-1,0), 5),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,1), (-1,-1), 7),
        ]))
        
        elements.append(pdf_table)
        doc.build(elements)
        buffer.seek(0)
        return buffer

    # -------------------------------------------------------------
    # VIEW 1: PRINT PREVIEW (ONLY NOS > 0 STRICTLY FILTERED & TOTAL)
    # -------------------------------------------------------------
    if "Print Preview" in report_type:
        st.subheader(f"📋 Daily Dispatch Print Preview — As of {selected_date.strftime('%d-%b-%Y')}")
        st.info("Note: Stock with Nos <= 0 (Zero or Negative Balance) is strictly hidden.")

        in_cum = build_ddr_pivot(df_upto[df_upto['Movement Direction'] == 'Inward']).rename(columns={'Val': 'Total Inward Till Date'})
        out_cum = build_ddr_pivot(df_upto[df_upto['Movement Direction'] == 'Outward']).rename(columns={'Val': 'Total Outward Till Date'})
        in_today = build_ddr_pivot(df_today[df_today['Movement Direction'] == 'Inward']).rename(columns={'Val': 'Today IN'})
        out_today = build_ddr_pivot(df_today[df_today['Movement Direction'] == 'Outward']).rename(columns={'Val': 'Today OUT'})

        m_keys = group_cols + ['Values']
        m_ddr = in_cum.merge(out_cum, on=m_keys, how='outer')
        m_ddr = m_ddr.merge(in_today, on=m_keys, how='left')
        m_ddr = m_ddr.merge(out_today, on=m_keys, how='left')
        m_ddr.fillna(0, inplace=True)

        m_ddr['Balance'] = m_ddr['Total Inward Till Date'] + m_ddr['Total Outward Till Date']

        # STRICT FILTER: NOS Balance MUST BE GREATER THAN 0 ONLY
        pcs_m = m_ddr[m_ddr['Values'] == 'Sum of No of PCS']
        strictly_positive_keys = pcs_m[pcs_m['Balance'] > 0][group_cols]

        if not strictly_positive_keys.empty:
            res_df = m_ddr.merge(strictly_positive_keys, on=group_cols)
        else:
            res_df = pd.DataFrame(columns=m_ddr.columns)

        res_df['Values_Cat'] = pd.Categorical(res_df['Values'], categories=metric_order, ordered=True)
        res_df.sort_values(by=group_cols + ['Values_Cat'], inplace=True)

        disp_cols = group_cols + ['Values', 'Today IN', 'Total Inward Till Date', 'Today OUT', 'Total Outward Till Date', 'Balance']
        
        total_rows = []
        for v in metric_order:
            sub_v = res_df[res_df['Values'] == v]
            t_row = {
                'Plnt': 'TOTAL',
                'Customer Name': 'GRAND TOTAL',
                'TO Sales Ord.': '',
                'TO SO item': '',
                'Valuation Class': '',
                'OD': '',
                'Thickness (mm)': '',
                'Values': v,
                'Today IN': sub_v['Today IN'].sum(),
                'Total Inward Till Date': sub_v['Total Inward Till Date'].sum(),
                'Today OUT': sub_v['Today OUT'].sum(),
                'Total Outward Till Date': sub_v['Total Outward Till Date'].sum(),
                'Balance': sub_v['Balance'].sum()
            }
            total_rows.append(t_row)

        df_totals = pd.DataFrame(total_rows)
        res_df_with_total = pd.concat([res_df[disp_cols], df_totals], ignore_index=True)

        formatted_df = res_df_with_total.copy()
        num_fields = ['Today IN', 'Total Inward Till Date', 'Today OUT', 'Total Outward Till Date', 'Balance']
        for col in num_fields:
            formatted_df[col] = formatted_df[col].apply(lambda x: "-" if abs(x) < 1e-4 else f"{x:,.2f}")

        st.dataframe(formatted_df, use_container_width=True, hide_index=True)

        col1, col2 = st.columns(2)
        with col1:
            excel_buf = generate_excel_report(res_df_with_total, "Daily Dispatch Print Preview")
            st.download_button("📥 Download Print Preview (Excel .xlsx)", data=excel_buf, file_name=f"DDR_Positive_Balance_{selected_date}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            pdf_buf = generate_pdf_report(formatted_df, "Daily Dispatch Print Preview")
            st.download_button("🖨️ Export / Print Complete Report as PDF", data=pdf_buf, file_name=f"DDR_Report_{selected_date}.pdf", mime="application/pdf")

    # -------------------------------------------------------------
    # VIEW 2: MONTHLY STOCK RECONCILIATION
    # -------------------------------------------------------------
    elif "Monthly Stock Reconciliation" in report_type:
        st.subheader("📅 Monthly Stock Reconciliation Report")

        df_raw['Effective_WT'] = df_raw['Coated Pipe Weight(TO)'].where(df_raw['Coated Pipe Weight(TO)'] > 0, df_raw['Bare Pipe Weight(TO)'])
        df_valid = df_raw.dropna(subset=['Pstng Date']).copy()
        df_valid['YearMonth'] = df_valid['Pstng Date'].dt.to_period('M')

        months = sorted(df_valid['YearMonth'].unique())
        recon_rows = []
        open_nos, open_mtr, open_mt = 0.0, 0.0, 0.0

        for m in months:
            m_data = df_valid[df_valid['YearMonth'] == m]
            in_m = m_data[m_data['Movement Direction'] == 'Inward']
            out_m = m_data[m_data['Movement Direction'] == 'Outward']

            in_nos = in_m['No of PCS'].sum()
            in_mtr = in_m['Total Length (Mtr )'].sum()
            in_mt = in_m['Effective_WT'].sum()

            out_nos = abs(out_m['No of PCS'].sum())
            out_mtr = abs(out_m['Total Length (Mtr )'].sum())
            out_mt = abs(out_m['Effective_WT'].sum())

            close_nos = open_nos + in_nos - out_nos
            close_mtr = open_mtr + in_mtr - out_mtr
            close_mt = open_mt + in_mt - out_mt

            recon_rows.append({
                'Month': str(m),
                'Opening Nos': open_nos,
                'Opening Mtr': open_mtr,
                'Opening MT': open_mt,
                'IN Nos': in_nos,
                'IN Mtr': in_mtr,
                'IN MT': in_mt,
                'OUT Nos': out_nos,
                'OUT Mtr': out_mtr,
                'OUT MT': out_mt,
                'Closing Nos': close_nos,
                'Closing Mtr': close_mtr,
                'Closing MT': close_mt
            })

            open_nos, open_mtr, open_mt = close_nos, close_mtr, close_mt

        df_recon = pd.DataFrame(recon_rows)

        formatted_recon = df_recon.copy()
        for c in formatted_recon.columns:
            if c != 'Month':
                formatted_recon[c] = formatted_recon[c].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)

        st.dataframe(formatted_recon, use_container_width=True, hide_index=True)

        col1, col2 = st.columns(2)
        with col1:
            excel_buf = generate_excel_report(df_recon, "Monthly Stock Reconciliation")
            st.download_button("📥 Download Monthly Reconciliation (Excel .xlsx)", data=excel_buf, file_name="Monthly_Stock_Reconciliation.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            pdf_buf = generate_pdf_report(formatted_recon, "Monthly Stock Reconciliation Report")
            st.download_button("🖨️ Export / Print Complete Report as PDF", data=pdf_buf, file_name="Monthly_Stock_Reconciliation.pdf", mime="application/pdf")

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

        col1, col2 = st.columns(2)
        with col1:
            excel_buf = generate_excel_report(p_summary, "Plant-wise Summary")
            st.download_button("📥 Download Summary (Excel .xlsx)", data=excel_buf, file_name=f"Plant_Summary_{selected_date}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            pdf_buf = generate_pdf_report(p_summary, "Plant-wise Summary")
            st.download_button("🖨️ Export / Print Complete Report as PDF", data=pdf_buf, file_name=f"Plant_Summary_{selected_date}.pdf", mime="application/pdf")

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
            else:
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

        col1, col2 = st.columns(2)
        with col1:
            excel_buf = generate_excel_report(prod_summary, "Product-wise Summary")
            st.download_button("📥 Download Summary (Excel .xlsx)", data=excel_buf, file_name=f"Product_Summary_{selected_date}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            pdf_buf = generate_pdf_report(prod_summary, "Product-wise Summary")
            st.download_button("🖨️ Export / Print Complete Report as PDF", data=pdf_buf, file_name=f"Product_Summary_{selected_date}.pdf", mime="application/pdf")

    # -------------------------------------------------------------
    # VIEW 5 & 6: INWARD / OUTWARD REGISTERS (STRICTLY NOS > 0)
    # -------------------------------------------------------------
    elif "Inward Register" in report_type or "Outward Register" in report_type:
        direction = "Inward" if "Inward" in report_type else "Outward"
        st.subheader(f"{'📥 Inward' if direction=='Inward' else '📤 Outward'} Register — As of {selected_date.strftime('%d-%b-%Y')}")
        
        reg_df = df_today[df_today['Movement Direction'] == direction].copy()
        reg_df['Recipient'] = reg_df['Recipient'].fillna("")
        reg_df['Customer Name'] = reg_df['Customer Name'].fillna("")
        reg_df['Outer Diameter (inches)'] = pd.to_numeric(reg_df['Outer Diameter (inches)'], errors='coerce').fillna(0)
        reg_df['Thickness (mm)'] = pd.to_numeric(reg_df['Thickness (mm)'], errors='coerce').fillna(0)

        grouped_reg = reg_df.groupby(['Plnt', 'Customer Name', 'Recipient', 'TO Sales Ord.', 'TO SO item', 'Outer Diameter (inches)', 'Thickness (mm)'], as_index=False).agg(
            Nos=('No of PCS', lambda x: abs(x.sum())),
            Length_Mtr=('Total Length (Mtr )', lambda x: abs(x.sum())),
            Coat_Weight=('Coated Pipe Weight(TO)', lambda x: abs(x.sum())),
            Bare_Weight=('Bare Pipe Weight(TO)', lambda x: abs(x.sum()))
        )

        # STRICT FILTER: NOS > 0 ONLY
        grouped_reg = grouped_reg[grouped_reg['Nos'] > 0].copy()

        grouped_reg.rename(columns={
            'Recipient': 'Destination',
            'TO Sales Ord.': 'Sales Doc.',
            'TO SO item': 'SO Item',
            'Outer Diameter (inches)': 'OD (inch)',
            'Thickness (mm)': 'Thickness (mm)',
            'Length_Mtr': 'Length (Mtr)',
            'Coat_Weight': 'Coat Weight (Mt)',
            'Bare_Weight': 'Bare Weight (Mt)'
        }, inplace=True)

        grouped_reg['NET Weight (Mt)'] = grouped_reg['Coat Weight (Mt)'].where(grouped_reg['Coat Weight (Mt)'] > 0, grouped_reg['Bare Weight (Mt)'])
        grouped_reg['Trip'] = grouped_reg['Nos']

        total_row = {
            'Plnt': 'TOTAL',
            'Customer Name': 'GRAND TOTAL',
            'Destination': '',
            'Sales Doc.': '',
            'SO Item': '',
            'OD (inch)': '',
            'Thickness (mm)': '',
            'Nos': grouped_reg['Nos'].sum(),
            'Length (Mtr)': grouped_reg['Length (Mtr)'].sum(),
            'Coat Weight (Mt)': grouped_reg['Coat Weight (Mt)'].sum(),
            'Bare Weight (Mt)': grouped_reg['Bare Weight (Mt)'].sum(),
            'NET Weight (Mt)': grouped_reg['NET Weight (Mt)'].sum(),
            'Trip': grouped_reg['Trip'].sum()
        }
        display_reg = pd.concat([grouped_reg, pd.DataFrame([total_row])], ignore_index=True)
        st.dataframe(display_reg, use_container_width=True, hide_index=True)

        col1, col2 = st.columns(2)
        with col1:
            excel_buf = generate_excel_report(display_reg, f"{direction} Register")
            st.download_button("📥 Download Register (Excel .xlsx)", data=excel_buf, file_name=f"{direction}_Register_{selected_date}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            pdf_buf = generate_pdf_report(display_reg, f"{direction} Register")
            st.download_button("🖨️ Export / Print Complete Report as PDF", data=pdf_buf, file_name=f"{direction}_Register_{selected_date}.pdf", mime="application/pdf")

    # -------------------------------------------------------------
    # VIEW 7: PIPE BALANCE REPORT (STRICTLY NOS > 0)
    # -------------------------------------------------------------
    elif "Carbon Steel Pipe Balance" in report_type:
        st.subheader(f"📦 Carbon Steel Pipe Balance — As of {selected_date.strftime('%d-%b-%Y')}")
        bal_df = df_upto.groupby(['Customer Name', 'TO Sales Ord.', 'TO SO item', 'OD', 'Thickness (mm)', 'Material Number'], as_index=False).agg(
            NOS=('No of PCS', 'sum'),
            MTR=('Total Length (Mtr )', 'sum'),
            MT=('Bare Pipe Weight(TO)', 'sum')
        )

        # STRICT FILTER: NOS > 0 ONLY
        bal_df = bal_df[bal_df['NOS'] > 0].copy()

        bal_df.rename(columns={
            'Customer Name': 'CUST NAME',
            'TO Sales Ord.': 'S.O. NUMBER',
            'TO SO item': 'SO ITEM',
            'OD': 'OD(MM)',
            'Thickness (mm)': 'THK',
            'Material Number': 'MATERIAL NUMBER'
        }, inplace=True)

        st.dataframe(bal_df, use_container_width=True, hide_index=True)

        col1, col2 = st.columns(2)
        with col1:
            excel_buf = generate_excel_report(bal_df, "Carbon Steel Pipe Balance")
            st.download_button("📥 Download Pipe Balance (Excel .xlsx)", data=excel_buf, file_name=f"Pipe_Balance_{selected_date}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            pdf_buf = generate_pdf_report(bal_df, "Carbon Steel Pipe Balance")
            st.download_button("🖨️ Export / Print Complete Report as PDF", data=pdf_buf, file_name=f"{direction}_Register_{selected_date}.pdf" if 'direction' in locals() else f"Pipe_Balance_{selected_date}.pdf", mime="application/pdf")

except Exception as e:
    st.error(f"Error executing application: {e}")
    st.info("Ensure `Inward_Outward_FIXED_and_FAST.xlsx` is uploaded in the GitHub repository root folder.")
