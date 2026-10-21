# Helper: PDF Generation Engine (Optimized for wrapping & clean layout)
    def generate_pdf_report(df_to_pdf, title_text):
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=15, leftMargin=15, topMargin=15, bottomMargin=15)
        elements = []
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=12, leading=14, textColor=colors.HexColor('#0B2545'))
        
        elements.append(Paragraph(f"<b>JINDAL SAW LIMITED</b> - {title_text}", title_style))
        elements.append(Paragraph(f"Report Date: {selected_date.strftime('%d-%b-%Y')}", styles['Normal']))
        elements.append(Spacer(1, 8))
        
        # Cell style for text wrapping inside table
        cell_style = ParagraphStyle(
            'TableCell',
            fontName='Helvetica',
            fontSize=6,
            leading=7,
            alignment=1 # Center aligned
        )
        header_style = ParagraphStyle(
            'TableHeader',
            fontName='Helvetica-Bold',
            fontSize=7,
            leading=8,
            textColor=colors.whitesmoke,
            alignment=1
        )
        
        # Format table data with Paragraphs for auto-wrapping
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
