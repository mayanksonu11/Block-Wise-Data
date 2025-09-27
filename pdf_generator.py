import io
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from datetime import date

def generate_pdf_bytes(all_data_df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
    styles = getSampleStyleSheet()
    elements = []

    today_date = date.today()
    formatted_date = today_date.strftime("%d.%m.%Y")

    elements.append(Paragraph("Schoolwise Sapling Data - " + str(formatted_date), styles['Title']))
    elements.append(Spacer(1, 12))

    table_data = [all_data_df.columns.tolist()] + all_data_df.astype(str).values.tolist()

    total_width = landscape(A4)[0] - 40  # 40 for margins
    excel_widths = [5, 20, 12, 14, 12]
    if len(all_data_df.columns) > 5:
        excel_widths += [12] * (len(all_data_df.columns) - 5)
    width_sum = sum(excel_widths)
    col_widths = [w / width_sum * total_width for w in excel_widths]

    table = Table(table_data, repeatRows=1, colWidths=col_widths)

    # Style: Red text for rows where Saplings < 70
    saplings_col_idx = all_data_df.columns.get_loc('Saplings') + 1  # +1 due to 'Sl No.' insert
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
    ])

    # Apply red text color for rows where Saplings < 70
    for i, row in enumerate(all_data_df.itertuples(index=False), start=1):  # start=1 for header offset
        try:
            saplings = int(getattr(row, 'Saplings'))
            if saplings < 70:
                style.add('TEXTCOLOR', (0, i), (-1, i), colors.red)
            else:
                style.add('TEXTCOLOR', (0, i), (-1, i), colors.green)
        except Exception:
            pass

    table.setStyle(style)
    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer
