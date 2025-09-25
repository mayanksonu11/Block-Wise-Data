import streamlit as st
import pandas as pd
import io
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from datetime import date

st.title('Sapling Data Processing')

schools_per_block = {
    "BOCHAHA": 201, "MINAPUR": 247, "KURHANI": 291, "PAROO": 283, "MORAUL": 83,
    "SAKRA": 231, "MOTIPUR": 262, "AURAI": 217, "KANTI": 181, "SAHEBGANJ": 197,
    "MARWAN": 116, "BANDRA": 108, "MUSHARI": 249, "SARAIYA": 282, "GAIGHAT": 244, "KATRA": 177
}

uploaded_file = st.file_uploader("Upload your sapling.xlsx file", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # --- Data Processing ---
    muzaffarpur_df = df[df['District'] == 'MUZAFFARPUR'].copy()
    muzaffarpur_df['Block'] = muzaffarpur_df['Block'].ffill()  # Updated to use .ffill()

    block_stats = muzaffarpur_df.groupby('Block').agg(
        Number_of_Schools=('School', 'count'),
        Total_Saplings=('Saplings', 'sum')
    ).reset_index()

    # --- Prepare CSV bytes for download ---
    csv_bytes = block_stats.to_csv(index=False).encode('utf-8')  # CSV button uses bytes[2][1]

    # Add totals and expectations
    block_stats['Total no. of Schools'] = block_stats['Block'].apply(
        lambda x: schools_per_block.get(str(x).upper(), 0)
    )
    block_stats['Expected Saplings per Block'] = block_stats['Total no. of Schools'] * 70
    block_stats['Percentage'] = round(
        (block_stats['Total_Saplings'] / block_stats['Expected Saplings per Block']) * 100, 2
    )
    block_stats['Rank'] = block_stats['Percentage'].rank(ascending=False, method='dense').astype(int)
    block_stats = block_stats.sort_values(by='Rank')


    # --- Build Block-wise counts for exact saplings 1..50 ---
    ks = list(range(1, 51))
    indicators = {k: (muzaffarpur_df['Saplings'] == k).astype(int) for k in ks}
    exact_counts_df = pd.DataFrame(indicators)
    exact_counts_df['Block'] = muzaffarpur_df['Block'].values

    exact_block_counts = exact_counts_df.groupby('Block', as_index=False)[ks].sum()
    # Sort by Block name before adding grand total
    exact_block_counts = exact_block_counts.sort_values(by='Block')
    # Add Grand Total row
    grand_total = exact_block_counts[ks].sum()
    grand_total_row = pd.DataFrame(
        [['Grand Total'] + grand_total.tolist()],
        columns=['Block'] + [str(k) for k in ks]
    )
    # Convert column names to string for consistency
    exact_block_counts.columns = ['Block'] + [str(k) for k in ks]
    # Append grand total row
    exact_sheet_with_total = pd.concat([exact_block_counts, grand_total_row], ignore_index=True)

    # --- Prepare All_Data sheet: remove 'District', add 'Sl No.', sort block-wise by Saplings ---
    all_data_df = muzaffarpur_df.drop(columns=['District']).copy()
    # Sort block-wise by number of saplings (ascending)
    all_data_df = all_data_df.sort_values(['Block', 'Saplings'], ascending=[True, True])
    all_data_df.insert(0, 'Sl No.', range(1, len(all_data_df) + 1))

    # --- Create Excel (multi-sheet) into BytesIO for download ---
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        # 1. Block_Stats sheet
        block_stats.to_excel(writer, sheet_name='Block_Stats', index=False)
        # 2. Exact_1_to_50 sheet (sorted by Block name, with Grand Total)
        exact_sheet_with_total.to_excel(writer, sheet_name='Exact_1_to_50', index=False)
        # 3. All block data (with Sl No. and without District, sorted)
        all_data_df.to_excel(writer, sheet_name='All_Data', index=False)

        workbook  = writer.book
        border_fmt = workbook.add_format({'border': 1})

        # Apply border to all cells in each sheet
        for sheet_name in ['Block_Stats', 'Exact_1_to_50', 'All_Data']:
            worksheet = writer.sheets[sheet_name]
            df_to_format = {
                'Block_Stats': block_stats,
                'Exact_1_to_50': exact_sheet_with_total,
                'All_Data': all_data_df
            }[sheet_name]
            nrows, ncols = df_to_format.shape
            # +1 for header row
            worksheet.conditional_format(0, 0, nrows, ncols - 1, {'type': 'no_blanks', 'format': border_fmt})
            worksheet.conditional_format(0, 0, nrows, ncols - 1, {'type': 'blanks', 'format': border_fmt})

        # Set custom column widths for All_Data sheet
        worksheet = writer.sheets['All_Data']
        col_widths = [5, 20, 12, 14, 12]
        for i, width in enumerate(col_widths):
            worksheet.set_column(i, i, width)
        for i in range(len(col_widths), all_data_df.shape[1]):
            worksheet.set_column(i, i, 12)

    excel_buffer.seek(0)

    # --- Convert all_data_df to PDF ---
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=landscape(A4))
    styles = getSampleStyleSheet()
    elements = []

    today_date = date.today()
    formatted_date = today_date.strftime("%d.%m.%Y")

    elements.append(Paragraph("Schoolwise Sapling Data - "+ str(formatted_date), styles['Title']))
    elements.append(Spacer(1, 12))

    table_data = [all_data_df.columns.tolist()] + all_data_df.astype(str).values.tolist()

    total_width = landscape(A4)[0] - 40  # 40 for margins
    excel_widths = [5, 20, 12, 14, 12]
    if len(all_data_df.columns) > 5:
        excel_widths += [12] * (len(all_data_df.columns) - 5)
    width_sum = sum(excel_widths)
    col_widths = [w / width_sum * total_width for w in excel_widths]

    table = Table(table_data, repeatRows=1, colWidths=col_widths)

    # --- Style: Red text for rows where Saplings < 70 ---
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
    pdf_buffer.seek(0)

    st.success("Processing complete!")

    st.markdown("**This file contains block-wise current status of Total Saplings**")
    st.download_button(
        label="Download Blockwise Data",
        data=csv_bytes,
        file_name="muzaffarpur_block_stats.csv",
        mime="text/csv"
    )

    st.markdown("**This file contains Rank data, sapling counts (1-50) per block, and detailed data for each block.**")
    st.download_button(
        label="Download Detailed Report",
        data=excel_buffer.getvalue(),
        file_name="muzaffarpur_blockwise_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown("**This file contains all block data as PDF.**")
    st.download_button(
        label="Download All Block Data (PDF)",
        data=pdf_buffer.getvalue(),
        file_name="muzaffarpur_all_block_data.pdf",
        mime="application/pdf"
    )

    st.write("Block Statistics:")
    st.dataframe(block_stats)
