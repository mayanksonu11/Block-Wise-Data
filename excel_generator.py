import pandas as pd
import io

def generate_block_stats_excel_bytes(block_stats):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Create a copy and remove underscores from column names
        block_stats_clean = block_stats.copy()
        block_stats_clean.columns = block_stats_clean.columns.str.replace('_', ' ')
        block_stats_clean.to_excel(writer, sheet_name='Block_Stats', index=False)

        workbook = writer.book
        border_fmt = workbook.add_format({'border': 1})
        header_fmt = workbook.add_format({'border': 1, 'text_wrap': True, 'bold': True})

        worksheet = writer.sheets['Block_Stats']
        nrows, ncols = block_stats.shape
        worksheet.conditional_format(0, 0, nrows, ncols - 1, {'type': 'no_blanks', 'format': border_fmt})
        worksheet.conditional_format(0, 0, nrows, ncols - 1, {'type': 'blanks', 'format': border_fmt})

        # Set row height for first row
        worksheet.set_row(0, 45)

        # Apply text wrap format to header cells
        for col in range(ncols):
            worksheet.write(0, col, block_stats_clean.columns[col], header_fmt)

        # Set column widths
        worksheet.set_column(0, 0, 8)  # Sl. No.
        worksheet.set_column(1, 1, 11)  # Block
        for i in range(2, ncols):
            worksheet.set_column(i, i, 11)  # Rest of columns

    buffer.seek(0)
    return buffer

def generate_block_rank_stats_excel_bytes(block_stats_with_ranks):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Block_Rankings sheet
        block_stats_with_ranks.to_excel(writer, sheet_name='Block_Rankings', index=False)

        workbook = writer.book
        border_fmt = workbook.add_format({'border': 1})

        # Apply border to all cells in the sheet
        worksheet = writer.sheets['Block_Rankings']
        nrows, ncols = block_stats_with_ranks.shape
        worksheet.conditional_format(0, 0, nrows, ncols - 1, {'type': 'no_blanks', 'format': border_fmt})
        worksheet.conditional_format(0, 0, nrows, ncols - 1, {'type': 'blanks', 'format': border_fmt})

    buffer.seek(0)
    return buffer
