import streamlit as st
import pandas as pd
import io

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
    muzaffarpur_df['Block'] = muzaffarpur_df['Block'].fillna(method='ffill')

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
    block_stats['PERCENTAGE'] = round(
        (block_stats['Total_Saplings'] / block_stats['Expected Saplings per Block']) * 100, 2
    )
    block_stats['Rank'] = block_stats['PERCENTAGE'].rank(ascending=False, method='dense').astype(int)
    block_stats = block_stats.sort_values(by='Rank')

    # --- Create Excel (multi-sheet) into BytesIO for download ---
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        block_stats.to_excel(writer, sheet_name='Block_Stats', index=False)
        for block in muzaffarpur_df['Block'].unique():
            block_df = muzaffarpur_df[muzaffarpur_df['Block'] == block].sort_values(by='Saplings', ascending=False)
            block_df.to_excel(writer, sheet_name=str(block), index=False)
    excel_buffer.seek(0)  # rewind to start for download[6][11]

    st.success("Processing complete!")

    # --- Two download buttons: CSV and Excel ---
    st.download_button(
        label="Download Blockwise Data",
        data=csv_bytes,
        file_name="muzaffarpur_block_stats.csv",
        mime="text/csv"
    )  # Provide CSV correctly as bytes[2][1]

    st.download_button(
        label="Download Rank Data",
        data=excel_buffer,
        file_name="muzaffarpur_blockwise_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )  # Multi-sheet Excel from BytesIO[6][11]

    st.write("Block Statistics:")
    st.dataframe(block_stats)
