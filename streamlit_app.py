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
    block_stats['PERCENTAGE'] = round(
        (block_stats['Total_Saplings'] / block_stats['Expected Saplings per Block']) * 100, 2
    )
    block_stats['Rank'] = block_stats['PERCENTAGE'].rank(ascending=False, method='dense').astype(int)
    block_stats = block_stats.sort_values(by='Rank')


    # --- Build Block-wise counts for exact saplings 1..50 ---
    ks = list(range(1, 51))
    indicators = {k: (muzaffarpur_df['Saplings'] == k).astype(int) for k in ks}
    exact_counts_df = pd.DataFrame(indicators)
    exact_counts_df['Block'] = muzaffarpur_df['Block'].values

    exact_block_counts = exact_counts_df.groupby('Block', as_index=False)[ks].sum()
    exact_block_counts = exact_block_counts.set_index('Block').reindex(block_stats['Block']).fillna(0).astype(int).reset_index()

    # --- Create Excel (multi-sheet) into BytesIO for download ---
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        # 1. Block_Stats sheet
        block_stats.to_excel(writer, sheet_name='Block_Stats', index=False)
        # 2. Exact_1_to_50 sheet (sorted by Block name)
        exact_sheet = exact_block_counts.rename(columns={k: str(k) for k in ks})
        exact_sheet_sorted = exact_sheet.sort_values(by='Block')
        exact_sheet_sorted.to_excel(writer, sheet_name='Exact_1_to_50', index=False)
        # 3. Individual block sheets
        for block in muzaffarpur_df['Block'].unique():
            block_df = muzaffarpur_df[muzaffarpur_df['Block'] == block].sort_values(by='Saplings', ascending=False)
            block_df.to_excel(writer, sheet_name=str(block), index=False)
    excel_buffer.seek(0)

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

    st.write("Block Statistics:")
    st.dataframe(block_stats)
