import streamlit as st
import pandas as pd

st.title('Sapling Data Processing')

uploaded_file = st.file_uploader("Upload your sapling.xlsx file", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # --- Data Processing (from previous notebook cells) ---
    # Filter for Muzaffarpur district
    muzaffarpur_df = df[df['District'] == 'MUZAFFARPUR'].copy()

    # Impute missing 'Block' values (using ffill as in the original notebook)
    muzaffarpur_df['Block'] = muzaffarpur_df['Block'].fillna(method='ffill')

    # Calculate block statistics
    block_stats = muzaffarpur_df.groupby('Block').agg(
        Number_of_Schools=('School', 'count'),
        Total_Saplings=('Saplings', 'sum')
    ).reset_index()

    # Rank blocks by Total_Saplings
    block_stats['Rank'] = block_stats['Total_Saplings'].rank(ascending=False, method='dense').astype(int)
    block_stats = block_stats.sort_values(by='Rank')
    block_stats = block_stats.drop(['Rank'], axis=1)

    # --- Create Excel file with multiple sheets ---
    excel_file = 'muzaffarpur_blockwise_data.xlsx'
    with pd.ExcelWriter(excel_file) as writer:
        # Write block_stats as the first sheet
        block_stats.to_excel(writer, sheet_name='Block_Stats', index=False)

        # Write individual block data
        for block in muzaffarpur_df['Block'].unique():
            block_df = muzaffarpur_df[muzaffarpur_df['Block'] == block].sort_values(by='Saplings', ascending=False)
            block_df.to_excel(writer, sheet_name=block, index=False)

    st.success("Processing complete!")

    # --- Provide download link ---
    with open(excel_file, "rb") as file:
        st.download_button(
            label="Download Muzaffarpur_blockwise_data.xlsx",
            data=file,
            file_name=excel_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.write("Block Statistics:")
    st.dataframe(block_stats)