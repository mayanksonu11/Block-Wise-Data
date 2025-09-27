import streamlit as st
import pandas as pd
from data_processing import generate_block_stats, add_totals_and_rankings, generate_exact_counts, generate_all_data
from csv_generator import generate_csv_bytes
from excel_generator import generate_block_stats_excel_bytes, generate_block_rank_stats_excel_bytes
from pdf_generator import generate_pdf_bytes
from data_fetcher import get_previous_data
from data_sender import send_data_to_storage

st.title('Sapling Data Processing')

uploaded_file = st.file_uploader("Upload your sapling.xlsx file", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Get previous data
    prev_data = get_previous_data()

    # Get previous data date
    previous_data_date = None
    if not prev_data.empty:
        latest_timestamp = pd.to_datetime(prev_data['Timestamp']).max()
        previous_data_date = latest_timestamp.strftime("%d.%m.%Y")

    # Process data
    block_stats = generate_block_stats(df, prev_data)
    block_stats_with_rankings = add_totals_and_rankings(block_stats.copy())
    exact_sheet_with_total = generate_exact_counts(df)
    all_data_df = generate_all_data(df)

    # Generate files
    block_stats_excel_buffer = generate_block_stats_excel_bytes(block_stats)
    excel_buffer = generate_block_rank_stats_excel_bytes(block_stats_with_rankings)
    pdf_buffer = generate_pdf_bytes(all_data_df)

    st.success("Processing complete!")

    if previous_data_date:
        st.write(f"**Previous data as of: {previous_data_date}**")

    st.markdown("**This file contains block-wise stats as Excel.**")
    st.download_button(
        label="Download Blockwise Data (Excel)",
        data=block_stats_excel_buffer.getvalue(),
        file_name="muzaffarpur_block_stats.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown("**This file contains block rankings data.**")
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

    st.markdown("**Use this button to update today's data in database.**")
    # Button to store data persistently
    if st.button("Store Data"):
        # Prepare data: blockwise stats with schools and saplings
        data_to_send = {
            "blockwise_data": block_stats[['Block', 'Number_of_Schools', 'Total_Saplings']].to_dict(orient='records')
        }
        result = send_data_to_storage(data_to_send)
        if "successfully" in result:
            st.success(result)
        else:
            st.error(result)
