import pandas as pd

schools_per_block = {
    "BOCHAHA": 201, "MINAPUR": 247, "KURHANI": 291, "PAROO": 283, "MORAUL": 83,
    "SAKRA": 231, "MOTIPUR": 262, "AURAI": 217, "KANTI": 181, "SAHEBGANJ": 197,
    "MARWAN": 116, "BANDRA": 108, "MUSHARI": 249, "SARAIYA": 282, "GAIGHAT": 244, "KATRA": 177
}

def get_filtered_df(df):
    """Filter dataframe for Muzaffarpur district and forward fill Block column."""
    muzaffarpur_df = df[df['District'] == 'MUZAFFARPUR'].copy()
    muzaffarpur_df['Block'] = muzaffarpur_df['Block'].ffill()
    return muzaffarpur_df

def generate_block_stats(df, prev_data):
    """Generate block statistics dataframe."""
    muzaffarpur_df = get_filtered_df(df)

    # Compute block stats
    block_stats = muzaffarpur_df.groupby('Block').agg(
        Number_of_Schools=('School', 'count'),
        Total_Saplings=('Saplings', 'sum')
    ).reset_index()

    # Merge previous data
    if not prev_data.empty:
        block_stats = block_stats.merge(prev_data, on='Block', how='left')
        block_stats['Previous_Number_of_Schools'] = block_stats['Previous_Number_of_Schools'].fillna(0).astype(int)
        block_stats['Previous_Total_Saplings'] = block_stats['Previous_Total_Saplings'].fillna(0).astype(int)
        block_stats = block_stats.drop(columns=['Timestamp'], errors='ignore')
    else:
        block_stats['Previous_Number_of_Schools'] = 0
        block_stats['Previous_Total_Saplings'] = 0

    # Calculate progression
    block_stats['Increase_Number_of_Schools'] = block_stats['Number_of_Schools'] - block_stats['Previous_Number_of_Schools']
    block_stats['Increase_Total_Saplings'] = block_stats['Total_Saplings'] - block_stats['Previous_Total_Saplings']

    # Reorder columns
    block_stats = block_stats[['Block', 'Previous_Number_of_Schools', 'Previous_Total_Saplings', 'Number_of_Schools', 'Total_Saplings', 'Increase_Number_of_Schools', 'Increase_Total_Saplings'] + [col for col in block_stats.columns if col not in ['Block', 'Previous_Number_of_Schools', 'Previous_Total_Saplings', 'Number_of_Schools', 'Total_Saplings', 'Increase_Number_of_Schools', 'Increase_Total_Saplings']]]

    # Add Sl No.
    block_stats.insert(0, 'Sl No.', range(1, len(block_stats) + 1))

    # Add total row
    total_dict = {'Sl No.': '', 'Block': 'Total'}
    for col in block_stats.columns:
        if col not in ['Sl No.', 'Block']:
            total_dict[col] = block_stats[col].sum()
    block_stats = pd.concat([block_stats, pd.DataFrame([total_dict])], ignore_index=True)

    return block_stats

def add_totals_and_rankings(block_stats):
    """Add totals, expectations, percentages and rankings to block stats."""
    # Add totals and expectations
    block_stats['Total no. of Schools'] = block_stats['Block'].apply(
        lambda x: schools_per_block.get(str(x).upper(), 0)
    )
    block_stats['Expected Saplings'] = block_stats['Total no. of Schools'] * 70
    block_stats['Current_Sapling_Count'] = block_stats['Total_Saplings']
    block_stats['Percentage'] = round(
        (block_stats['Current_Sapling_Count'] / block_stats['Expected Saplings']) * 100, 2
    )
    block_stats['Rank'] = block_stats['Percentage'].rank(ascending=False, method='dense').astype(int)

    # Reorder columns
    block_stats = block_stats[['Sl No.', 'Block', 'Total no. of Schools', 'Expected Saplings', 'Current_Sapling_Count', 'Percentage', 'Rank']]
    block_stats = block_stats.sort_values(by='Rank')

    return block_stats

def generate_exact_counts(df):
    """Generate exact sapling counts dataframe."""
    muzaffarpur_df = get_filtered_df(df)

    # Build exact counts
    ks = list(range(1, 51))
    indicators = {k: (muzaffarpur_df['Saplings'] == k).astype(int) for k in ks}
    exact_counts_df = pd.DataFrame(indicators)
    exact_counts_df['Block'] = muzaffarpur_df['Block'].values

    exact_block_counts = exact_counts_df.groupby('Block', as_index=False)[ks].sum()
    exact_block_counts = exact_block_counts.sort_values(by='Block')
    grand_total = exact_block_counts[ks].sum()
    grand_total_row = pd.DataFrame(
        [['Grand Total'] + grand_total.tolist()],
        columns=['Block'] + [str(k) for k in ks]
    )
    exact_block_counts.columns = ['Block'] + [str(k) for k in ks]
    exact_sheet_with_total = pd.concat([exact_block_counts, grand_total_row], ignore_index=True)

    return exact_sheet_with_total

def generate_all_data(df):
    """Generate all data dataframe."""
    muzaffarpur_df = get_filtered_df(df)

    # All data df
    all_data_df = muzaffarpur_df.drop(columns=['District']).copy()
    all_data_df = all_data_df.sort_values(['Block', 'Saplings'], ascending=[True, True])
    all_data_df.insert(0, 'Sl No.', range(1, len(all_data_df) + 1))

    return all_data_df

# Keep the old function for backward compatibility
def process_data(df, prev_data):
    block_stats = generate_block_stats(df, prev_data)
    exact_sheet_with_total = generate_exact_counts(df)
    all_data_df = generate_all_data(df)
    return block_stats, exact_sheet_with_total, all_data_df
