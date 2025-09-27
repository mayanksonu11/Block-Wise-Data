import pandas as pd

def generate_csv_bytes(block_stats):
    return block_stats.to_csv(index=False).encode('utf-8')
