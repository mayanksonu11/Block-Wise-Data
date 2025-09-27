import requests
import pandas as pd

def get_previous_data():
    url = "https://script.google.com/macros/s/AKfycbxQ1jTTUzv3ch47YAA0ZwnK6iKOBlk1PDQ3mUF9nGNA-KJUTFYmGBRPmsFNLSdLMq_6xQ/exec"  # Replace with actual Apps Script URL
    try:
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            if data.get('status') == 'success' and 'data' in data:
                prev_df = pd.DataFrame(data['data'])
                if not prev_df.empty:
                    # Convert Timestamp to datetime, localizing to GMT+5:30 (IST)
                    prev_df['Timestamp'] = pd.to_datetime(prev_df['Timestamp']).dt.tz_convert('Asia/Kolkata')
                    # Group by Block and take the latest entry
                    prev_df = prev_df.sort_values('Timestamp').groupby('Block').last().reset_index()
                    # Select and rename columns
                    prev_df = prev_df[['Block', 'Number_of_Schools', 'Total_Saplings', 'Timestamp']].rename(columns={
                        'Number_of_Schools': 'Previous_Number_of_Schools',
                        'Total_Saplings': 'Previous_Total_Saplings'
                    })
                    return prev_df
        return pd.DataFrame()  # Empty if no data or error
    except Exception as e:
        print(f"Could not fetch previous data: {str(e)}")
        return pd.DataFrame()
