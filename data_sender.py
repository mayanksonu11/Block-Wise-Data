import requests

def send_data_to_storage(data):
    url = "https://script.google.com/macros/s/AKfycbxQ1jTTUzv3ch47YAA0ZwnK6iKOBlk1PDQ3mUF9nGNA-KJUTFYmGBRPmsFNLSdLMq_6xQ/exec"  # Replace with actual API endpoint
    headers = {'Content-Type': 'application/json'}
    try:
        response = requests.post(url, json=data, headers=headers)
        if response.status_code == 200:
            return "Data stored successfully!"
        else:
            return f"Failed to store data: {response.status_code}"
    except Exception as e:
        return f"Error sending data: {str(e)}"
