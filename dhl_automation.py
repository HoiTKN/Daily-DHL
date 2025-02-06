import os
import time
import requests
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime
import numpy as np

# Constants
GOOGLE_SHEET_ID = "16blaF86ky_4Eu4BK8AyXajohzpMsSyDaoPPKGVDYqWw"
SHEET_NAME = "DHL"
SERVICE_ACCOUNT_FILE = 'service-account.json'
LOGIN_URL = "https://ecommerceportal.dhl.com/Portal/pages/customer/statisticsdashboard.xhtml"
DOWNLOAD_URL = "https://ecommerceportal.dhl.com/Portal/pages/customer/statisticsdashboard.xhtml"  # Update this with correct download URL

def login_and_download_file(retry=3):
    """Login to DHL portal and download report using requests"""
    session = requests.Session()
    
    print("🔹 Accessing DHL portal...")
    login_payload = {
        "username": os.environ.get('DHL_USERNAME'),
        "password": os.environ.get('DHL_PASSWORD')
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    
    try:
        # Login
        response = session.post(LOGIN_URL, data=login_payload, headers=headers)
        if response.status_code == 200:
            print("✅ Login successful!")
        else:
            print(f"❌ Login failed! Status code: {response.status_code}")
            return None
            
        # Download report
        for attempt in range(retry):
            try:
                print(f"🔹 Downloading report (Attempt {attempt + 1})...")
                response = session.get(DOWNLOAD_URL, headers=headers)
                
                if response.status_code == 200:
                    output_path = "dhl_report.xlsx"
                    with open(output_path, 'wb') as f:
                        f.write(response.content)
                    print("✅ Report downloaded successfully!")
                    return output_path
            except Exception as e:
                print(f"❌ Download attempt {attempt + 1} failed: {str(e)}")
                if attempt < retry - 1:
                    time.sleep(2)
                    
    except Exception as e:
        print(f"❌ Error in login process: {str(e)}")
    return None

def process_data(file_path):
    """Process the downloaded DHL report"""
    print(f"🔹 Processing file: {file_path}")
    
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        
        processed_df = pd.DataFrame({
            'Order ID': df['Consignee Name'].str.extract(r'(\d{7})')[0].fillna(''),
            'Tracking Number': df['Tracking ID'].fillna(''),
            'Pickup DateTime': pd.to_datetime(df['Pickup Event DateTime'], errors='coerce').fillna(pd.NaT),
            'Delivery Date': pd.to_datetime(df['Delivery Date'], errors='coerce').fillna(pd.NaT),
            'Status': df['Last Status'].fillna('')
        })
        
        processed_df['Pickup DateTime'] = processed_df['Pickup DateTime'].apply(
            lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(x) else '')
        processed_df['Delivery Date'] = processed_df['Delivery Date'].apply(
            lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(x) else '')
        
        processed_df = processed_df.replace({np.nan: '', 'NaT': '', None: ''})
        
        print("✅ Data processing completed successfully")
        return processed_df
        
    except Exception as e:
        print(f"❌ Error processing data: {str(e)}")
        raise

def upload_to_google_sheets(df):
    """Upload processed data to Google Sheets"""
    print("🔹 Preparing to upload to Google Sheets...")
    
    try:
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            raise FileNotFoundError(f"Service account file not found at: {SERVICE_ACCOUNT_FILE}")
            
        print("🔹 Authenticating with Google Sheets...")
        creds = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        service = build("sheets", "v4", credentials=creds)
        
        print("🔹 Preparing data for upload...")
        headers = df.columns.tolist()
        data = df.astype(str).values.tolist()
        values = [headers] + data
        
        print("🔹 Clearing existing content...")
        service.spreadsheets().values().clear(
            spreadsheetId=GOOGLE_SHEET_ID,
            range=f"{SHEET_NAME}!A1:Z1000"
        ).execute()
        
        print("🔹 Uploading new data...")
        body = {
            'values': values
        }
        
        response = service.spreadsheets().values().update(
            spreadsheetId=GOOGLE_SHEET_ID,
            range=f"{SHEET_NAME}!A1",
            valueInputOption="USER_ENTERED",
            body=body
        ).execute()
        
        print("✅ Data uploaded successfully to Google Sheets")
        return True
        
    except Exception as e:
        print(f"❌ Error uploading to Google Sheets: {str(e)}")
        print("Debug information:")
        print(f"Sheet ID: {GOOGLE_SHEET_ID}")
        print(f"Sheet Name: {SHEET_NAME}")
        raise

def main():
    try:
        print("🚀 Starting DHL report automation process...")
        
        # Download report
        downloaded_file = login_and_download_file()
        if not downloaded_file:
            raise Exception("Failed to download report")
            
        # Process the downloaded file
        processed_df = process_data(downloaded_file)
        
        # Upload to Google Sheets
        upload_to_google_sheets(processed_df)
        
        print("🎉 Complete process finished successfully!")
        
    except Exception as e:
        print(f"❌ Process failed: {str(e)}")

if __name__ == "__main__":
    main()
