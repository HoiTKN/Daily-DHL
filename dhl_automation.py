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
BASE_URL = "https://ecommerceportal.dhl.com"
LOGIN_URL = f"{BASE_URL}/Portal/pages/customer/statisticsdashboard.xhtml"

def login_and_download_file(retry=3):
    """Login to DHL portal and download report using requests"""
    session = requests.Session()
    
    print("🔹 Accessing DHL portal...")
    
    # Initial GET request to get JSESSIONID and view state
    try:
        initial_response = session.get(LOGIN_URL)
        if initial_response.status_code != 200:
            print(f"❌ Failed to access portal: {initial_response.status_code}")
            return None
            
        # Extract ViewState from the response
        view_state = ""  # We'll need to extract this from the HTML
        # You might need to use regex or HTML parsing to get the viewstate
        
        # Login with proper form data
        login_payload = {
            "loginForm": "loginForm",
            "loginForm:email1": os.environ.get('DHL_USERNAME'),
            "loginForm:j_password": os.environ.get('DHL_PASSWORD'),
            "javax.faces.ViewState": view_state,
            "loginForm:j_idt40": "loginForm:j_idt40"
        }
        
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
        }
        
        login_response = session.post(LOGIN_URL, data=login_payload, headers=headers)
        
        if "Welcome" in login_response.text or "Dashboard" in login_response.text:
            print("✅ Login successful!")
        else:
            print("❌ Login failed!")
            return None
            
        # Now navigate to the report page and download
        for attempt in range(retry):
            try:
                print(f"🔹 Downloading report (Attempt {attempt + 1})...")
                
                # Make a GET request to the report page first
                report_url = f"{BASE_URL}/Portal/pages/customer/statisticsdashboard.xhtml"
                report_response = session.get(report_url)
                
                # Extract the new ViewState for the report download
                # You'll need to extract this from the response HTML
                
                # Prepare the download request with proper form data
                download_payload = {
                    "dashboardForm": "dashboardForm",
                    "javax.faces.ViewState": view_state,
                    "dashboardForm:xlsIcon": "dashboardForm:xlsIcon"
                }
                
                # Make the download request
                download_response = session.post(
                    report_url,
                    data=download_payload,
                    headers={
                        **headers,
                        "Accept": "application/vnd.ms-excel"
                    }
                )
                
                if download_response.headers.get('content-type', '').startswith('application/vnd.ms-excel'):
                    output_path = "dhl_report.xlsx"
                    with open(output_path, 'wb') as f:
                        f.write(download_response.content)
                    print("✅ Report downloaded successfully!")
                    return output_path
                else:
                    print(f"❌ Received wrong content type: {download_response.headers.get('content-type')}")
                    
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
