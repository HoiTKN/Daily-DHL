import os
import time
import requests
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime
import numpy as np
import re

# Constants
GOOGLE_SHEET_ID = "16blaF86ky_4Eu4BK8AyXajohzpMsSyDaoPPKGVDYqWw"
SHEET_NAME = "DHL"
SERVICE_ACCOUNT_FILE = 'service-account.json'
BASE_URL = "https://ecommerceportal.dhl.com/Portal"
LOGIN_URL = f"{BASE_URL}/login.xhtml"
DASHBOARD_URL = f"{BASE_URL}/pages/customer/statisticsdashboard.xhtml"

def get_viewstate(html_content):
    """Extract ViewState from HTML content"""
    print("Looking for ViewState...")
    # Print a small sample of the HTML to debug
    print("HTML sample:", html_content[:200])
    match = re.search(r'name="javax\.faces\.ViewState"\s+id="[^"]+"\s+value="([^"]+)"', html_content)
    if match:
        print("ViewState found!")
        return match.group(1)
    print("ViewState not found in HTML")
    return None

def login_and_download_file(retry=3):
    """Login to DHL portal and download report"""
    session = requests.Session()
    
    print("🔹 Accessing DHL portal...")
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.5",
        "Content-Type": "application/x-www-form-urlencoded",
    }
    
    try:
        # First get the login page
        print("Getting login page...")
        response = session.get(LOGIN_URL, headers=headers)
        if response.status_code != 200:
            print(f"❌ Failed to access login page: {response.status_code}")
            return None
            
        viewstate = get_viewstate(response.text)
        if not viewstate:
            print("❌ Failed to get ViewState from login page")
            return None
            
        # Login with credentials
        login_payload = {
            "loginForm": "loginForm",
            "loginForm:email1": os.environ.get('DHL_USERNAME'),
            "loginForm:j_password": os.environ.get('DHL_PASSWORD'),
            "javax.faces.ViewState": viewstate,
            "loginForm:j_idt40": "loginForm:j_idt40"
        }
        
        print("Attempting login...")
        login_response = session.post(LOGIN_URL, data=login_payload, headers=headers)
        
        if login_response.status_code == 200:
            print("✅ Login successful!")
        else:
            print("❌ Login failed!")
            return None
        
        # Navigate to dashboard
        print("Accessing dashboard...")
        dashboard_response = session.get(DASHBOARD_URL, headers=headers)
        
        # Download report
        for attempt in range(retry):
            try:
                print(f"🔹 Downloading report (Attempt {attempt + 1})...")
                
                # Get the download URL directly
                download_url = f"{BASE_URL}/report/total-received-report.xhtml"
                download_response = session.get(download_url, headers={
                    **headers,
                    "Accept": "application/vnd.ms-excel,application/octet-stream"
                })
                
                if download_response.status_code == 200:
                    output_file = "dhl_report.xlsx"
                    with open(output_file, 'wb') as f:
                        f.write(download_response.content)
                    print(f"✅ Report downloaded successfully!")
                    return output_file
                else:
                    print(f"⚠️ Download failed with status code: {download_response.status_code}")
                
            except Exception as e:
                print(f"❌ Download attempt {attempt + 1} failed: {str(e)}")
                if attempt < retry - 1:
                    time.sleep(2)
                    
    except Exception as e:
        print(f"❌ Error in process: {str(e)}")
        print(f"Error details: {type(e).__name__}, {str(e)}")
    return None

def process_data(file_path):
    """Process the downloaded DHL report"""
    print(f"🔹 Processing file: {file_path}")
    
    try:
        # Try different methods to read the file
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
        except Exception as e:
            print(f"⚠️ Failed to read with openpyxl: {e}, trying xlrd...")
            df = pd.read_excel(file_path, engine='xlrd')
        
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
        try:
            service.spreadsheets().values().clear(
                spreadsheetId=GOOGLE_SHEET_ID,
                range=f"{SHEET_NAME}!A1:Z1000"
            ).execute()
        except Exception as e:
            print(f"⚠️ Warning: Could not clear sheet: {str(e)}")
        
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
