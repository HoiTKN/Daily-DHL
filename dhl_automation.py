import os
import time
import requests
import pandas as pd
import mimetypes
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
        "Origin": "https://ecommerceportal.dhl.com",
        "Referer": "https://ecommerceportal.dhl.com/Portal/pages/customer/statisticsdashboard.xhtml"
    }
    
    try:
        # Get initial page to get viewstate
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
            "j_username": "truongcongdai4@gmail.com",
            "j_password": "Thavi@26052565",
            "javax.faces.ViewState": viewstate
        }
        
        print("Attempting login...")
        login_response = session.post(LOGIN_URL, data=login_payload, headers=headers)
        
        if login_response.status_code == 200:
            print("✅ Login successful!")
        else:
            print("❌ Login failed!")
            return None
        
        # Change language to English
        dashboard_response = session.get(DASHBOARD_URL, headers=headers)
        dashboard_viewstate = get_viewstate(dashboard_response.text)
        
        language_payload = {
            "j_idt19": "j_idt19",
            "j_idt19:j_idt21_input": "en",
            "javax.faces.ViewState": dashboard_viewstate
        }
        
        session.post(DASHBOARD_URL, data=language_payload, headers=headers)
        
        for attempt in range(retry):
            try:
                print(f"🔹 Downloading report (Attempt {attempt + 1})...")
                
                # Set date range and generate report
                dashboard_response = session.get(DASHBOARD_URL, headers=headers)
                dashboard_viewstate = get_viewstate(dashboard_response.text)
                
                date_payload = {
                    "dashboardForm": "dashboardForm",
                    "dashboardForm:frmDate_input": "01-01-2025",
                    "dashboardForm:toDate_input": datetime.now().strftime("%d-%m-%Y"),
                    "javax.faces.ViewState": dashboard_viewstate
                }
                
                session.post(DASHBOARD_URL, data=date_payload, headers=headers)
                
                # Get updated viewstate after date setting
                dashboard_response = session.get(DASHBOARD_URL, headers=headers)
                dashboard_viewstate = get_viewstate(dashboard_response.text)
                
                # Download Excel file
                download_payload = {
                    "dashboardForm": "dashboardForm",
                    "dashboardForm:j_idt116": "dashboardForm:j_idt116",
                    "javax.faces.ViewState": dashboard_viewstate
                }
                
                download_headers = {
                    **headers,
                    "Accept": "application/vnd.ms-excel",
                    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"
                }
                
                # Make the actual download request
                download_url = f"{BASE_URL}/report/total-received-report.xhtml"
                download_response = session.get(
                    download_url,
                    headers=download_headers,
                    allow_redirects=True
                )
                
                if download_response.status_code == 200:
                    content_type = download_response.headers.get('content-type', '')
                    print(f"Content type: {content_type}")
                    
                    if 'excel' in content_type.lower() or 'application/vnd.ms-excel' in content_type.lower():
                        output_file = "dhl_report.xlsx"
                        with open(output_file, 'wb') as f:
                            f.write(download_response.content)
                        print("✅ Report downloaded successfully!")
                        return output_file
                    else:
                        print("⚠️ Received non-Excel content")
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
    """Process the downloaded DHL report with smart format detection"""
    print(f"🔹 Processing file: {file_path}")
    
    try:
        # First, let's examine the file content
        with open(file_path, 'rb') as f:
            content = f.read()
            print(f"File size: {len(content)} bytes")
            print("First 200 bytes:", content[:200])
            f.seek(0)
        
        try:
            print("📄 Trying to read file...")
            df = pd.read_csv(file_path, encoding='utf-8')
            print("✅ Successfully read file")
        except Exception as e:
            print(f"⚠️ Failed to read as CSV: {str(e)}, trying other formats...")
            try:
                df = pd.read_excel(file_path)
                print("✅ Successfully read as Excel")
            except Exception as e:
                print(f"⚠️ Failed to read as Excel: {str(e)}")
                raise ValueError("Could not read file in any format")

        # Print actual columns we received
        print("\nActual columns in the file:")
        print(df.columns.tolist())
        print("\nFirst few rows of raw data:")
        print(df.head())

        print("🔹 Processing data...")
        
        # Check if required columns exist
        required_columns = ['Consignee Name', 'Tracking ID', 'Pickup Event DateTime', 'Delivery Date', 'Last Status']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"❌ Missing required columns: {missing_columns}")
            # Try to find similar column names
            for missing_col in missing_columns:
                similar_cols = [col for col in df.columns if missing_col.lower() in col.lower()]
                if similar_cols:
                    print(f"Found similar columns for '{missing_col}': {similar_cols}")

        # Extract Order ID (first 7 digits from Consignee Name)
        def extract_order_id(text):
            if pd.isna(text):
                return ''
            match = re.search(r'(\d{7})', str(text))
            return match.group(1) if match else ''

        processed_df = pd.DataFrame({
            'Order ID': df['Consignee Name'].apply(extract_order_id),
            'Tracking Number': df['Tracking ID'].fillna(''),
            'Pickup DateTime': pd.to_datetime(df['Pickup Event DateTime'], errors='coerce').fillna(pd.NaT),
            'Delivery Date': pd.to_datetime(df['Delivery Date'], errors='coerce').fillna(pd.NaT),
            'Status': df['Last Status'].fillna('')
        })
        
        # Format datetime columns
        processed_df['Pickup DateTime'] = processed_df['Pickup DateTime'].apply(
            lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(x) else '')
        processed_df['Delivery Date'] = processed_df['Delivery Date'].apply(
            lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(x) else '')
        
        # Clean up any remaining NaN values
        processed_df = processed_df.replace({np.nan: '', 'NaT': '', None: ''})
        
        print("✅ Data processing completed successfully")
        print(f"Processed {len(processed_df)} rows")
        
        return processed_df
        
    except Exception as e:
        print(f"❌ Error processing data: {str(e)}")
        # Add stack trace for better debugging
        import traceback
        print("Stack trace:")
        print(traceback.format_exc())
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
    finally:
        # Cleanup temporary file
        if 'downloaded_file' in locals() and os.path.exists(downloaded_file):
            try:
                os.remove(downloaded_file)
            except:
                pass

if __name__ == "__main__":
    main()
