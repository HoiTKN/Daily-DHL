# File: dhl_automation.py
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd
import os
from datetime import datetime
import time
import numpy as np

# Constants
GOOGLE_SHEET_ID = "16blaF86ky_4Eu4BK8AyXajohzpMsSyDaoPPKGVDYqWw"
SHEET_NAME = "DHL"
SERVICE_ACCOUNT_FILE = 'service-account.json'
DOWNLOAD_FOLDER = os.path.join(os.getcwd(), "downloads")

# Create downloads directory if it doesn't exist
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

def setup_chrome_driver():
    """Setup Chrome driver with necessary options for GitHub Actions"""
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--disable-gpu')
    
    # Set up download preferences
    chrome_options.add_experimental_option(
        'prefs', {
            'download.default_directory': DOWNLOAD_FOLDER,
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': True
        }
    )
    
    # Setup service with ChromeDriverManager
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)

def login_to_dhl(driver):
    """Login to DHL portal using environment variables"""
    try:
        print("🔹 Accessing DHL portal...")
        driver.get("https://ecommerceportal.dhl.com/Portal/pages/customer/statisticsdashboard.xhtml")
        
        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "email1"))
        )
        username_input.send_keys(os.environ.get('DHL_USERNAME'))
        
        password_input = driver.find_element(By.NAME, "j_password")
        password_input.send_keys(os.environ.get('DHL_PASSWORD'))
        
        login_button = driver.find_element(By.CLASS_NAME, "btn-login")
        login_button.click()
        
        print("✅ Login successful!")
        return True
        
    except Exception as e:
        print(f"❌ Login failed: {str(e)}")
        return False

[Previous functions remain the same: change_language_to_english, navigate_to_dashboard, download_report, get_latest_file, process_data]

def upload_to_google_sheets(df):
    """Upload processed data to Google Sheets using service account"""
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

[Main function remains the same]

if __name__ == "__main__":
    main()
