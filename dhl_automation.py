from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
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
    chrome_options.add_argument('--disable-extensions')
    
    # Set download preferences
    prefs = {
        'download.default_directory': DOWNLOAD_FOLDER,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': True,
        'profile.default_content_settings.popups': 0
    }
    chrome_options.add_experimental_option('prefs', prefs)
    
    # Create Chrome driver instance
    driver = webdriver.Chrome(options=chrome_options)
    
    return driver
    
    # Set download preferences
    prefs = {
        'download.default_directory': DOWNLOAD_FOLDER,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': True,
        'profile.default_content_settings.popups': 0
    }
    chrome_options.add_experimental_option('prefs', prefs)
    
    # Set up Chrome service
    service = Service('/usr/local/bin/chromedriver')  # Use system-installed ChromeDriver
    return webdriver.Chrome(service=service, options=chrome_options)

def login_to_dhl(driver):
    """Login to DHL portal using environment variables"""
    try:
        print("🔹 Accessing DHL portal...")
        driver.get("https://ecommerceportal.dhl.com/Portal/pages/customer/statisticsdashboard.xhtml")
        
        # Wait for username field and login
        username_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "email1"))
        )
        username_input.send_keys(os.environ.get('DHL_USERNAME'))
        
        # Find and fill password
        password_input = driver.find_element(By.NAME, "j_password")
        password_input.send_keys(os.environ.get('DHL_PASSWORD'))
        
        # Click login button
        login_button = driver.find_element(By.CLASS_NAME, "btn-login")
        login_button.click()
        
        # Wait for successful login
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Dashboard')]"))
        )
        
        print("✅ Login successful!")
        return True
        
    except Exception as e:
        print(f"❌ Login failed: {str(e)}")
        return False

def change_language_to_english(driver):
    """Change portal language to English"""
    try:
        time.sleep(5)  # Wait for page to fully load
        
        # Wait for language selector and click
        language_selector = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "j_idt19:j_idt21_label"))
        )
        language_selector.click()
        
        # Wait and click English option
        english_option = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//li[text()='English']"))
        )
        english_option.click()
        
        # Wait for language change to take effect
        time.sleep(5)
        
        print("✅ Language changed to English")
        return True
        
    except Exception as e:
        print(f"❌ Failed to change language: {str(e)}")
        return False

def navigate_to_dashboard(driver):
    """Navigate to dashboard and set date range"""
    try:
        time.sleep(5)  # Wait for page to fully load
        
        dashboard_link = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Dashboard')]"))
        )
        dashboard_link.click()
        
        # Wait for date inputs to be present
        start_date_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "dashboardForm:frmDate_input"))
        )
        driver.execute_script("arguments[0].value = '01-01-2025'", start_date_input)
        
        end_date_input = driver.find_element(By.ID, "dashboardForm:toDate_input")
        current_date = datetime.now().strftime("%d-%m-%Y")
        driver.execute_script(f"arguments[0].value = '{current_date}'", end_date_input)
        
        # Press Enter to confirm dates
        start_date_input.send_keys("\n")
        time.sleep(2)
        end_date_input.send_keys("\n")
        
        # Wait for data to load
        time.sleep(5)
        
        print("✅ Dashboard navigation and date setting complete")
        return True
        
    except Exception as e:
        print(f"❌ Failed to navigate to dashboard: {str(e)}")
        return False

def download_report(driver):
    """Download the Total Received report"""
    try:
        time.sleep(5)  # Wait for page to fully load
        
        # Wait for and click download button
        download_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//td[contains(.,'Total Received')]//img[@id='xlsIcon']"))
        )
        download_button.click()
        
        # Wait for download to complete
        time.sleep(10)
        
        print("✅ Report download initiated")
        return True
        
    except Exception as e:
        print(f"❌ Failed to download report: {str(e)}")
        return False

def get_latest_file(folder_path, max_attempts=5, delay=2):
    """Get the most recently downloaded file from the specified folder"""
    for attempt in range(max_attempts):
        try:
            print(f"🔍 Looking for downloaded file (Attempt {attempt + 1})...")
            files = [
                os.path.join(folder_path, f) for f in os.listdir(folder_path)
                if (f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.csv'))
                and not f.startswith('~$')
                and 'total_received_report' in f.lower()
            ]
            
            if not files:
                print(f"No matching files found. Attempt {attempt + 1}/{max_attempts}")
                time.sleep(delay)
                continue
                
            latest_file = max(files, key=os.path.getctime)
            
            # Verify file is readable
            with open(latest_file, 'rb') as f:
                f.read(1)
                
            print(f"✅ Found latest file: {latest_file}")
            return latest_file
            
        except (PermissionError, FileNotFoundError) as e:
            print(f"⚠️ Attempt {attempt + 1}: File access error: {str(e)}")
            if attempt < max_attempts - 1:
                print(f"Waiting {delay} seconds before retry...")
                time.sleep(delay)
            else:
                raise Exception("Could not access the report file after multiple attempts")

def process_data(file_path):
    """Process the downloaded DHL report with additional data cleaning"""
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
    """Upload processed data to Google Sheets with additional error handling"""
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
    driver = None
    try:
        print("🚀 Starting DHL report automation process...")
        
        # Step 1: Download report from DHL portal
        driver = setup_chrome_driver()
        
        if not login_to_dhl(driver):
            raise Exception("Login failed")
        
        if not change_language_to_english(driver):
            raise Exception("Language change failed")
        
        if not navigate_to_dashboard(driver):
            raise Exception("Dashboard navigation failed")
        
        if not download_report(driver):
            raise Exception("Report download failed")
        
        # Step 2: Process the downloaded file
        latest_file = get_latest_file(DOWNLOAD_FOLDER)
        processed_df = process_data(latest_file)
        
        # Step 3: Upload to Google Sheets
        upload_to_google_sheets(processed_df)
        
        print("🎉 Complete process finished successfully!")
        
    except Exception as e:
        print(f"❌ Process failed: {str(e)}")
    
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()
