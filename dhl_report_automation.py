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
SERVICE_ACCOUNT_FILE = 'service_account.json'  # Will be created by GitHub Actions
DOWNLOAD_FOLDER = os.getcwd()  # Use current directory for GitHub Actions

def setup_chrome_driver():
    """Setup Chrome driver with necessary options"""
    try:
        from selenium.webdriver.chrome.service import Service
        
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--window-size=1920,1080')
        
        # Set download preferences
        prefs = {
            "download.default_directory": DOWNLOAD_FOLDER,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        # Try different possible ChromeDriver locations
        chromedriver_paths = [
            'chromedriver',                    # System PATH
            '/usr/local/bin/chromedriver',     # Common Linux location
            '/usr/bin/chromedriver',           # Alternative Linux location
        ]
        
        driver = None
        last_error = None
        
        for driver_path in chromedriver_paths:
            try:
                service = Service(executable_path=driver_path)
                driver = webdriver.Chrome(service=service, options=chrome_options)
                print(f"✅ Successfully initialized ChromeDriver from: {driver_path}")
                break
            except Exception as e:
                last_error = e
                continue
        
        if driver is None:
            raise Exception(f"Could not initialize ChromeDriver. Last error: {last_error}")
        
        driver.set_page_load_timeout(30)
        return driver
        
    except Exception as e:
        print(f"❌ Chrome driver setup failed: {str(e)}")
        print("Debug information:")
        try:
            import subprocess
            chrome_version = subprocess.check_output(['chrome', '--version']).decode().strip()
            print(f"Chrome version: {chrome_version}")
            chromedriver_version = subprocess.check_output(['chromedriver', '--version']).decode().strip()
            print(f"ChromeDriver version: {chromedriver_version}")
        except Exception as debug_e:
            print(f"Could not get version info: {str(debug_e)}")
        raise

def login_to_dhl(driver):
    """Login to DHL portal"""
    try:
        print("🔹 Accessing DHL portal...")
        driver.get("https://ecommerceportal.dhl.com/Portal/pages/login/userlogin.xhtml")
        
        # Wait for and fill in login credentials
        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "email1"))
        )
        username_input.send_keys(os.environ.get('DHL_USERNAME', 'truongcongdai4@gmail.com'))
        
        password_input = driver.find_element(By.NAME, "j_password")
        password_input.send_keys(os.environ.get('DHL_PASSWORD', '@Thavi035@'))
        
        login_button = driver.find_element(By.CLASS_NAME, "btn-login")
        login_button.click()
        
        # Wait for login to complete
        time.sleep(5)
        
        # Take screenshot for debugging
        driver.save_screenshot("after_login.png")
        
        print("✅ Login successful!")
        return True
        
    except Exception as e:
        print(f"❌ Login failed: {str(e)}")
        return False

def change_language_to_english(driver):
    """Change portal language to English with improved error handling"""
    try:
        print("🔹 Attempting to change language to English...")
        
        # Try different approaches to find the language selector
        try:
            # First wait for page to fully load
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'ui-selectonemenu')]"))
            )
            
            # Take a screenshot for debugging
            driver.save_screenshot("page_before_language.png")
            
            # Approach 1: Try to find by any language selector by class
            selectors = driver.find_elements(By.XPATH, "//div[contains(@class, 'ui-selectonemenu')]//label")
            if selectors:
                print(f"Found {len(selectors)} potential language selectors")
                for i, selector in enumerate(selectors):
                    print(f"Selector {i}: {selector.text}")
                    if 'language' in selector.get_attribute('class').lower() or i == 0:
                        driver.execute_script("arguments[0].click();", selector)
                        break
            
            time.sleep(3)
            
            # Look for English in the dropdown that appeared
            english_options = driver.find_elements(By.XPATH, "//li[contains(text(), 'English')]")
            if english_options:
                driver.execute_script("arguments[0].click();", english_options[0])
                print("Selected English option")
            else:
                print("English option not found in dropdown")
            
            time.sleep(3)
            print("✅ Language change attempted")
            return True
            
        except Exception as inner_e:
            print(f"⚠️ First approach failed: {str(inner_e)}")
            
            # Approach 2: Try a more general approach - skip language change
            print("Skipping language change and continuing with script...")
            return True
        
    except Exception as e:
        print(f"❌ All language change approaches failed: {str(e)}")
        # Return true to continue with script despite language error
        return True

def navigate_to_dashboard(driver):
    """Navigate to dashboard and set date range"""
    try:
        # Take screenshot before navigation
        driver.save_screenshot("before_dashboard.png")
        
        # Click Dashboard link
        dashboard_link = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Dashboard')]"))
        )
        driver.execute_script("arguments[0].click();", dashboard_link)
        time.sleep(3)
        
        # Take screenshot after dashboard navigation
        driver.save_screenshot("after_dashboard_click.png")
        
        # Set start date (01/01/2025)
        start_date_input = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "dashboardForm:frmDate_input"))
        )
        driver.execute_script("arguments[0].value = '01-01-2025'", start_date_input)
        
        # Set end date (current date)
        end_date_input = driver.find_element(By.ID, "dashboardForm:toDate_input")
        current_date = datetime.now().strftime("%d-%m-%Y")
        driver.execute_script(f"arguments[0].value = '{current_date}'", end_date_input)
        
        time.sleep(2)
        print("✅ Dashboard navigation and date setting complete")
        return True
        
    except Exception as e:
        print(f"❌ Failed to navigate to dashboard: {str(e)}")
        return False

def download_report(driver):
    """Download the Total Received report"""
    try:
        # Take screenshot before download
        driver.save_screenshot("before_download.png")
        
        # Find and click the download button for Total Received
        download_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//td[contains(.,'Total Received')]//img[@id='xlsIcon']"))
        )
        driver.execute_script("arguments[0].click();", download_button)
        
        # Wait for download to complete
        time.sleep(15)  # Increased wait time for download
        
        print("✅ Report download initiated")
        return True
        
    except Exception as e:
        print(f"❌ Failed to download report: {str(e)}")
        return False

def get_latest_file(folder_path, max_attempts=5, delay=2):
    """Get the most recently downloaded file from the specified folder"""
    for attempt in range(max_attempts):
        try:
            files = [
                os.path.join(folder_path, f) for f in os.listdir(folder_path)
                if (f.endswith('.xlsx') or f.endswith('.csv'))
                and not f.startswith('~$')
                and 'total_received_report' in f.lower()
            ]
            
            if not files:
                print(f"No matching files found. Attempt {attempt + 1}/{max_attempts}")
                time.sleep(delay)
                continue
                
            latest_file = max(files, key=os.path.getctime)
            
            with open(latest_file, 'rb') as f:
                pass
                
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
        time.sleep(2)
        
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path, engine='openpyxl')
        
        # Create processed dataframe with cleaned data
        processed_df = pd.DataFrame({
            'Order ID': df['Consignee Name'].str.extract(r'(\d{7})')[0].fillna(''),
            'Tracking Number': df['Tracking ID'].fillna(''),
            'Pickup DateTime': pd.to_datetime(df['Pickup Event DateTime'], errors='coerce').fillna(pd.NaT),
            'Delivery Date': pd.to_datetime(df['Delivery Date'], errors='coerce').fillna(pd.NaT),
            'Status': df['Last Status'].fillna('')
        })
        
        # Convert datetime columns to string format
        processed_df['Pickup DateTime'] = processed_df['Pickup DateTime'].apply(
            lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(x) else '')
        processed_df['Delivery Date'] = processed_df['Delivery Date'].apply(
            lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(x) else '')
        
        # Replace any remaining NaN values with empty strings
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
        
        # MODIFIED: Continue even if language change fails
        change_language_to_english(driver)
        # Don't check the return value or raise an exception
        
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
