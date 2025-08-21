from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd
import os
from datetime import datetime
import time
import numpy as np
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Constants
GOOGLE_SHEET_ID = "16blaF86ky_4Eu4BK8AyXajohzpMsSyDaoPPKGVDYqWw"
SHEET_NAME = "DHL"
SERVICE_ACCOUNT_FILE = 'service_account.json'
DOWNLOAD_FOLDER = os.getcwd()

# Date range - DD-MM-YYYY format for DHL portal
START_DATE = "01-01-2025"
END_DATE = datetime.now().strftime("%d-%m-%Y")

# Credentials
DHL_USERNAME = os.getenv('DHL_USERNAME', 'truongcongdai4@gmail.com')
DHL_PASSWORD = os.getenv('DHL_PASSWORD', '@Thavi035@')

def setup_chrome_driver():
    """Setup Chrome driver"""
    try:
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--window-size=1920,1080')
        
        prefs = {
            "download.default_directory": DOWNLOAD_FOLDER,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        # Try to find ChromeDriver
        chromedriver_paths = ['/usr/bin/chromedriver', '/usr/local/bin/chromedriver', 'chromedriver']
        
        for driver_path in chromedriver_paths:
            try:
                if os.path.exists(driver_path) or driver_path == 'chromedriver':
                    service = Service(executable_path=driver_path)
                    driver = webdriver.Chrome(service=service, options=chrome_options)
                    logger.info(f"✅ ChromeDriver initialized from: {driver_path}")
                    return driver
            except Exception as e:
                logger.warning(f"Failed to use {driver_path}: {str(e)}")
                continue
        
        raise Exception("Could not initialize ChromeDriver")
        
    except Exception as e:
        logger.error(f"❌ Chrome driver setup failed: {str(e)}")
        raise

def wait_and_find(driver, by, value, timeout=15):
    """Wait for element and return it"""
    try:
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
    except TimeoutException:
        logger.error(f"Element not found: {by}={value}")
        return None

def login_to_dhl(driver):
    """Login to DHL portal"""
    try:
        logger.info("🔹 Logging into DHL portal...")
        
        # Go to login page
        driver.get("https://ecommerceportal.dhl.com/Portal/pages/login/userlogin.xhtml")
        time.sleep(5)
        
        # Find and fill username
        username_field = wait_and_find(driver, By.ID, "email1")
        if not username_field:
            return False
        
        username_field.clear()
        username_field.send_keys(DHL_USERNAME)
        logger.info("✅ Username entered")
        
        # Find and fill password
        password_field = wait_and_find(driver, By.NAME, "j_password")
        if not password_field:
            return False
        
        password_field.clear()
        password_field.send_keys(DHL_PASSWORD)
        logger.info("✅ Password entered")
        
        # Click login button
        login_button = wait_and_find(driver, By.CLASS_NAME, "btn-login")
        if not login_button:
            return False
        
        login_button.click()
        logger.info("✅ Login button clicked")
        
        # Wait for redirect
        time.sleep(10)
        
        # Check if login successful
        if "login" not in driver.current_url.lower():
            logger.info("✅ Login successful")
            return True
        else:
            logger.error("❌ Login failed")
            return False
            
    except Exception as e:
        logger.error(f"❌ Login error: {str(e)}")
        return False

def navigate_to_dashboard(driver):
    """Navigate to dashboard"""
    try:
        logger.info("🔹 Navigating to dashboard...")
        time.sleep(5)
        
        # Find dashboard link
        dashboard_link = wait_and_find(driver, By.XPATH, "//span[contains(@class, 'left-navigation-text') and contains(text(), 'Dashboard')]")
        if not dashboard_link:
            logger.error("❌ Dashboard link not found")
            return False
        
        dashboard_link.click()
        logger.info("✅ Clicked dashboard")
        time.sleep(8)
        
        return True
        
    except Exception as e:
        logger.error(f"❌ Dashboard navigation error: {str(e)}")
        return False

def set_datepicker_value(driver, element_id, date_value):
    """Set value for datepicker widget using JavaScript"""
    try:
        logger.info(f"🗓️ Setting {element_id} to {date_value}")
        
        # Use JavaScript to set datepicker value
        script = f"""
        var element = document.getElementById('{element_id}');
        if (element) {{
            // Set the value
            element.value = '{date_value}';
            
            // Trigger events to notify the datepicker
            var events = ['input', 'change', 'blur'];
            events.forEach(function(eventType) {{
                var event = new Event(eventType, {{ bubbles: true }});
                element.dispatchEvent(event);
            }});
            
            // Try jQuery events if available
            if (window.jQuery && jQuery.fn.datepicker) {{
                jQuery(element).datepicker('setDate', '{date_value}');
                jQuery(element).trigger('change');
            }}
            
            return element.value;
        }}
        return null;
        """
        
        result = driver.execute_script(script)
        logger.info(f"✅ Set {element_id} value: {result}")
        
        return result == date_value
        
    except Exception as e:
        logger.error(f"❌ Error setting {element_id}: {str(e)}")
        return False

def set_date_range(driver):
    """Set date range on dashboard"""
    try:
        logger.info(f"🔹 Setting date range: {START_DATE} to {END_DATE}")
        
        # Wait for page to load completely
        time.sleep(5)
        
        # Set start date
        from_success = set_datepicker_value(driver, "dashboardForm:frmDate_input", START_DATE)
        time.sleep(1)
        
        # Set end date  
        to_success = set_datepicker_value(driver, "dashboardForm:toDate_input", END_DATE)
        time.sleep(1)
        
        if from_success and to_success:
            logger.info("✅ Date range set successfully")
            return True
        else:
            logger.warning("⚠️ Date range setting may have failed")
            return False
            
    except Exception as e:
        logger.error(f"❌ Error setting date range: {str(e)}")
        return False

def click_generate_button(driver):
    """Click the GENERATE button"""
    try:
        logger.info("🔹 Looking for GENERATE button...")
        
        # Wait a bit for any UI updates
        time.sleep(3)
        
        # Find GENERATE button
        generate_button = wait_and_find(driver, By.XPATH, "//span[contains(@class, 'ui-button-text') and contains(text(), 'GENERATE')]")
        if not generate_button:
            # Try alternative selector
            generate_button = wait_and_find(driver, By.XPATH, "//button[contains(text(), 'GENERATE')] | //input[@value='GENERATE']")
        
        if generate_button:
            generate_button.click()
            logger.info("✅ Clicked GENERATE button")
            time.sleep(8)  # Wait for data to load
            return True
        else:
            logger.error("❌ GENERATE button not found")
            return False
            
    except Exception as e:
        logger.error(f"❌ Error clicking GENERATE: {str(e)}")
        return False

def download_report(driver):
    """Download the report by clicking Excel icon"""
    try:
        logger.info("🔹 Looking for download icon...")
        
        # Clear old files first
        clear_download_folder()
        
        # Find Excel download icon
        download_icon = wait_and_find(driver, By.ID, "xlsIcon")
        if not download_icon:
            # Try alternative selectors
            download_icon = wait_and_find(driver, By.XPATH, "//img[contains(@src, 'download_Pixel_30.png')] | //img[contains(@src, 'excel')] | //img[contains(@id, 'xls')]")
        
        if download_icon:
            download_icon.click()
            logger.info("✅ Clicked download icon")
            time.sleep(10)  # Wait for download
            
            # Check if file was downloaded
            if check_for_new_download():
                logger.info("✅ File downloaded successfully")
                return True
            else:
                logger.warning("⚠️ No file downloaded")
                return False
        else:
            logger.error("❌ Download icon not found")
            return False
            
    except Exception as e:
        logger.error(f"❌ Error downloading report: {str(e)}")
        return False

def clear_download_folder():
    """Clear old download files"""
    try:
        files_removed = 0
        for filename in os.listdir(DOWNLOAD_FOLDER):
            if filename.endswith(('.xlsx', '.csv', '.xls')) and not filename.startswith('~'):
                file_path = os.path.join(DOWNLOAD_FOLDER, filename)
                os.remove(file_path)
                files_removed += 1
        if files_removed > 0:
            logger.info(f"✅ Cleared {files_removed} old files")
    except Exception as e:
        logger.warning(f"Could not clear download folder: {str(e)}")

def check_for_new_download():
    """Check if new files were downloaded"""
    try:
        files = [f for f in os.listdir(DOWNLOAD_FOLDER) 
                if f.endswith(('.xlsx', '.csv', '.xls')) and not f.startswith('~')]
        return len(files) > 0
    except:
        return False

def get_latest_file(folder_path, max_attempts=10, delay=3):
    """Get the latest downloaded file"""
    logger.info(f"🔍 Looking for downloaded files...")
    
    for attempt in range(max_attempts):
        try:
            files = [
                os.path.join(folder_path, f) for f in os.listdir(folder_path)
                if (f.endswith('.xlsx') or f.endswith('.csv') or f.endswith('.xls'))
                and not f.startswith('~$')
            ]
            
            if not files:
                logger.info(f"No files found. Attempt {attempt + 1}/{max_attempts}")
                time.sleep(delay)
                continue
            
            latest_file = max(files, key=os.path.getctime)
            file_size = os.path.getsize(latest_file)
            
            logger.info(f"Found file: {latest_file} (Size: {file_size} bytes)")
            
            if file_size > 0:
                logger.info(f"✅ Valid file found: {latest_file}")
                return latest_file
            
            time.sleep(delay)
                
        except Exception as e:
            logger.warning(f"Error checking files: {str(e)}")
            time.sleep(delay)
    
    logger.warning("❌ No valid file found")
    return None

def process_data(file_path):
    """Process downloaded data"""
    if file_path is None:
        logger.warning("No file to process, creating empty DataFrame")
        return create_empty_data()
    
    logger.info(f"🔹 Processing file: {file_path}")
    
    try:
        # Read file
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path, encoding='utf-8')
        else:
            df = pd.read_excel(file_path, engine='openpyxl')
        
        logger.info(f"File loaded. Shape: {df.shape}")
        logger.info(f"Columns: {df.columns.tolist()}")
        
        if len(df) > 0:
            logger.info("Sample data:")
            logger.info(df.head(3).to_string())
        
        # Create processed DataFrame with flexible mapping
        processed_df = pd.DataFrame()
        
        # Map Order ID from Consignee Name
        if 'Consignee Name' in df.columns:
            processed_df['Order ID'] = df['Consignee Name'].astype(str).str.extract(r'(\d{7})')[0].fillna('')
        else:
            processed_df['Order ID'] = ''
        
        # Map Tracking Number
        tracking_cols = ['Tracking ID', 'Tracking Number', 'AWB', 'Waybill Number']
        for col in tracking_cols:
            if col in df.columns:
                processed_df['Tracking Number'] = df[col].fillna('')
                break
        else:
            processed_df['Tracking Number'] = ''
        
        # Map Pickup DateTime
        pickup_cols = ['Pickup Event DateTime', 'Pickup Date', 'Collection Date', 'Ship Date']
        for col in pickup_cols:
            if col in df.columns:
                processed_df['Pickup DateTime'] = pd.to_datetime(df[col], errors='coerce')
                break
        else:
            processed_df['Pickup DateTime'] = pd.NaT
        
        # Map Delivery Date
        delivery_cols = ['Delivery Date', 'Delivered Date', 'POD Date']
        for col in delivery_cols:
            if col in df.columns:
                processed_df['Delivery Date'] = pd.to_datetime(df[col], errors='coerce')
                break
        else:
            processed_df['Delivery Date'] = pd.NaT
        
        # Map Status
        status_cols = ['Last Status', 'Status', 'Current Status', 'Shipment Status']
        for col in status_cols:
            if col in df.columns:
                processed_df['Status'] = df[col].fillna('')
                break
        else:
            processed_df['Status'] = ''
        
        # Sort by Pickup DateTime (newest first)
        if not processed_df['Pickup DateTime'].isna().all():
            processed_df = processed_df.sort_values('Pickup DateTime', ascending=False, na_position='last')
        
        # Convert datetime to string
        processed_df['Pickup DateTime'] = processed_df['Pickup DateTime'].apply(
            lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(x) else ''
        )
        processed_df['Delivery Date'] = processed_df['Delivery Date'].apply(
            lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(x) else ''
        )
        
        # Clean data
        processed_df = processed_df.replace({np.nan: '', 'NaT': '', None: ''})
        
        logger.info(f"✅ Processing completed. Final shape: {processed_df.shape}")
        return processed_df
        
    except Exception as e:
        logger.error(f"❌ Error processing data: {str(e)}")
        return create_empty_data()

def create_empty_data():
    """Create empty DataFrame structure"""
    return pd.DataFrame({
        'Order ID': [],
        'Tracking Number': [],
        'Pickup DateTime': [],
        'Delivery Date': [],
        'Status': []
    })

def upload_to_google_sheets(df):
    """Upload data to Google Sheets"""
    logger.info("🔹 Uploading to Google Sheets...")
    
    try:
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            logger.error(f"❌ Service account file not found: {SERVICE_ACCOUNT_FILE}")
            return False
        
        creds = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        service = build("sheets", "v4", credentials=creds)
        
        # Prepare data
        headers = df.columns.tolist()
        data = df.astype(str).values.tolist()
        values = [headers] + data
        
        logger.info(f"Uploading {len(data)} rows")
        
        # Clear existing content
        service.spreadsheets().values().clear(
            spreadsheetId=GOOGLE_SHEET_ID,
            range=f"{SHEET_NAME}!A1:Z1000"
        ).execute()
        logger.info("✅ Cleared existing content")
        
        # Upload new data
        service.spreadsheets().values().update(
            spreadsheetId=GOOGLE_SHEET_ID,
            range=f"{SHEET_NAME}!A1",
            valueInputOption="USER_ENTERED",
            body={'values': values}
        ).execute()
        
        logger.info("✅ Data uploaded successfully to Google Sheets")
        return True
        
    except Exception as e:
        logger.error(f"❌ Upload failed: {str(e)}")
        return False

def main():
    """Main execution function"""
    driver = None
    try:
        logger.info("🚀 Starting DHL report automation...")
        logger.info(f"📅 Date range: {START_DATE} to {END_DATE}")
        
        # Setup driver
        driver = setup_chrome_driver()
        
        # Step 1: Login
        if not login_to_dhl(driver):
            logger.error("❌ Login failed")
            upload_to_google_sheets(create_empty_data())
            return
        
        # Step 2: Navigate to dashboard
        if not navigate_to_dashboard(driver):
            logger.error("❌ Dashboard navigation failed")
            upload_to_google_sheets(create_empty_data())
            return
        
        # Step 3: Set date range
        set_date_range(driver)  # Continue even if this fails
        
        # Step 4: Click generate
        if not click_generate_button(driver):
            logger.error("❌ Generate button click failed")
            upload_to_google_sheets(create_empty_data())
            return
        
        # Step 5: Download report
        if not download_report(driver):
            logger.error("❌ Download failed")
            upload_to_google_sheets(create_empty_data())
            return
        
        # Step 6: Process data
        latest_file = get_latest_file(DOWNLOAD_FOLDER)
        processed_df = process_data(latest_file)
        
        # Step 7: Upload to sheets
        upload_to_google_sheets(processed_df)
        
        logger.info("🎉 Process completed successfully!")
        
    except Exception as e:
        logger.error(f"❌ Main process failed: {str(e)}")
        upload_to_google_sheets(create_empty_data())
    
    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass

if __name__ == "__main__":
    main()
