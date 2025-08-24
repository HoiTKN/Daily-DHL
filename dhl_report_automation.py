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
import shutil
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

# Download folder - try multiple common paths
DOWNLOAD_FOLDER = os.path.expanduser("~/Downloads")
if not os.path.exists(DOWNLOAD_FOLDER):
    DOWNLOAD_FOLDER = os.getcwd()

# Date range - DD-MM-YYYY format for DHL portal (giữ nguyên theo yêu cầu)
START_DATE_STR = "01-01-2025"
END_DATE_STR = datetime.now().strftime("%d-%m-%Y")

# Credentials
DHL_USERNAME = os.getenv('DHL_USERNAME', 'truongcongdai4@gmail.com')
DHL_PASSWORD = os.getenv('DHL_PASSWORD', '@Thavi035@')

def setup_chrome_driver():
    """Setup Chrome driver with better error handling"""
    try:
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--window-size=1920,1080')
        chrome_options.add_argument('--disable-web-security')
        chrome_options.add_argument('--allow-running-insecure-content')
        chrome_options.add_argument('--ignore-certificate-errors')
        
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

def take_debug_screenshot(driver, filename):
    """Take screenshot for debugging"""
    try:
        screenshot_path = os.path.join(DOWNLOAD_FOLDER, f"debug_{filename}_{int(time.time())}.png")
        driver.save_screenshot(screenshot_path)
        logger.info(f"📸 Screenshot saved: {screenshot_path}")
        return screenshot_path
    except Exception as e:
        logger.warning(f"Could not take screenshot: {str(e)}")
        return None

def wait_and_find_multiple(driver, selectors, timeout=15):
    """Try multiple selectors until one works"""
    for by, value in selectors:
        try:
            element = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((by, value))
            )
            logger.info(f"✅ Found element with: {by}={value}")
            return element
        except TimeoutException:
            logger.warning(f"⚠️ Element not found: {by}={value}")
            continue
    
    logger.error(f"❌ No elements found from {len(selectors)} selectors")
    return None

def login_to_dhl(driver):
    """Login to DHL portal with improved error handling"""
    try:
        logger.info("🔹 Logging into DHL portal...")
        
        # Go to login page
        driver.get("https://ecommerceportal.dhl.com/Portal/pages/login/userlogin.xhtml")
        time.sleep(8)  # Longer wait for page load
        
        # Take screenshot for debugging
        take_debug_screenshot(driver, "login_page")
        
        # Find and fill username with multiple selectors
        username_selectors = [
            (By.ID, "email1"),
            (By.NAME, "email"),
            (By.XPATH, "//input[@type='email']"),
            (By.XPATH, "//input[contains(@placeholder, 'mail')]")
        ]
        
        username_field = wait_and_find_multiple(driver, username_selectors)
        if not username_field:
            return False
        
        username_field.clear()
        username_field.send_keys(DHL_USERNAME)
        logger.info("✅ Username entered")
        time.sleep(2)
        
        # Find and fill password with multiple selectors
        password_selectors = [
            (By.NAME, "j_password"),
            (By.ID, "j_password"),
            (By.XPATH, "//input[@type='password']"),
            (By.XPATH, "//input[contains(@name, 'password')]")
        ]
        
        password_field = wait_and_find_multiple(driver, password_selectors)
        if not password_field:
            return False
        
        password_field.clear()
        password_field.send_keys(DHL_PASSWORD)
        logger.info("✅ Password entered")
        time.sleep(2)
        
        # Click login button with multiple selectors
        login_selectors = [
            (By.CLASS_NAME, "btn-login"),
            (By.XPATH, "//button[contains(text(), 'Login')]"),
            (By.XPATH, "//input[@type='submit']"),
            (By.XPATH, "//button[@type='submit']")
        ]
        
        login_button = wait_and_find_multiple(driver, login_selectors)
        if not login_button:
            return False
        
        login_button.click()
        logger.info("✅ Login button clicked")
        
        # Wait longer for redirect
        time.sleep(15)
        
        # Take screenshot after login
        take_debug_screenshot(driver, "after_login")
        
        # Check if login successful
        current_url = driver.current_url.lower()
        if "login" not in current_url or "dashboard" in current_url or "portal" in current_url:
            logger.info("✅ Login successful")
            return True
        else:
            logger.error(f"❌ Login failed. Current URL: {driver.current_url}")
            return False
            
    except Exception as e:
        logger.error(f"❌ Login error: {str(e)}")
        take_debug_screenshot(driver, "login_error")
        return False

def navigate_to_dashboard(driver):
    """Navigate to dashboard with improved selectors"""
    try:
        logger.info("🔹 Navigating to dashboard...")
        time.sleep(8)
        
        # Take screenshot to see current state
        take_debug_screenshot(driver, "before_dashboard")
        
        # Multiple dashboard selectors to try
        dashboard_selectors = [
            # Original selector
            (By.XPATH, "//span[contains(@class, 'left-navigation-text') and contains(text(), 'Dashboard')]"),
            # Alternative text selectors
            (By.XPATH, "//span[contains(text(), 'Dashboard')]"),
            (By.XPATH, "//a[contains(text(), 'Dashboard')]"),
            (By.LINK_TEXT, "Dashboard"),
            (By.PARTIAL_LINK_TEXT, "Dashboard"),
            # ID selectors
            (By.ID, "dashboard"),
            (By.ID, "Dashboard"),
            # Class selectors
            (By.CLASS_NAME, "dashboard-link"),
            (By.CLASS_NAME, "nav-dashboard"),
            # Generic navigation selectors
            (By.XPATH, "//li[contains(@class, 'nav')]//*[contains(text(), 'Dashboard')]"),
            (By.XPATH, "//div[contains(@class, 'menu')]//*[contains(text(), 'Dashboard')]"),
            (By.XPATH, "//nav//*[contains(text(), 'Dashboard')]"),
            # Fallback - any element with Dashboard text
            (By.XPATH, "//*[contains(text(), 'Dashboard')]"),
        ]
        
        dashboard_link = wait_and_find_multiple(driver, dashboard_selectors, timeout=20)
        
        if dashboard_link:
            # Scroll element into view
            driver.execute_script("arguments[0].scrollIntoView();", dashboard_link)
            time.sleep(2)
            
            # Try to click
            try:
                dashboard_link.click()
            except Exception as e:
                # If normal click fails, try JavaScript click
                logger.warning(f"Normal click failed: {str(e)}, trying JavaScript click")
                driver.execute_script("arguments[0].click();", dashboard_link)
            
            logger.info("✅ Clicked dashboard")
            time.sleep(10)
            
            # Take screenshot after clicking
            take_debug_screenshot(driver, "after_dashboard_click")
            return True
        else:
            logger.error("❌ Dashboard link not found with any selector")
            
            # Log current page source for debugging
            logger.info("📄 Current page title: " + driver.title)
            logger.info("📄 Current URL: " + driver.current_url)
            
            # Try to find any navigation elements
            nav_elements = driver.find_elements(By.XPATH, "//*[contains(@class, 'nav') or contains(@class, 'menu')]")
            logger.info(f"Found {len(nav_elements)} navigation elements")
            
            return False
        
    except Exception as e:
        logger.error(f"❌ Dashboard navigation error: {str(e)}")
        take_debug_screenshot(driver, "dashboard_error")
        return False

def set_date_range_with_retry(driver, max_retries=3):
    """Set date range with retry logic - giữ nguyên từ 01-01-2025 đến hiện tại"""
    for attempt in range(max_retries):
        try:
            logger.info(f"🔹 Setting date range (attempt {attempt + 1}/{max_retries}): {START_DATE_STR} to {END_DATE_STR}")
            logger.info("📅 Note: Đang tải toàn bộ data từ đầu năm - có thể mất nhiều thời gian")
            
            # Wait for page to load completely
            time.sleep(5)
            
            # Take screenshot
            take_debug_screenshot(driver, f"date_range_attempt_{attempt + 1}")
            
            # Set start date
            from_success = set_datepicker_value(driver, "dashboardForm:frmDate_input", START_DATE_STR)
            time.sleep(2)
            
            # Set end date  
            to_success = set_datepicker_value(driver, "dashboardForm:toDate_input", END_DATE_STR)
            time.sleep(2)
            
            if from_success and to_success:
                logger.info("✅ Date range set successfully")
                return True
            elif attempt < max_retries - 1:
                logger.warning(f"⚠️ Date range setting failed, retrying...")
                time.sleep(3)
            
        except Exception as e:
            logger.error(f"❌ Error setting date range (attempt {attempt + 1}): {str(e)}")
            if attempt < max_retries - 1:
                time.sleep(5)
    
    logger.warning("⚠️ Date range setting failed after all retries, continuing anyway")
    return False

def set_datepicker_value(driver, element_id, date_value):
    """Set value for datepicker widget using JavaScript with better error handling"""
    try:
        logger.info(f"🗓️ Setting {element_id} to {date_value}")
        
        # Enhanced JavaScript to handle various datepicker implementations
        script = f"""
        var element = document.getElementById('{element_id}');
        if (element) {{
            // Clear existing value
            element.value = '';
            element.focus();
            
            // Set the value
            element.value = '{date_value}';
            
            // Trigger comprehensive events
            var events = ['input', 'change', 'blur', 'keyup', 'keydown'];
            events.forEach(function(eventType) {{
                var event = new Event(eventType, {{ bubbles: true, cancelable: true }});
                element.dispatchEvent(event);
            }});
            
            // Try PrimeFaces specific events
            if (window.PrimeFaces && PrimeFaces.widget) {{
                var widgetVar = element.getAttribute('data-widget') || element.id.replace(':', '_');
                var widget = PF(widgetVar);
                if (widget && widget.setDate) {{
                    try {{
                        widget.setDate(new Date('{date_value}'.split('-').reverse().join('/')));
                    }} catch(e) {{ console.log('PrimeFaces widget error:', e); }}
                }}
            }}
            
            // Try jQuery events if available
            if (window.jQuery) {{
                jQuery(element).trigger('change').trigger('blur');
                if (jQuery.fn.datepicker) {{
                    try {{
                        jQuery(element).datepicker('setDate', '{date_value}');
                    }} catch(e) {{ console.log('jQuery datepicker error:', e); }}
                }}
            }}
            
            return element.value;
        }} else {{
            return 'ELEMENT_NOT_FOUND';
        }}
        """
        
        result = driver.execute_script(script)
        
        if result == 'ELEMENT_NOT_FOUND':
            logger.error(f"❌ Element {element_id} not found")
            return False
        
        logger.info(f"✅ Set {element_id} value: {result}")
        return result == date_value or len(result) > 0
        
    except Exception as e:
        logger.error(f"❌ Error setting {element_id}: {str(e)}")
        return False

def click_generate_button(driver):
    """Click the GENERATE button with multiple selectors"""
    try:
        logger.info("🔹 Looking for GENERATE button...")
        
        # Wait a bit for any UI updates
        time.sleep(5)
        
        # Multiple selectors for GENERATE button
        generate_selectors = [
            (By.XPATH, "//span[contains(@class, 'ui-button-text') and contains(text(), 'GENERATE')]"),
            (By.XPATH, "//button[contains(text(), 'GENERATE')]"),
            (By.XPATH, "//input[@value='GENERATE']"),
            (By.XPATH, "//span[contains(text(), 'GENERATE')]"),
            (By.XPATH, "//*[contains(text(), 'Generate')]"),
            (By.XPATH, "//*[contains(text(), 'generate')]"),
            (By.ID, "generateBtn"),
            (By.ID, "generate"),
            (By.CLASS_NAME, "generate-btn"),
        ]
        
        generate_button = wait_and_find_multiple(driver, generate_selectors, timeout=20)
        
        if generate_button:
            # Scroll into view
            driver.execute_script("arguments[0].scrollIntoView();", generate_button)
            time.sleep(2)
            
            # Try clicking
            try:
                generate_button.click()
            except Exception as e:
                logger.warning(f"Normal click failed: {str(e)}, trying JavaScript click")
                driver.execute_script("arguments[0].click();", generate_button)
            
            logger.info("✅ Clicked GENERATE button")
            time.sleep(15)  # Wait longer for large dataset (8 months of data)
            
            take_debug_screenshot(driver, "after_generate")
            return True
        else:
            logger.error("❌ GENERATE button not found")
            take_debug_screenshot(driver, "generate_not_found")
            return False
            
    except Exception as e:
        logger.error(f"❌ Error clicking GENERATE: {str(e)}")
        return False

def download_report(driver):
    """Download the report by clicking Excel icon with better error handling"""
    try:
        logger.info("🔹 Looking for download icon...")
        
        # Clear old files first
        clear_download_folder()
        
        # Wait for data to be generated
        time.sleep(8)
        
        # Multiple selectors for download icon
        download_selectors = [
            (By.ID, "xlsIcon"),
            (By.XPATH, "//img[contains(@src, 'download_Pixel_30.png')]"),
            (By.XPATH, "//img[contains(@src, 'excel')]"),
            (By.XPATH, "//img[contains(@id, 'xls')]"),
            (By.XPATH, "//*[contains(@title, 'Excel')]"),
            (By.XPATH, "//*[contains(@alt, 'Excel')]"),
            (By.CLASS_NAME, "excel-icon"),
            (By.CLASS_NAME, "download-icon"),
            (By.XPATH, "//a[contains(@href, 'excel') or contains(@href, 'download')]"),
        ]
        
        download_icon = wait_and_find_multiple(driver, download_selectors, timeout=20)
        
        if download_icon:
            # Scroll into view
            driver.execute_script("arguments[0].scrollIntoView();", download_icon)
            time.sleep(2)
            
            # Try clicking
            try:
                download_icon.click()
            except Exception as e:
                logger.warning(f"Normal click failed: {str(e)}, trying JavaScript click")
                driver.execute_script("arguments[0].click();", download_icon)
            
            logger.info("✅ Clicked download icon")
            logger.info("⏳ Waiting for file download (45 seconds expected for large dataset)...")
            time.sleep(45)  # Wait longer for large file download (8 months of data)
            
            # Check if file was downloaded
            if check_for_new_download():
                logger.info("✅ File downloaded successfully")
                return True
            else:
                logger.warning("⚠️ No file downloaded, checking alternative download paths...")
                return check_alternative_download_paths()
        else:
            logger.error("❌ Download icon not found")
            take_debug_screenshot(driver, "download_not_found")
            return False
            
    except Exception as e:
        logger.error(f"❌ Error downloading report: {str(e)}")
        return False

def clear_download_folder():
    """Clear old download files from multiple possible locations"""
    try:
        clear_paths = [
            DOWNLOAD_FOLDER,
            os.path.expanduser("~/Downloads"),
            os.path.expanduser("~/Desktop")
        ]
        
        total_files_removed = 0
        
        for path in clear_paths:
            if os.path.exists(path):
                try:
                    files_removed = 0
                    for filename in os.listdir(path):
                        if (filename.endswith(('.xlsx', '.csv', '.xls')) and 
                            not filename.startswith('~')):
                            file_path = os.path.join(path, filename)
                            os.remove(file_path)
                            files_removed += 1
                    
                    if files_removed > 0:
                        logger.info(f"✅ Cleared {files_removed} Excel/CSV files from {path}")
                        total_files_removed += files_removed
                        
                except Exception as e:
                    logger.warning(f"Could not clear {path}: {str(e)}")
        
        if total_files_removed > 0:
            logger.info(f"✅ Total files cleared: {total_files_removed}")
            
    except Exception as e:
        logger.warning(f"Could not clear download folders: {str(e)}")

def check_for_new_download():
    """Check if new files were downloaded"""
    try:
        files = [f for f in os.listdir(DOWNLOAD_FOLDER) 
                if f.endswith(('.xlsx', '.csv', '.xls')) and not f.startswith('~')]
        
        if files:
            logger.info(f"Found {len(files)} files in download folder")
            recent_files = []
            current_time = time.time()
            
            for file in files:
                file_path = os.path.join(DOWNLOAD_FOLDER, file)
                file_time = os.path.getctime(file_path)
                
                if current_time - file_time < 180:  # 3 minutes = 180 seconds
                    recent_files.append(file)
                    logger.info(f"Recent file found: {file}")
            
            return len(recent_files) > 0
        
        return False
    except Exception as e:
        logger.warning(f"Error checking downloads: {str(e)}")
        return False

def check_alternative_download_paths():
    """Check for downloads in alternative paths"""
    try:
        download_paths = [
            os.path.expanduser("~/Downloads"),
            os.path.expanduser("~/Desktop"),
            os.path.join(os.path.expanduser("~"), "Downloads"),
            "/tmp",
            os.getcwd()
        ]
        
        logger.info("🔍 Checking alternative download paths...")
        
        for path in download_paths:
            if os.path.exists(path):
                logger.info(f"Checking: {path}")
                files = [f for f in os.listdir(path) 
                        if f.endswith(('.xlsx', '.csv', '.xls')) and not f.startswith('~')]
                
                if files:
                    recent_files = []
                    current_time = time.time()
                    
                    for file in files:
                        file_path = os.path.join(path, file)
                        file_time = os.path.getctime(file_path)
                        
                        if current_time - file_time < 300:  # 5 minutes
                            recent_files.append(file_path)
                            logger.info(f"Found recent file: {file} in {path}")
                    
                    if recent_files:
                        latest_file = max(recent_files, key=os.path.getctime)
                        destination = os.path.join(DOWNLOAD_FOLDER, os.path.basename(latest_file))
                        
                        shutil.copy2(latest_file, destination)
                        logger.info(f"✅ Copied file from {latest_file} to {destination}")
                        return True
        
        logger.warning("❌ No recent files found in any download path")
        return False
        
    except Exception as e:
        logger.error(f"Error checking alternative paths: {str(e)}")
        return False

def get_latest_file(folder_path, max_attempts=8, delay=5):
    """Get the latest downloaded file with more attempts"""
    logger.info(f"🔍 Looking for downloaded files in: {folder_path}")
    
    search_paths = [
        folder_path,
        os.path.expanduser("~/Downloads"),
        os.path.expanduser("~/Desktop"),
        "/tmp"
    ]
    
    for attempt in range(max_attempts):
        try:
            all_files = []
            
            for path in search_paths:
                if os.path.exists(path):
                    path_files = [
                        os.path.join(path, f) for f in os.listdir(path)
                        if (f.endswith('.xlsx') or f.endswith('.csv') or f.endswith('.xls'))
                        and not f.startswith('~$')
                    ]
                    all_files.extend(path_files)
            
            if not all_files:
                logger.info(f"No Excel/CSV files found. Attempt {attempt + 1}/{max_attempts}")
                time.sleep(delay)
                continue
            
            latest_file = max(all_files, key=os.path.getctime)
            file_size = os.path.getsize(latest_file)
            
            logger.info(f"Found file: {latest_file} (Size: {file_size} bytes)")
            
            if file_size > 1000:  # Minimum 1KB file
                if os.path.dirname(latest_file) != folder_path:
                    destination = os.path.join(folder_path, os.path.basename(latest_file))
                    shutil.copy2(latest_file, destination)
                    logger.info(f"Copied file to working directory: {destination}")
                    latest_file = destination
                
                logger.info(f"✅ Valid file found: {latest_file}")
                return latest_file
            
            time.sleep(delay)
                
        except Exception as e:
            logger.warning(f"Error checking files: {str(e)}")
            time.sleep(delay)
    
    logger.warning("❌ No valid file found")
    return None

def process_data(file_path):
    """Process downloaded data with better filtering - handles large dataset từ 01-01-2025"""
    if file_path is None:
        logger.warning("No file to process, creating empty DataFrame")
        return create_empty_data()
    
    logger.info(f"🔹 Processing file: {file_path}")
    logger.info("📊 Processing large dataset từ đầu năm - có thể mất vài phút...")
    
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
        
        # Filter out rows with missing critical information
        logger.info("🔹 Filtering data for completeness...")
        
        # Create processed DataFrame
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
        
        # Filter out only truly invalid rows (more lenient for large dataset)
        initial_count = len(processed_df)
        
        # Keep rows that have at least a valid tracking number (more lenient filtering)
        mask = (
            (processed_df['Tracking Number'].str.len() >= 10) &  # Valid tracking number length
            (~processed_df['Tracking Number'].str.contains(r'^\s*
        
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
        
        # Reset index
        processed_df = processed_df.reset_index(drop=True)
        
        logger.info(f"✅ Processing completed. Final shape: {processed_df.shape}")
        
        if len(processed_df) > 0:
            logger.info("Final sample data:")
            logger.info(processed_df.head(3).to_string())
        
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
    """Main execution function with enhanced error handling"""
    driver = None
    try:
        logger.info("🚀 Starting DHL report automation...")
        logger.info(f"📅 Date range: {START_DATE_STR} to {END_DATE_STR} (giữ nguyên theo yêu cầu)")
        logger.info(f"📁 Download folder: {DOWNLOAD_FOLDER}")
        
        # Setup driver
        driver = setup_chrome_driver()
        
        # Step 1: Login with retry
        login_success = False
        for attempt in range(3):
            logger.info(f"🔹 Login attempt {attempt + 1}/3")
            if login_to_dhl(driver):
                login_success = True
                break
            elif attempt < 2:
                logger.warning("⚠️ Login failed, retrying...")
                time.sleep(10)
        
        if not login_success:
            logger.error("❌ Login failed after all attempts")
            upload_to_google_sheets(create_empty_data())
            return
        
        # Step 2: Navigate to dashboard with retry
        dashboard_success = False
        for attempt in range(3):
            logger.info(f"🔹 Dashboard navigation attempt {attempt + 1}/3")
            if navigate_to_dashboard(driver):
                dashboard_success = True
                break
            elif attempt < 2:
                logger.warning("⚠️ Dashboard navigation failed, retrying...")
                time.sleep(10)
        
        if not dashboard_success:
            logger.error("❌ Dashboard navigation failed after all attempts")
            upload_to_google_sheets(create_empty_data())
            return
        
        # Step 3: Set date range (continue even if this fails)
        # Note: Giữ nguyên date range từ 01-01-2025 theo yêu cầu user
        set_date_range_with_retry(driver)
        
        # Step 4: Click generate with retry
        generate_success = False
        for attempt in range(3):
            logger.info(f"🔹 Generate button click attempt {attempt + 1}/3")
            if click_generate_button(driver):
                generate_success = True
                break
            elif attempt < 2:
                logger.warning("⚠️ Generate button click failed, retrying...")
                time.sleep(10)
        
        if not generate_success:
            logger.error("❌ Generate button click failed after all attempts")
            upload_to_google_sheets(create_empty_data())
            return
        
        # Step 5: Download report with retry
        download_success = False
        for attempt in range(2):
            logger.info(f"🔹 Download attempt {attempt + 1}/2")
            if download_report(driver):
                download_success = True
                break
            elif attempt < 1:
                logger.warning("⚠️ Download failed, retrying...")
                time.sleep(15)
        
        if not download_success:
            logger.error("❌ Download failed after all attempts")
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
                # Take final screenshot for debugging
                take_debug_screenshot(driver, "final_state")
                driver.quit()
            except:
                pass

if __name__ == "__main__":
    main(), na=True))  # Not empty/whitespace
        )
        
        # Optional: Also keep rows that have meaningful Order ID even with shorter tracking
        mask_alternative = (
            (processed_df['Order ID'].str.len() >= 7) &  # Valid order ID
            (processed_df['Tracking Number'].str.len() >= 5)  # Some tracking info
        )
        
        # Combine conditions - keep if either condition is true
        final_mask = mask | mask_alternative
        
        processed_df = processed_df[final_mask].copy()
        
        filtered_count = len(processed_df)
        logger.info(f"🔹 Filtered from {initial_count} to {filtered_count} rows")
        logger.info(f"📊 Removed {initial_count - filtered_count} rows with invalid/missing tracking info")
        
        # Log some statistics about the data
        if len(processed_df) > 0:
            complete_orders = processed_df[processed_df['Order ID'].str.len() > 0].shape[0] 
            delivered_orders = processed_df[processed_df['Status'].str.contains('DELIVERED', na=False)].shape[0]
            logger.info(f"📈 Data stats: {complete_orders} có Order ID, {delivered_orders} đã delivered")
        
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
        
        # Reset index
        processed_df = processed_df.reset_index(drop=True)
        
        logger.info(f"✅ Processing completed. Final shape: {processed_df.shape}")
        
        if len(processed_df) > 0:
            logger.info("Final sample data:")
            logger.info(processed_df.head(3).to_string())
        
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
    """Main execution function with enhanced error handling"""
    driver = None
    try:
        logger.info("🚀 Starting DHL report automation...")
        logger.info(f"📅 Date range: {START_DATE_STR} to {END_DATE_STR} (giữ nguyên theo yêu cầu)")
        logger.info(f"📁 Download folder: {DOWNLOAD_FOLDER}")
        
        # Setup driver
        driver = setup_chrome_driver()
        
        # Step 1: Login with retry
        login_success = False
        for attempt in range(3):
            logger.info(f"🔹 Login attempt {attempt + 1}/3")
            if login_to_dhl(driver):
                login_success = True
                break
            elif attempt < 2:
                logger.warning("⚠️ Login failed, retrying...")
                time.sleep(10)
        
        if not login_success:
            logger.error("❌ Login failed after all attempts")
            upload_to_google_sheets(create_empty_data())
            return
        
        # Step 2: Navigate to dashboard with retry
        dashboard_success = False
        for attempt in range(3):
            logger.info(f"🔹 Dashboard navigation attempt {attempt + 1}/3")
            if navigate_to_dashboard(driver):
                dashboard_success = True
                break
            elif attempt < 2:
                logger.warning("⚠️ Dashboard navigation failed, retrying...")
                time.sleep(10)
        
        if not dashboard_success:
            logger.error("❌ Dashboard navigation failed after all attempts")
            upload_to_google_sheets(create_empty_data())
            return
        
        # Step 3: Set date range (continue even if this fails)
        set_date_range_with_retry(driver)
        
        # Step 4: Click generate with retry
        generate_success = False
        for attempt in range(3):
            logger.info(f"🔹 Generate button click attempt {attempt + 1}/3")
            if click_generate_button(driver):
                generate_success = True
                break
            elif attempt < 2:
                logger.warning("⚠️ Generate button click failed, retrying...")
                time.sleep(10)
        
        if not generate_success:
            logger.error("❌ Generate button click failed after all attempts")
            upload_to_google_sheets(create_empty_data())
            return
        
        # Step 5: Download report with retry
        download_success = False
        for attempt in range(2):
            logger.info(f"🔹 Download attempt {attempt + 1}/2")
            if download_report(driver):
                download_success = True
                break
            elif attempt < 1:
                logger.warning("⚠️ Download failed, retrying...")
                time.sleep(15)
        
        if not download_success:
            logger.error("❌ Download failed after all attempts")
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
                # Take final screenshot for debugging
                take_debug_screenshot(driver, "final_state")
                driver.quit()
            except:
                pass

if __name__ == "__main__":
    main()
