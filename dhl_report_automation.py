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
DOWNLOAD_FOLDER = os.path.expanduser("~/Downloads")  # User's Downloads folder
if not os.path.exists(DOWNLOAD_FOLDER):
    DOWNLOAD_FOLDER = os.getcwd()  # Fallback to current directory

# Date range - DD-MM-YYYY format for DHL portal
START_DATE = "01-02-2026"
END_DATE = datetime.now().strftime("%d-%m-%Y")

# Credentials
DHL_USERNAME = os.getenv('DHL_USERNAME', 'truongcongdai4@gmail.com')
DHL_PASSWORD = os.getenv('DHL_PASSWORD', '@Love123123')

def setup_chrome_driver():
    """Setup Chrome driver with anti-detection"""
    try:
        chrome_options = Options()
        chrome_options.add_argument('--headless=new')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--window-size=1920,1080')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
        chrome_options.add_experimental_option('useAutomationExtension', False)

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
                    # Remove webdriver flag to avoid detection
                    driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
                        'source': 'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'
                    })
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
        time.sleep(8)

        logger.info(f"📄 Current URL: {driver.current_url}")
        logger.info(f"📄 Page title: {driver.title}")

        # Debug: log the login form HTML structure
        try:
            form_html = driver.execute_script("""
                var forms = document.querySelectorAll('form');
                var result = [];
                forms.forEach(function(f) {
                    result.push('FORM id=' + f.id + ' action=' + f.action + ' method=' + f.method);
                    var inputs = f.querySelectorAll('input, button, a.btn-login, span.btn-login');
                    inputs.forEach(function(i) {
                        result.push('  ' + i.tagName + ' id=' + i.id + ' name=' + i.name + ' type=' + i.type + ' class=' + i.className + ' onclick=' + i.getAttribute('onclick'));
                    });
                });
                return result.join('\\n');
            """)
            logger.info(f"📄 Form structure:\\n{form_html}")
        except Exception as e:
            logger.warning(f"Could not get form structure: {e}")

        # Find and fill username - try multiple selectors
        username_field = None
        username_selectors = [
            (By.ID, "email1"),
            (By.NAME, "j_username"),
            (By.XPATH, "//input[@type='email']"),
            (By.XPATH, "//input[@type='text' and contains(@id, 'email')]"),
            (By.CSS_SELECTOR, "input[id*='email']"),
            (By.CSS_SELECTOR, "input[name*='username']"),
        ]
        for by, value in username_selectors:
            username_field = wait_and_find(driver, by, value, timeout=5)
            if username_field:
                logger.info(f"✅ Found username field with: {by}={value}")
                break

        if not username_field:
            logger.error("❌ Username field not found")
            logger.info(f"📄 Page source (first 3000 chars): {driver.page_source[:3000]}")
            return False

        username_field.clear()
        username_field.send_keys(DHL_USERNAME)
        logger.info("✅ Username entered")

        # Find and fill password - try multiple selectors
        password_field = None
        password_selectors = [
            (By.NAME, "j_password"),
            (By.ID, "password1"),
            (By.XPATH, "//input[@type='password']"),
            (By.CSS_SELECTOR, "input[type='password']"),
        ]
        for by, value in password_selectors:
            password_field = wait_and_find(driver, by, value, timeout=5)
            if password_field:
                logger.info(f"✅ Found password field with: {by}={value}")
                break

        if not password_field:
            logger.error("❌ Password field not found")
            return False

        password_field.clear()
        password_field.send_keys(DHL_PASSWORD)
        logger.info("✅ Password entered")

        # Strategy 1: Press Enter on password field (most reliable for JSF forms)
        logger.info("🔄 Trying login via Enter key on password field...")
        password_field.send_keys(Keys.RETURN)
        time.sleep(15)

        current_url = driver.current_url.lower()
        logger.info(f"📄 URL after Enter key: {driver.current_url}")

        if "userlogin" not in current_url and "login.xhtml" not in current_url:
            logger.info("✅ Login successful (Enter key)")
            return True

        # Strategy 2: Find and click the actual submit button/link
        logger.info("🔄 Enter key didn't work, trying button click...")
        login_button = None
        login_selectors = [
            (By.CSS_SELECTOR, "a.btn-login"),
            (By.CSS_SELECTOR, "button.btn-login"),
            (By.CLASS_NAME, "btn-login"),
            (By.XPATH, "//a[contains(@class, 'btn-login')]"),
            (By.XPATH, "//input[@type='submit']"),
            (By.XPATH, "//button[@type='submit']"),
        ]
        for by, value in login_selectors:
            login_button = wait_and_find(driver, by, value, timeout=3)
            if login_button:
                tag = login_button.tag_name
                onclick = login_button.get_attribute('onclick') or ''
                href = login_button.get_attribute('href') or ''
                logger.info(f"✅ Found login button: tag={tag}, {by}={value}, onclick={onclick[:100]}, href={href[:100]}")
                break

        if login_button:
            # Try JS click (more reliable for overlays)
            driver.execute_script("arguments[0].click();", login_button)
            logger.info("✅ Login button JS-clicked")
            time.sleep(15)

            current_url = driver.current_url.lower()
            logger.info(f"📄 URL after button click: {driver.current_url}")

            if "userlogin" not in current_url and "login.xhtml" not in current_url:
                logger.info("✅ Login successful (button click)")
                return True

        # Strategy 3: Submit the form directly via JavaScript
        logger.info("🔄 Button click didn't work, trying JS form submit...")
        try:
            driver.execute_script("""
                var form = document.querySelector('form');
                if (form) {
                    // Try clicking any element with btn-login class
                    var btn = document.querySelector('.btn-login');
                    if (btn && btn.onclick) {
                        btn.onclick();
                    } else if (btn) {
                        btn.click();
                    } else {
                        form.submit();
                    }
                }
            """)
            time.sleep(15)

            current_url = driver.current_url.lower()
            logger.info(f"📄 URL after JS submit: {driver.current_url}")

            if "userlogin" not in current_url and "login.xhtml" not in current_url:
                logger.info("✅ Login successful (JS submit)")
                return True
        except Exception as e:
            logger.warning(f"JS submit failed: {e}")

        # Check for error messages on the page
        try:
            error_msgs = driver.execute_script("""
                var msgs = [];
                var elements = document.querySelectorAll('.ui-messages-error, .ui-message-error, .error, .alert-danger, .login-error, [class*=error], [class*=Error]');
                elements.forEach(function(el) {
                    if (el.textContent.trim()) msgs.push(el.textContent.trim());
                });
                return msgs.join(' | ');
            """)
            if error_msgs:
                logger.error(f"❌ Error messages on page: {error_msgs}")
        except Exception:
            pass

        logger.error("❌ Login failed - all strategies exhausted")
        logger.info(f"📄 Page source (first 3000 chars): {driver.page_source[:3000]}")
        return False

    except Exception as e:
        logger.error(f"❌ Login error: {str(e)}")
        return False

def navigate_to_dashboard(driver):
    """Navigate to dashboard - FIX 1: Multiple selectors cho schedule"""
    try:
        logger.info("🔹 Navigating to dashboard...")
        time.sleep(5)
        
        # Try multiple dashboard selectors
        dashboard_selectors = [
            "//span[contains(@class, 'left-navigation-text') and contains(text(), 'Dashboard')]",
            "//span[contains(text(), 'Dashboard')]",
            "//a[contains(text(), 'Dashboard')]",
            "//*[contains(text(), 'Dashboard')]"
        ]
        
        dashboard_link = None
        for selector in dashboard_selectors:
            try:
                dashboard_link = wait_and_find(driver, By.XPATH, selector, timeout=8)
                if dashboard_link:
                    logger.info(f"✅ Found dashboard with selector: {selector}")
                    break
            except:
                continue
        
        if not dashboard_link:
            logger.error("❌ Dashboard link not found")
            return False
        
        # Try clicking with JavaScript if normal click fails
        try:
            dashboard_link.click()
        except:
            logger.info("🔄 Normal click failed, trying JavaScript click")
            driver.execute_script("arguments[0].click();", dashboard_link)
        
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

def clear_download_folder():
    """Clear old download files from multiple possible locations"""
    try:
        # Paths to clear
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
                        # Remove any Excel/CSV files that look like reports
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
            # Check if any file was created recently (last 2 minutes)
            recent_files = []
            current_time = time.time()
            
            for file in files:
                file_path = os.path.join(DOWNLOAD_FOLDER, file)
                file_time = os.path.getctime(file_path)
                
                if current_time - file_time < 120:  # 2 minutes = 120 seconds
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
        # Common download folders to check
        download_paths = [
            os.path.expanduser("~/Downloads"),  # Linux/Mac default
            os.path.expanduser("~/Desktop"),    # Sometimes downloads go here
            os.path.join(os.path.expanduser("~"), "Downloads"),  # Alternative path
            "/tmp",  # Temporary folder
            os.getcwd()  # Current working directory
        ]
        
        logger.info("🔍 Checking alternative download paths...")
        
        for path in download_paths:
            if os.path.exists(path):
                logger.info(f"Checking: {path}")
                files = [f for f in os.listdir(path) 
                        if f.endswith(('.xlsx', '.csv', '.xls')) and not f.startswith('~')]
                
                if files:
                    # Look for recently created files (last 5 minutes)
                    recent_files = []
                    current_time = time.time()
                    
                    for file in files:
                        file_path = os.path.join(path, file)
                        file_time = os.path.getctime(file_path)
                        
                        # File created within last 5 minutes
                        if current_time - file_time < 300:  # 5 minutes = 300 seconds
                            recent_files.append(file_path)
                            logger.info(f"Found recent file: {file} in {path}")
                    
                    if recent_files:
                        # Copy the most recent file to our working directory
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
            logger.info("⏳ Waiting for file download (15-20 seconds expected)...")
            time.sleep(25)  # Wait longer for download (15-20s + buffer)
            
            # Check if file was downloaded in multiple possible locations
            if check_for_new_download():
                logger.info("✅ File downloaded successfully")
                return True
            else:
                logger.warning("⚠️ No file downloaded, checking alternative download paths...")
                return check_alternative_download_paths()
        else:
            logger.error("❌ Download icon not found")
            return False
            
    except Exception as e:
        logger.error(f"❌ Error downloading report: {str(e)}")
        return False

def get_latest_file(folder_path, max_attempts=5, delay=5):
    """Get the latest downloaded file"""
    logger.info(f"🔍 Looking for downloaded files in: {folder_path}")
    
    # Also check common download folders
    search_paths = [
        folder_path,
        os.path.expanduser("~/Downloads"),
        os.path.expanduser("~/Desktop"),
        "/tmp"
    ]
    
    for attempt in range(max_attempts):
        try:
            all_files = []
            
            # Search in all possible paths
            for path in search_paths:
                if os.path.exists(path):
                    path_files = [
                        os.path.join(path, f) for f in os.listdir(path)
                        if (f.endswith('.xlsx') or f.endswith('.csv') or f.endswith('.xls'))
                        and not f.startswith('~$')
                        # Look for any Excel/CSV files
                    ]
                    all_files.extend(path_files)
            
            if not all_files:
                logger.info(f"No Excel/CSV files found. Attempt {attempt + 1}/{max_attempts}")
                time.sleep(delay)
                continue
            
            # Get the most recent file
            latest_file = max(all_files, key=os.path.getctime)
            file_size = os.path.getsize(latest_file)
            
            logger.info(f"Found file: {latest_file} (Size: {file_size} bytes)")
            
            if file_size > 0:
                # Copy to working directory if it's in a different location
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
    """Process downloaded data - FIX 2: Chỉ lấy tracking numbers hợp lệ"""
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
        
        # Map Order ID from Consignee Name (first 7 characters)
        if 'Consignee Name' in df.columns:
            processed_df['Order ID'] = df['Consignee Name'].fillna('').astype(str).str[:7]
        else:
            processed_df['Order ID'] = ''
        
        # Map Tracking Number
        tracking_cols = ['Tracking ID', 'Tracking Number', 'AWB', 'Waybill Number']
        for col in tracking_cols:
            if col in df.columns:
                processed_df['Tracking Number'] = df[col].fillna('').astype(str)
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
                processed_df['Status'] = df[col].fillna('').astype(str)
                break
        else:
            processed_df['Status'] = ''

        # Map Last Failure Reason
        failure_cols = ['Last Failure Reason', 'Failure Reason', 'Reason']
        for col in failure_cols:
            if col in df.columns:
                processed_df['Last Failure Reason'] = df[col].fillna('').astype(str)
                break
        else:
            processed_df['Last Failure Reason'] = ''
        
        # FIX 2: Chỉ giữ tracking numbers hợp lệ (>= 13 digits)
        # Ensure all string columns are properly converted
        processed_df['Order ID'] = processed_df['Order ID'].astype(str)
        processed_df['Tracking Number'] = processed_df['Tracking Number'].astype(str)
        processed_df['Status'] = processed_df['Status'].astype(str)
        processed_df['Last Failure Reason'] = processed_df['Last Failure Reason'].astype(str)
        
        initial_count = len(processed_df)
        processed_df = processed_df[processed_df['Tracking Number'].str.len() >= 13].copy()
        final_count = len(processed_df)
        
        logger.info(f"🔧 Filtered {initial_count - final_count} invalid tracking numbers")
        logger.info(f"✅ Kept {final_count} rows with valid tracking numbers")
        
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
        'Status': [],
        'Last Failure Reason': []
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
