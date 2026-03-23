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

# Download folder
DOWNLOAD_FOLDER = os.path.expanduser("~/Downloads")
if not os.path.exists(DOWNLOAD_FOLDER):
    DOWNLOAD_FOLDER = os.getcwd()

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

        chromedriver_paths = ['/usr/bin/chromedriver', '/usr/local/bin/chromedriver', 'chromedriver']

        for driver_path in chromedriver_paths:
            try:
                if os.path.exists(driver_path) or driver_path == 'chromedriver':
                    service = Service(executable_path=driver_path)
                    driver = webdriver.Chrome(service=service, options=chrome_options)
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

def check_login_error(driver):
    """=== FIX 1: Kiểm tra thông báo lỗi trên trang login ==="""
    error_selectors = [
        # PrimeFaces error messages
        (By.CSS_SELECTOR, ".ui-messages-error"),
        (By.CSS_SELECTOR, ".ui-message-error"),
        (By.CSS_SELECTOR, ".ui-growl-message"),
        # Common error containers
        (By.CSS_SELECTOR, ".error-message"),
        (By.CSS_SELECTOR, ".alert-danger"),
        (By.CSS_SELECTOR, ".login-error"),
        (By.CSS_SELECTOR, "[class*='error']"),
        (By.CSS_SELECTOR, "[class*='Error']"),
        # JSF messages
        (By.CSS_SELECTOR, ".ui-messages"),
        (By.XPATH, "//span[contains(@class, 'ui-messages-error-summary')]"),
        (By.XPATH, "//div[contains(@class, 'ui-message')]"),
    ]
    
    for by, value in error_selectors:
        try:
            elements = driver.find_elements(by, value)
            for el in elements:
                text = el.text.strip()
                if text:
                    logger.error(f"🔴 LOGIN ERROR MESSAGE: '{text}'")
                    return text
        except:
            continue
    
    # Also check page source for common error keywords
    try:
        page_text = driver.find_element(By.TAG_NAME, "body").text
        error_keywords = [
            "Invalid", "incorrect", "failed", "wrong", "locked", 
            "disabled", "expired", "captcha", "verify", "suspended",
            "không hợp lệ", "sai"
        ]
        for keyword in error_keywords:
            if keyword.lower() in page_text.lower():
                # Get surrounding context
                idx = page_text.lower().find(keyword.lower())
                start = max(0, idx - 50)
                end = min(len(page_text), idx + 100)
                context = page_text[start:end].strip()
                logger.error(f"🔴 Found error keyword '{keyword}' in page: ...{context}...")
                return context
    except:
        pass
    
    return None

def login_to_dhl(driver):
    """Login to DHL portal - with error detection and multiple submit methods"""
    try:
        logger.info("🔹 Logging into DHL portal...")

        driver.get("https://ecommerceportal.dhl.com/Portal/pages/login/userlogin.xhtml")
        time.sleep(8)

        logger.info(f"📄 Current URL: {driver.current_url}")
        logger.info(f"📄 Page title: {driver.title}")

        # === FIX: Check for CAPTCHA or additional verification ===
        captcha_selectors = [
            (By.CSS_SELECTOR, "[class*='captcha']"),
            (By.CSS_SELECTOR, "[id*='captcha']"),
            (By.CSS_SELECTOR, "iframe[src*='recaptcha']"),
            (By.CSS_SELECTOR, "iframe[src*='captcha']"),
            (By.CSS_SELECTOR, ".g-recaptcha"),
        ]
        for by, value in captcha_selectors:
            try:
                captcha = driver.find_element(by, value)
                if captcha:
                    logger.error("🔴 CAPTCHA detected on login page! Automated login cannot proceed.")
                    logger.error("🔴 Bạn cần kiểm tra xem DHL có yêu cầu CAPTCHA không.")
                    return False
            except NoSuchElementException:
                continue

        # Find and fill username
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
            return False

        username_field.clear()
        username_field.send_keys(DHL_USERNAME)
        logger.info("✅ Username entered")

        # Find and fill password
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

        # === FIX 3: Try multiple submit methods for JSF form ===
        login_success = False
        
        # Method 1: Press ENTER on password field (most reliable for JSF)
        logger.info("🔄 Method 1: Submitting via ENTER key...")
        password_field.send_keys(Keys.ENTER)
        time.sleep(10)
        
        if "userlogin" not in driver.current_url.lower():
            login_success = True
            logger.info("✅ Login successful via ENTER key")
        else:
            # Check for error message before trying next method
            error_msg = check_login_error(driver)
            if error_msg:
                logger.error(f"🔴 Login failed with error: {error_msg}")
                return False
            
            # Method 2: Click login button
            logger.info("🔄 Method 2: Clicking login button...")
            
            # Re-enter credentials (page might have refreshed)
            try:
                username_field = driver.find_element(By.ID, "email1")
                password_field = driver.find_element(By.NAME, "j_password")
                username_field.clear()
                username_field.send_keys(DHL_USERNAME)
                password_field.clear()
                password_field.send_keys(DHL_PASSWORD)
                time.sleep(1)
            except:
                pass
            
            login_button = None
            login_selectors = [
                (By.CLASS_NAME, "btn-login"),
                (By.XPATH, "//button[contains(@class, 'btn-login')]"),
                (By.XPATH, "//input[@type='submit']"),
                (By.XPATH, "//button[@type='submit']"),
                (By.XPATH, "//button[contains(text(), 'Login')]"),
                (By.XPATH, "//button[contains(text(), 'Sign')]"),
            ]
            for by, value in login_selectors:
                login_button = wait_and_find(driver, by, value, timeout=5)
                if login_button:
                    logger.info(f"✅ Found login button with: {by}={value}")
                    break

            if login_button:
                try:
                    login_button.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", login_button)
                logger.info("✅ Login button clicked")
                time.sleep(15)
                
                if "userlogin" not in driver.current_url.lower():
                    login_success = True
                else:
                    # Method 3: JSF form submit via JavaScript
                    logger.info("🔄 Method 3: JSF form submit via JavaScript...")
                    try:
                        driver.execute_script("""
                            var forms = document.getElementsByTagName('form');
                            for (var i = 0; i < forms.length; i++) {
                                if (forms[i].querySelector('input[type="password"]')) {
                                    forms[i].submit();
                                    break;
                                }
                            }
                        """)
                        time.sleep(15)
                        
                        if "userlogin" not in driver.current_url.lower():
                            login_success = True
                    except Exception as e:
                        logger.warning(f"JS submit failed: {e}")

        if login_success:
            logger.info(f"✅ Login successful! URL: {driver.current_url}")
            return True
        
        # === FIX 1: Final error check ===
        error_msg = check_login_error(driver)
        if error_msg:
            logger.error(f"🔴 Login failed. Error on page: {error_msg}")
        else:
            logger.error("❌ Login failed - still on login page (no error message found)")
            logger.error("❌ Possible causes: wrong credentials, account locked, CAPTCHA, or IP blocked")
        
        # Print more page info for debugging
        try:
            body_text = driver.find_element(By.TAG_NAME, "body").text[:1000]
            logger.info(f"📄 Page text: {body_text}")
        except:
            pass
        
        return False

    except Exception as e:
        logger.error(f"❌ Login error: {str(e)}")
        return False

def navigate_to_dashboard(driver):
    """Navigate to dashboard"""
    try:
        logger.info("🔹 Navigating to dashboard...")
        time.sleep(5)
        
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
        
        script = f"""
        var element = document.getElementById('{element_id}');
        if (element) {{
            element.value = '{date_value}';
            var events = ['input', 'change', 'blur'];
            events.forEach(function(eventType) {{
                var event = new Event(eventType, {{ bubbles: true }});
                element.dispatchEvent(event);
            }});
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
        time.sleep(5)
        
        from_success = set_datepicker_value(driver, "dashboardForm:frmDate_input", START_DATE)
        time.sleep(1)
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
        time.sleep(3)
        
        generate_button = wait_and_find(driver, By.XPATH, "//span[contains(@class, 'ui-button-text') and contains(text(), 'GENERATE')]")
        if not generate_button:
            generate_button = wait_and_find(driver, By.XPATH, "//button[contains(text(), 'GENERATE')] | //input[@value='GENERATE']")
        
        if generate_button:
            generate_button.click()
            logger.info("✅ Clicked GENERATE button")
            time.sleep(8)
            return True
        else:
            logger.error("❌ GENERATE button not found")
            return False
            
    except Exception as e:
        logger.error(f"❌ Error clicking GENERATE: {str(e)}")
        return False

def clear_download_folder():
    """Clear old download files"""
    try:
        clear_paths = [DOWNLOAD_FOLDER, os.path.expanduser("~/Downloads"), os.path.expanduser("~/Desktop")]
        total_files_removed = 0
        
        for path in clear_paths:
            if os.path.exists(path):
                try:
                    files_removed = 0
                    for filename in os.listdir(path):
                        if filename.endswith(('.xlsx', '.csv', '.xls')) and not filename.startswith('~'):
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
                if current_time - file_time < 120:
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
            "/tmp",
            os.getcwd()
        ]
        
        logger.info("🔍 Checking alternative download paths...")
        
        for path in download_paths:
            if os.path.exists(path):
                files = [f for f in os.listdir(path) 
                        if f.endswith(('.xlsx', '.csv', '.xls')) and not f.startswith('~')]
                
                if files:
                    recent_files = []
                    current_time = time.time()
                    
                    for file in files:
                        file_path = os.path.join(path, file)
                        file_time = os.path.getctime(file_path)
                        if current_time - file_time < 300:
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

def download_report(driver):
    """Download the report by clicking Excel icon"""
    try:
        logger.info("🔹 Looking for download icon...")
        clear_download_folder()
        
        download_icon = wait_and_find(driver, By.ID, "xlsIcon")
        if not download_icon:
            download_icon = wait_and_find(driver, By.XPATH, "//img[contains(@src, 'download_Pixel_30.png')] | //img[contains(@src, 'excel')] | //img[contains(@id, 'xls')]")
        
        if download_icon:
            download_icon.click()
            logger.info("✅ Clicked download icon")
            logger.info("⏳ Waiting for file download...")
            time.sleep(25)
            
            if check_for_new_download():
                logger.info("✅ File downloaded successfully")
                return True
            else:
                logger.warning("⚠️ Checking alternative download paths...")
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
    
    search_paths = [folder_path, os.path.expanduser("~/Downloads"), os.path.expanduser("~/Desktop"), "/tmp"]
    
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
                logger.info(f"No files found. Attempt {attempt + 1}/{max_attempts}")
                time.sleep(delay)
                continue
            
            latest_file = max(all_files, key=os.path.getctime)
            file_size = os.path.getsize(latest_file)
            
            if file_size > 0:
                if os.path.dirname(latest_file) != folder_path:
                    destination = os.path.join(folder_path, os.path.basename(latest_file))
                    shutil.copy2(latest_file, destination)
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
    """Process downloaded data"""
    if file_path is None:
        logger.warning("No file to process")
        return None  # FIX: Return None instead of empty DataFrame
    
    logger.info(f"🔹 Processing file: {file_path}")
    
    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path, encoding='utf-8')
        else:
            df = pd.read_excel(file_path, engine='openpyxl')
        
        logger.info(f"File loaded. Shape: {df.shape}")
        logger.info(f"Columns: {df.columns.tolist()}")
        
        if len(df) > 0:
            logger.info("Sample data:")
            logger.info(df.head(3).to_string())
        
        processed_df = pd.DataFrame()
        
        if 'Consignee Name' in df.columns:
            processed_df['Order ID'] = df['Consignee Name'].fillna('').astype(str).str[:7]
        else:
            processed_df['Order ID'] = ''
        
        tracking_cols = ['Tracking ID', 'Tracking Number', 'AWB', 'Waybill Number']
        for col in tracking_cols:
            if col in df.columns:
                processed_df['Tracking Number'] = df[col].fillna('').astype(str)
                break
        else:
            processed_df['Tracking Number'] = ''
        
        pickup_cols = ['Pickup Event DateTime', 'Pickup Date', 'Collection Date', 'Ship Date']
        for col in pickup_cols:
            if col in df.columns:
                processed_df['Pickup DateTime'] = pd.to_datetime(df[col], errors='coerce')
                break
        else:
            processed_df['Pickup DateTime'] = pd.NaT
        
        delivery_cols = ['Delivery Date', 'Delivered Date', 'POD Date']
        for col in delivery_cols:
            if col in df.columns:
                processed_df['Delivery Date'] = pd.to_datetime(df[col], errors='coerce')
                break
        else:
            processed_df['Delivery Date'] = pd.NaT
        
        status_cols = ['Last Status', 'Status', 'Current Status', 'Shipment Status']
        for col in status_cols:
            if col in df.columns:
                processed_df['Status'] = df[col].fillna('').astype(str)
                break
        else:
            processed_df['Status'] = ''

        failure_cols = ['Last Failure Reason', 'Failure Reason', 'Reason']
        for col in failure_cols:
            if col in df.columns:
                processed_df['Last Failure Reason'] = df[col].fillna('').astype(str)
                break
        else:
            processed_df['Last Failure Reason'] = ''
        
        processed_df['Order ID'] = processed_df['Order ID'].astype(str)
        processed_df['Tracking Number'] = processed_df['Tracking Number'].astype(str)
        processed_df['Status'] = processed_df['Status'].astype(str)
        processed_df['Last Failure Reason'] = processed_df['Last Failure Reason'].astype(str)
        
        initial_count = len(processed_df)
        processed_df = processed_df[processed_df['Tracking Number'].str.len() >= 13].copy()
        final_count = len(processed_df)
        
        logger.info(f"🔧 Filtered {initial_count - final_count} invalid tracking numbers")
        logger.info(f"✅ Kept {final_count} rows with valid tracking numbers")
        
        if not processed_df['Pickup DateTime'].isna().all():
            processed_df = processed_df.sort_values('Pickup DateTime', ascending=False, na_position='last')
        
        processed_df['Pickup DateTime'] = processed_df['Pickup DateTime'].apply(
            lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(x) else ''
        )
        processed_df['Delivery Date'] = processed_df['Delivery Date'].apply(
            lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(x) else ''
        )
        
        processed_df = processed_df.replace({np.nan: '', 'NaT': '', None: ''})
        
        logger.info(f"✅ Processing completed. Final shape: {processed_df.shape}")
        return processed_df
        
    except Exception as e:
        logger.error(f"❌ Error processing data: {str(e)}")
        return None

def upload_to_google_sheets(df):
    """=== FIX 2: KHÔNG upload nếu không có dữ liệu (tránh xóa dữ liệu cũ) ==="""
    
    # CRITICAL: If df is None or empty, DO NOT clear the sheet
    if df is None or len(df) == 0:
        logger.warning("⚠️ SKIPPING upload - no data to upload")
        logger.warning("⚠️ Existing data in Google Sheet is PRESERVED (not cleared)")
        return False
    
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
        
        headers = df.columns.tolist()
        data = df.astype(str).values.tolist()
        values = [headers] + data
        
        logger.info(f"Uploading {len(data)} rows")
        
        # Only clear and upload when we have actual data
        service.spreadsheets().values().clear(
            spreadsheetId=GOOGLE_SHEET_ID,
            range=f"{SHEET_NAME}!A1:Z10000"
        ).execute()
        logger.info("✅ Cleared existing content")
        
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
        
        driver = setup_chrome_driver()
        
        # Step 1: Login
        if not login_to_dhl(driver):
            logger.error("❌ Login failed - STOPPING (existing data preserved)")
            # FIX 2: Do NOT upload empty data when login fails
            return
        
        # Step 2: Navigate to dashboard
        if not navigate_to_dashboard(driver):
            logger.error("❌ Dashboard navigation failed - STOPPING (existing data preserved)")
            return
        
        # Step 3: Set date range
        set_date_range(driver)
        
        # Step 4: Click generate
        if not click_generate_button(driver):
            logger.error("❌ Generate button click failed - STOPPING (existing data preserved)")
            return
        
        # Step 5: Download report
        if not download_report(driver):
            logger.error("❌ Download failed - STOPPING (existing data preserved)")
            return
        
        # Step 6: Process data
        latest_file = get_latest_file(DOWNLOAD_FOLDER)
        processed_df = process_data(latest_file)
        
        # Step 7: Upload to sheets (only if we have data)
        if processed_df is not None and len(processed_df) > 0:
            upload_to_google_sheets(processed_df)
            logger.info("🎉 Process completed successfully!")
        else:
            logger.warning("⚠️ No valid data to upload. Existing data preserved.")
        
    except Exception as e:
        logger.error(f"❌ Main process failed: {str(e)}")
        # FIX 2: Do NOT upload empty data on error
    
    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass

if __name__ == "__main__":
    main()
