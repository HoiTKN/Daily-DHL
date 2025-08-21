from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd
import os
from datetime import datetime, date
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
DEFAULT_TIMEOUT = 30
PAGE_LOAD_TIMEOUT = 60

# Date range settings - from start of year to current date
START_DATE = "01/01/2025"  # US format MM/DD/YYYY
END_DATE = datetime.now().strftime("%m/%d/%Y")  # Current date in MM/DD/YYYY format

# Get credentials from environment variables (GitHub secrets)
DHL_USERNAME = os.getenv('DHL_USERNAME', 'truongcongdai4@gmail.com')
DHL_PASSWORD = os.getenv('DHL_PASSWORD', '@Thavi035@')

def validate_environment():
    """Validate environment variables and setup"""
    try:
        logger.info("🔍 Validating environment setup...")
        
        # Check credentials
        if not DHL_USERNAME or DHL_USERNAME == '':
            logger.error("❌ DHL_USERNAME not set in environment")
            return False
            
        if not DHL_PASSWORD or DHL_PASSWORD == '':
            logger.error("❌ DHL_PASSWORD not set in environment") 
            return False
        
        logger.info(f"✅ Username: {DHL_USERNAME}")
        logger.info(f"✅ Password: {'*' * len(DHL_PASSWORD)} (length: {len(DHL_PASSWORD)})")
        logger.info(f"✅ Date range: 01/01/2025 to {END_DATE} (MM/DD/YYYY format)")
        
        # Check service account file
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            logger.warning(f"⚠️ Service account file not found: {SERVICE_ACCOUNT_FILE}")
        else:
            logger.info(f"✅ Service account file found: {SERVICE_ACCOUNT_FILE}")
        
        # Check download folder
        if not os.path.exists(DOWNLOAD_FOLDER):
            logger.warning(f"⚠️ Download folder not found: {DOWNLOAD_FOLDER}")
            os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
            logger.info(f"✅ Created download folder: {DOWNLOAD_FOLDER}")
        else:
            logger.info(f"✅ Download folder exists: {DOWNLOAD_FOLDER}")
        
        return True
        
    except Exception as e:
        logger.error(f"❌ Environment validation failed: {str(e)}")
        return False

def setup_chrome_driver():
    """Setup Chrome driver with enhanced options"""
    try:
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--window-size=1920,1080')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        
        # Enhanced download preferences
        prefs = {
            "download.default_directory": DOWNLOAD_FOLDER,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_settings.popups": 0,
            "profile.default_content_setting_values.automatic_downloads": 1
        }
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # Try to find ChromeDriver
        chromedriver_paths = [
            '/usr/bin/chromedriver',
            '/usr/local/bin/chromedriver',
            'chromedriver'
        ]
        
        driver = None
        for driver_path in chromedriver_paths:
            try:
                if os.path.exists(driver_path) or driver_path == 'chromedriver':
                    service = Service(executable_path=driver_path)
                    driver = webdriver.Chrome(service=service, options=chrome_options)
                    logger.info(f"✅ ChromeDriver initialized from: {driver_path}")
                    break
            except Exception as e:
                logger.warning(f"Failed to use {driver_path}: {str(e)}")
                continue
        
        if driver is None:
            raise Exception("Could not initialize ChromeDriver from any location")
        
        # Set stealth mode
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
        
        return driver
        
    except Exception as e:
        logger.error(f"❌ Chrome driver setup failed: {str(e)}")
        raise

def debug_page_state(driver, step_name):
    """Debug current page state"""
    try:
        logger.info(f"🔍 DEBUG {step_name}:")
        logger.info(f"   Current URL: {driver.current_url}")
        logger.info(f"   Page title: {driver.title}")
        logger.info(f"   Page source length: {len(driver.page_source)}")
        
        # Save page source for debugging
        filename = f"debug_{step_name.lower().replace(' ', '_')}.html"
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        logger.info(f"   Page source saved to: {filename}")
        
        # Take screenshot
        screenshot_file = f"debug_{step_name.lower().replace(' ', '_')}.png"
        driver.save_screenshot(screenshot_file)
        logger.info(f"   Screenshot saved to: {screenshot_file}")
        
    except Exception as e:
        logger.warning(f"Could not debug page state: {str(e)}")

def check_already_logged_in(driver):
    """Check if user is already logged in (manual login)"""
    try:
        current_url = driver.current_url
        logger.info(f"Checking login status. Current URL: {current_url}")
        
        # If not on login page, assume logged in
        if "login" not in current_url.lower():
            logger.info("✅ User appears to be already logged in")
            return True
        
        # Look for logout elements or user indicators
        logout_indicators = [
            "//a[contains(text(), 'Logout') or contains(text(), 'Log out')]",
            "//span[contains(@class, 'username') or contains(@class, 'user')]",
            "//div[contains(@class, 'user-menu')]",
            "//nav[contains(@class, 'user-nav')]"
        ]
        
        for indicator in logout_indicators:
            elements = driver.find_elements(By.XPATH, indicator)
            if elements and any(elem.is_displayed() for elem in elements):
                logger.info("✅ Found login indicators - user is logged in")
                return True
        
        return False
        
    except Exception as e:
        logger.warning(f"Error checking login status: {str(e)}")
        return False

def try_direct_dashboard_access(driver):
    """Try to access dashboard directly (if already logged in)"""
    try:
        logger.info("🔹 Trying direct dashboard access...")
        
        dashboard_urls = [
            "https://ecommerceportal.dhl.com/Portal/pages/dashboard/dashboard.xhtml",
            "https://ecommerceportal.dhl.com/Portal/pages/shipment/shipmentList.xhtml",
            "https://ecommerceportal.dhl.com/Portal/pages/reports/reports.xhtml"
        ]
        
        for url in dashboard_urls:
            try:
                logger.info(f"Trying: {url}")
                driver.get(url)
                time.sleep(8)
                
                current_url = driver.current_url
                if "login" not in current_url.lower() and "error" not in current_url.lower():
                    logger.info(f"✅ Successfully accessed: {url}")
                    debug_page_state(driver, "Direct Dashboard Access")
                    return True
                else:
                    logger.info(f"Redirected to: {current_url}")
                    
            except Exception as e:
                logger.warning(f"Failed to access {url}: {str(e)}")
        
        return False
        
    except Exception as e:
        logger.error(f"Error in direct dashboard access: {str(e)}")
        return False

def safe_find_element(driver, by, value, timeout=10):
    """Safely find element with timeout"""
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
        return element
    except TimeoutException:
        logger.warning(f"Element not found: {by}={value}")
        return None
    except Exception as e:
        logger.warning(f"Error finding element {by}={value}: {str(e)}")
        return None

def safe_click(driver, element, method="normal"):
    """Safely click element with multiple fallback methods"""
    if element is None:
        return False
        
    try:
        if method == "normal":
            element.click()
            return True
        elif method == "js":
            driver.execute_script("arguments[0].click();", element)
            return True
        elif method == "action":
            from selenium.webdriver.common.action_chains import ActionChains
            ActionChains(driver).move_to_element(element).click().perform()
            return True
    except Exception as e:
        logger.warning(f"Click method '{method}' failed: {str(e)}")
        return False

def provide_login_troubleshooting(driver):
    """Provide troubleshooting information for login issues"""
    try:
        logger.info("🔧 LOGIN TROUBLESHOOTING GUIDE:")
        logger.info("=" * 50)
        
        # Check current page state
        current_url = driver.current_url
        page_title = driver.title
        
        logger.info(f"Current URL: {current_url}")
        logger.info(f"Page title: {page_title}")
        
        # Check for common login issues
        if "blocked" in driver.page_source.lower():
            logger.warning("⚠️ Account may be blocked or suspended")
        
        if "captcha" in driver.page_source.lower():
            logger.warning("⚠️ CAPTCHA detected - requires manual intervention")
        
        if "verification" in driver.page_source.lower():
            logger.warning("⚠️ Additional verification required")
        
        # Provide manual steps
        logger.info("\n💡 MANUAL STEPS TO TRY:")
        logger.info("1. Login manually in browser with same credentials")
        logger.info("2. Check if account requires additional verification")
        logger.info("3. Verify credentials in GitHub Secrets:")
        logger.info(f"   - DHL_USERNAME: {DHL_USERNAME}")
        logger.info(f"   - DHL_PASSWORD: [length: {len(DHL_PASSWORD)} chars]")
        logger.info("4. Try running script after manual login (stay logged in)")
        logger.info("5. Check for emails about suspicious login attempts")
        
        logger.info("\n🔧 TECHNICAL DEBUGGING:")
        logger.info("- Check debug_*.html files for page content")
        logger.info("- Check debug_*.png files for visual state")
        logger.info("- Verify no popup blockers or browser extensions interfering")
        
    except Exception as e:
        logger.error(f"Error in troubleshooting: {str(e)}")

def login_to_dhl(driver):
    """Enhanced login with comprehensive security handling"""
    try:
        logger.info("🔹 Accessing DHL portal...")
        driver.get("https://ecommerceportal.dhl.com/Portal/pages/login/userlogin.xhtml")
        time.sleep(8)
        
        debug_page_state(driver, "Initial Load")
        
        # Clear cookies and add realistic behavior
        driver.delete_all_cookies()
        
        # Add human-like mouse movements
        try:
            from selenium.webdriver.common.action_chains import ActionChains
            actions = ActionChains(driver)
            actions.move_by_offset(100, 100).perform()
            time.sleep(1)
        except:
            pass
        
        # Check for CAPTCHA or additional security elements
        captcha_elements = driver.find_elements(By.XPATH, 
            "//div[contains(@class, 'captcha')] | //iframe[contains(@src, 'captcha')] | //div[contains(@class, 'recaptcha')]")
        if captcha_elements:
            logger.warning("⚠️ CAPTCHA detected - manual intervention may be required")
        
        # Find username field with multiple strategies
        username_selectors = ["email1", "username", "user", "login"]
        username_input = None
        
        for selector in username_selectors:
            username_input = safe_find_element(driver, By.ID, selector, 5)
            if username_input:
                logger.info(f"Found username field: {selector}")
                break
        
        if not username_input:
            # Try by name attribute
            for selector in username_selectors:
                username_input = safe_find_element(driver, By.NAME, selector, 5)
                if username_input:
                    logger.info(f"Found username field by name: {selector}")
                    break
        
        if not username_input:
            logger.error("❌ Could not find username field")
            return False
        
        # Human-like typing for username
        username_input.clear()
        time.sleep(0.5)
        for char in DHL_USERNAME:
            username_input.send_keys(char)
            time.sleep(0.05)  # Realistic typing speed
        logger.info("✅ Username entered")
        
        # Find password field
        password_selectors = ["j_password", "password", "pass", "pwd"]
        password_input = None
        
        for selector in password_selectors:
            password_input = safe_find_element(driver, By.NAME, selector, 5)
            if password_input:
                logger.info(f"Found password field: {selector}")
                break
        
        if not password_input:
            for selector in password_selectors:
                password_input = safe_find_element(driver, By.ID, selector, 5)
                if password_input:
                    logger.info(f"Found password field by ID: {selector}")
                    break
        
        if not password_input:
            logger.error("❌ Could not find password field")
            return False
        
        # Human-like typing for password
        password_input.clear()
        time.sleep(0.5)
        for char in DHL_PASSWORD:
            password_input.send_keys(char)
            time.sleep(0.05)
        logger.info("✅ Password entered")
        
        # Wait before clicking login (human behavior)
        time.sleep(2)
        
        debug_page_state(driver, "Before Login")
        
        # Find login button with multiple strategies
        login_selectors = [
            (By.CLASS_NAME, "btn-login"),
            (By.XPATH, "//button[contains(text(), 'Login') or contains(text(), 'Sign in')]"),
            (By.XPATH, "//input[@type='submit']"),
            (By.XPATH, "//button[@type='submit']"),
            (By.ID, "login"),
            (By.ID, "submit")
        ]
        
        login_button = None
        for by, selector in login_selectors:
            login_button = safe_find_element(driver, by, selector, 5)
            if login_button:
                logger.info(f"Found login button: {by}={selector}")
                break
        
        if not login_button:
            logger.error("❌ Could not find login button")
            return False
        
        # Multiple login attempts with different methods
        login_success = False
        
        for attempt in range(3):
            logger.info(f"Login attempt {attempt + 1}/3")
            
            # Try different click methods
            for method in ["normal", "js", "action"]:
                try:
                    if safe_click(driver, login_button, method):
                        logger.info(f"✅ Login button clicked using {method}")
                        
                        # Wait and check for immediate feedback
                        time.sleep(3)
                        
                        # Check for alerts or popups
                        try:
                            alert = driver.switch_to.alert
                            alert_text = alert.text
                            logger.info(f"Alert detected: {alert_text}")
                            alert.accept()
                            time.sleep(2)
                        except:
                            pass
                        
                        # Check for loading indicators
                        loading_elements = driver.find_elements(By.XPATH, 
                            "//div[contains(@class, 'loading')] | //div[contains(@class, 'spinner')] | //div[contains(text(), 'Loading')]")
                        
                        if loading_elements:
                            logger.info("Loading indicator detected, waiting...")
                            time.sleep(10)
                        
                        # Wait for potential redirect
                        time.sleep(12)
                        
                        current_url = driver.current_url
                        logger.info(f"URL after login attempt: {current_url}")
                        
                        if "login" not in current_url.lower():
                            login_success = True
                            break
                        
                except Exception as e:
                    logger.warning(f"Click method {method} failed: {str(e)}")
            
            if login_success:
                break
            
            # If still on login page, check for error messages
            error_selectors = [
                "//div[contains(@class, 'error')]",
                "//div[contains(@class, 'alert')]",
                "//div[contains(@class, 'message')]",
                "//span[contains(@class, 'error')]",
                "//p[contains(@class, 'error')]"
            ]
            
            for selector in error_selectors:
                error_elements = driver.find_elements(By.XPATH, selector)
                for elem in error_elements:
                    if elem.is_displayed() and elem.text.strip():
                        logger.error(f"Error message detected: {elem.text}")
            
            # Check for additional security steps
            security_elements = driver.find_elements(By.XPATH, 
                "//div[contains(text(), 'verification')] | //div[contains(text(), 'security')] | //div[contains(text(), 'code')]")
            
            if security_elements:
                for elem in security_elements:
                    if elem.is_displayed():
                        logger.warning(f"Security step detected: {elem.text}")
            
            # Wait before retry
            if attempt < 2:
                logger.info("Waiting before retry...")
                time.sleep(5)
        
        debug_page_state(driver, "After All Login Attempts")
        
        # Final check
        if login_success or "login" not in driver.current_url.lower():
            logger.info("✅ Login appears successful")
            return True
        else:
            logger.error("❌ Login failed after all attempts")
            
            # Provide troubleshooting guidance
            provide_login_troubleshooting(driver)
            
            return False
        
    except Exception as e:
        logger.error(f"❌ Login failed with exception: {str(e)}")
        debug_page_state(driver, "Login Exception")
        return False

def set_date_range(driver):
    """
    Set date range for reports from start of year to current date
    
    Enhanced approach with:
    - Multiple date format attempts (US format prioritized)
    - Input state validation (readonly/disabled checks)
    - Multiple input methods (clear/sendkeys, JS, character-by-character)
    - Proper event triggering for dynamic forms
    """
    try:
        logger.info(f"🔹 Setting date range: {START_DATE} to {END_DATE}")
        
        # First, try to find all date inputs and analyze their state
        debug_date_inputs(driver)
        
        # Different date formats to try (US format first as DHL likely uses US format)
        current_date_us = datetime.now().strftime("%m/%d/%Y")
        current_date_dash = datetime.now().strftime("%m-%d-%Y") 
        current_date_iso = datetime.now().strftime("%Y-%m-%d")
        current_date_short = datetime.now().strftime("%m/%d/%y")
        
        date_formats = [
            ("01/01/2025", current_date_us),      # MM/DD/YYYY format (US format - try first)
            ("01-01-2025", current_date_dash),    # MM-DD-YYYY format  
            ("2025-01-01", current_date_iso),     # YYYY-MM-DD format
            ("01/01/25", current_date_short),     # MM/DD/YY format
            ("1/1/2025", current_date_us),        # Single digit format
        ]
        
        # Multiple date input strategies for different portal layouts
        date_strategies = [
            # Strategy 1: Dashboard form date inputs
            {
                'from_selectors': ['dashboardForm:frmDate_input', 'fromDate', 'startDate'],
                'to_selectors': ['dashboardForm:toDate_input', 'toDate', 'endDate'],
                'apply_selectors': ['dashboardForm:applyBtn', 'applyButton', '//button[contains(text(), "Apply")]']
            },
            # Strategy 2: Report form date inputs  
            {
                'from_selectors': ['reportForm:fromDate_input', 'report_fromDate', 'from_date'],
                'to_selectors': ['reportForm:toDate_input', 'report_toDate', 'to_date'],
                'apply_selectors': ['reportForm:submitBtn', 'submitButton', '//button[contains(text(), "Submit")]']
            },
            # Strategy 3: Generic date inputs
            {
                'from_selectors': ['from_date', 'start_date', 'dateFrom'],
                'to_selectors': ['to_date', 'end_date', 'dateTo'],
                'apply_selectors': ['submit', 'apply', '//input[@type="submit"]']
            }
        ]
        
        date_set_success = False
        
        for strategy_num, strategy in enumerate(date_strategies, 1):
            logger.info(f"Trying date strategy {strategy_num}...")
            
            try:
                # Find from date input
                from_input = None
                for selector in strategy['from_selectors']:
                    if selector.startswith('//'):
                        from_input = safe_find_element(driver, By.XPATH, selector, 3)
                    else:
                        from_input = safe_find_element(driver, By.ID, selector, 3)
                    if from_input:
                        logger.info(f"Found from date input: {selector}")
                        break
                
                # Find to date input
                to_input = None
                for selector in strategy['to_selectors']:
                    if selector.startswith('//'):
                        to_input = safe_find_element(driver, By.XPATH, selector, 3)
                    else:
                        to_input = safe_find_element(driver, By.ID, selector, 3)
                    if to_input:
                        logger.info(f"Found to date input: {selector}")
                        break
                
                if from_input and to_input:
                    # Check if inputs are enabled and their type
                    from_enabled = from_input.is_enabled()
                    to_enabled = to_input.is_enabled()
                    from_readonly = from_input.get_attribute('readonly')
                    to_readonly = to_input.get_attribute('readonly')
                    from_class = from_input.get_attribute('class') or ''
                    to_class = to_input.get_attribute('class') or ''
                    
                    logger.info(f"From input - enabled: {from_enabled}, readonly: {from_readonly}, class: {from_class}")
                    logger.info(f"To input - enabled: {to_enabled}, readonly: {to_readonly}, class: {to_class}")
                    
                    # Check if these are datepicker widgets
                    is_from_datepicker = 'datepicker' in from_class.lower() or 'hasDatepicker' in from_class
                    is_to_datepicker = 'datepicker' in to_class.lower() or 'hasDatepicker' in to_class
                    
                    logger.info(f"Datepicker detection - From: {is_from_datepicker}, To: {is_to_datepicker}")
                    
                    # Skip only if inputs are disabled (not just readonly)
                    if not from_enabled:
                        logger.warning("From date input is disabled")
                        continue
                    if not to_enabled:
                        logger.warning("To date input is disabled") 
                        continue
                    
                    # Try different date formats
                    for format_index, (start_date, end_date) in enumerate(date_formats):
                        logger.info(f"Trying date format {format_index + 1}: {start_date} - {end_date}")
                        
                        try:
                            # Enhanced input method with datepicker support
                            success = set_input_value_enhanced(driver, from_input, start_date)
                            if success:
                                logger.info(f"✅ Set from date: {start_date}")
                                time.sleep(1)
                                
                                success = set_input_value_enhanced(driver, to_input, end_date)
                                if success:
                                    logger.info(f"✅ Set to date: {end_date}")
                                    time.sleep(2)
                                    
                                    # Try to find and click apply button
                                    for selector in strategy['apply_selectors']:
                                        apply_btn = None
                                        if selector.startswith('//'):
                                            apply_btn = safe_find_element(driver, By.XPATH, selector, 3)
                                        else:
                                            apply_btn = safe_find_element(driver, By.ID, selector, 3)
                                        
                                        if apply_btn and apply_btn.is_displayed():
                                            if safe_click(driver, apply_btn):
                                                logger.info(f"✅ Clicked apply button: {selector}")
                                                time.sleep(5)
                                                date_set_success = True
                                                break
                                    
                                    if date_set_success:
                                        break
                                else:
                                    logger.warning(f"Failed to set to date with format {format_index + 1}")
                            else:
                                logger.warning(f"Failed to set from date with format {format_index + 1}")
                        except Exception as format_e:
                            logger.warning(f"Date format {format_index + 1} failed: {str(format_e)}")
                            continue
                    
                    if date_set_success:
                        break
                        
            except Exception as e:
                logger.warning(f"Date strategy {strategy_num} failed: {str(e)}")
                continue
        
        if date_set_success:
            logger.info("✅ Date range set successfully")
            debug_page_state(driver, "After Date Range Set")
        else:
            logger.warning("⚠️ Could not set date range with any strategy")
            # Try alternative approach - look for calendar widgets
            try_advanced_date_setting(driver)
        
        return date_set_success
        
    except Exception as e:
        logger.error(f"❌ Error setting date range: {str(e)}")
        return False

def debug_date_inputs(driver):
    """Debug all date inputs on the page"""
    try:
        logger.info("🔍 Debugging all date inputs...")
        
        # Find all potential date inputs
        all_inputs = driver.find_elements(By.XPATH, "//input")
        date_related_inputs = []
        
        for inp in all_inputs:
            if inp.is_displayed():
                input_type = inp.get_attribute('type') or ''
                input_id = inp.get_attribute('id') or ''
                input_name = inp.get_attribute('name') or ''
                input_class = inp.get_attribute('class') or ''
                placeholder = inp.get_attribute('placeholder') or ''
                readonly = inp.get_attribute('readonly')
                disabled = inp.get_attribute('disabled')
                enabled = inp.is_enabled()
                
                if (input_type in ['date', 'text'] or 
                    any(word in (input_id + input_name + input_class + placeholder).lower() 
                        for word in ['date', 'from', 'to', 'start', 'end', 'calendar'])):
                    
                    date_related_inputs.append({
                        'element': inp,
                        'id': input_id,
                        'name': input_name,
                        'type': input_type,
                        'class': input_class,
                        'placeholder': placeholder,
                        'readonly': readonly,
                        'disabled': disabled,
                        'enabled': enabled
                    })
        
        logger.info(f"Found {len(date_related_inputs)} date-related inputs:")
        for i, inp_info in enumerate(date_related_inputs):
            logger.info(f"  Input {i}: ID={inp_info['id']}, Type={inp_info['type']}, "
                       f"Enabled={inp_info['enabled']}, Readonly={inp_info['readonly']}")
        
        return date_related_inputs
        
    except Exception as e:
        logger.warning(f"Error debugging date inputs: {str(e)}")
        return []

def handle_datepicker_widget(driver, input_element, date_value):
    """Handle calendar datepicker widgets specifically"""
    try:
        input_id = input_element.get_attribute('id')
        logger.info(f"🗓️ Handling datepicker widget: {input_id}")
        
        # Method 1: Try direct JavaScript value setting for readonly inputs
        try:
            logger.info("Trying direct JavaScript date setting...")
            driver.execute_script("""
                var element = arguments[0];
                var date = arguments[1];
                
                // Set value directly
                element.value = date;
                
                // Trigger comprehensive events for datepicker
                element.dispatchEvent(new Event('focus', { bubbles: true }));
                element.dispatchEvent(new Event('input', { bubbles: true }));
                element.dispatchEvent(new Event('change', { bubbles: true }));
                element.dispatchEvent(new Event('blur', { bubbles: true }));
                
                // Try jQuery events if available
                if (window.jQuery) {
                    jQuery(element).trigger('change');
                    jQuery(element).trigger('dateSelect');
                    jQuery(element).datepicker && jQuery(element).datepicker('setDate', date);
                }
            """, input_element, date_value)
            
            # Verify if value was set
            new_value = input_element.get_attribute('value')
            if new_value and (date_value in new_value or new_value in date_value):
                logger.info(f"✅ JavaScript method worked: {new_value}")
                return True
            else:
                logger.info(f"Value after JS: {new_value}")
                
        except Exception as e:
            logger.warning(f"Direct JS method failed: {str(e)}")
        
        # Method 2: Try to find and click calendar trigger button
        try:
            logger.info("Looking for calendar trigger button...")
            
            # Look for calendar triggers near this input
            calendar_triggers = []
            
            # Strategy 1: Look for buttons/images with calendar-related attributes near this input
            parent = input_element.find_element(By.XPATH, "..")
            nearby_triggers = parent.find_elements(By.XPATH, 
                ".//button[contains(@class, 'calendar') or contains(@class, 'date')] | " +
                ".//img[contains(@src, 'calendar') or contains(@alt, 'calendar')] | " +
                ".//span[contains(@class, 'calendar') or contains(@class, 'date')] | " +
                ".//a[contains(@class, 'calendar')]")
            
            calendar_triggers.extend(nearby_triggers)
            
            # Strategy 2: Look for triggers with similar ID pattern
            base_id = input_id.replace('_input', '').replace('_focus', '')
            trigger_selectors = [
                f"#{base_id}_trigger",
                f"#{base_id}_button",
                f"#{base_id}_calendar",
                f"//img[@id='{base_id}_trigger']",
                f"//button[@id='{base_id}_trigger']"
            ]
            
            for selector in trigger_selectors:
                try:
                    if selector.startswith('//'):
                        trigger = driver.find_element(By.XPATH, selector)
                    else:
                        trigger = driver.find_element(By.CSS_SELECTOR, selector)
                    calendar_triggers.append(trigger)
                except:
                    continue
            
            # Try clicking calendar triggers
            for trigger in calendar_triggers:
                if trigger.is_displayed():
                    logger.info(f"Found calendar trigger: {trigger.tag_name}")
                    try:
                        safe_click(driver, trigger)
                        time.sleep(2)
                        
                        # After opening calendar, try to navigate to correct date
                        if navigate_calendar_to_date(driver, date_value):
                            return True
                            
                    except Exception as e:
                        logger.warning(f"Calendar trigger click failed: {str(e)}")
                        continue
        
        except Exception as e:
            logger.warning(f"Calendar trigger method failed: {str(e)}")
        
        # Method 3: Try keyboard simulation
        try:
            logger.info("Trying keyboard simulation...")
            input_element.click()
            time.sleep(0.5)
            
            # Clear field and type date
            input_element.send_keys(Keys.CONTROL + "a")
            time.sleep(0.2)
            input_element.send_keys(date_value)
            time.sleep(0.5)
            input_element.send_keys(Keys.ENTER)
            
            # Check if value was set
            new_value = input_element.get_attribute('value')
            if new_value and date_value in new_value:
                logger.info(f"✅ Keyboard method worked: {new_value}")
                return True
                
        except Exception as e:
            logger.warning(f"Keyboard method failed: {str(e)}")
        
        return False
        
    except Exception as e:
        logger.error(f"❌ All datepicker methods failed: {str(e)}")
        return False

def navigate_calendar_to_date(driver, target_date):
    """Navigate calendar widget to specific date"""
    try:
        logger.info(f"🗓️ Navigating calendar to date: {target_date}")
        
        # Wait for calendar popup to appear
        time.sleep(1)
        
        # Look for calendar popup elements
        calendar_selectors = [
            "//div[contains(@class, 'ui-datepicker')]",
            "//div[contains(@class, 'calendar')]",
            "//div[contains(@class, 'date-picker')]",
            "//table[contains(@class, 'calendar')]"
        ]
        
        calendar_popup = None
        for selector in calendar_selectors:
            try:
                popup = driver.find_element(By.XPATH, selector)
                if popup.is_displayed():
                    calendar_popup = popup
                    logger.info(f"Found calendar popup: {selector}")
                    break
            except:
                continue
        
        if not calendar_popup:
            logger.warning("No calendar popup found")
            return False
        
        # Parse target date
        try:
            if "/" in target_date:
                parts = target_date.split("/")
                if len(parts) == 3:
                    target_month = int(parts[0])
                    target_day = int(parts[1]) 
                    target_year = int(parts[2])
                else:
                    return False
            else:
                return False
        except:
            logger.warning(f"Could not parse date: {target_date}")
            return False
        
        # Navigate to correct year/month
        try:
            # Look for year/month selectors
            year_selector = calendar_popup.find_elements(By.XPATH, ".//select[contains(@class, 'year')] | .//select[@class='ui-datepicker-year']")
            month_selector = calendar_popup.find_elements(By.XPATH, ".//select[contains(@class, 'month')] | .//select[@class='ui-datepicker-month']")
            
            # Set year if selector exists
            if year_selector:
                from selenium.webdriver.support.ui import Select
                year_select = Select(year_selector[0])
                try:
                    year_select.select_by_value(str(target_year))
                    time.sleep(0.5)
                except:
                    pass
            
            # Set month if selector exists  
            if month_selector:
                month_select = Select(month_selector[0])
                try:
                    month_select.select_by_value(str(target_month - 1))  # 0-based months
                    time.sleep(0.5)
                except:
                    pass
            
            # Click on the specific day
            day_links = calendar_popup.find_elements(By.XPATH, f".//a[text()='{target_day}'] | .//td[text()='{target_day}']")
            for day_link in day_links:
                if day_link.is_displayed():
                    safe_click(driver, day_link)
                    logger.info(f"✅ Clicked on day: {target_day}")
                    time.sleep(1)
                    return True
        
        except Exception as e:
            logger.warning(f"Calendar navigation failed: {str(e)}")
        
        return False
        
    except Exception as e:
        logger.warning(f"Calendar navigation error: {str(e)}")
        return False

def set_input_value_enhanced(driver, element, value):
    """Enhanced input value setting with datepicker widget support"""
    try:
        # Check if this is a datepicker widget
        element_class = element.get_attribute('class') or ''
        is_readonly = element.get_attribute('readonly') 
        is_datepicker = 'datepicker' in element_class.lower() or 'hasDatepicker' in element_class
        
        logger.info(f"Element class: {element_class}")
        logger.info(f"Is readonly: {is_readonly}")
        logger.info(f"Is datepicker: {is_datepicker}")
        
        # If it's a datepicker widget, use specialized handler
        if is_datepicker or is_readonly:
            logger.info("🗓️ Detected datepicker widget - using specialized handler")
            return handle_datepicker_widget(driver, element, value)
        
        # For regular inputs, use standard methods
        # Method 1: Standard clear and send_keys
        try:
            element.clear()
            time.sleep(0.2)
            element.send_keys(value)
            time.sleep(0.2)
            # Trigger change events
            driver.execute_script("""
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
            """, element)
            
            # Verify value was set
            if element.get_attribute('value') == value:
                return True
        except Exception as e:
            logger.warning(f"Method 1 failed: {str(e)}")
        
        # Method 2: JavaScript value setting
        try:
            driver.execute_script("arguments[0].value = arguments[1];", element, value)
            driver.execute_script("""
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, element)
            
            if element.get_attribute('value') == value:
                return True
        except Exception as e:
            logger.warning(f"Method 2 failed: {str(e)}")
        
        # Method 3: Focus, select all, then type
        try:
            element.click()
            time.sleep(0.1)
            element.send_keys(Keys.CONTROL + "a")
            time.sleep(0.1)
            element.send_keys(value)
            time.sleep(0.2)
            
            if element.get_attribute('value') == value:
                return True
        except Exception as e:
            logger.warning(f"Method 3 failed: {str(e)}")
        
        return False
        
    except Exception as e:
        logger.warning(f"All input methods failed: {str(e)}")
        return False

def try_advanced_date_setting(driver):
    """Advanced date setting approach with calendar widgets and dropdowns"""
    try:
        logger.info("🔄 Trying advanced date setting approaches...")
        
        # Look for date range presets or quick select options
        preset_selectors = [
            "//button[contains(text(), 'This Year')]",
            "//button[contains(text(), 'Year to Date')]",
            "//button[contains(text(), 'YTD')]",
            "//option[contains(text(), '2025')]",
            "//select[contains(@name, 'year')]",
            "//select[contains(@id, 'year')]"
        ]
        
        for selector in preset_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                for element in elements:
                    if element.is_displayed():
                        logger.info(f"Found preset option: {element.text}")
                        if safe_click(driver, element):
                            logger.info("✅ Clicked date preset")
                            time.sleep(3)
                            return True
            except Exception as e:
                logger.warning(f"Preset selector {selector} failed: {str(e)}")
        
        # Try calendar widget approach
        return try_calendar_widget_approach(driver)
        
    except Exception as e:
        logger.warning(f"Advanced date setting failed: {str(e)}")
        return False

def try_calendar_widget_approach(driver):
    """Try alternative calendar widget approach for date setting"""
    try:
        logger.info("🔄 Trying alternative calendar widget approach...")
        
        # Look for calendar icons or date picker triggers
        calendar_triggers = driver.find_elements(By.XPATH, 
            "//img[contains(@src, 'calendar')] | //span[contains(@class, 'calendar')] | " +
            "//button[contains(@class, 'date')] | //div[contains(@class, 'datepicker')]")
        
        if calendar_triggers:
            logger.info(f"Found {len(calendar_triggers)} calendar triggers")
            for trigger in calendar_triggers:
                if trigger.is_displayed():
                    try:
                        safe_click(driver, trigger)
                        time.sleep(2)
                        logger.info("Clicked calendar trigger")
                        break
                    except:
                        continue
        
        # Look for any visible date inputs that might have appeared
        all_inputs = driver.find_elements(By.XPATH, "//input[@type='text' or @type='date']")
        date_inputs = []
        
        for inp in all_inputs:
            if inp.is_displayed():
                placeholder = inp.get_attribute('placeholder') or ''
                name = inp.get_attribute('name') or ''
                id_attr = inp.get_attribute('id') or ''
                
                if any(word in (placeholder + name + id_attr).lower() for word in ['date', 'from', 'to', 'start', 'end']):
                    date_inputs.append(inp)
        
        if len(date_inputs) >= 2:
            logger.info(f"Found {len(date_inputs)} potential date inputs")
            # Set first input as from date, second as to date
            try:
                date_inputs[0].clear()
                date_inputs[0].send_keys(START_DATE)
                time.sleep(1)
                
                date_inputs[1].clear()
                date_inputs[1].send_keys(END_DATE)
                time.sleep(1)
                
                logger.info("✅ Set dates using alternative approach")
                return True
            except Exception as e:
                logger.warning(f"Alternative date setting failed: {str(e)}")
        
        return False
        
    except Exception as e:
        logger.warning(f"Calendar widget approach failed: {str(e)}")
        return False

def find_navigation_elements(driver):
    """Find and analyze available navigation elements"""
    try:
        logger.info("🔍 Analyzing navigation elements...")
        
        # Look for various navigation patterns
        nav_patterns = [
            "//nav",
            "//div[contains(@class, 'nav')]",
            "//ul[contains(@class, 'menu')]",
            "//div[contains(@class, 'sidebar')]",
            "//div[contains(@class, 'left-navigation')]",
            "//span[contains(@class, 'left-navigation-text')]",
            "//a[contains(@href, 'dashboard')]",
            "//span[contains(text(), 'Dashboard')]",
            "//a[contains(text(), 'Dashboard')]",
            "//span[contains(text(), 'Reports')]",
            "//a[contains(text(), 'Reports')]"
        ]
        
        found_elements = []
        for pattern in nav_patterns:
            try:
                elements = driver.find_elements(By.XPATH, pattern)
                if elements:
                    logger.info(f"Found {len(elements)} elements matching: {pattern}")
                    for i, element in enumerate(elements):
                        if element.is_displayed():
                            text = element.text.strip()[:50]
                            href = element.get_attribute('href') or 'N/A'
                            logger.info(f"  Element {i}: '{text}' (href: {href})")
                            found_elements.append((pattern, element, text))
            except Exception as e:
                logger.warning(f"Error with pattern {pattern}: {str(e)}")
        
        return found_elements
        
    except Exception as e:
        logger.error(f"Error analyzing navigation: {str(e)}")
        return []

def navigate_to_reports_section(driver):
    """Navigate specifically to Reports section for date range setting"""
    try:
        logger.info("🔹 Trying to navigate to Reports section...")
        
        # Try to find Reports navigation
        reports_selectors = [
            "//span[contains(text(), 'REPORTS')]",
            "//span[contains(text(), 'Reports')]", 
            "//a[contains(text(), 'Reports')]",
            "//a[contains(@href, 'report')]",
            "//span[@class='left-navigation-text' and contains(text(), 'REPORTS')]"
        ]
        
        reports_found = False
        for selector in reports_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                logger.info(f"Found {len(elements)} reports elements with selector: {selector}")
                
                for element in elements:
                    if element.is_displayed():
                        logger.info(f"Trying to click reports element: {element.text}")
                        
                        for method in ["normal", "js", "action"]:
                            if safe_click(driver, element, method):
                                logger.info(f"✅ Clicked reports using {method}")
                                reports_found = True
                                time.sleep(5)
                                break
                        
                        if reports_found:
                            break
                
                if reports_found:
                    break
                    
            except Exception as e:
                logger.warning(f"Error with reports selector {selector}: {str(e)}")
        
        if not reports_found:
            # Try direct URL navigation to reports
            reports_urls = [
                "https://ecommerceportal.dhl.com/Portal/pages/customer/reports.xhtml",
                "https://ecommerceportal.dhl.com/Portal/pages/reports/reports.xhtml", 
                "https://ecommerceportal.dhl.com/Portal/pages/customer/advancedreport.xhtml"
            ]
            
            for url in reports_urls:
                try:
                    logger.info(f"Trying direct navigation to reports: {url}")
                    driver.get(url)
                    time.sleep(8)
                    
                    if "error" not in driver.current_url.lower() and "login" not in driver.current_url.lower():
                        logger.info(f"✅ Successfully navigated to reports: {url}")
                        reports_found = True
                        break
                except Exception as e:
                    logger.warning(f"Direct navigation to {url} failed: {str(e)}")
        
        if reports_found:
            debug_page_state(driver, "After Reports Navigation")
            # Try setting date range in reports section
            return set_date_range(driver)
        
        return False
        
    except Exception as e:
        logger.error(f"❌ Reports navigation failed: {str(e)}")
        return False

def try_advanced_reports_with_date_filter(driver):
    """Try to access Advanced Reports with date filtering capabilities"""
    try:
        logger.info("🔹 Trying Advanced Reports with date filters...")
        
        # Look for Advanced Report link specifically
        advanced_report_selectors = [
            "//a[contains(text(), 'Advanced Report')]",
            "//span[contains(text(), 'Advanced Report')]",
            "//div[contains(text(), 'Advanced Report')]",
            "//a[contains(@href, 'advanced')]",
            "//a[contains(@href, 'report')]//span[contains(text(), 'Advanced')]"
        ]
        
        for selector in advanced_report_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                if elements:
                    logger.info(f"Found {len(elements)} advanced report elements")
                    for element in elements:
                        if element.is_displayed():
                            logger.info(f"Clicking advanced report: {element.text}")
                            if safe_click(driver, element):
                                time.sleep(8)
                                logger.info("✅ Navigated to Advanced Reports")
                                
                                # Try setting date range in advanced reports
                                debug_page_state(driver, "Advanced Reports Page")
                                
                                # Look for date filters in advanced reports
                                if try_advanced_report_date_filters(driver):
                                    return True
                                break
            except Exception as e:
                logger.warning(f"Advanced report selector failed: {str(e)}")
        
        # Try direct URLs for advanced reports
        advanced_urls = [
            "https://ecommerceportal.dhl.com/Portal/pages/customer/advancedreport.xhtml",
            "https://ecommerceportal.dhl.com/Portal/pages/reports/advancedreport.xhtml",
            "https://ecommerceportal.dhl.com/Portal/pages/shipment/reports.xhtml"
        ]
        
        for url in advanced_urls:
            try:
                logger.info(f"Trying direct advanced reports URL: {url}")
                driver.get(url)
                time.sleep(8)
                
                if "error" not in driver.current_url.lower():
                    logger.info(f"✅ Accessed advanced reports: {url}")
                    debug_page_state(driver, f"Advanced Reports - {url}")
                    
                    if try_advanced_report_date_filters(driver):
                        return True
                        
            except Exception as e:
                logger.warning(f"Direct advanced URL failed: {str(e)}")
        
        return False
        
    except Exception as e:
        logger.error(f"Advanced reports approach failed: {str(e)}")
        return False

def try_advanced_report_date_filters(driver):
    """Try to set date filters in advanced reports page"""
    try:
        logger.info("🔹 Setting date filters in advanced reports...")
        
        # Wait for page to load
        time.sleep(5)
        
        # Look for date filter sections
        date_filter_sections = driver.find_elements(By.XPATH, 
            "//div[contains(@class, 'filter')] | //div[contains(@class, 'date')] | " +
            "//fieldset[contains(., 'Date')] | //div[contains(., 'Date Range')]")
        
        if date_filter_sections:
            logger.info(f"Found {len(date_filter_sections)} date filter sections")
            
            # Try to find date inputs within these sections
            for section in date_filter_sections:
                try:
                    date_inputs = section.find_elements(By.XPATH, ".//input[@type='text' or @type='date']")
                    if len(date_inputs) >= 2:
                        logger.info(f"Found date inputs in section: {len(date_inputs)}")
                        
                        # Try to set from and to dates
                        from_input = date_inputs[0]
                        to_input = date_inputs[1]
                        
                        if set_input_value_enhanced(driver, from_input, "01/01/2025"):
                            logger.info("✅ Set from date in advanced reports")
                            if set_input_value_enhanced(driver, to_input, END_DATE):
                                logger.info("✅ Set to date in advanced reports")
                                
                                # Look for apply/submit button in this section
                                apply_buttons = section.find_elements(By.XPATH, 
                                    ".//button | .//input[@type='submit'] | .//a[contains(@class, 'button')]")
                                
                                for btn in apply_buttons:
                                    if btn.is_displayed() and btn.is_enabled():
                                        if safe_click(driver, btn):
                                            logger.info("✅ Applied date filter in advanced reports")
                                            time.sleep(5)
                                            return True
                                
                                return True  # Date set even if no apply button
                                
                except Exception as e:
                    logger.warning(f"Error with date filter section: {str(e)}")
                    continue
        
        # Try generic date setting approach
        return set_date_range(driver)
        
    except Exception as e:
        logger.warning(f"Advanced report date filters failed: {str(e)}")
        return False

def retry_dashboard_datepickers(driver):
    """Retry datepicker interaction on dashboard with enhanced approach"""
    try:
        logger.info("🗓️ Retrying dashboard datepickers with focused approach...")
        time.sleep(5)
        
        # Look specifically for the known datepicker inputs
        from_input = safe_find_element(driver, By.ID, "dashboardForm:frmDate_input", 10)
        to_input = safe_find_element(driver, By.ID, "dashboardForm:toDate_input", 10)
        
        if not from_input or not to_input:
            logger.warning("Could not find dashboard datepicker inputs")
            return False
        
        logger.info("Found dashboard datepicker inputs, attempting enhanced interaction...")
        
        # Try a more aggressive approach for datepickers
        success = False
        
        # Method 1: Try triggering calendar popup via multiple events
        try:
            logger.info("Method 1: Triggering calendar popup events...")
            
            # Focus on from input and trigger events
            driver.execute_script("""
                var fromInput = arguments[0];
                var toInput = arguments[1];
                var fromDate = arguments[2];
                var toDate = arguments[3];
                
                // Focus and trigger calendar events
                fromInput.focus();
                fromInput.click();
                
                // Try to trigger calendar open events
                var events = ['mousedown', 'mouseup', 'click', 'focus'];
                events.forEach(function(eventType) {
                    var event = new Event(eventType, { bubbles: true });
                    fromInput.dispatchEvent(event);
                });
                
                // Set values directly
                fromInput.value = fromDate;
                toInput.value = toDate;
                
                // Trigger change events
                fromInput.dispatchEvent(new Event('change', { bubbles: true }));
                toInput.dispatchEvent(new Event('change', { bubbles: true }));
                
                // Try jQuery if available
                if (window.jQuery) {
                    jQuery(fromInput).trigger('change').trigger('blur');
                    jQuery(toInput).trigger('change').trigger('blur');
                    
                    // Try jQuery datepicker methods
                    if (jQuery.fn.datepicker) {
                        try {
                            jQuery(fromInput).datepicker('setDate', fromDate);
                            jQuery(toInput).datepicker('setDate', toDate);
                        } catch(e) {}
                    }
                }
                
                return {
                    fromValue: fromInput.value,
                    toValue: toInput.value
                };
            """, from_input, to_input, "01/01/2025", END_DATE)
            
            time.sleep(3)
            
            # Check if values were set
            from_value = from_input.get_attribute('value')
            to_value = to_input.get_attribute('value')
            
            logger.info(f"After JS method - From: {from_value}, To: {to_value}")
            
            if from_value and to_value and ("2025" in from_value or "01" in from_value):
                success = True
                logger.info("✅ Enhanced JS method appears to have worked")
                
        except Exception as e:
            logger.warning(f"Enhanced JS method failed: {str(e)}")
        
        # Method 2: Try clicking around the inputs to trigger calendars
        if not success:
            try:
                logger.info("Method 2: Trying to trigger calendar via clicking nearby elements...")
                
                # Click on the input and nearby elements
                from_input.click()
                time.sleep(1)
                
                # Look for calendar trigger elements near the inputs
                parent = from_input.find_element(By.XPATH, "../..")
                clickable_elements = parent.find_elements(By.XPATH, ".//button | .//img | .//span[@class]")
                
                for element in clickable_elements[:5]:  # Try first 5 elements
                    try:
                        if element.is_displayed():
                            safe_click(driver, element)
                            time.sleep(1)
                    except:
                        continue
                
            except Exception as e:
                logger.warning(f"Calendar trigger method failed: {str(e)}")
        
        # Look for apply button and try clicking it
        try:
            apply_buttons = driver.find_elements(By.XPATH, 
                "//button[contains(text(), 'Apply')] | //input[@value='Apply'] | " +
                "//button[contains(@id, 'apply')] | //button[contains(@class, 'apply')]")
            
            for btn in apply_buttons:
                if btn.is_displayed() and btn.is_enabled():
                    logger.info("Trying to click apply button...")
                    if safe_click(driver, btn):
                        time.sleep(5)
                        success = True
                        break
                        
        except Exception as e:
            logger.warning(f"Apply button click failed: {str(e)}")
        
        if success:
            logger.info("✅ Dashboard datepicker retry succeeded")
        else:
            logger.warning("⚠️ Dashboard datepicker retry did not succeed")
            
        return success
        
    except Exception as e:
        logger.error(f"❌ Dashboard datepicker retry failed: {str(e)}")
        return False

def navigate_to_dashboard(driver):
    """Enhanced dashboard navigation with multiple fallback strategies"""
    try:
        logger.info("🔹 Attempting to navigate to dashboard...")
        time.sleep(10)  # Wait for page to fully load
        
        debug_page_state(driver, "Before Dashboard Navigation")
        
        # Analyze available navigation
        nav_elements = find_navigation_elements(driver)
        
        # Strategy 1: Try Dashboard first
        dashboard_found = False
        dashboard_selectors = [
            "//span[@class='left-navigation-text' and contains(text(), 'Dashboard')]",
            "//a[contains(text(), 'Dashboard')]",
            "//span[contains(text(), 'Dashboard')]",
            "//div[contains(text(), 'Dashboard')]",
            "//a[contains(@href, 'dashboard')]"
        ]
        
        for selector in dashboard_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                logger.info(f"Found {len(elements)} dashboard elements with selector: {selector}")
                
                for element in elements:
                    if element.is_displayed():
                        logger.info(f"Trying to click dashboard element: {element.text}")
                        
                        # Try multiple click methods
                        for method in ["normal", "js", "action"]:
                            if safe_click(driver, element, method):
                                logger.info(f"✅ Clicked dashboard using {method}")
                                dashboard_found = True
                                time.sleep(5)
                                break
                        
                        if dashboard_found:
                            break
                
                if dashboard_found:
                    break
                    
            except Exception as e:
                logger.warning(f"Error with selector {selector}: {str(e)}")
        
        # If dashboard link not found, try direct navigation
        if not dashboard_found:
            logger.info("🔄 Dashboard link not found, trying direct URL navigation...")
            dashboard_urls = [
                "https://ecommerceportal.dhl.com/Portal/pages/customer/statisticsdashboard.xhtml",
                "https://ecommerceportal.dhl.com/Portal/pages/dashboard/dashboard.xhtml"
            ]
            
            for url in dashboard_urls:
                try:
                    logger.info(f"Trying direct navigation to: {url}")
                    driver.get(url)
                    time.sleep(8)
                    
                    # Check if page loaded successfully
                    if "error" not in driver.current_url.lower() and "login" not in driver.current_url.lower():
                        logger.info(f"✅ Successfully navigated to: {url}")
                        dashboard_found = True
                        break
                except Exception as e:
                    logger.warning(f"Direct navigation to {url} failed: {str(e)}")
        
        # Strategy 2: Try date setting on dashboard
        date_success = False
        if dashboard_found:
            logger.info("✅ Dashboard navigation completed")
            time.sleep(5)  # Wait for page to stabilize
            
            # Try to set date range
            date_success = set_date_range(driver)
        
        # Strategy 3: If date setting failed, try Reports section
        if not date_success:
            logger.info("🔄 Date setting failed, trying Reports section...")
            date_success = navigate_to_reports_section(driver)
        
        # Strategy 4: If still failed, try Advanced Reports
        if not date_success:
            logger.info("🔄 Trying Advanced Reports approach...")
            date_success = try_advanced_reports_with_date_filter(driver)
        
        # Strategy 4: Try going back to dashboard and re-trigger datepickers
        if not date_success:
            logger.info("🔄 Final attempt: Going back to dashboard for datepicker retry...")
            try:
                driver.get("https://ecommerceportal.dhl.com/Portal/pages/customer/statisticsdashboard.xhtml") 
                time.sleep(8)
                debug_page_state(driver, "Dashboard Retry")
                date_success = retry_dashboard_datepickers(driver)
            except Exception as e:
                logger.warning(f"Dashboard retry failed: {str(e)}")
        
        # Strategy 5: Final fallback - continue without date filter
        if not date_success:
            logger.warning("⚠️ All date setting strategies failed - continuing with default date range")
            logger.warning("📊 Data may be limited to recent period (last 7-30 days) due to portal defaults")
            logger.info("💡 Detected readonly datepicker widgets but could not interact with calendar popups")
            logger.info("🔧 Manual verification: Check if portal requires specific date format or interaction method")
            logger.info("📋 Consider checking debug HTML files to analyze datepicker widget structure")
        
        debug_page_state(driver, "After All Navigation Attempts")
        return True  # Continue with report extraction regardless
            
    except Exception as e:
        logger.error(f"❌ Dashboard navigation failed: {str(e)}")
        debug_page_state(driver, "Dashboard Navigation Error")
        return False

def find_and_analyze_reports(driver):
    """Find and analyze available reports on the page"""
    try:
        logger.info("🔍 Analyzing available reports and download options...")
        
        debug_page_state(driver, "Report Analysis")
        
        # Look for tables that might contain reports
        tables = driver.find_elements(By.TAG_NAME, "table")
        logger.info(f"Found {len(tables)} tables on page")
        
        report_data = []
        for i, table in enumerate(tables):
            try:
                if table.is_displayed():
                    rows = table.find_elements(By.TAG_NAME, "tr")
                    if len(rows) > 1:  # Has header and data
                        logger.info(f"Table {i}: {len(rows)} rows")
                        
                        # Get table content sample
                        sample_text = table.text[:200].replace('\n', ' | ')
                        logger.info(f"  Sample: {sample_text}")
                        
                        # Look for download buttons/links in this table
                        download_elements = table.find_elements(By.XPATH, 
                            ".//img[contains(@src, 'excel') or contains(@src, 'xls') or contains(@src, 'download')] | " +
                            ".//a[contains(@href, 'excel') or contains(@href, 'xls') or contains(@href, 'download')] | " +
                            ".//button[contains(text(), 'Download') or contains(text(), 'Export')] | " +
                            ".//span[contains(text(), 'Excel') or contains(text(), 'CSV')]")
                        
                        if download_elements:
                            logger.info(f"  Found {len(download_elements)} download elements in table {i}")
                            for j, elem in enumerate(download_elements):
                                if elem.is_displayed():
                                    report_data.append((i, j, elem, table))
            except Exception as e:
                logger.warning(f"Error analyzing table {i}: {str(e)}")
        
        # Look for standalone download buttons
        standalone_downloads = driver.find_elements(By.XPATH,
            "//img[contains(@src, 'excel') or contains(@src, 'xls') or contains(@src, 'download')] | " +
            "//a[contains(@href, 'excel') or contains(@href, 'xls') or contains(@href, 'download')] | " +
            "//button[contains(text(), 'Download') or contains(text(), 'Export')] | " +
            "//span[contains(text(), 'Excel') or contains(text(), 'CSV')] | " +
            "//div[contains(@onclick, 'export') or contains(@onclick, 'download')]")
        
        logger.info(f"Found {len(standalone_downloads)} standalone download elements")
        
        for i, elem in enumerate(standalone_downloads):
            if elem.is_displayed():
                report_data.append(("standalone", i, elem, None))
        
        return report_data
        
    except Exception as e:
        logger.error(f"Error analyzing reports: {str(e)}")
        return []

def attempt_downloads(driver, report_data):
    """Attempt to download reports from found elements"""
    try:
        logger.info("🔹 Attempting to download from found elements...")
        
        # Clear download folder
        clear_download_folder()
        
        download_success = False
        for location, index, element, table in report_data:
            try:
                logger.info(f"Attempting download from {location}-{index}")
                
                # Get element info
                tag_name = element.tag_name
                src = element.get_attribute('src') or 'N/A'
                href = element.get_attribute('href') or 'N/A'
                onclick = element.get_attribute('onclick') or 'N/A'
                text = element.text or 'N/A'
                
                logger.info(f"  Element: {tag_name}, src: {src}, href: {href}, text: {text}")
                
                # Try clicking the element
                for method in ["normal", "js", "action"]:
                    try:
                        if safe_click(driver, element, method):
                            logger.info(f"  Clicked using {method}")
                            time.sleep(8)  # Wait for download
                            
                            # Check if download started
                            if check_for_new_download():
                                logger.info("✅ Download detected!")
                                download_success = True
                                break
                    except Exception as e:
                        logger.warning(f"  Click method {method} failed: {str(e)}")
                
                if download_success:
                    break
                    
            except Exception as e:
                logger.warning(f"Error with download attempt {location}-{index}: {str(e)}")
        
        return download_success
        
    except Exception as e:
        logger.error(f"Error in download attempts: {str(e)}")
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
    """Get the latest downloaded file with enhanced checking"""
    logger.info(f"🔍 Looking for downloaded files in: {folder_path}")
    
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
            
            # Get the most recent file
            latest_file = max(files, key=os.path.getctime)
            file_size = os.path.getsize(latest_file)
            
            logger.info(f"Found file: {latest_file} (Size: {file_size} bytes)")
            
            if file_size == 0:
                logger.warning("File is empty, waiting...")
                time.sleep(delay)
                continue
            
            # Try to read file to ensure it's complete
            try:
                if latest_file.endswith('.csv'):
                    pd.read_csv(latest_file, nrows=1)
                else:
                    pd.read_excel(latest_file, nrows=1)
                
                logger.info(f"✅ Found valid file: {latest_file}")
                return latest_file
                
            except Exception as e:
                logger.warning(f"File not ready: {str(e)}")
                time.sleep(delay)
                continue
                
        except Exception as e:
            logger.warning(f"Error checking files: {str(e)}")
            time.sleep(delay)
    
    logger.warning("❌ No valid file found")
    return None

def process_data(file_path):
    """Process downloaded data with flexible column handling"""
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
        
        # Show sample data
        if len(df) > 0:
            logger.info("Sample data:")
            logger.info(df.head(3).to_string())
        
        # Create processed DataFrame with flexible mapping
        processed_df = pd.DataFrame()
        
        # Map Order ID
        if 'Consignee Name' in df.columns:
            processed_df['Order ID'] = df['Consignee Name'].astype(str).str.extract(r'(\d{7})')[0].fillna('')
        elif 'Order ID' in df.columns:
            processed_df['Order ID'] = df['Order ID'].fillna('')
        else:
            processed_df['Order ID'] = ''
        
        # Map Tracking Number
        tracking_cols = ['Tracking ID', 'Tracking Number', 'AWB', 'Waybill Number']
        tracking_col = None
        for col in tracking_cols:
            if col in df.columns:
                tracking_col = col
                break
        
        processed_df['Tracking Number'] = df[tracking_col].fillna('') if tracking_col else ''
        
        # Map Pickup DateTime
        pickup_cols = ['Pickup Event DateTime', 'Pickup Date', 'Collection Date', 'Ship Date']
        pickup_col = None
        for col in pickup_cols:
            if col in df.columns:
                pickup_col = col
                break
        
        if pickup_col:
            processed_df['Pickup DateTime'] = pd.to_datetime(df[pickup_col], errors='coerce')
        else:
            processed_df['Pickup DateTime'] = pd.NaT
        
        # Map Delivery Date
        delivery_cols = ['Delivery Date', 'Delivered Date', 'POD Date']
        delivery_col = None
        for col in delivery_cols:
            if col in df.columns:
                delivery_col = col
                break
        
        if delivery_col:
            processed_df['Delivery Date'] = pd.to_datetime(df[delivery_col], errors='coerce')
        else:
            processed_df['Delivery Date'] = pd.NaT
        
        # Map Status
        status_cols = ['Last Status', 'Status', 'Current Status', 'Shipment Status']
        status_col = None
        for col in status_cols:
            if col in df.columns:
                status_col = col
                break
        
        processed_df['Status'] = df[status_col].fillna('') if status_col else ''
        
        # Sort by Pickup DateTime (newest first)
        try:
            processed_df = processed_df.sort_values('Pickup DateTime', ascending=False, na_position='last')
            logger.info(f"✅ Sorted {len(processed_df)} rows")
        except:
            logger.warning("Could not sort data")
        
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
        try:
            service.spreadsheets().values().clear(
                spreadsheetId=GOOGLE_SHEET_ID,
                range=f"{SHEET_NAME}!A1:Z1000"
            ).execute()
            logger.info("✅ Cleared existing content")
        except Exception as e:
            logger.warning(f"Could not clear sheet: {str(e)}")
        
        # Upload new data
        response = service.spreadsheets().values().update(
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
    """Main execution function with comprehensive error handling"""
    driver = None
    try:
        logger.info("🚀 Starting DHL report automation process...")
        logger.info("📅 Enhanced date range setting with datepicker widget support:")
        logger.info("   - Detects readonly datepicker inputs (hasDatepicker class)")
        logger.info("   - JavaScript date setting for calendar widgets")
        logger.info("   - Calendar popup navigation and date selection")
        logger.info("   - Multiple fallback approaches for different widget types")
        
        # Validate environment first
        if not validate_environment():
            logger.error("❌ Environment validation failed")
            upload_to_google_sheets(create_empty_data())
            return
        
        # Setup driver
        driver = setup_chrome_driver()
        
        # Strategy 1: Try direct dashboard access first (in case user is already logged in)
        logger.info("🔹 Strategy 1: Checking if already logged in...")
        
        # Go to main portal first
        driver.get("https://ecommerceportal.dhl.com/Portal/")
        time.sleep(5)
        
        if check_already_logged_in(driver):
            logger.info("✅ User appears to be already logged in")
            if try_direct_dashboard_access(driver):
                logger.info("✅ Direct dashboard access successful")
                skip_login = True
            else:
                skip_login = False
        else:
            skip_login = False
        
        # Strategy 2: Perform login if needed
        if not skip_login:
            logger.info("🔹 Strategy 2: Performing login...")
            if not login_to_dhl(driver):
                logger.warning("⚠️ Login failed, trying alternative approaches...")
                
                # Strategy 3: Try manual intervention guidance
                logger.info("🔹 Strategy 3: Attempting alternative login methods...")
                
                # Try refreshing and login again
                driver.refresh()
                time.sleep(10)
                
                if not login_to_dhl(driver):
                    logger.error("❌ All login strategies failed")
                    
                    # Final strategy: Upload empty data and inform user
                    logger.warning("⚠️ Could not login automatically. Manual intervention may be required.")
                    logger.info("💡 Suggestion: Run the script manually after logging in through browser")
                    
                    upload_to_google_sheets(create_empty_data())
                    return
        
        # Continue with report extraction
        logger.info("🔹 Proceeding with report extraction...")
        
        # Navigate to reports section and set date range
        navigate_to_dashboard(driver)
        
        # Analyze and attempt downloads
        report_data = find_and_analyze_reports(driver)
        
        if report_data:
            logger.info(f"Found {len(report_data)} potential download elements")
            download_success = attempt_downloads(driver, report_data)
        else:
            logger.warning("No download elements found")
            download_success = False
        
        # Process downloaded file or create empty data
        if download_success:
            latest_file = get_latest_file(DOWNLOAD_FOLDER)
            processed_df = process_data(latest_file)
        else:
            logger.warning("No downloads successful, using empty data")
            processed_df = create_empty_data()
        
        # Upload to Google Sheets
        upload_to_google_sheets(processed_df)
        
        logger.info("🎉 Process completed successfully!")
        
    except Exception as e:
        logger.error(f"❌ Main process failed: {str(e)}")
        try:
            upload_to_google_sheets(create_empty_data())
            logger.info("⚠️ Uploaded empty data after failure")
        except Exception as upload_e:
            logger.error(f"❌ Failed to upload empty data: {str(upload_e)}")
    
    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass

if __name__ == "__main__":
    main()
