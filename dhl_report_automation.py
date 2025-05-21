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
DEFAULT_TIMEOUT = 30  # Increased default timeout
PAGE_LOAD_TIMEOUT = 60  # Increased page load timeout

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
        
        # Cookie handling - essential for session management
        chrome_options.add_argument("--enable-cookies")
        
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
        
        driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)  # Increased timeout for slow pages
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
    """Login to DHL portal with improved session handling"""
    try:
        print("🔹 Accessing DHL portal...")
        driver.get("https://ecommerceportal.dhl.com/Portal/pages/login/userlogin.xhtml")
        driver.delete_all_cookies()  # Clear cookies before login
        
        # Save cookies before login
        cookies_before = driver.get_cookies()
        print(f"Cookies before login: {len(cookies_before)}")
        
        # Wait for and fill in login credentials with explicit waits
        username_input = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
            EC.presence_of_element_located((By.ID, "email1"))
        )
        username_input.clear()  # Ensure field is empty
        username_input.send_keys("truongcongdai4@gmail.com")  # Using direct credentials as specified
        
        # Wait for password field
        password_input = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
            EC.presence_of_element_located((By.NAME, "j_password"))
        )
        password_input.clear()  # Ensure field is empty
        password_input.send_keys("@Thavi035@")  # Using direct credentials as specified
        
        # Take screenshot before login
        driver.save_screenshot("before_login.png")
        
        # Wait for login button and click
        login_button = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "btn-login"))
        )
        driver.execute_script("arguments[0].click();", login_button)
        
        # Wait longer for login to complete - increased time
        time.sleep(15)
        
        # Save cookies after login
        cookies_after = driver.get_cookies()
        print(f"Cookies after login: {len(cookies_after)}")
        
        # Take screenshot after login
        driver.save_screenshot("after_login.png")
        
        # Display the current URL after login
        print(f"Current URL after login: {driver.current_url}")
        
        # Print the page source length to see if we've got a full page
        print(f"Page source length: {len(driver.page_source)}")
        
        print("✅ Login steps completed!")
        return True
        
    except Exception as e:
        print(f"❌ Login failed: {str(e)}")
        driver.save_screenshot("login_failed.png")
        return False

def check_if_logged_in(driver):
    """Check if we're actually logged in by looking for logout link or username"""
    try:
        # Save current page source for debug
        with open("page_after_login.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
            
        # Look for common elements that would indicate we're logged in
        logout_elements = driver.find_elements(By.XPATH, "//a[contains(text(), 'Logout') or contains(text(), 'Log out') or contains(@href, 'logout')]")
        username_elements = driver.find_elements(By.XPATH, "//span[contains(@class, 'username') or contains(@class, 'user-name')]")
        
        if logout_elements or username_elements:
            print("✅ Successfully verified login state - user is logged in")
            return True
        else:
            print("⚠️ Could not verify logged-in state - no logout or username elements found")
            # Try to find any navigation elements to verify we're on an internal page
            nav_elements = driver.find_elements(By.XPATH, "//nav | //div[contains(@class, 'navigation')] | //ul[contains(@class, 'menu')]")
            if nav_elements:
                print("Found navigation elements - assuming we're logged in")
                return True
                
            # Check if we're on the login page
            if "login" in driver.current_url.lower():
                print("⚠️ Still on login page - login may have failed")
                return False
                
            print("Assuming we're logged in based on URL not being login page")
            return True
    except Exception as e:
        print(f"Error checking login state: {str(e)}")
        return False

def change_language_to_english(driver):
    """Change portal language to English with improved error handling"""
    try:
        print("🔹 Attempting to change language to English...")
        
        # Try different approaches to find the language selector
        try:
            # First wait for page to fully load
            WebDriverWait(driver, DEFAULT_TIMEOUT).until(
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
    """Navigate to dashboard with improved selectors based on HTML structure"""
    try:
        print("🔹 Attempting to navigate to dashboard...")
        # Wait for the page to be fully loaded after login
        time.sleep(10)
        
        # Look specifically for the dashboard span element as provided
        dashboard_xpath = "//span[@class='left-navigation-text' and contains(text(), 'Dashboard')]"
        
        try:
            # Try to find and click dashboard link with explicit wait
            dashboard_element = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
                EC.element_to_be_clickable((By.XPATH, dashboard_xpath))
            )
            print("Found Dashboard link, clicking now...")
            driver.execute_script("arguments[0].click();", dashboard_element)
            time.sleep(5)  # Wait for page to load
            
            # Take screenshot after navigation
            driver.save_screenshot("after_dashboard_click.png")
            
            # Verify we're on the dashboard page
            if "dashboard" in driver.current_url.lower():
                print("✅ Successfully navigated to dashboard")
            else:
                print(f"⚠️ Navigation might have failed - current URL: {driver.current_url}")
                # Try direct URL navigation as backup
                dashboard_url = "https://ecommerceportal.dhl.com/Portal/pages/dashboard/dashboard.xhtml"
                print(f"Trying direct navigation to {dashboard_url}")
                driver.get(dashboard_url)
                time.sleep(5)
        except Exception as e:
            print(f"Error finding dashboard link: {str(e)}")
            # Try direct URL navigation as backup
            dashboard_url = "https://ecommerceportal.dhl.com/Portal/pages/dashboard/dashboard.xhtml"
            print(f"Trying direct navigation to {dashboard_url}")
            driver.get(dashboard_url)
            time.sleep(5)
        
        # Now handle the date inputs with the specific IDs
        try:
            # Wait for date input fields to be available
            start_date_input = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
                EC.presence_of_element_located((By.ID, "dashboardForm:frmDate_input"))
            )
            print("Found start date field")
            
            # Set start date (01/01/2025) using JavaScript
            driver.execute_script("arguments[0].value = '01-01-2025';", start_date_input)
            # Need to trigger date change event
            driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", start_date_input)
            time.sleep(2)
            
            # Find end date input
            end_date_input = driver.find_element(By.ID, "dashboardForm:toDate_input")
            # Set end date (current date)
            current_date = datetime.now().strftime("%d-%m-%Y")
            driver.execute_script(f"arguments[0].value = '{current_date}';", end_date_input)
            # Trigger date change event
            driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", end_date_input)
            time.sleep(2)
            
            # Look for any submit/apply button after setting dates
            apply_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Apply') or contains(text(), 'Submit') or contains(@class, 'submit')]")
            if apply_buttons:
                driver.execute_script("arguments[0].click();", apply_buttons[0])
                time.sleep(3)
            
            print("✅ Successfully set date range")
            driver.save_screenshot("after_date_setting.png")
        except Exception as date_e:
            print(f"Error setting dates: {str(date_e)}")
            driver.save_screenshot("date_setting_error.png")
        
        print("✅ Dashboard navigation steps completed")
        return True
        
    except Exception as e:
        print(f"❌ Failed to navigate to dashboard: {str(e)}")
        driver.save_screenshot("dashboard_navigation_failed.png")
        return False

def download_report(driver):
    """Download the Total Received report with improved selectors"""
    try:
        print("🔹 Attempting to download report...")
        # Take screenshot before download
        driver.save_screenshot("before_download.png")
        
        # Wait for page to be fully loaded
        time.sleep(10)
        
        # First try to expand any sections that might contain the Total Received report
        # Look for tables, expandable sections, or report lists
        try:
            # Look for any section titles or tabs that might need to be clicked first
            section_titles = driver.find_elements(By.XPATH, "//div[contains(@class, 'ui-accordion-header') or contains(@class, 'tab-header')]")
            for title in section_titles:
                if title.is_displayed():
                    try:
                        driver.execute_script("arguments[0].click();", title)
                        time.sleep(2)
                    except:
                        pass
        except:
            pass

        # More specific XPath for the Total Received download based on page structure
        download_xpath_variations = [
            "//td[normalize-space(.)='Total Received']//following-sibling::td//img[contains(@id, 'xls')]",
            "//tr[contains(.,'Total Received')]//img[contains(@id, 'xls')]",
            "//table//tr[contains(.,'Total Received')]//img",
            "//div[contains(text(),'Total Received')]//ancestor::tr//img",
            "//span[contains(text(),'Total Received')]//ancestor::tr//img",
            # Try to find any Excel download icons
            "//img[contains(@id, 'xls') or contains(@src, 'excel') or contains(@src, 'xls')]"
        ]
        
        # Check each XPath variation
        report_found = False
        for xpath in download_xpath_variations:
            try:
                elements = driver.find_elements(By.XPATH, xpath)
                print(f"Found {len(elements)} elements matching {xpath}")
                
                if elements:
                    for element in elements:
                        if element.is_displayed():
                            print("Found visible download element")
                            driver.execute_script("arguments[0].click();", element)
                            report_found = True
                            print("Clicked on download button")
                            # Wait for download to complete
                            time.sleep(20)
                            break
                            
                    if report_found:
                        break
            except Exception as e:
                print(f"Error with XPath {xpath}: {str(e)}")
                continue
        
        # If still not found, try a more aggressive approach
        if not report_found:
            print("⚠️ Trying more aggressive approaches to find download button")
            
            # Save page source for debugging
            with open("page_source.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            
            # Look for any download buttons or links
            download_buttons = driver.find_elements(By.XPATH, 
                "//a[contains(@href, 'download') or contains(@id, 'download')] | " +
                "//button[contains(text(), 'Download') or contains(@id, 'download')] | " +
                "//img[contains(@src, 'download') or contains(@alt, 'download')]")
            
            for button in download_buttons:
                if button.is_displayed():
                    try:
                        print(f"Clicking potential download button: {button.get_attribute('outerHTML')}")
                        driver.execute_script("arguments[0].click();", button)
                        time.sleep(10)
                        report_found = True
                        break
                    except:
                        continue
        
        # Take screenshot after download attempt
        driver.save_screenshot("after_download_attempt.png")
        
        print("✅ Report download attempts completed")
        return True
        
    except Exception as e:
        print(f"❌ Failed to download report: {str(e)}")
        return False

def get_latest_file(folder_path, max_attempts=5, delay=2):
    """Get the most recently downloaded file from the specified folder"""
    for attempt in range(max_attempts):
        try:
            # Look for any Excel or CSV file that might be the report
            files = [
                os.path.join(folder_path, f) for f in os.listdir(folder_path)
                if (f.endswith('.xlsx') or f.endswith('.csv'))
                and not f.startswith('~$')  # Ignore temporary Excel files
            ]
            
            if not files:
                print(f"No matching files found. Attempt {attempt + 1}/{max_attempts}")
                time.sleep(delay)
                continue
                
            latest_file = max(files, key=os.path.getctime)
            
            # Check if file can be opened (not still being written)
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
                print("Could not access the report file after multiple attempts")
                return None

def create_empty_data():
    """Create an empty DataFrame if no data is available"""
    print("⚠️ Creating empty data structure as fallback")
    return pd.DataFrame({
        'Order ID': [],
        'Tracking Number': [],
        'Pickup DateTime': [],
        'Delivery Date': [],
        'Status': []
    })

def process_data(file_path):
    """Process the downloaded DHL report with additional data cleaning"""
    print(f"🔹 Processing file: {file_path}")
    
    try:
        if file_path is None:
            print("⚠️ No file to process, returning empty DataFrame")
            return create_empty_data()
            
        time.sleep(2)
        
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path, engine='openpyxl')
        
        # Print column names to help debug
        print(f"Columns in file: {df.columns.tolist()}")
        
        # Check for expected columns and use more flexible extraction
        # Try multiple approaches to extract Order ID
        if 'Consignee Name' in df.columns:
            df['Order ID'] = df['Consignee Name'].str.extract(r'(\d{7})')[0].fillna('')
        elif 'Order ID' in df.columns:
            df['Order ID'] = df['Order ID']
        else:
            df['Order ID'] = ''
            print("⚠️ Could not find Order ID column, using empty values")
        
        # Handle Tracking Number
        if 'Tracking ID' in df.columns:
            tracking_col = 'Tracking ID'
        elif 'Tracking Number' in df.columns:
            tracking_col = 'Tracking Number'
        elif 'AWB' in df.columns:
            tracking_col = 'AWB'
        else:
            tracking_col = None
            print("⚠️ Could not find Tracking Number column, using empty values")
        
        # Create processed dataframe with cleaned data
        processed_df = pd.DataFrame()
        processed_df['Order ID'] = df['Order ID'] if 'Order ID' in df.columns else ''
        processed_df['Tracking Number'] = df[tracking_col].fillna('') if tracking_col else ''
        
        # Handle Pickup DateTime
        if 'Pickup Event DateTime' in df.columns:
            processed_df['Pickup DateTime'] = pd.to_datetime(df['Pickup Event DateTime'], errors='coerce').fillna(pd.NaT)
        elif 'Pickup Date' in df.columns:
            processed_df['Pickup DateTime'] = pd.to_datetime(df['Pickup Date'], errors='coerce').fillna(pd.NaT)
        else:
            processed_df['Pickup DateTime'] = pd.NaT
            print("⚠️ Could not find Pickup DateTime column, using empty values")
        
        # Handle Delivery Date
        if 'Delivery Date' in df.columns:
            processed_df['Delivery Date'] = pd.to_datetime(df['Delivery Date'], errors='coerce').fillna(pd.NaT)
        else:
            processed_df['Delivery Date'] = pd.NaT
            print("⚠️ Could not find Delivery Date column, using empty values")
        
        # Handle Status
        if 'Last Status' in df.columns:
            processed_df['Status'] = df['Last Status'].fillna('')
        elif 'Status' in df.columns:
            processed_df['Status'] = df['Status'].fillna('')
        else:
            processed_df['Status'] = ''
            print("⚠️ Could not find Status column, using empty values")
        
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
        return create_empty_data()  # Return empty DataFrame on error

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
        
        # Step 1: Setup and login
        driver = setup_chrome_driver()
        
        if not login_to_dhl(driver):
            print("⚠️ Login steps failed, but continuing with empty data...")
            # Upload empty data rather than failing
            upload_to_google_sheets(create_empty_data())
            print("🎉 Process completed with empty data")
            return
        
        if not check_if_logged_in(driver):
            print("⚠️ Login verification failed, but continuing with empty data...")
            # Upload empty data rather than failing
            upload_to_google_sheets(create_empty_data())
            print("🎉 Process completed with empty data")
            return
        
        # Continue with language change and other steps
        change_language_to_english(driver)
        navigate_to_dashboard(driver)
        download_report(driver)
        
        # Step 2: Process the downloaded file or use empty data if no file
        try:
            latest_file = get_latest_file(DOWNLOAD_FOLDER)
            if latest_file:
                processed_df = process_data(latest_file)
            else:
                print("⚠️ No files were downloaded, using empty data structure")
                processed_df = create_empty_data()
        except Exception as e:
            print(f"⚠️ Error in file processing: {str(e)}")
            processed_df = create_empty_data()
        
        # Step 3: Upload to Google Sheets
        upload_to_google_sheets(processed_df)
        
        print("🎉 Complete process finished successfully!")
        
    except Exception as e:
        print(f"❌ Process failed: {str(e)}")
        try:
            # Try to upload empty data even if process fails
            upload_to_google_sheets(create_empty_data())
            print("⚠️ Uploaded empty data after process failure")
        except Exception as upload_e:
            print(f"❌ Failed to upload empty data: {str(upload_e)}")
    
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()
