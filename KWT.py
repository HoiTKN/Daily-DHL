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
GOOGLE_SHEET_ID = "1bxu6NWsG5dbYzhNsn5-bZUM6K6jB5-XpcNPGHUq1HJs"
SHEET_NAME = "Sheet1"  # Update this if your sheet has a different name
SERVICE_ACCOUNT_FILE = 'service_account.json'  # Will be created by GitHub Actions
DOWNLOAD_FOLDER = os.getcwd()  # Use current directory for GitHub Actions
DEFAULT_TIMEOUT = 30  # Default timeout
PAGE_LOAD_TIMEOUT = 60  # Page load timeout
IMPLICIT_WAIT = 10  # Implicit wait time

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
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        
        # Additional options for stability
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-infobars')
        chrome_options.add_argument('--ignore-certificate-errors')
        
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
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
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
        
        driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
        
        # Execute script to remove webdriver flag
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
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

def login_to_postaplus(driver):
    """Login to PostaPlus portal"""
    try:
        print("🔹 Accessing PostaPlus portal...")
        driver.get("https://etrack.postaplus.net/CustomerPortal/Login.aspx")
        
        # Wait for page to load completely
        time.sleep(5)
        
        # Take screenshot to see what page we're on
        driver.save_screenshot("login_page_loaded.png")
        print(f"Current URL: {driver.current_url}")
        print(f"Page title: {driver.title}")
        
        # Clear cookies after page load
        driver.delete_all_cookies()
        
        # Try multiple approaches to find username field
        username_input = None
        try:
            # First try by ID
            username_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "txtusername"))
            )
            print("✅ Found username field by ID")
        except:
            try:
                # Try by name attribute
                username_input = driver.find_element(By.NAME, "txtusername")
                print("✅ Found username field by name")
            except:
                try:
                    # Try by placeholder
                    username_input = driver.find_element(By.XPATH, "//input[@placeholder='User ID']")
                    print("✅ Found username field by placeholder")
                except:
                    print("❌ Could not find username field")
                    # Log page source for debugging
                    with open("page_source.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
                    raise Exception("Username field not found")
        
        # Fill username
        username_input.clear()
        username_input.send_keys("CR25005121")
        print("✅ Entered username")
        
        # Find password field
        password_input = None
        try:
            password_input = driver.find_element(By.ID, "txtpass")
            print("✅ Found password field by ID")
        except:
            try:
                password_input = driver.find_element(By.NAME, "txtpass")
                print("✅ Found password field by name")
            except:
                password_input = driver.find_element(By.XPATH, "//input[@placeholder='Password']")
                print("✅ Found password field by placeholder")
        
        # Fill password
        password_input.clear()
        password_input.send_keys("levelupvn@1234")
        print("✅ Entered password")
        
        # Take screenshot before login
        driver.save_screenshot("before_login.png")
        
        # Find and click login button
        login_button = None
        try:
            login_button = driver.find_element(By.ID, "btnLogin")
            print("✅ Found login button by ID")
        except:
            try:
                login_button = driver.find_element(By.NAME, "btnLogin")
                print("✅ Found login button by name")
            except:
                login_button = driver.find_element(By.XPATH, "//input[@value='Login']")
                print("✅ Found login button by value")
        
        # Click login button using JavaScript
        driver.execute_script("arguments[0].click();", login_button)
        print("✅ Clicked login button")
        
        # Wait for login to complete
        time.sleep(10)
        
        # Check if login was successful by looking for URL change or login elements
        current_url = driver.current_url
        if "login" not in current_url.lower() or current_url != "https://etrack.postaplus.net/CustomerPortal/Login.aspx":
            print("✅ Login appears successful - URL changed")
        else:
            print("⚠️ Still on login page, checking for error messages")
            # Check for any error messages
            error_elements = driver.find_elements(By.XPATH, "//span[contains(@class, 'error')] | //div[contains(@class, 'error')] | //label[contains(@class, 'error')]")
            if error_elements:
                for elem in error_elements:
                    if elem.text:
                        print(f"Error message found: {elem.text}")
        
        # Take screenshot after login
        driver.save_screenshot("after_login.png")
        
        # Display the current URL after login
        print(f"Current URL after login: {driver.current_url}")
        print(f"Page title after login: {driver.title}")
        
        print("✅ Login steps completed!")
        return True
        
    except Exception as e:
        print(f"❌ Login failed: {str(e)}")
        print(f"Error type: {type(e).__name__}")
        driver.save_screenshot("login_failed.png")
        
        # Additional debugging info
        try:
            print(f"Current URL at failure: {driver.current_url}")
            print(f"Page title at failure: {driver.title}")
            # Save page source for debugging
            with open("error_page_source.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
        except:
            pass
            
        return False

def navigate_to_reports(driver):
    """Navigate to reports section and download shipment report"""
    try:
        print("🔹 Navigating to reports section...")
        time.sleep(5)
        
        # Click on REPORTS in sidebar
        reports_xpath = "//a[contains(text(), 'REPORTS')]"
        try:
            reports_element = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
                EC.element_to_be_clickable((By.XPATH, reports_xpath))
            )
            driver.execute_script("arguments[0].click();", reports_element)
            time.sleep(3)
            print("✅ Clicked on REPORTS menu")
        except Exception as e:
            print(f"⚠️ Could not find REPORTS menu: {str(e)}")
            driver.save_screenshot("reports_menu_error.png")
        
        # Click on My Shipments Report
        shipments_xpath = "//a[contains(text(), 'My Shipments Report')]"
        try:
            shipments_element = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
                EC.element_to_be_clickable((By.XPATH, shipments_xpath))
            )
            driver.execute_script("arguments[0].click();", shipments_element)
            time.sleep(5)
            print("✅ Clicked on My Shipments Report")
        except Exception as e:
            print(f"⚠️ Could not find My Shipments Report: {str(e)}")
            # Try direct navigation to the report page
            driver.get("https://etrack.postaplus.net/CustomerPortal/CustCustomerExcelExportReport.aspx")
            time.sleep(5)
        
        driver.save_screenshot("after_navigation.png")
        print("✅ Navigation to reports completed")
        return True
        
    except Exception as e:
        print(f"❌ Failed to navigate to reports: {str(e)}")
        return False

def set_dates_and_download(driver):
    """Set date range and download the report"""
    try:
        print("🔹 Setting date range and downloading report...")
        
        # Set from date (01/05/2025)
        try:
            from_date_input = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtfromdate_I"))
            )
            driver.execute_script("arguments[0].value = '01/05/2025';", from_date_input)
            driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", from_date_input)
            time.sleep(2)
            print("✅ Set from date to 01/05/2025")
        except Exception as e:
            print(f"⚠️ Error setting from date: {str(e)}")
        
        # Set to date (current date)
        try:
            to_date_input = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txttodate_I"))
            )
            current_date = datetime.now().strftime("%d/%m/%Y")
            driver.execute_script(f"arguments[0].value = '{current_date}';", to_date_input)
            driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", to_date_input)
            time.sleep(2)
            print(f"✅ Set to date to {current_date}")
        except Exception as e:
            print(f"⚠️ Error setting to date: {str(e)}")
        
        # Click Load button
        try:
            load_button = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
                EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnload"))
            )
            driver.execute_script("arguments[0].click();", load_button)
            print("✅ Clicked Load button")
            
            # Wait for loading to complete
            time.sleep(15)  # Give time for data to load
            driver.save_screenshot("after_load.png")
        except Exception as e:
            print(f"⚠️ Error clicking Load button: {str(e)}")
        
        # Click Export button
        try:
            export_button = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
                EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnexport"))
            )
            driver.execute_script("arguments[0].click();", export_button)
            print("✅ Clicked Export button")
            
            # Wait for download to complete
            time.sleep(20)
            driver.save_screenshot("after_export.png")
        except Exception as e:
            print(f"⚠️ Error clicking Export button: {str(e)}")
        
        print("✅ Date setting and download steps completed")
        return True
        
    except Exception as e:
        print(f"❌ Failed to set dates and download: {str(e)}")
        return False

def get_latest_file(folder_path, max_attempts=5, delay=2):
    """Get the most recently downloaded file from the specified folder"""
    for attempt in range(max_attempts):
        try:
            # Look for any Excel or CSV file that might be the report
            files = [
                os.path.join(folder_path, f) for f in os.listdir(folder_path)
                if (f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.csv'))
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
        'Airway Bill': [],
        'Create Date': [],
        'Reference 1': [],
        'Last Event': [],
        'Last Event Date': [],
        'Calling Status': [],
        'Cash/Cod Amt': []
    })

def process_data(file_path):
    """Process the downloaded PostaPlus report"""
    print(f"🔹 Processing file: {file_path}")
    
    try:
        if file_path is None:
            print("⚠️ No file to process, returning empty DataFrame")
            return create_empty_data()
            
        time.sleep(2)
        
        # Read the file
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
        
        # Print column names to help debug
        print(f"Columns in file: {df.columns.tolist()}")
        
        # Create a list of columns to keep
        columns_to_keep = ['Airway Bill', 'Create Date', 'Reference 1', 'Last Event', 
                          'Last Event Date', 'Calling Status', 'Cash/Cod Amt']
        
        # Create processed dataframe with only required columns
        processed_df = pd.DataFrame()
        
        for col in columns_to_keep:
            if col in df.columns:
                processed_df[col] = df[col]
            else:
                # Try to find similar column names (case-insensitive)
                found = False
                for df_col in df.columns:
                    if col.lower() == df_col.lower():
                        processed_df[col] = df[df_col]
                        found = True
                        break
                if not found:
                    print(f"⚠️ Column '{col}' not found, using empty values")
                    processed_df[col] = ''
        
        # Convert Airway Bill column to text type (string)
        if 'Airway Bill' in processed_df.columns:
            processed_df['Airway Bill'] = processed_df['Airway Bill'].astype(str)
        
        # Sort by Create Date (newest first) if it exists
        if 'Create Date' in processed_df.columns and processed_df['Create Date'].notna().any():
            try:
                # Try to convert to datetime for sorting
                processed_df['temp_date'] = pd.to_datetime(processed_df['Create Date'], errors='coerce')
                processed_df = processed_df.sort_values(by='temp_date', ascending=False)
                processed_df = processed_df.drop('temp_date', axis=1)
                print("✅ Sorted data by Create Date (newest first)")
            except Exception as sort_e:
                print(f"⚠️ Could not sort by Create Date: {str(sort_e)}")
        
        # Replace any NaN values with empty strings
        processed_df = processed_df.fillna('')
        
        print(f"✅ Data processing completed successfully - {len(processed_df)} rows")
        return processed_df
        
    except Exception as e:
        print(f"❌ Error processing data: {str(e)}")
        return create_empty_data()

def upload_to_google_sheets(df):
    """Upload processed data to Google Sheets"""
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
        print("🚀 Starting PostaPlus report automation process...")
        
        # Step 1: Setup and login
        driver = setup_chrome_driver()
        driver.implicitly_wait(IMPLICIT_WAIT)  # Add implicit wait
        
        if not login_to_postaplus(driver):
            print("⚠️ Login failed, but continuing with empty data...")
            upload_to_google_sheets(create_empty_data())
            print("🎉 Process completed with empty data")
            return
        
        # Step 2: Navigate to reports
        navigate_to_reports(driver)
        
        # Step 3: Set dates and download report
        set_dates_and_download(driver)
        
        # Step 4: Process the downloaded file
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
        
        # Step 5: Upload to Google Sheets
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
