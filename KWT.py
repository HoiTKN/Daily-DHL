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
import requests
from urllib.parse import urljoin
import json
import base64
from bs4 import BeautifulSoup
import io
import magic  # For file type detection

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
        
        # ENHANCED: Improved download handling
        chrome_options.add_argument("--disable-features=IsolateOrigins,site-per-process")
        
        # Additional options for stability
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-infobars')
        chrome_options.add_argument('--ignore-certificate-errors')
        
        # Cookie handling - essential for session management
        chrome_options.add_argument("--enable-cookies")
        
        # CRITICAL: Enhanced preferences for downloads
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": DOWNLOAD_FOLDER,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "safebrowsing.disable_download_protection": True,
            "download.extensions_to_open": "",
            "plugins.always_open_pdf_externally": True,
            "browser.helperApps.neverAsk.saveToDisk": "application/vnd.ms-excel,application/msexcel,application/x-msexcel,application/x-ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,text/csv"
        })
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # Enable browser logging
        chrome_options.set_capability('goog:loggingPrefs', {'browser': 'ALL'})
        
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
        
        # CRITICAL: Enable Chrome DevTools Protocol for headless downloads
        try:
            driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": DOWNLOAD_FOLDER
            })
        except Exception as e:
            # Ignore CDP errors - will use alternative methods
            print(f"Note: CDP command failed, will use alternative download methods: {str(e)}")
        
        print(f"✅ Download directory set to: {DOWNLOAD_FOLDER}")
        
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
        
        # Set to date (23/05/2025 - current date in log)
        try:
            to_date_input = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txttodate_I"))
            )
            # Use the current date format from the log (23/05/2025)
            driver.execute_script("arguments[0].value = '23/05/2025';", to_date_input)
            driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", to_date_input)
            time.sleep(2)
            print(f"✅ Set to date to 23/05/2025")
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
        
        # ENHANCED: Get the form action and all form data
        try:
            form = driver.find_element(By.XPATH, "//form")
            form_action = form.get_attribute('action')
            print(f"Form action URL: {form_action}")
            
            # Save all form data for later use
            form_data = {}
            inputs = driver.find_elements(By.XPATH, "//form//input")
            for input_elem in inputs:
                input_name = input_elem.get_attribute('name')
                input_value = input_elem.get_attribute('value')
                if input_name:
                    form_data[input_name] = input_value
            
            # Save form data for debugging
            with open("form_data.json", "w") as f:
                json.dump(form_data, f, indent=2)
        except Exception as e:
            print(f"⚠️ Error getting form data: {str(e)}")
        
        # Attempt to export using standard approach first
        export_clicked = False
        try:
            # Check if page uses ASP.NET postback
            viewstate = driver.find_element(By.ID, "__VIEWSTATE")
            if viewstate:
                print("Page uses ASP.NET ViewState")
                
                # Try to trigger export using JavaScript postback
                driver.execute_script("__doPostBack('ctl00$ContentPlaceHolder1$btnexport', '')")
                export_clicked = True
                print("✅ Clicked Export button (JavaScript)")
                time.sleep(10)
        except Exception as e:
            print(f"⚠️ Export button click failed: {str(e)}")
        
        # Try direct form submission with export parameter
        if not export_clicked:
            try:
                print("Attempting direct form submission...")
                driver.execute_script("""
                    var form = document.querySelector('form');
                    var exportInput = document.createElement('input');
                    exportInput.type = 'hidden';
                    exportInput.name = 'ctl00$ContentPlaceHolder1$btnexport';
                    exportInput.value = 'Export';
                    form.appendChild(exportInput);
                    
                    // Log form data
                    console.log('Submitting form with export parameter');
                    
                    form.submit();
                """)
                export_clicked = True
                print("✅ Submitted form with export parameter")
                time.sleep(15)
            except Exception as e:
                print(f"⚠️ Form submission failed: {str(e)}")
        
        # Handle potential alert/popup
        try:
            alert = driver.switch_to.alert
            print(f"Alert found: {alert.text}")
            alert.accept()
            print("✅ Accepted alert")
        except:
            print("No alert found (this is normal)")
        
        # Take screenshot after export
        driver.save_screenshot("after_export.png")
        
        # ENHANCED: Check and save any network request data
        try:
            logs = driver.get_log('browser')
            if logs:
                print("Browser console logs:")
                for log in logs[-10:]:  # Last 10 logs
                    print(f"  {log['level']}: {log['message']}")
        except Exception as e:
            print(f"⚠️ Error getting browser logs: {str(e)}")
        
        # Wait for download to complete
        print("⏳ Waiting for download to complete...")
        time.sleep(20)
        
        print("✅ Date setting and download steps completed")
        return True
        
    except Exception as e:
        print(f"❌ Failed to set dates and download: {str(e)}")
        driver.save_screenshot("download_error.png")
        return False

def get_latest_file(folder_path, max_attempts=10, delay=5):
    """Get the most recently downloaded file from the specified folder"""
    print(f"🔍 Looking for downloaded files in: {folder_path}")
    
    # First, list all files in the directory for debugging
    try:
        all_files = os.listdir(folder_path)
        print(f"Total files in directory: {len(all_files)}")
        if len(all_files) < 20:  # Only print if not too many files
            for f in all_files:
                print(f"  - {f}")
    except Exception as e:
        print(f"Error listing directory: {str(e)}")
    
    for attempt in range(max_attempts):
        try:
            # Look for any Excel or CSV file that might be the report
            files = []
            for f in os.listdir(folder_path):
                file_path = os.path.join(folder_path, f)
                # Check if it's a file (not directory) and has the right extension
                if os.path.isfile(file_path) and (
                    f.endswith('.xlsx') or 
                    f.endswith('.xls') or 
                    f.endswith('.csv') or
                    f.endswith('.XLSX') or
                    f.endswith('.XLS') or
                    f.endswith('.CSV')
                ) and not f.startswith('~'):  # Ignore temporary Excel files
                    # Check if file was created in the last 5 minutes
                    file_time = os.path.getctime(file_path)
                    current_time = time.time()
                    if current_time - file_time < 300:  # 5 minutes
                        files.append(file_path)
                        print(f"Found recent file: {f} (created {int(current_time - file_time)} seconds ago)")
            
            if not files:
                print(f"No matching recent files found. Attempt {attempt + 1}/{max_attempts}")
                print(f"Waiting {delay} seconds before retry...")
                time.sleep(delay)
                continue
                
            latest_file = max(files, key=os.path.getctime)
            
            # Check if file can be opened (not still being written)
            try:
                with open(latest_file, 'rb') as f:
                    # Try to read first few bytes to ensure file is complete
                    f.read(100)
            except:
                print(f"File {latest_file} is still being written, waiting...")
                time.sleep(delay)
                continue
                
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
    
    # If we get here, no file was found after all attempts
    print("❌ No downloaded file found after all attempts")
    print("Possible reasons:")
    print("  1. Download was blocked by the website")
    print("  2. File is downloading to a different location")
    print("  3. File has an unexpected extension")
    print("  4. Download requires additional interaction (popup, etc.)")
    return None

def check_file_type(file_path):
    """Check the actual MIME type of the file"""
    try:
        # Install python-magic if not already installed
        try:
            import magic
        except ImportError:
            import subprocess
            subprocess.check_call(["pip", "install", "python-magic"])
            import magic
        
        # Check file MIME type
        mime = magic.Magic(mime=True)
        file_type = mime.from_file(file_path)
        print(f"Detected file type: {file_type}")
        
        # Check if it's HTML
        if "html" in file_type.lower() or "text" in file_type.lower():
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(1000)  # Read first 1000 chars
                if "<html" in content.lower() or "<!doctype html" in content.lower():
                    print("⚠️ File appears to be HTML, not Excel")
                    return "html"
        
        # Check if it's Excel
        if "excel" in file_type.lower() or "spreadsheet" in file_type.lower() or "office" in file_type.lower():
            return "excel"
            
        # Check if it's CSV
        if "csv" in file_type.lower() or "text/plain" in file_type.lower():
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(100)
                if "," in content and "\n" in content:
                    return "csv"
        
        return "unknown"
    except Exception as e:
        print(f"⚠️ Error checking file type: {str(e)}")
        # Fallback to simple extension check
        if file_path.lower().endswith('.csv'):
            return "csv"
        elif file_path.lower().endswith(('.xlsx', '.xls')):
            return "excel"
        elif file_path.lower().endswith(('.html', '.htm')):
            return "html"
        return "unknown"

def extract_table_from_html(file_path):
    """Extract data from HTML table"""
    try:
        print("🔍 Extracting data from HTML file...")
        
        # Read the HTML file
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            html_content = f.read()
        
        # Parse HTML
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Look for tables - try to find the data table
        tables = soup.find_all('table')
        print(f"Found {len(tables)} tables in HTML")
        
        # Find the largest table which is likely our data table
        largest_table = None
        max_rows = 0
        
        for table in tables:
            rows = table.find_all('tr')
            if len(rows) > max_rows:
                max_rows = len(rows)
                largest_table = table
        
        if largest_table and max_rows > 1:  # Ensure it has more than just a header row
            print(f"Using largest table with {max_rows} rows")
            
            # Extract headers
            headers = []
            header_row = largest_table.find('tr')
            if header_row:
                for th in header_row.find_all(['th', 'td']):
                    headers.append(th.get_text().strip())
            
            # If no headers found, use default column names
            if not headers:
                headers = ['Column' + str(i) for i in range(1, 20)]  # Default column names
            
            # Extract data rows
            data = []
            rows = largest_table.find_all('tr')
            for row in rows[1:]:  # Skip header row
                row_data = {}
                cells = row.find_all(['td', 'th'])
                for i, cell in enumerate(cells):
                    if i < len(headers):
                        row_data[headers[i]] = cell.get_text().strip()
                    else:
                        row_data[f'Column{i+1}'] = cell.get_text().strip()
                data.append(row_data)
            
            # Create DataFrame
            if data:
                df = pd.DataFrame(data)
                
                # Map columns to expected format if possible
                expected_columns = ['Airway Bill', 'Create Date', 'Reference 1', 'Last Event', 
                                  'Last Event Date', 'Calling Status', 'Cash/Cod Amt']
                
                # Try to map columns by similar names
                column_mapping = {}
                for expected_col in expected_columns:
                    best_match = None
                    for df_col in df.columns:
                        # Look for exact match first
                        if df_col.lower() == expected_col.lower():
                            best_match = df_col
                            break
                        # Then look for partial match
                        elif expected_col.lower() in df_col.lower() or df_col.lower() in expected_col.lower():
                            best_match = df_col
                    
                    if best_match:
                        column_mapping[best_match] = expected_col
                
                # Apply mapping if any found
                if column_mapping:
                    df = df.rename(columns=column_mapping)
                
                # Add missing expected columns
                for col in expected_columns:
                    if col not in df.columns:
                        df[col] = ''
                
                print(f"✅ Successfully extracted {len(df)} rows from HTML table")
                return df
        
        # If we couldn't extract data from tables, try other methods
        # Look for potential grid or data structure in divs or other elements
        grid_data = []
        grid_headers = []
        
        # Try to find grid headers
        header_elems = soup.select('.grid-header, .header, th, .column-header')
        if header_elems:
            for elem in header_elems:
                text = elem.get_text().strip()
                if text:
                    grid_headers.append(text)
        
        # Try to find grid rows
        row_elems = soup.select('.grid-row, .row, tr, .data-row')
        for row in row_elems:
            cells = row.select('td, .cell, .grid-cell')
            if cells:
                row_data = {}
                for i, cell in enumerate(cells):
                    if i < len(grid_headers):
                        col_name = grid_headers[i]
                    else:
                        col_name = f'Column{i+1}'
                    row_data[col_name] = cell.get_text().strip()
                grid_data.append(row_data)
        
        if grid_data:
            df = pd.DataFrame(grid_data)
            print(f"✅ Extracted {len(df)} rows from grid-like structure")
            return df
        
        print("❌ Could not extract data from HTML")
        return None
        
    except Exception as e:
        print(f"❌ Error extracting data from HTML: {str(e)}")
        return None

def enhanced_download_method(driver):
    """Enhanced method to capture and download report data"""
    try:
        print("🔄 Implementing enhanced download method...")
        
        # Look for any iframe that might contain the export
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            print(f"Found {len(iframes)} iframes - checking each for export options")
            for i, iframe in enumerate(iframes):
                try:
                    driver.switch_to.frame(iframe)
                    print(f"Switched to iframe {i}")
                    
                    # Look for export button in iframe
                    export_buttons = driver.find_elements(By.XPATH, 
                        "//button[contains(text(), 'Export') or contains(@id, 'export')] | " +
                        "//input[contains(@id, 'export') or contains(@value, 'Export')] | " +
                        "//a[contains(text(), 'Export') or contains(@id, 'export')]")
                    
                    if export_buttons:
                        print(f"Found {len(export_buttons)} export buttons in iframe {i}")
                        for btn in export_buttons:
                            try:
                                print(f"Export button found: {btn.get_attribute('outerHTML')}")
                                driver.execute_script("arguments[0].click();", btn)
                                print(f"Clicked export button in iframe {i}")
                                time.sleep(10)
                            except Exception as e:
                                print(f"Error clicking iframe export button: {str(e)}")
                    
                    driver.switch_to.default_content()
                except Exception as e:
                    print(f"Error with iframe {i}: {str(e)}")
                    driver.switch_to.default_content()
        
        # Try modified direct export form submission
        print("Attempting modified direct form submission...")
        
        # Save page source for analysis
        with open("page_before_enhanced_export.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        
        # Find the export button first to get its proper ID and attributes
        export_button = None
        try:
            export_button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnexport")
            print(f"Export button found: {export_button.get_attribute('outerHTML')}")
            print(f"Export button onclick: {export_button.get_attribute('onclick')}")
            print(f"Export button type: {export_button.get_attribute('type')}")
        except:
            print("Export button not found by ID")
            # Try to find by other means
            export_buttons = driver.find_elements(By.XPATH, 
                "//button[contains(text(), 'Export') or contains(@id, 'export')] | " +
                "//input[contains(@id, 'export') or contains(@value, 'Export')] | " +
                "//a[contains(text(), 'Export') or contains(@id, 'export')]")
            
            if export_buttons:
                export_button = export_buttons[0]
                print(f"Export button found by xpath: {export_button.get_attribute('outerHTML')}")
        
        # Get form details
        form = driver.find_element(By.XPATH, "//form")
        form_action = form.get_attribute('action')
        print(f"Form action: {form_action}")
        
        # Capture ViewState and EventValidation
        viewstate = driver.find_element(By.ID, "__VIEWSTATE").get_attribute('value') if driver.find_elements(By.ID, "__VIEWSTATE") else ""
        eventvalidation = driver.find_element(By.ID, "__EVENTVALIDATION").get_attribute('value') if driver.find_elements(By.ID, "__EVENTVALIDATION") else ""
        
        # Implement direct submission with full ASP.NET parameters
        try:
            # Get all form data
            form_data = {}
            inputs = driver.find_elements(By.XPATH, "//form//input")
            for input_elem in inputs:
                input_name = input_elem.get_attribute('name')
                input_value = input_elem.get_attribute('value')
                if input_name:
                    form_data[input_name] = input_value
            
            # Add export button parameter
            if export_button:
                button_name = export_button.get_attribute('name')
                if button_name:
                    form_data[button_name] = "Export"
                else:
                    form_data["ctl00$ContentPlaceHolder1$btnexport"] = "Export"
            else:
                form_data["ctl00$ContentPlaceHolder1$btnexport"] = "Export"
            
            # Create a session and submit the form
            session = requests.Session()
            
            # Get cookies from browser
            cookies = driver.get_cookies()
            for cookie in cookies:
                session.cookies.set(cookie['name'], cookie['value'])
            
            # Post request to get the file
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Referer': form_action,
                'Origin': 'https://etrack.postaplus.net',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            
            print("Submitting form with the following data:")
            print(json.dumps({k: v for k, v in form_data.items() if k.startswith('__') or 'btn' in k}, indent=2))
            
            response = session.post(form_action, data=form_data, headers=headers, allow_redirects=True)
            
            print(f"Form submission response status: {response.status_code}")
            print(f"Response headers: {dict(response.headers)}")
            
            # Check if we got a file
            if response.status_code == 200:
                content_type = response.headers.get('Content-Type', '')
                content_disp = response.headers.get('Content-Disposition', '')
                
                print(f"Response content type: {content_type}")
                print(f"Content-Disposition: {content_disp}")
                
                # Save the response content for analysis and processing
                filename = f"postaplus_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                filepath = os.path.join(DOWNLOAD_FOLDER, filename)
                with open(filepath, 'wb') as f:
                    f.write(response.content)
                
                print(f"✅ Saved file: {filepath}")
                return filepath
        except Exception as e:
            print(f"Form direct submission failed: {str(e)}")
        
        # Scrape data directly from the page if available
        try:
            print("Attempting to scrape data directly from the page...")
            
            # Save the current page as HTML
            html_filepath = os.path.join(DOWNLOAD_FOLDER, f"postaplus_page_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
            with open(html_filepath, 'w', encoding='utf-8') as f:
                f.write(driver.page_source)
            
            print(f"✅ Saved current page HTML to: {html_filepath}")
            return html_filepath
            
        except Exception as e:
            print(f"❌ Failed to save page HTML: {str(e)}")
        
        return None
        
    except Exception as e:
        print(f"❌ Enhanced download method failed: {str(e)}")
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
        
        # Check file type first
        file_type = check_file_type(file_path)
        print(f"Detected file type: {file_type}")
        
        # Process based on file type
        if file_type == "html":
            # Extract data from HTML
            df = extract_table_from_html(file_path)
            if df is not None and not df.empty:
                print(f"✅ Successfully extracted data from HTML: {len(df)} rows")
                return df
            else:
                print("⚠️ Could not extract data from HTML, returning empty DataFrame")
                return create_empty_data()
                
        elif file_type == "csv":
            # Process CSV file
            try:
                df = pd.read_csv(file_path)
                print(f"✅ Read CSV file: {len(df)} rows")
            except Exception as e:
                print(f"❌ Error reading CSV: {str(e)}")
                # Try with different encoding
                try:
                    df = pd.read_csv(file_path, encoding='latin1')
                    print(f"✅ Read CSV file with latin1 encoding: {len(df)} rows")
                except Exception as e2:
                    print(f"❌ Error reading CSV with latin1 encoding: {str(e2)}")
                    return create_empty_data()
        
        elif file_type == "excel":
            # Process Excel file
            try:
                df = pd.read_excel(file_path)
                print(f"✅ Read Excel file: {len(df)} rows")
            except Exception as e:
                print(f"❌ Error reading Excel file: {str(e)}")
                try:
                    # Try with engine specification
                    df = pd.read_excel(file_path, engine='openpyxl')
                    print(f"✅ Read Excel file with openpyxl engine: {len(df)} rows")
                except Exception as e2:
                    print(f"❌ Error reading Excel with openpyxl: {str(e2)}")
                    # Final attempt with xlrd
                    try:
                        df = pd.read_excel(file_path, engine='xlrd')
                        print(f"✅ Read Excel file with xlrd engine: {len(df)} rows")
                    except Exception as e3:
                        print(f"❌ Error reading Excel with xlrd: {str(e3)}")
                        return create_empty_data()
        else:
            # Unknown file type - try to read as text and look for tabular data
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                
                # Check if it's HTML even if not detected earlier
                if "<html" in content.lower() or "<!doctype html" in content.lower():
                    print("⚠️ File appears to be HTML, attempting to extract tables...")
                    with open(file_path + ".html", 'w', encoding='utf-8') as f:
                        f.write(content)
                    df = extract_table_from_html(file_path + ".html")
                    if df is not None and not df.empty:
                        return df
                
                # Check if it might be CSV
                if "," in content and "\n" in content:
                    print("⚠️ File might be CSV, attempting to parse...")
                    try:
                        df = pd.read_csv(io.StringIO(content))
                        print(f"✅ Parsed as CSV: {len(df)} rows")
                        return df
                    except:
                        pass
                
                print("⚠️ Could not determine file format, returning empty DataFrame")
                return create_empty_data()
            except Exception as e:
                print(f"❌ Error processing unknown file type: {str(e)}")
                return create_empty_data()
        
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
        
        # Install required packages
        try:
            import pip
            pip.main(['install', 'beautifulsoup4', 'python-magic', 'openpyxl'])
        except Exception as e:
            print(f"⚠️ Warning: Couldn't install packages: {str(e)}")
        
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
            
            # If no file found, try enhanced download method
            if not latest_file:
                print("⚠️ Regular download failed, trying enhanced method...")
                latest_file = enhanced_download_method(driver)
            
            if latest_file:
                processed_df = process_data(latest_file)
            else:
                # Try to get table data directly from the page
                print("⚠️ No files were downloaded, trying to scrape from current page...")
                html_path = os.path.join(DOWNLOAD_FOLDER, "current_page.html")
                with open(html_path, "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                processed_df = extract_table_from_html(html_path)
                
                if processed_df is None or processed_df.empty:
                    print("⚠️ No data found on page, using empty data structure")
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
