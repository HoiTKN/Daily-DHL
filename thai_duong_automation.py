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
SHEET_NAME = "Thai Duong"
SERVICE_ACCOUNT_FILE = 'service_account.json'
DOWNLOAD_FOLDER = os.getcwd()
DEFAULT_TIMEOUT = 30
PAGE_LOAD_TIMEOUT = 60

def setup_chrome_driver():
    """Setup Chrome driver with necessary options"""
    try:
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--window-size=1920,1080')
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
            'chromedriver',
            '/usr/local/bin/chromedriver',
            '/usr/bin/chromedriver',
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
        return driver
        
    except Exception as e:
        print(f"❌ Chrome driver setup failed: {str(e)}")
        raise

def login_to_thai_duong(driver):
    """Login to Thai Duong portal"""
    try:
        print("🔹 Accessing Thai Duong portal...")
        driver.get("https://tdffm.com/login")
        driver.delete_all_cookies()
        
        # Wait for and fill in username
        username_input = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
            EC.presence_of_element_located((By.NAME, "email"))
        )
        username_input.clear()
        username_input.send_keys("THA262")
        
        # Wait for and fill in password
        password_input = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
            EC.presence_of_element_located((By.NAME, "password"))
        )
        password_input.clear()
        password_input.send_keys("THA262")
        
        # Take screenshot before login
        driver.save_screenshot("thai_duong_before_login.png")
        
        # Find and click login button
        login_button = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']//span[text()='Đăng nhập']"))
        )
        driver.execute_script("arguments[0].click();", login_button)
        
        # Wait for login to complete
        time.sleep(10)
        
        # Take screenshot after login
        driver.save_screenshot("thai_duong_after_login.png")
        
        print(f"Current URL after login: {driver.current_url}")
        print("✅ Login steps completed!")
        return True
        
    except Exception as e:
        print(f"❌ Login failed: {str(e)}")
        driver.save_screenshot("thai_duong_login_failed.png")
        return False

def navigate_to_orders(driver):
    """Navigate to orders section"""
    try:
        print("🔹 Navigating to orders section...")
        time.sleep(5)
        
        # Click on "Đơn hàng" menu item
        orders_menu = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@class='ant-menu-title-content' and text()='Đơn hàng']"))
        )
        driver.execute_script("arguments[0].click();", orders_menu)
        
        time.sleep(5)
        driver.save_screenshot("thai_duong_orders_page.png")
        
        print("✅ Successfully navigated to orders section")
        return True
        
    except Exception as e:
        print(f"❌ Failed to navigate to orders: {str(e)}")
        driver.save_screenshot("thai_duong_orders_failed.png")
        return False

def export_orders(driver):
    """Export orders data"""
    try:
        print("🔹 Starting orders export...")
        
        # Click on "Xuất đơn hàng" button
        export_button = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='Xuất đơn hàng']"))
        )
        driver.execute_script("arguments[0].click();", export_button)
        
        time.sleep(3)
        driver.save_screenshot("thai_duong_export_popup.png")
        
        # Click "Xuất" button in the popup
        xuat_button = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='Xuất']"))
        )
        driver.execute_script("arguments[0].click();", xuat_button)
        
        # Wait for download to complete
        time.sleep(20)
        
        driver.save_screenshot("thai_duong_after_export.png")
        print("✅ Export completed")
        return True
        
    except Exception as e:
        print(f"❌ Failed to export orders: {str(e)}")
        driver.save_screenshot("thai_duong_export_failed.png")
        return False

def get_latest_file(folder_path, max_attempts=5, delay=2):
    """Get the most recently downloaded file from the specified folder"""
    for attempt in range(max_attempts):
        try:
            files = [
                os.path.join(folder_path, f) for f in os.listdir(folder_path)
                if (f.endswith('.xlsx') or f.endswith('.csv'))
                and not f.startswith('~$')
            ]
            
            if not files:
                print(f"No matching files found. Attempt {attempt + 1}/{max_attempts}")
                time.sleep(delay)
                continue
                
            latest_file = max(files, key=os.path.getctime)
            
            # Check if file can be opened
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
        'Ngày tạo đơn': [],
        'Tên khách hàng': [],
        'SĐT khách hàng': [],
        'Địa chỉ giao hàng': [],
        'Sản phẩm bán': [],
        'Số lượng': [],
        'Tiền COD': [],
        'Ngày gửi đơn': [],
        'Mã vận đơn': [],
        'Tình trạng giao hàng': [],
        'Phí vận chuyển': [],
        'Phí vùng xa': [],
        'Phí COD': [],
        'Phí tài khoản': [],
        'Phí xử lý hàng khó': [],
        'Phí dịch vụ FFM': []
    })

def process_data(file_path):
    """Process the downloaded Thai Duong report"""
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
        
        print(f"File shape: {df.shape}")
        print(f"Columns in file: {df.columns.tolist()}")
        
        # Extract specific columns by their positions (Excel column letters to numbers)
        # D=3, Q=16, R=17, S=18, T=19, AA=26, AB=27, AC=28, AI=34, AM=38, AN=39, AP=41, AR=43, AS=44, AT=45, BF=57
        
        column_mapping = {
            3: 'Ngày tạo đơn',      # Column D
            16: 'Ngày gửi đơn',     # Column Q  
            17: 'Tên khách hàng',   # Column R
            18: 'SĐT khách hàng',   # Column S
            19: 'Địa chỉ giao hàng', # Column T
            26: 'Sản phẩm bán',     # Column AA
            27: 'Số lượng',         # Column AB
            28: 'Tiền COD',         # Column AC
            34: 'Mã vận đơn',       # Column AI
            38: 'Phí vận chuyển',   # Column AM
            39: 'Phí vùng xa',      # Column AN
            41: 'Phí COD',          # Column AP
            43: 'Phí tài khoản',    # Column AR
            44: 'Phí xử lý hàng khó', # Column AS
            45: 'Phí dịch vụ FFM',  # Column AT
            57: 'Tình trạng giao hàng' # Column BF
        }
        
        # Create processed dataframe
        processed_df = pd.DataFrame()
        
        for col_index, col_name in column_mapping.items():
            if col_index < len(df.columns):
                processed_df[col_name] = df.iloc[:, col_index].fillna('')
            else:
                processed_df[col_name] = ''
                print(f"⚠️ Column {col_index} not found, using empty values for {col_name}")
        
        # Clean and format data
        processed_df = processed_df.replace({np.nan: '', None: ''})
        
        # Sort by creation date if available
        if 'Ngày tạo đơn' in processed_df.columns and not processed_df['Ngày tạo đơn'].empty:
            try:
                processed_df['temp_sort'] = pd.to_datetime(processed_df['Ngày tạo đơn'], errors='coerce')
                processed_df = processed_df.sort_values(by='temp_sort', ascending=False)
                processed_df = processed_df.drop('temp_sort', axis=1)
                print("✅ Successfully sorted by creation date")
            except Exception as sort_e:
                print(f"⚠️ Could not sort by date: {str(sort_e)}")
        
        print(f"✅ Data processing completed. Processed {len(processed_df)} rows")
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
        print("🚀 Starting Thai Duong automation process...")
        
        # Step 1: Setup and login
        driver = setup_chrome_driver()
        
        if not login_to_thai_duong(driver):
            print("⚠️ Login failed, uploading empty data...")
            upload_to_google_sheets(create_empty_data())
            print("🎉 Process completed with empty data")
            return
        
        # Step 2: Navigate to orders and export
        if not navigate_to_orders(driver):
            print("⚠️ Navigation failed, uploading empty data...")
            upload_to_google_sheets(create_empty_data())
            print("🎉 Process completed with empty data")
            return
            
        if not export_orders(driver):
            print("⚠️ Export failed, uploading empty data...")
            upload_to_google_sheets(create_empty_data())
            print("🎉 Process completed with empty data")
            return
        
        # Step 3: Process the downloaded file
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
        
        # Step 4: Upload to Google Sheets
        upload_to_google_sheets(processed_df)
        
        print("🎉 Thai Duong automation completed successfully!")
        
    except Exception as e:
        print(f"❌ Process failed: {str(e)}")
        try:
            upload_to_google_sheets(create_empty_data())
            print("⚠️ Uploaded empty data after process failure")
        except Exception as upload_e:
            print(f"❌ Failed to upload empty data: {str(upload_e)}")
    
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()
