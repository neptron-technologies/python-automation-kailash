
import time
import os
import pandas as pd
import shutil
import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# -------------------- CONFIGURATION --------------------
#enter valid path below for chromedriver.exe
CHROME_DRIVER_PATH = "D:/xxxxxxxx/Application-Dev/chromedriver.exe"
DOWNLOAD_DIR = os.path.abspath("downloads")
CSV_FILE = "1000_to_1999_commodities_pending.csv" 

# Create download directory if it doesn't exist
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# Chrome options for automatic download
options = Options()
prefs = {"download.default_directory": DOWNLOAD_DIR}
options.add_experimental_option("prefs", prefs)

# Read HS Codes from CSV
df = pd.read_csv(CSV_FILE, dtype={'HSCode': str})
hs_codes = df['HSCode'].tolist()

# Setup Selenium WebDriver
service = Service(CHROME_DRIVER_PATH)
driver = webdriver.Chrome(service=service, options=options)

# -------------------- RENAMING FUNCTION --------------------
def convert_xls_to_xlsx(file_path):
    """ Convert .xls file to .xlsx using Excel application """
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False 

        wb = excel.Workbooks.Open(file_path)
        new_file_path = file_path + "x"
        wb.SaveAs(new_file_path, FileFormat=51)  
        wb.Close()
        excel.Quit()

        os.remove(file_path)  
        print(f"Converted and removed: {file_path}")
        return new_file_path

    except Exception as e:
        print(f"Conversion failed for {file_path}: {e}")
        return None

def wait_and_rename_download(hs_code):
    """ Wait for the file to download and rename it with HS Code """
    before_download = set(os.listdir(DOWNLOAD_DIR))
    start_time = time.time()
    downloaded_file = None

    while time.time() - start_time < 60: 
        after_download = set(os.listdir(DOWNLOAD_DIR))
        new_files = after_download - before_download

        for file_name in new_files:
            if file_name.endswith(".xls") and not file_name.startswith(hs_code):
                downloaded_file = os.path.join(DOWNLOAD_DIR, file_name)
                break

        if downloaded_file and os.path.exists(downloaded_file):
            new_name = os.path.join(DOWNLOAD_DIR, f"{hs_code}.xls")
            try:
                shutil.move(downloaded_file, new_name)
                print(f"Renamed file to: {new_name}")
                convert_xls_to_xlsx(new_name) 
                return
            except Exception as e:
                print(f" Rename failed for HS Code {hs_code}: {e}")
                return
        time.sleep(2)

    print(f"File for HS Code {hs_code} not renamed within timeout.")

# -------------------- MAIN DOWNLOAD FUNCTION --------------------
def download_excel_for_hs_code(hs_code):
    """ Download the Excel file for the given HS Code """
    #enter valid url below - https://tradestat.commerce.gov.in/eidb/ecomcntq.asp
    driver.get("valid-url")

    try:
        # Enter HS Code
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "hscode")))
        hs_code_field = driver.find_element(By.NAME, "hscode")
        hs_code_field.clear()
        ActionChains(driver).move_to_element(hs_code_field).click().send_keys(hs_code).perform()

        # Click Submit
        submit_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.NAME, "button1"))
        )
        try:
            submit_button.click()
        except:
            driver.execute_script("arguments[0].click();", submit_button)
        print(f" Submitted HS Code: {hs_code}")
        time.sleep(5)
    except Exception as e:
        print(f" Error submitting HS Code {hs_code}: {e}")
        return

    try:
        # Click Export to Excel
        download_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "button1"))
        )
        download_button.click()
        print(f" Download started for HS Code: {hs_code}")
        wait_and_rename_download(hs_code)

    except Exception as e:
        print(f" Failed to download for HS Code: {hs_code} due to {e}")

# -------------------- LOOP THROUGH ALL CODES --------------------
for code in hs_codes:
    download_excel_for_hs_code(code)

# -------------------- DONE --------------------
driver.quit()
print("All downloads complete.")
