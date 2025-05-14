import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import shutil
# Set download directory (folder inside the automation folder)
download_folder = os.path.join(os.getcwd(), "downloads_uncomtrade_import1")
if not os.path.exists(download_folder):
    os.makedirs(download_folder)
 
# Setup WebDriver
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
 
# Set preferences for downloading files to the desired directory
prefs = {
    "download.default_directory": download_folder,  # Specify the download directory
    "download.prompt_for_download": False,  # Disable the prompt asking where to save the file
    "directory_upgrade": True  # Allow the directory to be upgraded if needed
}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=options)
prefs["safebrowsing.enabled"] = True
wait = WebDriverWait(driver, 30)
 
# Step 1: Open the website
driver.get("https://comtradeplus.un.org/")
 
# Step 2: Click on Login
login_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[text()='Login']")))
login_btn.click()
time.sleep(2)
 
# Step 3: Wait for login form & fill credentials
email_input = wait.until(EC.visibility_of_element_located((By.ID, "email")))
password_input = driver.find_element(By.ID, "password")
email_input.send_keys("siddharth.p@neptrontech.com")
password_input.send_keys("Admin@123456")
 
# Function to close modal if present
# def close_modal_if_present():
#     try:
#         time.sleep(1)
#         modal_close_button = driver.find_element(By.XPATH, "//div[contains(@class, 'modal')]//button[contains(@class, 'close')]")
#         if modal_close_button.is_displayed():
#             modal_close_button.click()
#             time.sleep(1)
#             print("Closed the modal.")
#     except Exception:
#         print("No modal to close or already handled.")
def close_modal_if_present():
    try:
        modals = driver.find_elements(By.XPATH, "//div[contains(@class, 'modal') and contains(@style, 'display: block')]")
        for modal in modals:
            close_button = modal.find_element(By.XPATH, ".//button[contains(@class, 'close')]")
            if close_button.is_displayed():
                driver.execute_script("arguments[0].click();", close_button)
                print("Closed a modal popup.")
                time.sleep(1)
    except Exception as e:
        print("No modal or error handling modal:", str(e))
 
# Function to select HS Code
def select_hscode_dropdown(value_to_select):
    try:
        close_modal_if_present()
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "modal")))
        time.sleep(2)
        close_modal_if_present()
        hscode_dropdown_input = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "react-dropdown-select-input")))
        driver.execute_script("arguments[0].scrollIntoView(true);", hscode_dropdown_input)
        time.sleep(1)
 
        try:
            remove_button = driver.find_element(By.XPATH, "//span[contains(@class, 'react-dropdown-select-option-remove')]")
            if remove_button.is_displayed():
                remove_button.click()
                time.sleep(1)
                print("Removed 'TOTAL' option.")
        except:
            pass
 
        hscode_dropdown_input = driver.find_element(By.CLASS_NAME, "react-dropdown-select-input.css-1ngods0-InputComponent")
        hscode_dropdown_input.click()
        time.sleep(1)
        hscode_dropdown_input.send_keys(value_to_select)
        time.sleep(2)
 
        wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "react-dropdown-select-content")))
        options = driver.find_elements(By.CLASS_NAME, "react-dropdown-select-option")
        for option in options:
            if value_to_select in option.text.strip():
                option.click()
                time.sleep(1)
                print(f"HS Code selected: {option.text.strip()}")
                return
 
        hscode_dropdown_input.send_keys(Keys.ARROW_DOWN)
        time.sleep(1)
        hscode_dropdown_input.send_keys(Keys.ENTER)
        time.sleep(1)
 
    except Exception as e:
        print("Error selecting HS Code:", e)
        driver.save_screenshot(f"hscode_error_{value_to_select}.png")
   
#import
def select_trade_flow_to_import():
    try:
        dropdown_wrapper = driver.find_element(By.XPATH, "//input[@name='TradeFlows']/parent::div[contains(@class,'react-dropdown-select')]")
        dropdown_input = dropdown_wrapper.find_element(By.CLASS_NAME, "react-dropdown-select-input")
        dropdown_input.click()
        time.sleep(1.5)
 
        selected_items = dropdown_wrapper.find_elements(By.XPATH, ".//span[@role='listitem']")
        for item in selected_items:
            remove_btn = item.find_element(By.CLASS_NAME, "react-dropdown-select-option-remove")
            driver.execute_script("arguments[0].click();", remove_btn)
            time.sleep(0.5)
 
        dropdown_input.send_keys("Import")
        time.sleep(1.5)
        dropdown_input.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.5)
        dropdown_input.send_keys(Keys.ENTER)
        print("Selected 'Import' successfully.")
        time.sleep(2)
 
    except Exception as e:
        print("Error selecting 'Import':", e)
 
#download
def wait_and_rename_download(hs_code, files_before):
    """ Wait for a new file to appear and rename it to HS Code """
    timeout = 90
    start_time = time.time()
 
    while time.time() - start_time < timeout:
        current_files = set(os.listdir(download_folder))
        new_files = current_files - files_before
 
        downloading = [f for f in current_files if f.endswith(".crdownload")]
        if downloading:
            time.sleep(2)
            continue
 
        new_xlsx = [f for f in new_files if f.endswith(".xlsx")]
        if new_xlsx:
            downloaded_file = os.path.join(download_folder, new_xlsx[0])
            new_name = os.path.join(download_folder, f"{hs_code}.xlsx")
 
            try:
                # Check if the file exists and is not in use before renaming
                if not os.path.exists(downloaded_file):
                    print(f"File {downloaded_file} does not exist. Retrying.")
                    time.sleep(2)
                    continue
 
                # Try renaming with a retry mechanism
                attempts = 3
                while attempts > 0:
                    try:
                        shutil.move(downloaded_file, new_name)
                        print(f"Renamed file to: {new_name}")
                        return
                    except PermissionError as e:
                        print(f"Permission denied while renaming. Retry in 2 seconds. Attempt {4 - attempts}")
                        time.sleep(2)
                        attempts -= 1
 
                print(f"Failed to rename file after multiple attempts: {new_name}")
                return
 
            except Exception as e:
                print(f"Rename failed: {e}")
                time.sleep(2)
        else:
            time.sleep(2)
 
    print(f"Timeout: File for HS Code {hs_code} not renamed.")
 
# Function to click the Download button and select format
value_to_download = "Excel"
def trigger_download(hs_code, value_to_download):
    try:
        files_before = set(os.listdir(download_folder))
        close_modal_if_present()
        download_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".ant-btn.css-usln0u")))
        download_button.click()
        print("Download button clicked.")
 
        options = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "ant-dropdown-menu-item")))
        for option in options:
            if value_to_download in option.text.strip():
                option.click()
                print(f"Option selected: {option.text.strip()}")
                break
 
        time.sleep(2)  # Give time to start downloading
        close_modal_if_present()
        wait_and_rename_download(hs_code, files_before)
 
        driver.refresh()
        time.sleep(2)
 
    except Exception as e:
        print(f"Error during download trigger: {e}")
 
# Read HS Codes from CSV file
hs_codes_df = pd.read_csv("1_to_999_commodities.csv", dtype={'HSCode': str})
hs_codes_list = hs_codes_df['HSCode'].tolist()
 
# Loop through each HS Code and perform actions
for hs_code in hs_codes_list:
    print(f"\nProcessing HS Code: {hs_code}")
    select_hscode_dropdown(hs_code)
    time.sleep(2)
    select_trade_flow_to_import()
    time.sleep(2)
    close_modal_if_present()
    trigger_download(hs_code, value_to_download)
    time.sleep(7)  # Adjust delay if needed
 
# Close browser
print("All downloads attempted. Closing browser.")
driver.quit()
 

 