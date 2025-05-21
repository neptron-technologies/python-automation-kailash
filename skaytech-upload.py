import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
 
# Step 1: Set up Chrome WebDriver
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 20)
#enter valid url below for upload path
driver.get("valid-url")
 
# Step 2: Log in and navigate to "Upload Excel"
login_sidebar = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "toggle-login-sidebar")))
login_sidebar.click()
 
wait.until(EC.visibility_of_element_located((By.ID, "username"))).send_keys("7208006566")
driver.find_element(By.ID, "password").send_keys("Admin@123")
driver.find_element(By.CLASS_NAME, "btn").click()
driver.maximize_window()
 
hs_code_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='HS Code Data Import']")))
hs_code_btn.click()
 
graph2_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[text()='Graph 2']")))
graph2_link.click()
 
time.sleep(3)  # Wait for the page to load
 
# Step 3: Define the folder and file list
downloads_folder = r"C:/Users/Admin/Desktop/Automation/downloads"
excel_files = [f for f in os.listdir(downloads_folder) if f.endswith((".xls", ".xlsx"))]
 
if not excel_files:
    print("No Excel files found in folder.")
    driver.quit()
    exit()
 
# Step 4: Loop through each file and upload it
for excel_file in excel_files:
    file_path = os.path.join(downloads_folder, excel_file)
    print(f"Uploading file: {file_path}")
 
    # Wait and click the "Upload Excel" button
    try:
        upload_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='button' and contains(., 'Upload Excel')]")))
        upload_button.click()  # Click the button once it's clickable
        print("Successfully clicked 'Upload Excel' button.")
    except Exception as e:
        print(f"Error clicking 'Upload Excel' button: {e}")
        driver.quit()
        exit()
 
    # Upload file
    time.sleep(4)  # Wait for input to load
    file_input = wait.until(EC.presence_of_element_located((By.ID, "file")))
    file_input.send_keys(file_path)
 
    # Submit file
    submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(@class, 'btn-primary')]")))
    submit_btn.click()
    print(f"File '{excel_file}' submitted successfully.")
 
    # Wait for the page to reload and go back to the "Upload Excel" page
    time.sleep(5)  # Increased timer to visually track progress
    print("Waiting for the page to reload...")
 
    # Optionally, wait for a specific element that confirms you're back on the upload page
    # For example, we can check if the "Upload Excel" button is clickable again:
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='button' and contains(., 'Upload Excel')]")))
        print(" Back on 'Upload Excel' page.")
    except Exception as e:
        print(f"Error confirming page reload: {e}")
        driver.quit()
        exit()
 
# Close the driver when done
driver.quit()
 
 
