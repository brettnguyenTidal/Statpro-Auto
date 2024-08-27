from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import time
import os
import shutil
import openpyxl
import csv
from datetime import datetime

# Function to convert text to number
def convert_to_number(value):
    try:
        # Try to convert to integer
        return int(value)
    except ValueError:
        try:
            # convert to float
            return float(value)
        except ValueError:
            # Return the original value if conversion fails
            return value

# Set up headless mode for GitHub Actions environment
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Set download path within the repository
download_path = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(download_path):
    os.makedirs(download_path)

prefs = {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)

# WebDriver
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)

username = os.getenv("STATPRO_USERNAME")
password = os.getenv("STATPRO_PASSWORD")

# Define the naming dictionary for the portfolios
portfolio_naming = {
    "Return Stacked Global Stocks & Bonds ETF": "RSSB",
    "Return Stacked Bonds & Futures Yield ETF" : "RSBY",
    "Return Stacked U.S. Stocks & Futures Yield ETF": "RSSY",
    "Return Stacked U.S. Stocks & Managed Futures ETF": "RSST",
    "Return StackedTM Bonds & Managed Futures ETF": "RSBT"
}

# Get the current date in the desired format
current_date = datetime.now().strftime("%m.%d.%y")

try:
    # Open the website
    driver.get("https://revolution.statpro.com/")

    # Login process
    username_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "email")))
    password_field = driver.find_element(By.ID, "password")
    username_field.send_keys(username)
    password_field.send_keys(password)
    login_button = driver.find_element(By.ID, "loginBtn")
    login_button.click()

    # Check if login was successful
    time.sleep(20)
    current_url = driver.current_url
    if "analytics/#Home/Dashboard" in current_url:
        print("Login successful!")
    else:
        test = driver.find_element(By.CSS_SELECTOR, "input[value='Stop the other session and login']")
        test.click()

    # Navigate and export files for each portfolio
    portfolios = [
        "Return Stacked Global Stocks & Bonds ETF",
        "Return Stacked Bonds & Futures Yield ETF",
        "Return Stacked U.S. Stocks & Futures Yield ETF",
        "Return Stacked U.S. Stocks & Managed Futures ETF",
        "Return StackedTM Bonds & Managed Futures ETF"    
    ]

    for portfolio in portfolios:
        # Navigate to the risk section
        risk_tab = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "sitePageLink_Compliance_Index"))
        )
        risk_tab.click()

        time.sleep(10)

        # Open the "Risk Management" dropdown
        risk_management_dropdown = driver.find_element(By.CSS_SELECTOR, "li[data-menu-service-id='MS_RiskManagement_Section']")
        risk_management_dropdown.click()

        # Click on "Absolute Risk"
        absolute_risk_link = driver.find_element(By.CSS_SELECTOR, "a[href='/analytics/Risk#risk/risk-dashboard']")
        absolute_risk_link.click()

        # Select portfolio
        portfolio_button = driver.find_element(By.CSS_SELECTOR, "a[class='pull-left stat-analysis-toolbar-setting stat-analysis-toolbar-btn-select-portfolio']")
        portfolio_button.click()

        time.sleep(10)
        search_bar = driver.find_element(By.ID, "s2id_autogen18")
        search_bar_value_element = driver.find_element(By.CLASS_NAME, "select2-search-choice")
        search_bar_value = search_bar_value_element.text
        if search_bar_value != "return":
            search_bar.send_keys("return")
            search_bar.send_keys(Keys.ENTER)
            time.sleep(10)  # Wait for search results to load
            
        time.sleep(10)
        portfolio_option = driver.find_element(By.CSS_SELECTOR, f"td[data-title='{portfolio}']")
        portfolio_option.click()

        time.sleep(10)

        # Select categories and export
        category_button = driver.find_element(By.ID, "s2id_autogen73")
        category_button.click()
        risk_decomposition_option = driver.find_element(By.XPATH, "//div[contains(text(),'Risk Decomposition')]")
        risk_decomposition_option.click()

        category_button2 = driver.find_element(By.ID, "s2id_autogen71")
        category_button2.click()
        security_option = driver.find_element(By.XPATH, "//div[contains(text(),'Security Level')]")
        security_option.click()

        time.sleep(10)
        
        # Capture existing files before export
        files_before = set(os.listdir(download_path))
        
        export_button = driver.find_element(By.CSS_SELECTOR, "button[class='btn btn-small export-button hvrbl ']")
        export_button.click()
        print(f"Export for {portfolio} started successfully.")
        
        time.sleep(20)  # Wait for the download to complete
        
        # Capture new files after export
        files_after = set(os.listdir(download_path))
        new_files = files_after - files_before
        if len(new_files) == 0:
            print(f"No new file downloaded for {portfolio}")
        else:
            for new_file in new_files:
                if new_file == "Measures.csv":
                    # Rename the file to include the portfolio name
                    portfolio_code = portfolio_naming[portfolio]
                    new_name_csv = f"{portfolio_code} Risk Decomp {current_date}.csv"
                    os.rename(os.path.join(download_path, new_file), os.path.join(download_path, new_name_csv))
                    print(f"CSV file renamed to {new_name_csv} for {portfolio}")
                    
                    # Convert the CSV to Excel
                    csv_file_path = os.path.join(download_path, new_name_csv)
                    excel_file_path = os.path.join(download_path, f"{portfolio_code} Risk Decomp {current_date}.xlsx")
                    
                    # Reading CSV and writing to Excel with number conversion
                    workbook = openpyxl.Workbook()
                    sheet = workbook.active
                    with open(csv_file_path, 'r') as f:
                        reader = csv.reader(f)
                        for row in reader:
                            # Apply number conversion to each cell in the row
                            converted_row = [convert_to_number(cell) for cell in row]
                            sheet.append(converted_row)
                    workbook.save(excel_file_path)
                    print(f"Converted {new_name_csv} to Excel file {excel_file_path}")

                    # Delete the original CSV file
                    os.remove(csv_file_path)
                    print(f"Deleted the original CSV file {csv_file_path}")

    # Logout
    exit_btn = driver.find_element(By.CSS_SELECTOR, "a[class='btn btn-small dropdown-toggle']")
    exit_btn.click()
    logout = driver.find_element(By.XPATH, "//a[contains(text(),'Logout')]")
    logout.click()
    time.sleep(5)

except Exception as e:
    print(f"An error occurred: {e}")
finally:
    driver.quit()

