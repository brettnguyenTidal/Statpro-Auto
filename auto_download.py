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

# Set up headless mode for GitHub Actions environment
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Initialize WebDriver
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)

try:
    # Open the website
    driver.get("https://revolution.statpro.com/")

    # Login process
    username_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "email")))
    password_field = driver.find_element(By.ID, "password")
    username_field.send_keys("pshah@tidalfg.com")
    password_field.send_keys("Tidal12!")
    login_button = driver.find_element(By.ID, "loginBtn")
    login_button.click()

    # Check if login was successful
    time.sleep(5)
    current_url = driver.current_url
    if "analytics/#Home/Dashboard" in current_url:
        print("Login successful!")
    else:
        test = driver.find_element(By.CSS_SELECTOR, "input[value='Stop the other session and login']")
        test.click()

    # Navigate and export files for each portfolio
    portfolios = [
        "Return Stacked Global Stocks & Bonds ETF",
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

        time.sleep(5)

        # Open the "Risk Management" dropdown
        risk_management_dropdown = driver.find_element(By.CSS_SELECTOR, "li[data-menu-service-id='MS_RiskManagement_Section']")
        risk_management_dropdown.click()

        # Click on "Absolute Risk"
        absolute_risk_link = driver.find_element(By.CSS_SELECTOR, "a[href='/analytics/Risk#risk/risk-dashboard']")
        absolute_risk_link.click()

        # Select portfolio
        portfolio_button = driver.find_element(By.CSS_SELECTOR, "a[class='pull-left stat-analysis-toolbar-setting stat-analysis-toolbar-btn-select-portfolio']")
        portfolio_button.click()

        time.sleep(5)
        search_bar = driver.find_element(By.ID, "s2id_autogen18")
        search_bar.send_keys(portfolio)
        search_bar.send_keys(Keys.ENTER)

        time.sleep(5)
        portfolio_option = driver.find_element(By.CSS_SELECTOR, f"td[data-title='{portfolio}']")
        portfolio_option.click()

        time.sleep(5)

        # Select categories and export
        category_button = driver.find_element(By.ID, "s2id_autogen73")
        category_button.click()
        risk_decomposition_option = driver.find_element(By.XPATH, "//div[contains(text(),'Risk Decomposition')]")
        risk_decomposition_option.click()

        category_button2 = driver.find_element(By.ID, "s2id_autogen71")
        category_button2.click()
        security_option = driver.find_element(By.XPATH, "//div[contains(text(),'Security Level')]")
        security_option.click()

        time.sleep(5)
        export_button = driver.find_element(By.CSS_SELECTOR, "button[class='btn btn-small export-button hvrbl ']")
        export_button.click()
        print(f"Export for {portfolio} completed successfully.")

        time.sleep(5)

    # Logout
    exit_btn = driver.find_element(By.CSS_SELECTOR, "a[class='btn btn-small dropdown-toggle']")
    exit_btn.click()
    logout = driver.find_element(By.XPATH, "//a[contains(text(),'Logout')]")
    logout.click()
    time.sleep(5)

    # Move files to the repository directory
    download_dir = "/home/runner/work/Statpro-Auto/Statpro-Auto/downloads"
    repo_dir = "/home/runner/work/Statpro-Auto/Statpro-Auto/repository_folder"

    # List all files in the download directory and move them to the repository folder
    files_to_move = [f for f in os.listdir(download_dir) if f.endswith('.csv')]

    for file in files_to_move:
        shutil.move(os.path.join(download_dir, file), os.path.join(repo_dir, file))

    print("Files moved to the repository folder.")

except Exception as e:
    print(f"An error occurred: {e}")
finally:
    driver.quit()
