# In this file DATA is read from the Excel file and those DATA are passed(one by one) in the username and password field
# Validation is done by checking the dashboard URL with the Current URL and test result is updated in the XLS Sheet

from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.firefox import GeckoDriverManager
from locators import Locators
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from excel_function import Excel_Functions
from selenium.webdriver.common.action_chains import ActionChains
from data import Default_Data

url = Default_Data().url
excel_file = Default_Data.excel_file
sheet_number = Default_Data.sheet_number

s = Excel_Functions(Default_Data.excel_file, Default_Data.sheet_number)
driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))
driver.get(Default_Data.url)
driver.implicitly_wait(5)
wait = WebDriverWait(driver, 10)
rows = s.row_count()

# Read the Excel file
for row in range(2, rows+1):
    username_xls = s.read_data(row, column_number=6)
    password_xls = s.read_data(row, column_number=7)
    driver.implicitly_wait(10)

# Do the automation using Python Selenium
    username_we = wait.until(EC.presence_of_element_located((By.NAME, Locators.username)))
    username_we.send_keys(username_xls)

    password_we = wait.until(EC.presence_of_element_located((By.NAME, Locators.password)))
    password_we.send_keys(password_xls)

    submit_we = wait.until(EC.presence_of_element_located((By.XPATH, Locators.submit_button)))
    submit_we.click()

# write the Testcase Results into the Excel file

    if Default_Data().dashboard_url in driver.current_url:
        # If the login is successful with username and password
        s.write_data(row, column_number=8, data="TEST PASSED")
        action = ActionChains(driver)
        dropdown_we = wait.until(EC.presence_of_element_located((By.XPATH, Locators.dropdown_button)))
        dropdown_we.click()
        logout_we = wait.until(EC.presence_of_element_located((By.XPATH, Locators.logout_button)))
        action.move_to_element(logout_we).click().perform()

    elif Default_Data.url in driver.current_url:
        # If the login is un-successful with username and password
        s.write_data(row, column_number=8, data="TEST FAILED")
        driver.refresh()


# Close the DDTF automation testing
driver.quit()
