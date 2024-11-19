from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time


excel_file_path = "C:\\Users\\Sourav Narvekar\\Desktop\\daksh\\Membership_data_updater\\Data1.xlsx"
data = pd.read_excel(excel_file_path)
# Initialize WebDriver
# driver = webdriver.Chrome(service="C:\\Users\\Sourav Narvekar\\Downloads\\chromedriver-win64\\chromedriver-win64")
driver=webdriver.Chrome()
driver.maximize_window()

# Step 1: Login to the website
try:
    # Open the login page
    driver.get("https://services.acm.org/public/chapters/loadmembers/add.cfm?CFID=37792209&CFTOKEN=170494a2b0ab5488-A4F053AB-D1FC-44BD-1BA18E27CA9D0B81&g=99809")

    # # Locate the input fields using their 'name' attributes
    # username_field = driver.find_element(By.NAME, "username")
    # password_field = driver.find_element(By.NAME, "password")

    # # Enter login credentials
    # username_field.send_keys("trdrjhufd")
    # password_field.send_keys("pccoe123")

    # # Submit the form (assuming hitting 'Enter' submits the login)
    # password_field.send_keys(Keys.RETURN)

    # Wait for the next page to load
    # WebDriverWait(driver, 20).until(
    #     EC.presence_of_element_located((By.LINK_TEXT, "Add Chapter Members"))
    # )
 # Step 2: Navigate to Add Chapter Members
    # driver.find_element(By.LINK_TEXT, "Add Chapter Members").click()

    # Step 3: Fill the form with data from Excel
    for index, row in data.iterrows():
        # Fill in each field
        # acm_number_field = driver.find_element(By.ID, "acm_membership_number")  # Update field IDs
        # prefix_field = driver.find_element(By.ID, "prefix")
        # first_name_field = driver.find_element(By.ID, "first_name")
        # middle_name_field = driver.find_element(By.ID, "middle_name")
        # last_name_field = driver.find_element(By.ID, "last_name")
        # suffix_field = driver.find_element(By.ID, "suffix")
        # email_field = driver.find_element(By.ID, "email")
        # affiliation_field = driver.find_element(By.ID, "affiliation")

        # # Enter values from the Excel file
        # acm_number_field.send_keys(row["ACM Membership Number"])
        # prefix_field.send_keys(row["Prefix"])
        # first_name_field.send_keys(row["First Name"])
        # middle_name_field.send_keys(row["Middle Name"])
        # last_name_field.send_keys(row["Last Name"])
        # suffix_field.send_keys(row["Suffix"])
        # email_field.send_keys(row["E-mail"])
        # affiliation_field.send_keys(row["Affiliation"])

        driver.find_element(By.NAME, "member_number_ret").send_keys(row["ACM Membership Number"])  # ACM Membership Number
# driver.find_element(By.NAME, "member_prefix").send_keys(row["Prefix"])          # Prefix
        driver.find_element(By.NAME, "member_first_name").send_keys(row["First Name"])     # First Name
        # driver.find_element(By.NAME, "member_middle_name").send_keys(row["Middle Name"])      # Middle Name
        driver.find_element(By.NAME, "member_last_name").send_keys(row["Last Name"])       # Last Name
# driver.find_element(By.NAME, "member_suffix").send_keys()          # Suffix
        driver.find_element(By.NAME, "member_email").send_keys(row["E-Mail"])  # Email
        driver.find_element(By.NAME, "member_affiliation").send_keys("PCCoE")  # Affiliation

# Click the "Retrieve Information" button (optional, if needed)
        retrieve_button = driver.find_element(By.ID, "bri")
        retrieve_button.click()

# Wait for any dynamic changes triggered by the "Retrieve Information" button
        time.sleep(2)

# Submit the form
        submit_button = driver.find_element(By.XPATH, "//input[@value='Add member']")
        submit_button.click()
        
        # Wait for a short while to avoid overwhelming the server
        time.sleep(2)

    print("Login successful!")

except Exception as e:
    print("An error occurred during login:", e)

finally:
    # Close the browser
    driver.quit()
