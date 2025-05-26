from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Replace these with your actual credentials
EMAIL = "sarita.ishwarkar@learningmate.com"
PASSWORD = "sarita@Ved8812"

# Path to your chromedriver.exe
CHROMEDRIVER_PATH = "C:\\Vicky\\dev\\python\\EXE-Files\\chromedriver-win64\\chromedriver.exe"

# Start Chrome browser using Service
service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service)

try:
    # Open the login page
    driver.get("https://learningmate.instructure.com/login/canvas")

    # Wait for the email field to be present
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "pseudonym_session_unique_id"))
    )

    # Find the email and password fields and fill them
    email_field = driver.find_element(By.ID, "pseudonym_session_unique_id")
    password_field = driver.find_element(By.ID, "pseudonym_session_password")

    email_field.send_keys(EMAIL)
    password_field.send_keys(PASSWORD)

   # Wait for the login button to be clickable, then click
    login_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="login_form"]/div[3]/div[2]/input'))
    )
    login_button.click()

    # Wait to see the result
    time.sleep(5)

finally:
    driver.quit()