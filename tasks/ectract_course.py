from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl

EMAIL = "sarita.ishwarkar@learningmate.com"
PASSWORD = "sarita@Ved8812"
CHROMEDRIVER_PATH = "C:\Vicky\dev\python\playground\EXE-Files\chromedriver-win64\chromedriver.exe"

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

    # Wait for login to complete
    time.sleep(5)

    # Redirect to the course page
    driver.get("https://learningmate.instructure.com/courses/10123")
    time.sleep(5)  # Wait for the page to load
    
    # Wait for the custom-course-home element to be present
    course_home = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "custom-course-home"))
    )

    # Find all <a> tags inside the custom-course-home element
    elements = course_home.find_elements(By.TAG_NAME, "a")


   # Prepare data for Excel
    data = []
    for elem in elements:
        title = elem.text.strip()
        url = elem.get_attribute('href')
        if title and url:
            data.append((title, url))

    # Write to Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Course Links"
    ws.append(["Title", "Link"])
    for row in data:
        ws.append(row)
    wb.save("course_links.xlsx")
    print("Excel file 'course_links.xlsx' has been saved in the current directory.")

finally:
    driver.quit()