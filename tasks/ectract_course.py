from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl
import os
import requests
import urllib3
import datetime
import shutil


# Disable SSL warnings for self-signed certificates
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

EMAIL = "sarita.ishwarkar@learningmate.com"
PASSWORD = "sarita@Ved8812"
CHROMEDRIVER_PATH = "C:\\Vicky\\dev\\python\\playground\\EXE-Files\\chromedriver-win64\\chromedriver.exe"

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

    # Write to Excel in the specified directory with date and time in filename
    download_dir = r"C:\Vicky\dev\python\playground\py-files\tasks\downloadables"
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
    now = datetime.datetime.now().strftime("%m-%d-%Y-%H-%M")
    excel_filename = f"course_links_{now}.xlsx"
    excel_path = os.path.join(download_dir, excel_filename)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Course Links"
    ws.append(["Title", "Link"])
    for row in data:
        ws.append(row)
    wb.save(excel_path)
    print(f"Excel file '{excel_filename}' has been saved in {download_dir}.")

finally:
    # After all previous actions, check all img tags on the current page
    imgs = driver.find_elements(By.TAG_NAME, "img")
    alt_not_found_dir = r"C:\Vicky\dev\python\playground\py-files\tasks\alt_not_found"
    if not os.path.exists(alt_not_found_dir):
        os.makedirs(alt_not_found_dir)

    # Path to your default downloads folder (adjust if needed)
    user_download_dir = os.path.join(os.path.expanduser("~"), "Downloads")

    for img in imgs:
        alt = img.get_attribute("alt")
        src = img.get_attribute("src")
        if alt:
            print(f"Image alt: {alt}")
        else:
            if src:
                try:
                    # Open image in a new tab
                    driver.execute_script(f"window.open('{src}', '_blank');")
                    print(f"Opened {src} in a new tab for download.")
                    time.sleep(5)  # Wait for the download to start

                    # Move the most recently downloaded image file
                    image_extensions = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp')
                    files = [os.path.join(user_download_dir, f) for f in os.listdir(user_download_dir)
                             if os.path.isfile(os.path.join(user_download_dir, f)) and f.lower().endswith(image_extensions)]
                    if files:
                        latest_file = max(files, key=os.path.getmtime)
                        dest_file = os.path.join(alt_not_found_dir, os.path.basename(latest_file))
                        shutil.move(latest_file, dest_file)
                        print(f"Moved downloaded image to {dest_file}")
                    else:
                        print("No image files found in Downloads folder.")

                    # Optionally, close the new tab and switch back
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])

                except Exception as e:
                    print(f"Failed to process image {src}: {e}")

    driver.quit()