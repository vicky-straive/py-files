from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import openpyxl
import os
import urllib3
import datetime
from course_utils import login_and_navigate, extract_links, download_missing_alt_images

# Disable SSL warnings for self-signed certificates
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

EMAIL = "sarita.ishwarkar@learningmate.com"
PASSWORD = "sarita@Ved8812"
CHROMEDRIVER_PATH = "C:\\Vicky\\dev\\python\\playground\\EXE-Files\\chromedriver-win64\\chromedriver.exe"
COURSE_URL = "https://learningmate.instructure.com/courses/10123"

service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service)

try:
    # 1. Login and navigate
    login_and_navigate(driver, EMAIL, PASSWORD, COURSE_URL)

    # 2. Extract links
    data = extract_links(driver)

    # 3. Write to Excel
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

    # 4. Download missing alt images
    alt_not_found_dir = r"C:\Vicky\dev\python\playground\py-files\tasks\alt_not_found"
    download_missing_alt_images(driver, alt_not_found_dir)

finally:
    driver.quit()