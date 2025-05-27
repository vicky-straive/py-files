import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def login_and_navigate(driver, email, password, course_url):
    driver.get("https://learningmate.instructure.com/login/canvas")
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "pseudonym_session_unique_id"))
    )
    email_field = driver.find_element(By.ID, "pseudonym_session_unique_id")
    password_field = driver.find_element(By.ID, "pseudonym_session_password")
    email_field.send_keys(email)
    password_field.send_keys(password)
    login_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="login_form"]/div[3]/div[2]/input'))
    )
    login_button.click()
    time.sleep(5)
    driver.get(course_url)
    time.sleep(5)


def extract_links(driver):
    course_home = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "custom-course-home"))
    )
    elements = course_home.find_elements(By.TAG_NAME, "a")
    data = []
    for elem in elements:
        title = elem.text.strip()
        url = elem.get_attribute('href')
        if title and url:
            data.append((title, url))
    return data


def download_missing_alt_images(driver, alt_not_found_dir):
    import os
    import time
    import shutil
    imgs = driver.find_elements(By.TAG_NAME, "img")
    user_download_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    if not os.path.exists(alt_not_found_dir):
        os.makedirs(alt_not_found_dir)
    for img in imgs:
        alt = img.get_attribute("alt")
        src = img.get_attribute("src")
        if alt:
            print(f"Image alt: {alt}")
        else:
            if src:
                try:
                    driver.execute_script(f"window.open('{src}', '_blank');")
                    print(f"Opened {src} in a new tab for download.")
                    time.sleep(5)
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
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                except Exception as e:
                    print(f"Failed to process image {src}: {e}")
