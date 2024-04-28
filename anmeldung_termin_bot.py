# Import the packages
import os
import re
import time

import pandas as pd
import yagmail
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

pd.set_option("display.max_rows", 100)

# Set the Chrome options
chrome_options = Options()
chrome_options.add_argument("start-maximized") # Required for a maximized Viewport
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation', 'disable-popup-blocking']) # Disable pop-ups to speed up browsing
chrome_options.add_experimental_option("detach", True) # Keeps the Chrome window open after all the Selenium commands/operations are performed 
chrome_options.add_experimental_option('prefs', {'intl.accept_languages': 'en,en_US'}) # Operate Chrome using English as the main language
chrome_options.add_argument('--blink-settings=imagesEnabled=false') # Disable images
chrome_options.add_argument('--disable-extensions') # Disable extensions
chrome_options.add_argument('--no-sandbox') # Disables the sandbox for all process types that are normally sandboxed. Meant to be used as a browser-level switch for testing purposes only
chrome_options.add_argument('--disable-gpu') # An additional Selenium setting for headless to work properly, although for newer Selenium versions, it's not needed anymore
chrome_options.add_argument("enable-features=NetworkServiceInProcess") # Combats the renderer timeout problem
chrome_options.add_argument("disable-features=NetworkService") # Combats the renderer timeout problem
chrome_options.add_experimental_option('extensionLoadTimeout', 45000) # Fixes the problem of renderer timeout for a slow PC
chrome_options.add_argument("--window-size=1920x1080") # Set the Chrome window size to 1920 x 1080

# Inputs
STD_WAITING_TIME = 10
TARGET_URL = "https://service.berlin.de/dienstleistung/120686/"
IS_BOOK_APPT = True
FULL_NAME = "Omar ElMaria"
EMAIL = "omarmoataz6@gmail.com"

# Instantiate the yagmail object
yag = yagmail.SMTP("omarmoataz6@gmail.com", oauth2_file=os.path.expanduser("~")+"/email_authentication.json", smtp_ssl=False)

def main():
    # Instantiate the web driver
    driver = webdriver.Chrome(options=chrome_options)

    # Navigate to the target website
    driver.get(TARGET_URL)

    # Maximizing the Chrome window
    driver.maximize_window()

    # Wait until the desired checkbox is clickable
    WebDriverWait(driver, STD_WAITING_TIME).until(EC.element_to_be_clickable((By.XPATH, "//input[@id='checkbox_overall']")))

    # Scroll down to the checkbox webelement
    driver.execute_script("arguments[0].scrollIntoView();", driver.find_element(By.XPATH, "//h2[text()='Hinweise zur Zuständigkeit']"))

    # Click the checkbox
    driver.find_element(By.XPATH, "//input[@id='checkbox_overall']").click()

    # Wait until the "An diesem Standort einen Termin buchen" button is clickable
    WebDriverWait(driver, STD_WAITING_TIME).until(EC.element_to_be_clickable((By.XPATH, "//button[@id='appointment_submit']"))).click()

    # Wait until this element is visible --> //tr/td
    try:
        WebDriverWait(driver, STD_WAITING_TIME).until(EC.visibility_of_element_located((By.XPATH, "//tr/td")))
        
        # If this element is visible, then the calendar view will show up. Now, try to extract the Termine
        try:
            results = driver.find_elements(By.XPATH, "//tr/td/a")
            termine = []
            for res in results:
                termine.append(res.get_attribute("aria-label"))
            
            # Change termine to a DataFrame
            df_termine = pd.DataFrame(termine, columns=["termine"])

            # Remove None rows
            df_termine = df_termine.dropna()

            # Extract the date from the termine string
            df_termine["termine_cleaned"] = df_termine.apply(
                lambda x: re.findall(pattern=r".*(?=\s-)", string=x["termine"])[0] if " - " in x["termine"] else None, axis=1
            )

            # Change the date string to a datetime object. The format of the string is dd.mm.yyyy
            df_termine["dates"] = pd.to_datetime(df_termine["termine_cleaned"], format="%d.%m.%Y")

            # Extract the month from the dates
            df_termine["month"] = df_termine.apply(lambda x: x["dates"].month, axis=1)

            # Check if any of the dates fall in April or May and if the column termine contains this substring "An diesem Tag einen Termin buchen"
            df_termine["flag"] = df_termine.apply(
                lambda x: True if (x["month"] == 4 or x["month"] == 5) and "An diesem Tag einen Termin buchen" in x["termine"] else False, axis=1
            )

            # Check if any of the rows satisfy the above condition
            flag = df_termine["flag"].values.any()

            # If any of the rows satisfy the above condition, then send an email
            # Otherwise, print a message saying that appointments are available but not in April or May
            if flag:
                # Extract the minimum date with an appointment from the DataFrame
                min_date = str(df_termine[df_termine["flag"]]["dates"].min().date())
                max_date = str(df_termine[df_termine["flag"]]["dates"].max().date())

                # Express min_date in the format dd.mm.yyyy
                min_date_reformatted = ".".join(min_date.split("-")[::-1])

                # Send an email
                contents = [
                    f"""This is an automated notification to inform you that Anmeldung appointments in April or May have been found.
                    The earliest date is {min_date} and the latest date is {max_date}. Go book --> {TARGET_URL}"""
                ]
                subject = f"Anmeldung Appointments in April or May have been found. Earliest date is {min_date}"

                # Attempt to book the appointment
                if IS_BOOK_APPT:
                    # Click on the earliest date
                    driver.find_element(By.XPATH, f"//tr/td/a[contains(@aria-label, '{min_date_reformatted}')]").click()
                    
                    # Wait until the available times appear and click on the earliest available time
                    WebDriverWait(driver, STD_WAITING_TIME).until(EC.element_to_be_clickable((By.XPATH, "//tr/td[@class='frei']/a"))).click()
                    
                    # Wait until the input fields appear and input your personal details
                    WebDriverWait(driver, STD_WAITING_TIME).until(EC.presence_of_element_located((By.XPATH, "//input[@id='familyName']")))
                    driver.find_element(By.XPATH, "//input[@id='familyName']").send_keys(FULL_NAME)
                    time.sleep(0.5)
                    driver.find_element(By.XPATH, "//input[@id='email']").send_keys(EMAIL)
                    time.sleep(0.5)
                    driver.find_element(By.XPATH, "//input[@id='emailequality']").send_keys(EMAIL)
                    time.sleep(0.5)
                    
                    # Click on "AGB gelesen"
                    driver.find_element(By.XPATH, "//input[@id='agbgelesen']").click()
                    time.sleep(0.5)

                    # Wait until "Termin eintragen" is clickable and click on it
                    WebDriverWait(driver, STD_WAITING_TIME).until(EC.element_to_be_clickable((By.XPATH, "//button[@id='register_submit']"))).click()
            else:
                contents = ["Appointments are available but not in April or May"]
                subject = "Appointments are available but not in April or May"
        except NoSuchElementException:
            contents = ["Calendar view exists but no appointments available"]
            subject = "Calendar view exists but no appointments available"
    except TimeoutException:
        contents = ["Calendar view does not exist. No Anmeldung appointments available"]
        subject = "Calendar view does not exist. No Anmeldung appointments available"

    # Send the E-mail
    yag.send(["omarmoataz6@gmail.com"], subject, contents)

    # Quit the driver
    driver.quit()

if __name__ == "__main__":
    while True:
        try:
            main()
        except Exception as e:
            yag.send(["omarmoataz6@gmail.com"], "An Error Occurred While Looking for Anmeldung Appointments", [f"Error: {str(e)}"])
        
        time.sleep(90)