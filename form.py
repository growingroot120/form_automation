# import pandas as pd
# from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
# from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from webdriver_manager.chrome import ChromeDriverManager
# import time

# # Load the Excel file
# df = pd.read_excel('form.xls')

# # Set Chrome options
# chrome_options = Options()

# # Initialize the WebDriver using webdriver-manager
# driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# for index, row in df.iterrows():
#     # Get the URL and other form data
#     url = row['url']
#     schedule = row['schedule']
#     saluation = row['saluation']
#     firstname = row['firstname']
#     lastname = row['lastname']
#     email = row['email']
#     birthdate = row['birthdate']

#     # Open a new tab
#     driver.execute_script("window.open('');")
#     # Switch to the new tab
#     driver.switch_to.window(driver.window_handles[-1])
#     # Move to that URL
#     driver.get(url)

#     # Wait for the form to load (adjust the waiting conditions as needed)
#     try:
#         WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.CLASS_NAME, 'push_popup'))  # Waiting for an element with class 'push_popup'
#         )
#     except:
#         print(f"Form did not load in time for {url}")
#         continue

#     # Click the button with id 'maximizly-push-denied'
#     try:
#         deny_button = WebDriverWait(driver, 10).until(
#             EC.element_to_be_clickable((By.ID, 'maximizly-push-denied'))
#         )
#         deny_button.click()
#     except:
#         print(f"Could not click the deny button for {url}")
#         continue

#     # # Wait for 100 seconds
#     # time.sleep(10)

#     # # Fill the form with updated field names
#     # saluation_input = driver.find_element(By.CLASS_NAME, 'vs__search')
#     # saluation_input.send_keys(saluation)
#     # saluation_input.send_keys('\n')  # Press Enter after input 
#     # driver.find_element(By.ID, 'reservation_firstname').send_keys(firstname)
#     # driver.find_element(By.ID, 'reservation_lastname').send_keys(lastname)
#     # driver.find_element(By.ID, 'reservation_email').send_keys(email)
#     # driver.find_element(By.ID, 'reservation_dob').send_keys(birthdate)

#     # Wait for 100 seconds
#     time.sleep(10)
#     # Find the button with class 'ml-auto mt-2 btn-primary' and click it
#     try:
#         submit_button = driver.find_element(By.CLASS_NAME, 'ml-auto')
#         print("Submit button found.")
#         submit_button.click()
#     except:
#         print(f"Could not click the submit button for {url}")
#         continue

#     # Wait for some time after submission before moving to the next form (adjust as needed)
#     time.sleep(20)

# # Close the driver
# driver.quit()


import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
from datetime import datetime, timedelta

# Load the Excel file
df = pd.read_excel('form.xls')

# Set Chrome options
chrome_options = Options()

# Initialize the WebDriver using webdriver-manager
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

def wait_until_schedule(schedule_time_str):
    schedule_time = datetime.strptime(schedule_time_str, '%H:%M:%S').time()
    pre_schedule_time = (datetime.combine(datetime.today(), schedule_time) - timedelta(seconds=50)).time()
    
    while True:
        now = datetime.now()
        if now.time() >= pre_schedule_time and now.weekday() in [4, 5, 6]:  # Friday = 4, Saturday = 5, Sunday = 6
            return True
        time.sleep(1)

def wait_until_exact_schedule(schedule_time_str):
    schedule_time = datetime.strptime(schedule_time_str, '%H:%M:%S').time()
    
    while True:
        now = datetime.now()
        if now.time() >= schedule_time and now.weekday() in [4, 5, 6]:  # Friday = 4, Saturday = 5, Sunday = 6
            return True
        time.sleep(1)

while True:
    for index, row in df.iterrows():
        # Get the URL and other form data
        url = row['url']
        schedule = row['schedule']  # Assuming schedule is in 'HH:MM:SS' format
        saluation = row['saluation']
        firstname = row['firstname']
        lastname = row['lastname']
        email = row['email']
        birthdate = row['birthdate']

        # Print the read values
        print(f"Row {index} values: URL={url}, Schedule={schedule}, Saluation={saluation}, Firstname={firstname}, Lastname={lastname}, Email={email}, Birthdate={birthdate}")

        # Wait until 50 seconds before the scheduled time on Friday, Saturday, or Sunday
        # if wait_until_schedule(schedule):
        if True:
            # Open the browser and navigate to the URL
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
            driver.get(url)

            # Wait for the form to load (adjust the waiting conditions as needed)
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, 'push_popup'))  # Waiting for an element with class 'push_popup'
                )
            except:
                print(f"Form did not load in time for {url}")
                driver.quit()
                continue

            # Click the button with id 'maximizly-push-denied'
            try:
                deny_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, 'maximizly-push-denied'))
                )
                deny_button.click()
            except:
                print(f"Could not click the deny button for {url}")
                driver.quit()
                continue

            # Fill the form with updated field names
            try:
                saluation_input = driver.find_element(By.CLASS_NAME, 'vs__search')
                saluation_input.send_keys(saluation)
                saluation_input.send_keys('\n')  # Press Enter after input

                driver.find_element(By.ID, 'reservation_firstname').send_keys(firstname)
                driver.find_element(By.ID, 'reservation_lastname').send_keys(lastname)
                driver.find_element(By.ID, 'reservation_email').send_keys(email)
                driver.find_element(By.ID, 'reservation_dob').send_keys(birthdate)
            except:
                print(f"Could not fill the form for {url}")
                driver.quit()
                continue

            # Wait until the exact scheduled time
            # wait_until_exact_schedule(schedule)
            time.sleep(100)
            # Find the button with class 'ml-auto' and click it at the scheduled time
            try:
                submit_button = driver.find_element(By.CLASS_NAME, 'ml-auto')
                print("Submit button found.")
                # submit_button.click()
                print(f"Submit button clicked at scheduled time: {datetime.now()}.")
            except:
                print(f"Could not click the submit button for {url}")

            # Close the browser
            driver.quit()
