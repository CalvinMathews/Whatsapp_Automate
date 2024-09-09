from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import urllib.parse 
import os

chrome_options = Options()
chrome_options.binary_location = 'C:/Program Files/BraveSoftware/Brave-Browser/Application/brave.exe'
chrome_options.add_argument('--user-data-dir=C:/Users/Calvin\'s Acer/AppData/Local/BraveSoftware/Brave-Browser/User Data')
chrome_options.add_argument('--profile-directory=Default')

service = Service('C:\\Users\\Calvin\'s Acer\\Downloads\\chromedriver-win64\\chromedriver.exe')
driver = webdriver.Chrome(service=service, options=chrome_options)

driver.get('https://web.whatsapp.com')
time.sleep(10)

df = pd.read_excel('contacts-copy.xlsx')

for index, row in df.iterrows():
    name = row['Name']
    number = row['Phone']
    image_filename = row['Image Filename']
    image_path = os.path.abspath(os.path.join('images', image_filename))
    
    if os.path.exists(image_path):
        message = f"Thank you {name}, for confirming your presence at Converge 2024! 🎉"
        
        try:
            chat_url = f"https://web.whatsapp.com/send?phone={number}&text={urllib.parse.quote(message)}"
            driver.get(chat_url)
            time.sleep(5)
            
            # send_button = WebDriverWait(driver, 30).until(
            #     EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
            # )
            # send_button.click()
            # time.sleep(2)
            
            attachment_box = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//div[@title="Attach"]'))
            )
            attachment_box.click()
            time.sleep(2)
            
            image_input = driver.find_element(By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
            image_input.send_keys(image_path)
            time.sleep(2)
            
            send_image_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
            )
            send_image_button.click()
            
            print(f"Message and image sent to {name} ({number})")
            
        except Exception as e:
            print(f"Failed to send message to {name} ({number}): {e}")
        
        time.sleep(7)
    else:
        print(f"Image file for {name} not found. Skipping...")