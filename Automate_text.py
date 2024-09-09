from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import urllib.parse

chrome_options = Options()
chrome_options.binary_location = 'C:/Program Files/BraveSoftware/Brave-Browser/Application/brave.exe'
chrome_options.add_argument('--user-data-dir=C:/Users/Calvin\'s Acer/AppData/Local/BraveSoftware/Brave-Browser/User Data')
chrome_options.add_argument('--profile-directory=Default')

service = Service('C:\\Users\\Calvin\'s Acer\\Downloads\\chromedriver-win64\\chromedriver.exe')
driver = webdriver.Chrome(service=service, options=chrome_options)

driver.get('https://web.whatsapp.com')

time.sleep(10)

df = pd.read_excel('contacts.xlsx')

for index, row in df.iterrows():
    name = row['Name']
    number = row['Phone']
    
    message = f"Thank you {name}, for confirming your presence 2024! 🎉"
    
    try:
        encoded_message = urllib.parse.quote(message)
        chat_url = f"https://web.whatsapp.com/send?phone={number}&text={encoded_message}"
        driver.get(chat_url)
        time.sleep(2)
        send_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
        )
        
        send_button.click()
        
        print(f"Message sent to {name} ({number})")
        
    except Exception as e:
        print(f"Failed to send message to {name} ({number}): {e}")
    
    time.sleep(5)

driver.quit()