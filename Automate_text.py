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

df = pd.read_excel('contacts2.xlsx')

for index, row in df.iterrows():
    # name = row['Name']
    number = row['Phone']
    
    message =f"""*Plan your itinerary* \n\n*Location: Impact International Camp Centre, Periyanaickenpalayam*\n\n _30 km from International Airport_\n _24 km from Railway Station_ \n_23 km from Gandhipuram Bus Stand_\n\n*Camp Center Access:* Opens on the 12th at 02:00 PM.\n\n*Program Timings:*\nStarts at 04:00 PM on the 12th and ends at 05:00 PM on the 13th.\n\n*Meals Included:*\nDinner, breakfast, and lunch.\n\n*Important Notes:*\n- We will depart from Talmid House by bus at 02:00 PM on the 12th. If you're arriving after 2 PM, kindly plan your transportation accordingly.\n- Kindly let us know about early arrival or late departure for us to plan your accommodation.\n- If you require any assistance for extra transportation, meals, or accommodation outside the program, please notify us in advance.\n\n*Thank you*"""
    
    try:
        encoded_message = urllib.parse.quote(message)
        
        chat_url = f"https://web.whatsapp.com/send?phone={number}&text={encoded_message}"
        driver.get(chat_url)
        time.sleep(2)
        send_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
        )
        
        send_button.click()
        
        # print(f"Message sent to {name} ({number})")
        print(f"Message sent to ({number})")
        
    except Exception as e:
        print(f"Failed to send message to  ({number}): {e}")
    
    time.sleep(5)

driver.quit()
