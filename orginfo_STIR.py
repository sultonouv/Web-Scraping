# Imports
import selenium
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
import pandas as pd
import random
import os
from tqdm import tqdm

# URL and Codes list
url = 'https://orginfo.uz'
ids = pd.read_excel("ids.xlsx")
codes = ids['ids'].tolist()
length = len(codes)
print(f"Number of companies to be checked: {length}")
print("<<<<<< STARTING >>>>>>")

# driver
driver = webdriver.Chrome(ChromeDriverManager().install())
address = []
MANZIL_2= []
stirs = []
TELEFON = []
sum = 0

for code in tqdm(codes):
    sum = sum + 1

    driver.get(url)
    driver.maximize_window()

    # Find Searchbar --> Enter the code --> Press "Enter"
    searchbar = driver.find_element("name","q")
    searchbar.send_keys(code)
    searchbar.send_keys(Keys.RETURN)
    time.sleep(1)
    

    # if step is divisible by 50, print step number and sleep for 60 seconds
    

    # find Output --> Find organization link if exists --> open the link
    output  = driver.find_element(By.CLASS_NAME, "oranization-search-result")

    if output.text == "Hech narsa topilmadi":
        stirs.append(code)
        address.append("Korxona topilmadi")
        MANZIL_2.append("Manzil topilmadi")
        TELEFON.append("Telefon topilmadi")
        continue

    else:
        link = driver.find_element("xpath", '//a[@class="organization-page-link"]')
        link = link.get_attribute('href')
        driver.get(link)
        time.sleep(1)

        # get the address, only the region part until <,>
        manzil = driver.find_element("xpath", "/html/body/div/div[2]/div/main/div[3]/dl/dd[9]").text
        telefon = driver.find_element("xpath", "/html/body/div/div[2]/div/main/div[3]/dl/dd[10]/a").text
        manzil_1 = manzil.split(",")[0]
        manzil_2 = manzil.split(",")[1]


        stirs.append(code)
        address.append(manzil_1)
        MANZIL_2.append(manzil_2)
        TELEFON.append(telefon)

        # Save file every 100th time or if it is the last line
    if sum % 10 == 0 or sum == length-1:
        pd.DataFrame({
            "STIR": stirs,
            "MANZIL": address,
            "Manzil_tuman": MANZIL_2,
            "Telefon": TELEFON}).to_excel(f"hududlar_{sum}.xlsx")

print("<<<<<< Finished successfully! >>>>>>")
