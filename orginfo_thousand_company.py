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
ids = pd.read_excel("cooperation_company.xlsx")
codes = ids['ids'].tolist()
length = len(codes)
print(f"Number of companies to be checked: {length}")
print("<<<<<< STARTING >>>>>>")

# driver
driver = webdriver.Chrome(ChromeDriverManager().install())

desired_output = {
    "inn1": {
        "founder1":{
            "company_name": "some_compa
            ny",
            "inn": "inn_num",
            "name": "XYZ",
            "rate": "25",
        },
         "founder2":{
            "inn": "inn_num",
            "name": "XYZ",
            "rate": "25",
        }
    },
    "inn2": {
        "company_name": "some_company",
        "founder1":{
            "inn": "inn_num",
            "name": "XYZ",
            "rate": "25",
        },
         "founder2":{
            "inn": "inn_num",
            "name": "XYZ",
            "rate": "25",
        }
    },
    "inn3": {"company_name": "some_company",
             "founder1": "no_founder_found"},
    "inn4": "no_company_found",
    "inn5": {
        "company_name": "some_company",
        "founder1":{
            "inn": "inn_num",
            "name": "XYZ",
            "rate": "25",
    },
        "founder2":{
            "inn": "inn_num",
            "name": "XYZ",
            "rate": "25",
    },
        "founder3":{
            "inn": "inn_num",
            "name": "XYZ",
            "rate": "25",
    }
},
}

TOTAL_OUTPUT = {}
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

    # find Output --> Find organization link if exists --> open the link
    output  = driver.find_element(By.CLASS_NAME, "oranization-search-result")

    if output.text == "Hech narsa topilmadi":
        TOTAL_OUTPUT[code]="No Company Found"
        continue

    else:

        company_name = driver.find_element("xpath", '(//h6)[2]').text()
        TOTAL_OUTPUT[code] = {"company_name": company_name, "founders":[]} 

        link = driver.find_element("xpath", '//a[@class="organization-page-link"]')
        link = link.get_attribute('href')
        driver.get(link)
        time.sleep(1)

        try:
            elements = driver.find_elements(By.XPATH, '(//h5)[3]/../following-sibling::*')
            num_elements = len(elements)
            my_elements=[]
            for element in elements:
                if element.get_attribute("class") == "col-sm-8":
                    element = driver.find_element("xpath", f"{element}/a").text()
                    my_elements.append(element)
                else:
                    element = element.text()
                    my_elements.append(element)
            count = 1
            FOUNDERS = []
            while len(my_elements) > 0:
                founder = {}
                company_name = my_elements.pop(0)
                # Now we need to get the inn number of the company
                founder["company_name"]=company_name
                if " %" in my_elements[0]:
                    contribution = my_elements.pop(0)
                    founder["rate"] = contribution
                else:
                    founder["rate"] = "None"
                FOUNDERS.append(founder)
                count += 1
            TOTAL_OUTPUT[code]["founders"] = FOUNDERS

        except:
            TOTAL_OUTPUT[code]["founders"] = "No Founder Found" 
            continue


        # Save file every 100th time or if it is the last line
    if sum % 10 == 0 or sum == length-1:
        pd.DataFrame({
            "STIR": stirs,
            "MANZIL": address,
            "Manzil_tuman": MANZIL_2,
            "Telefon": TELEFON}).to_excel(f"hududlar_{sum}.xlsx")

print("<<<<<< Finished successfully! >>>>>>")
