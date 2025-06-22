import time
import json
import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from openpyxl import load_workbook


changed_cells = []
change_count = 0
timeout_for_pageload = 30


def login(driver, config):
    with open('credentials.json') as creads_file:
        creds = json.load(creads_file)
    username = creds["username"]
    password = creds["password"]

    driver.get(config["loginUrl"])
    time.sleep(2)
    driver.find_element(By.XPATH, config["loginUsernameField"]).send_keys(username)
    driver.find_element(By.XPATH, config["loginPasswordField"]).send_keys(password)
    driver.find_element(By.XPATH, config["loginSubmitButton"]).click()
    time.sleep(3)
 


def navigate_to_entry(driver, entry):
    if "url" in entry:
        driver.get(entry["url"])

    WebDriverWait(driver, timeout_for_pageload).until(EC.presence_of_element_located((By.TAG_NAME, "body")))


    if "iFrameToSwitchTo" in entry:
        driver.switch_to.frame(entry["iFrameToSwitchTo"])


    for key in ["elementOneToClickOn", "elementTwoToClickOn", "elementThreeToClickOn"]:
        if key in entry:
            WebDriverWait(driver, timeout_for_pageload).until(EC.presence_of_element_located((By.XPATH, entry[key])))
            driver.find_element(By.XPATH, entry[key]).click()



def save_changed_content(config, excel_file):
    if (len(changed_cells) > 0):
        excel_file.save(config["excelFilePath"])

        print(f"Es wurde(n) {len(changed_cells)} Zelle(n) aktualisiert:")
        for changed_cell in changed_cells:
            print(f"{changed_cell[0]}: {changed_cell[1]} -> {changed_cell[2]}")
    else:
        print("Es wurden keine Zellen aktualisiert.")

# Beispiel für die Initialisierung des Chrome WebDrivers mit unterdrückter Ausgabe:
def get_silent_chrome_driver(chrome_options=None):
    service = Service(log_path=os.devnull)
    if chrome_options is None:
        chrome_options = Options()
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver