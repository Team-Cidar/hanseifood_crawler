import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

import pickle


def crawling():
    ...


options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {'intl.accept_languages': 'ko'})  # driver language setting

driver = webdriver.Chrome(executable_path="./drivers/chromedriver", options=options)
driver.get("https://www.naver.com")

while True:
    if input() == "q":
        driver.close()
        driver.quit()
        exit(1)
