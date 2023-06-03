import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import pickle


def crawling():
    ...


# Chrome Driver Path, Options
path = "./drivers/chromedriver"
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {'intl.accept_languages': 'ko'})  # driver language setting

# Chrome Driver
driver = webdriver.Chrome(executable_path=path, options=options)

# Logic
driver.get("https://portal.hansei.ac.kr/portal/default/gnb/hanseiTidings/weekMenuTable.page")

tableNumberElement = "first C"
print(driver.find_elements(By.CLASS_NAME, tableNumberElement))


# 종료 트리거
while True:
    if input() == "q":
        driver.close()
        driver.quit()
        exit(1)
