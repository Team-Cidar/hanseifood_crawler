from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC


def crawling(driver):
    driver.switch_to.frame("IframePortlet_13444")
    a_tag = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div/table/tbody/tr[1]/td[2]/span/a")))
    a_tag.click()
    download_a = WebDriverWait(driver, 10).until(EC.visibility_of_element_located(
        (By.XPATH, "/html/body/div[2]/div[3]/div[1]/table[1]/tbody/tr[4]/td/span/a[2]")))
    download_a.click()


# Chrome Driver Path, Options
path = "./drivers/chromedriver"
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {'intl.accept_languages': 'ko'})  # driver language setting

driver = webdriver.Chrome(executable_path="./drivers/chromedriver", options=options)
driver.get("https://portal.hansei.ac.kr/portal/default/gnb/hanseiTidings/weekMenuTable.page")

crawling(driver)  # download xlsx file from site

while True:
    if input() == "q":
        driver.close()
        driver.quit()
        exit(1)
