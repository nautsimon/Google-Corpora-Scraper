# imports
from selenium import webdriver
import pandas as pd
import time
from random import randint
from pandas import DataFrame

# Set your username and password
usr =
pwd =


# Init Selenium
driver = webdriver.Chrome()
driver.get("https://www.english-corpora.org/login.asp")

# Locate login form elements
loginE = driver.find_element_by_xpath(
    "/html/body/div/div/table/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input")
loginP = driver.find_element_by_xpath(
    "/html/body/div/div/table/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[2]/td[2]/input")
loginB = driver.find_element_by_xpath(
    "/html/body/div/div/table/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[3]/td[2]/input[1]")

# Login
loginE.send_keys(usr)
loginP.send_keys(pwd)
loginB.click()

time.sleep(2)

# Go to corpora dashboard
driver.get("https://www.english-corpora.org/googlebooks/x.asp")
driver.find_element_by_xpath(
    "/html/body/div/table/tbody/tr[5]/td/table/tbody/tr[2]/td[1]/p/a").click()

time.sleep(2)


# enter the path to an excel file with a list of the words you want to search
excelPath = 'name.xlsx'
df = pd.read_excel(excelPath, index_col=0)
count = 0
data = []

# if scraping gets interrupted change this variable to the excel index
startAt = 0

# Scrape
for row in df.iterrows():
    count += 1
    if count < startAt:
        continue
    time.sleep(randint(1, 2))

    # Navigate page frames
    driver.switch_to.frame(3)

    # Conduct search
    queryI = driver.find_element_by_xpath('//*[@id="p"]')
    queryI.clear()
    queryI.send_keys(str(row[0]))
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="B7"]').click()
    driver.find_element_by_xpath('//*[@id="B7"]').click()
    time.sleep(randint(1, 2))

    # Collect frequency data
    driver.switch_to.default_content()
    time.sleep(randint(1, 2))
    driver.switch_to.frame(5)
    time.sleep(randint(1, 2))
    newRow = []
    newRow.append(row[0])
    for year in range(17, 26):
        try:
            newRow.append(driver.find_element_by_xpath(
                '//tbody/tr[2]/td['+str(year)+']/a/font').text)
        except:
            newRow.append("0")
    print(newRow)
    data.append(newRow)
    time.sleep(randint(1, 2))
    driver.switch_to.default_content()

# Export csv with frequency data
outDf = DataFrame(data, columns=[
                  'name', '1920', '1930', '1940', '1950', '1960', '1970', '1980', '1990', '2000']
outDf.to_csv('freqData.csv', encoding='utf-8')
