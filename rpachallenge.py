import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import os
from tkinter import messagebox
import logging

try:
    logging.basicConfig(format='%(asctime)s %(message)s', datefmt='%H:%M:%S',level=logging.INFO)
    logging.info('Started')

    dirname = os.path.abspath('.')
    inputExcelFile =  os.path.join(dirname,'challenge.xlsx')
    outputSS = os.path.join(dirname,'result.png')
    chromedriver = os.path.join(dirname,'chromedriver.exe')

    df = pd.read_excel(inputExcelFile,'Sheet1')

    options = Options()
    options.add_argument("start-maximized")
    options.headless = True

    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    except Exception as e:
        driver = webdriver.Chrome(chromedriver, options=options)

    driver.get('https://www.rpachallenge.com/')

    startBtn = driver.find_element(By.XPATH,'.//*[@class="waves-effect col s12 m12 l12 btn-large uiColorButton"]')
    submitBtn = driver.find_element(By.XPATH,'.//*[@class="btn uiColorButton"]')

    startBtn.click()
    for idx, row in df.iterrows():
        firstNameField = driver.find_element(By.XPATH,'.//*[@ng-reflect-name="labelFirstName"]')
        LastNameField = driver.find_element(By.XPATH,'.//*[@ng-reflect-name="labelLastName"]')
        companyNameField = driver.find_element(By.XPATH,'.//*[@ng-reflect-name="labelCompanyName"]')
        roleField = driver.find_element(By.XPATH,'.//*[@ng-reflect-name="labelRole"]')
        addressField = driver.find_element(By.XPATH,'.//*[@ng-reflect-name="labelAddress"]')
        emailField = driver.find_element(By.XPATH,'.//*[@ng-reflect-name="labelEmail"]')
        phoneField = driver.find_element(By.XPATH,'.//*[@ng-reflect-name="labelPhone"]')
        firstNameField.send_keys(row[0])
        LastNameField.send_keys(row[1])
        companyNameField.send_keys(row[2])
        roleField.send_keys(row[3])
        addressField.send_keys(row[4])
        emailField.send_keys(row[5])
        phoneField.send_keys(row[6])

        submitBtn.click()

    driver.save_screenshot(outputSS)
    driver.close()
    driver.quit()

    logging.info('Finished')

except Exception as e:
    logging.error(e)
    messagebox.showerror("Error",e)
