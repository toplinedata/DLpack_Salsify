# -*- coding: utf-8 -*-
"""
Created on Wed Oct 16 09:57:57 2019

@author: Ossang
"""

import os
from os.path import exists
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

# Today's date
date_label = time.strftime('%Y%m%d')
try:
    #local
    os.chdir('C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Salsify\\')
except:
    #0047
    os.chdir('N:\\E Commerce\\Public Share\\Salsify Assortment\\')

#storage Directory
if 'Desktop\\0047Automate_Script' in os.getcwd():
    driver_path = 'C:\\Users\\User\\Anaconda3\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Salsify\\'
    Download_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Salsify\\'
else:      
    driver_path = 'C:\\Users\\raymond.hung\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\raymond.hung\\Documents\\Automate_Script\\DLpack_Salsify\\'
    Download_dir = 'N:\\E Commerce\\Public Share\\Salsify Assortment\\'


# Account and Password
login_info = pd.read_csv(work_dir+ 'Account & Password.csv',index_col=0)
username = login_info.loc['Account', 'CONTENT']
password = login_info.loc['Password', 'CONTENT']

# Chrome driver setting
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {'download.default_directory' : Download_dir}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(driver_path,chrome_options=options)
    
# Supplier Oasis website turn into login page
Supplier_Oasis = 'https://app.salsify.com/users/sign_in'
driver.get(Supplier_Oasis)
LoadingChecker = (By.CSS_SELECTOR, 'body > div > form > button')
WebDriverWait(driver, 30).until(EC.presence_of_element_located(LoadingChecker))
driver.find_element_by_css_selector('body > div > form > button').click()

LoadingChecker = (By.CSS_SELECTOR, '.login-button')
WebDriverWait(driver, 30).until(EC.presence_of_element_located(LoadingChecker))
# Input username and password and login
driver.find_element_by_id('user_email').send_keys(username)
driver.find_element_by_id('user_password').send_keys(password)
driver.find_element_by_css_selector('.login-button').click()
time.sleep(10)

# click on products > View All > Actions >Download All Columns
driver.find_element_by_css_selector('body > div.ember-view > div > nav > div.ember-view > div > div > div').click()
time.sleep(3)

driver.find_element_by_css_selector('body > div > div.drop-content > div > div > div > a > div').click()
time.sleep(3)

driver.find_element_by_css_selector('#content > div > div._header_3hfy1l > div > div > div > button').click()
time.sleep(3)

driver.find_element_by_css_selector('#content > div > div._header_3hfy1l > div > div > div > ul > li:nth-child(5) > a').click()
# #ember549 > ul > li:nth-child(5) > a
time.sleep(60)

if exists('Salsify export_'+date_label+'.xlsx'):
    os.remove('Salsify export_'+date_label+'.xlsx')
os.rename('export.xlsx', 'Salsify export_'+date_label+'.xlsx')
driver.quit()