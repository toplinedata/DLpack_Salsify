# -*- coding: utf-8 -*-
"""
Created on Wed Oct 16 09:57:57 2019

@author: Ossang Ou
"""

import os
import shutil
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# Today's date
date_label = time.strftime('%Y%m%d')
try:
    #local
    os.chdir('C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Salsify\\DateCheck\\')
except:
    #0047
    os.chdir('N:\\E Commerce\\Public Share\\Salsify Assortment\\')

#storage Directory
if 'Desktop\\0047Automate_Script' in os.getcwd():
    driver_path = 'C:\\Users\\User\\Anaconda3\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Salsify\\'
    Download_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Salsify\\DateCheck\\'
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

for i in range(5):
    try:
        # Supplier Oasis website turn into login page
        Supplier_Oasis = 'https://app.salsify.com/users/sign_in'
        driver.get(Supplier_Oasis)
        LoadingChecker = (By.CSS_SELECTOR, 'body > div > form > button') # Allow all cookies
        WebDriverWait(driver, 30).until(EC.presence_of_element_located(LoadingChecker))
        driver.find_element_by_css_selector('body > div > form > button').click()
        
        LoadingChecker = (By.CSS_SELECTOR, '.login-button') # Log In
        WebDriverWait(driver, 30).until(EC.presence_of_element_located(LoadingChecker))
        # Input username and password and login
        driver.find_element_by_id('user_email').send_keys(username)
        driver.find_element_by_id('user_password').send_keys(password)
        driver.find_element_by_css_selector('.login-button').click()
        
        # Transfer to Tasks Page 
        LoadingChecker = (By.CSS_SELECTOR, 'body > div.ember-view > div > nav > div.ember-view > div > div > div') # More
        WebDriverWait(driver, 30).until(EC.presence_of_element_located(LoadingChecker))
        datecheck_link ='https://app.salsify.com/app/orgs/s-69d82988-277e-4ffd-ab76-b3b97ca474b8/content-flow/tasks?all=true'
        driver.get(datecheck_link)
               
        # Assignee Filter
        ASSIGNEE = {
                    'Danny Hsu' : 'Marketing Content'
                    , 'Lisa Majerchin' : 'Digital Asset'
                    , 'Claudia' : 'Product Info'
                    , 'Dot com AE' : 'AE Final Review'
                    }

        for assign in ASSIGNEE:
            main_assign = '#content > div._content-flow_dvjdnj > div._content_ffkuhn > div._sidebar_ffkuhn > div > div:nth-child(1) > '
            LoadingChecker = (By.CSS_SELECTOR, main_assign+'div:nth-child(1) > div > div:nth-child(2)') # ASSIGNEE
            WebDriverWait(driver, 60).until(EC.presence_of_element_located(LoadingChecker))
            
            assign_no = driver.find_elements_by_css_selector(main_assign+'div')
            for i in range(2,len(assign_no)+1):
                if driver.find_element_by_css_selector(main_assign+'div:nth-child('+str(i)+') > div > label > div > div:nth-child(2)').text == assign:
                    driver.find_element_by_css_selector(main_assign+'div:nth-child('+str(i)+') > div > label > div').click()
            
            time.sleep(10)
            # Download Filtered Activity
            LoadingChecker = (By.CSS_SELECTOR, '#content > div._content-flow_dvjdnj > div:nth-child(3) > div:nth-child(2) > div:nth-child(1) > button') # Actions
            WebDriverWait(driver, 30).until(EC.presence_of_element_located(LoadingChecker)) 
            driver.find_element_by_css_selector('#content > div._content-flow_dvjdnj > div:nth-child(3) > div:nth-child(2) > div:nth-child(1) > button').click()
            
            LoadingChecker = (By.CSS_SELECTOR, '#content > div._content-flow_dvjdnj > div:nth-child(3) > div:nth-child(2) > div:nth-child(2) > div > div > a:nth-child(5)') # Download Filtered Activity
            WebDriverWait(driver, 30).until(EC.presence_of_element_located(LoadingChecker))
            driver.find_element_by_css_selector('#content > div._content-flow_dvjdnj > div:nth-child(3) > div:nth-child(2) > div:nth-child(2) > div > div > a:nth-child(5)').click()
            time.sleep(60)
            driver.get(datecheck_link) # Filter resetup
                                                
            if os.path.exists(Download_dir+'tasks.csv'):
                shutil.move(Download_dir+'tasks.csv', Download_dir+ASSIGNEE[assign]+' tasks_'+date_label+'.csv')

        if len(os.listdir(Download_dir)) >= len(ASSIGNEE): # Confirm download file numbers
            break
        
    except Exception as e:
        print(e)
        driver.get(datecheck_link)
        time.sleep(60)

driver.quit()

# Define Send Mail function
def sendreport(mailinfo, infoSerial, AttachFile):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email import encoders
    import pandas
    import time
    # import os

    # pwd = os.path.abspath('C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Salsify\\')
    pwd = Download_dir
    info = pandas.read_csv(mailinfo)
    MyMail = pandas.read_csv(pwd + "\\UIDandPW.csv").ID[1]
    MyMailPW = pandas.read_csv(pwd + "\\UIDandPW.csv").PW[1]
    mailto = info.To[infoSerial].split("/")
    Cc = info.cc[infoSerial].split("/")

    emailMsg = MIMEMultipart('alternative')
    emailMsg['Subject'] = info.Subject[infoSerial]
    emailMsg['From'] = MyMail
    emailMsg['To'] = ", ".join(mailto)
    emailMsg['Cc'] = ", ".join(Cc)
    
    Attachs=AttachFile.split("|")
    for files in Attachs:
        att = MIMEBase('application', "octet-stream")
        att.set_payload(open(files, "rb").read())
        encoders.encode_base64(att)
        att.add_header('Content-Disposition', 'attachment', filename=files)
        emailMsg.attach(att)

    smtpObj = smtplib.SMTP('smtp.office365.com',587)
    smtpObj.ehlo()
    smtpObj.starttls()
    smtpObj.login(MyMail,MyMailPW)
    smtpObj.sendmail(MyMail,mailto+Cc,emailMsg.as_string())
    time.sleep(10)
    smtpObj.quit()


# Send Mail
filelist = ""
count = 1
for file in os.listdir(Download_dir):
    if count < len(os.listdir(Download_dir)):
        filelist+=file+"|"
        count+=1
    else:
        filelist+=file
sendreport(work_dir+'mailinfo.csv',0,filelist)
