# -*- coding: utf-8 -*-
"""
Created on Wed Jun 16 11:02:24 2021

@author: tchow
"""
## Import Packages ##
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from bs4 import BeautifulSoup
import win32clipboard
import pandas as pd
import my_config as my_c
from datetime import datetime
import os, fnmatch
import shutil
import re

ChromeDriverPATH = my_c.ChromeDriverPATH
Product_Path = my_c.Product_Path

driver = webdriver.Chrome(ChromeDriverPATH)
driver.get(Product_Path)


time.sleep(1)
driver.find_element_by_xpath("/html/body/div/div/div[1]/div/div/div[1]/div/div/input").send_keys(my_c.FL_login)
#Enter Password
driver.find_element_by_xpath("/html/body/div/div/div[1]/div/div/div[2]/div/div/input").send_keys(my_c.FL_password)
#Click Login Button
driver.find_element_by_xpath("/html/body/div/div/div[1]/div/div/div[3]/button/span[1]").click()

time.sleep(5)
item_links = []

more_buttons = driver.find_elements_by_xpath('//*[@title="Next page" and @class="MuiButtonBase-root MuiIconButton-root MuiIconButton-colorInherit"]')

def GetLinks(dps):
    soup = BeautifulSoup(dps, features="lxml")
    for x in soup.find_all('a', href=re.compile('^/products/*'), class_="MuiTypography-root MuiLink-root MuiLink-underlineHover MuiTypography-colorPrimary"):
        item_links.append(x['href'])
        #print(x.get_text())
    return item_links

GetLinks(driver.page_source)
while len(more_buttons) != 0:
    GetLinks(driver.page_source)
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)
    more_buttons = driver.find_elements_by_xpath('//*[@title="Next page" and @class="MuiButtonBase-root MuiIconButton-root MuiIconButton-colorInherit"]')
    if len(more_buttons) == 0:
        break
    more_buttons[0].click()
    time.sleep(1)


inventoryCode = []
name = []
subName = []
price = []
FD_ID_LINK = []

for link in item_links:
    driver.get(my_c.server_path + link)
    FD_ID_LINK.append(my_c.server_path + link)
    time.sleep(4)
    #inventoryCode //*[@id="variant-sku"]
    driver.find_element_by_xpath('//input[@name="variant.sku"]')
    driver.find_element_by_xpath('//input[@name="variant.sku"]').send_keys(Keys.CONTROL, 'a') #highlight all in box
    driver.find_element_by_xpath('//input[@name="variant.sku"]').send_keys(Keys.CONTROL, 'c') #copy
    win32clipboard.OpenClipboard()
    inventoryCode.append(win32clipboard.GetClipboardData()) #paste
    win32clipboard.CloseClipboard()
    time.sleep(.5)
    #Item Name
    driver.find_element_by_xpath('//input[@name="name"]')
    driver.find_element_by_xpath('//input[@name="name"]').send_keys(Keys.CONTROL, 'a') #highlight all in box
    driver.find_element_by_xpath('//input[@name="name"]').send_keys(Keys.CONTROL, 'c') #copy
    win32clipboard.OpenClipboard()
    name.append(win32clipboard.GetClipboardData()) #paste
    win32clipboard.CloseClipboard()
    time.sleep(.5)
    #subName
    driver.find_element_by_xpath('//input[@name="variant.name"]')
    driver.find_element_by_xpath('//input[@name="variant.name"]').send_keys(Keys.CONTROL, 'a') #highlight all in box
    driver.find_element_by_xpath('//input[@name="variant.name"]').send_keys(Keys.CONTROL, 'c') #copy
    win32clipboard.OpenClipboard()
    subName.append(win32clipboard.GetClipboardData()) #paste
    win32clipboard.CloseClipboard()
    time.sleep(.5)
    #price
    driver.find_element_by_xpath('//input[@name="variant.price"]')
    driver.find_element_by_xpath('//input[@name="variant.price"]').send_keys(Keys.CONTROL, 'a') #highlight all in box
    driver.find_element_by_xpath('//input[@name="variant.price"]').send_keys(Keys.CONTROL, 'c') #copy
    win32clipboard.OpenClipboard()
    price.append(win32clipboard.GetClipboardData()) #paste
    win32clipboard.CloseClipboard()
    
path = 'M:\\TC\\Freshline\\'

for files in os.listdir(path):
    if fnmatch.fnmatch(files, 'datascraping_FreshDelish*.xlsx'):
        shutil.move(os.path.join(path, files),os.path.join(path + 'Archived IM\\', files))
    
df = pd.DataFrame(list(zip(inventoryCode, name, subName, price, FD_ID_LINK)),columns =['inventoryCode', 'name', 'subName', 'price', 'FD_ID_LINK'])
df.to_excel('M:\\TC\\Freshline\\datascraping_FreshDelish ' + str(datetime.now().strftime("%Y-%m-%d")) +'.xlsx') 
#print(df)

driver.quit()

