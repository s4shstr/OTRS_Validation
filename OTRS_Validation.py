#!/usr/bin/env python
# -*- coding: utf-8 -*-
from __future__ import print_function
import io
import codecs
import html2text
import codecs
import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from getpass import getpass
import pyautogui,time
import pandas as pd
import numpy as np


BLOCKSIZE = 1048576
str1 = '['
str2 = '.'
str3 = '<SelectedValues>'
k = 0
j = 0
firstClient = 'randomword'
flag = False
temppath = os.getenv('TEMP')
browser = webdriver.Chrome('chromedriver.exe')


# OTRS staff info parsing and tables preparing (preview info)
company_name = str(input('Введите сокращенное наименование компании: '))
try:
    os.mkdir(temppath + "\OTRS_Validation");
except OSError:
    shutil.rmtree(temppath + "\OTRS_Validation");
    os.mkdir(temppath + "\OTRS_Validation");
browser.get('https://hd.itfb.ru/index.pl?Action=AdminCustomerUser;Nav=' + company_name)
s_username = browser.find_element_by_id('User')
s_password = browser.find_element_by_id('Password')
s_continue = browser.find_element_by_id('LoginButton')
s_username.send_keys(str(input('Введите логин: ')))
s_password.send_keys(str(getpass('Введите пароль: ')))
s_continue.click()
company_name_field = browser.find_element_by_xpath('//*[@id="Search"]')
company_search_button = browser.find_element_by_xpath('//form[contains(@class, "SearchBox")]//button[contains(@title,"Поиск")]')
browser.execute_script("arguments[0].click();", company_name_field)
browser.execute_script("arguments[0].value = '" + company_name + "';", company_name_field);
browser.execute_script("arguments[0].click();", company_search_button)
page = browser.page_source
otrs_staff_raw_html = pd.read_html(browser.find_element_by_id("CustomerTable").get_attribute('outerHTML'))
otrs_staff_raw_xlsx = otrs_staff_raw_html[0].drop(otrs_staff_raw_html[0].columns[[3, 4, 5]], axis='columns')
otrs_staff_raw_xlsx.to_excel(temppath + '\OTRS_Validation\\' + company_name + '_otrs_staff_raw.xlsx', index=False)

# OTRS staff info parsing, tables preparing, collisions searching (extended info)
digit_count = 0
otrs_staff = pd.read_excel(temppath + '\OTRS_Validation\\' + company_name + '_otrs_staff_raw.xlsx', encoding = 'utf-8')
for index, row in otrs_staff.iterrows():
    otrs_login = row['Логин'].split('@')[0]
    print(otrs_login + ' (' + str(index + 1) + ' of ' + str(len(otrs_staff)) + ')')
    browser.get('https://hd.itfb.ru/index.pl?Action=AdminCustomerUser;Subaction=Change;ID='+ otrs_login + '%40' + company_name + ';Search=' + company_name + ';Nav=Agent')
    browser.find_element_by_xpath('//*[@id="UserVip"]').click()
    time.sleep(0.5)
    #browser.find_element_by_xpath('//*[@id="ValidID"]').click()
    #phone_extension_raw = Select(browser.find_element_by_xpath('//*[@id="ValidID"]'))
    pyautogui.press('tab')  
    pyautogui.press('down')
    pyautogui.press('down') 
    pyautogui.press('down') 
    pyautogui.press('Enter')
    browser.find_element_by_xpath('//*[@id="SubmitAndContinue"]').click()


browser.close()
browser.quit()

# Deleting temp files
os.remove(temppath + '\OTRS_Validation\\' + company_name + '_otrs_staff_raw.xlsx')
