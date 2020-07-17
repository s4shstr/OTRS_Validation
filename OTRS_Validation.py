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


# Creating console menu
print('Выберите действие, которое будет произведено со всеми пользователями.\n1 - всем пользователям будет присвоен статус "недействительный"\n2 - всем пользователям будет присвоен статус "действительный"\n')
validID_value = input('Введите число: ')
while not validID_value.isdigit():
    print('Введенное значение не является числом.')
    validID_value = input('Введите число: ')
while int(validID_value) != 1 and int(validID_value) != 2:
    print('Введено некорректное значение, пожалуйста введите числа "1" или "2".')
    validID_value = input('Введите "1" или "2": ')
if int(validID_value) == 1:
    validID = 'недействительный'
elif int(validID_value) == 2:
    validID = 'действительный'
company_name = str(input('Введите ID клиента (сокращенное наименование компании): '))
username = str(input('Введите логин от OTRS: '))
password = str(getpass('Введите пароль от OTRS: '))

# OTRS company searching --> number of users counting
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')
browser = webdriver.Chrome('chromedriver.exe', options = options)
browser.get('https://hd.itfb.ru/index.pl?Action=AdminCustomerUser;Nav=' + company_name)
browser.find_element_by_id('User').send_keys(username)
browser.find_element_by_id('Password').send_keys(password)
browser.find_element_by_id('LoginButton').click()
company_name_field = browser.find_element_by_xpath('//*[@id="Search"]')
company_search_button = browser.find_element_by_xpath('//form[contains(@class, "SearchBox")]//button[contains(@title,"Поиск")]')
browser.execute_script("arguments[0].click();", company_name_field)
browser.execute_script("arguments[0].value = '" + company_name + "';", company_name_field);
browser.execute_script("arguments[0].click();", company_search_button)
page = browser.page_source
otrs_staff_raw_html = pd.read_html(browser.find_element_by_id("CustomerTable").get_attribute('outerHTML'))
otrs_staff_raw_xlsx = otrs_staff_raw_html[0].drop(otrs_staff_raw_html[0].columns[[3, 4, 5]], axis='columns')

# Users profile editing
digit_count = 0
for index, row in otrs_staff_raw_xlsx.iterrows():
    otrs_login = row['Логин']
    browser.get('https://hd.itfb.ru/index.pl?Action=AdminCustomerUser;Subaction=Change;ID=' + otrs_login + ';Search=' + company_name + ';Nav=Agent')
    otrs_login = row['Логин'].split('@')[0]
    print(otrs_login + ' (' + str(index + 1) + ' of ' + str(len(otrs_staff_raw_xlsx)) + ')')
    if int(validID_value) == 1 and validID == browser.find_element_by_xpath('/html/body/div[1]/div[3]/div[4]/div/div[2]/form/fieldset/div[33]/div[1]/div/div/div').text:
        continue
    elif int(validID_value) == 2 and validID == browser.find_element_by_xpath('/html/body/div[1]/div[3]/div[4]/div/div[2]/form/fieldset/div[33]/div[1]/div/div/div').text:
        continue
    browser.find_element_by_xpath('//*[@id="ValidID_Search"]').click()
    wait = WebDriverWait(browser, 3)
    wait.until(EC.text_to_be_present_in_element((By.XPATH, "/html/body/div[4]/div[1]/div/ul/li[3]/a"), validID))
    browser.find_element_by_xpath(str("//a[text()='" + validID + "']")).click()
    browser.find_element_by_xpath('//*[@id="SubmitAndContinue"]').click()
print('\nИдет завершение программы, пожалуйста, подождите...\n')

# Browser closing
browser.close()
browser.quit()