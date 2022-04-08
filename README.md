# Project Title
Case_Study2: Web Crawling for printing automation.

## Getting Started

These instructions will get you running on your local machine for development and testing purposes.

## Imports

from selenium import webdriver

from selenium.webdriver.common.keys import Keys

from selenium.webdriver.support.ui import WebDriverWait

from selenium.webdriver.common.by import By

from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import ElementClickInterceptedException

from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.chrome.service import Service

from selenium.webdriver.remote.webelement import WebElement

from selenium.common.exceptions import NoSuchElementException

import numpy as np

import pandas as pd

import time

## Installing

A step by step series of examples that tell you how to get a development enviromnent running

Say what the step will be

## Driver configuration.
Selenium 
Chromedriver

## Start the chromedriver on your local machine

path_to_chromedriver ='C:\\Users\\Ye Ryoung Kim\\Desktop\\chromedriver_win32\\chromedriver'

browser = webdriver.Chrome(executable_path = path_to_chromedriver)

## Set the url of "lims3.psomagen.com"

url = 'https://lims3usqas.psomagen.com/main.do'

browser.get(url)

## Send Username &  Password.

browser.find_element_by_name('username')

browser.find_element_by_name('username').clear() # Make sure that the form is empty

browser.find_element_by_name('username').send_keys('yrkim') # writing ID 

browser.find_element_by_name('password') # find the 'passwd' element by name

browser.find_element_by_name('password').clear()# Make sure that the form is empty

browser.find_element_by_name('password').send_keys('Initial0)')# writing password

browser.find_element_by_xpath('/html/body/div[1]/div[1]/form/div/div[1]/button').click()# Searching the login button's xpath and click it.

## Web scraping - read table of the listed rxn sheet numbers on "rxn sheet"

rxn_sheet_table_df = pd.read_html(browser.find_element(by=By.XPATH,value ="/html/body/div[2]/div[2]/div/form/div[2]/div/div/div/div/div/table/tbody/tr[3]").get_attribute('outerHTML'))

rxn_plate_dept = rxn_sheet_table_df[0][3]

rxn_plate_id_col = rxn_sheet_table_df[0][5] # Reading reaction sheets id column.

## Functions

### F_1 :Check the name of the department that you want to print the sheets.

def check_dept(table):
    list_dept = []  # Get the row number of designated department.
    for i in range(len(rxn_plate_dept)):
        if rxn_plate_dept[i] == "Massachusetts":
            list_dept.append(i)
    list_dept.reverse() #reverse the list numbers to prioritize old rxn sheets to print out.
    return list_dept
    
### F_2 Click linked rxn sheet number to print out

def click_print_sheet(pos): # 'pos' is a list that consist of row numbers that designated department rxn sheets.

    click_sheet = browser.find_element(by=By.XPATH, value ="/html/body/div[2]/div[2]/div/form/div[2]/div/div/div/div/div/table/tbody/tr[3]/td/div/div[1]/table/tbody/tr[%d]/td[4]/b/u"%(pos+1))
    
    return click_sheet.click()
    
## Get the rxn sheet row numbers to click

positions = check_dept(rxn_plate_dept)

list_position = positions

## Using for loop to print out repeatedly

for p in list_position:

    try:
    
        sheet_button = browser.find_element(by=By.XPATH, value ="/html/body/div[2]/div[2]/div/form/div[2]/div/div/div/div/div/table/tbody/tr[3]/td/div/div[1]/table/tbody/tr[%d]/td[4]/b/u"%(p+1))
        
        sheet_button.click()
        
    # break
    
    except NoSuchElementException:
    
        print('No element of that id present!')
        
    try:
    
        wait = WebDriverWait(browser, 10)
        
        element = wait.until(EC.element_to_be_clickable((By.ID, 'btnSheetPrint')))   # click the "Rxn Sheet" button.
        
    finally:
    
        element.click()
        
        print("Print dialog is popped up")
        
    print("Please wait for 10 sec until the page is fully loaded")
    
    time.sleep(6)

    click_ces = browser.find_element(by=By.XPATH, value = "/html/body/div[2]/aside/section/div[1]/ul/li[2]/a")
    
    click_ces.click()
        
    click_order_management = browser.find_element(by=By.XPATH, value = "/html/body/div[2]/aside/section/div[2]/div[2]/ul/li[3]/a")
    
    click_order_management.click()
    
    click_rxn_sheet = browser.find_element(by=By.XPATH, value = "/html/body/div[2]/aside/section/div[2]/div[2]/ul/li[3]/ul/li[14]/a")
    
    click_rxn_sheet.click()
    
    click_seq_dept = browser.find_element(by=By.XPATH, value = "/html/body/div[2]/div[2]/div/form/div[1]/div[2]/ul/li[2]/div/ul/li[3]/label")
    
    click_seq_dept.click()
    
    click_status = browser.find_element(by=By.XPATH, value = "/html/body/div[2]/div[2]/div/form/div[1]/div[2]/ul/li[3]/div/select")
    
    click_status.click()
    
    
    click_ready = browser.find_element(by=By.XPATH, value = "/html/body/div[2]/div[2]/div/form/div[1]/div[2]/ul/li[3]/div/select/option[2]")
    
    click_ready.click()
    
    
    click_search = browser.find_element(by=By.XPATH, value = "/html/body/div[2]/div[2]/div/form/div[1]/div[2]/div/div/button")
    
    click_search.click()
    
    print("{}".format(p)+" "+"loop is finished")
    
    try:
        rxn_sheet_table_df = pd.read_html(browser.find_element(by=By.XPATH,value ="/html/body/div[2]/div[2]/div/form/div[2]/div/div/div/div/div/table/tbody/tr[3]").get_attribute('outerHTML'))
        
        rxn_plate_dept = rxn_sheet_table_df[0][3]
        
        rxn_plate_id_col = rxn_sheet_table_df[0][5]
        
        sheet_button = browser.find_element(by=By.XPATH, value ="/html/body/div[2]/div[2]/div/form/div[2]/div/div/div/div/div/table/tbody/tr[3]/td/div/div[1]/table/tbody/tr[%d]/td[4]/b/u"%(p+1))
    
    finally:
    
        time.sleep(5)
        
        print("New DOM !!!!!")
    
    















