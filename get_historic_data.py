import json
from msilib.schema import tables
from select import select
from shutil import ExecError
from tkinter import W
import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import date
from datetime import timedelta
import re
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook

MAIN_BTNS = {
    'Audience' : "//span[text()='Audience']",
    'Overview' : "//span[text()='Overview']",
    'Behavior' : "//span[text()='Behavior']",
    'Events' : "//span[text()='Events']",
    'Top events' : "//span[text()='Top Events']",
    'Evolok' : "//span[text()='EVOLOK']",
    'Mobile' : "//span[text()='Mobile']",
    'Acquisition': "//span[text()='Acquisition']",
    'All Traffic' : "//span[text()='All Traffic']",
    'Channels' : "//span[text()='Channels']"
}

METRICS = {
    'Active users' : "//span[text()='Active Users']",
    'New vs Returning' : "//span[text()='New vs Returning']",
    'New users' : "/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[1]/td[4]/div",
    'Returning users': "/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[2]/td[4]/div"
}

AUDIENCE_METRICS = {
    'Avg session duration' : 0,
    'Bounce rate' : 1,
    'New users' : 2,
    'Sessions per user' : 3,
    'Pages per session' : 4,
    'Pageviews' : 5,
    'Sessions' : 6,
    'Users' : 7

}

DROPDOWN = {
    'Dropdown' : '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div/div[1]/div[1]/div[2]/div[1]/div',
    'Item container' : '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div/div[1]/div[1]/div[2]/div[2]/div/div',
    'Items' : 'div'
}

DATE_RANGE = {
    'Date selector' : '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[1]/table/tbody/tr',
    'Input start' : '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[2]/table/tbody/tr/td[2]/div/div[2]/input[1]',
    'Input end' : '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[2]/table/tbody/tr/td[2]/div/div[2]/input[2]',
    'Apply' : '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[2]/table/tbody/tr/td[2]/div/input',
    
}

EXPORT = {
    'Export' : "//span[text()='Export']",
    'GS' : "//span[text()='Google Sheets']"
}

OTHERS = {
    'Login entitlement' : "//span[text()='Login/Entitlement']",
    'Select metric' : "//div[text()='Select a metric']",
    'Metric input' : "ID-searchBox",
    'Loading' : "//div[text()='Loading']"
}

DATE = {
    
    'DD btn' : '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[1]/div/div/div[1]/ul/li[4]/div/div[2]',
    'DD input' : '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[1]/div/div/div[4]/div/ul/li[2]/input', 
    'Date' : '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/thead/tr[1]/th[3]/div[1]',
    'Select': '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[3]/div[1]/div/span[1]/span[1]/select',
    'Max rows' : '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[3]/div[1]/div/span[1]/span[1]/select/option[9]',
    'Pagination': '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[3]/div[1]/div/span[3]/ul/li[2]',
}

EXCEL_HEADS = ['Login', 'Entitlement', 'Login/Entitlement']


def start_driver(url):

    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get(url)

    sleep(30)

    return driver

def get_audience_overview_metrics():

    driver = start_driver()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Audience'])))

    audience = driver.find_element(By.XPATH, MAIN_BTNS['Audience']) 

    audience.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Overview'])))

    overview = driver.find_element(By.XPATH, MAIN_BTNS['Overview']) 

    overview.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'galaxyIframe')))

    iframe = driver.find_element(By.ID, 'galaxyIframe')

    driver.switch_to.frame(iframe)

    sleep(3)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Date selector'])))

    date_selector = driver.find_element(By.XPATH, DATE_RANGE['Date selector'])

    date_selector.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input start'])))

    start = driver.find_element(By.XPATH, DATE_RANGE['Input start'])

    sleep(1)

    start.clear()

    sleep(1)

    start.send_keys(f'Jan 4, 2009')

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input end'])))

    end = driver.find_element(By.XPATH, DATE_RANGE['Input end'])

    sleep(1)

    end.clear()

    sleep(1)

    end.send_keys(f'Jan 3, 2023')

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Apply'])))

    apply = driver.find_element(By.XPATH, DATE_RANGE['Apply'])

    apply.click()

    sleep(3)

    WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

    sleep(3)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DROPDOWN['Dropdown'])))

    ov_dropdown = driver.find_element(By.XPATH, DROPDOWN['Dropdown'])

    ov_dropdown.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DROPDOWN['Item container'])))

    item_cont = driver.find_element(By.XPATH, DROPDOWN['Item container'])

    sleep(1)

    ov_items = item_cont.find_elements(By.TAG_NAME, DROPDOWN['Items'])

    print(len(ov_items))

    i = 0

    while True:

        ov_items[i].click()

        WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, EXPORT['Export'])))

        export = driver.find_element(By.XPATH, EXPORT['Export'])

        sleep(1)

        export.click()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, EXPORT['GS'])))

        gsheets = driver.find_element(By.XPATH, EXPORT['GS'])

        sleep(1)

        gsheets.click()

        current_handles = driver.window_handles

        google_sheets_save(driver, current_handles)

        driver.refresh()

        sleep(15)

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'galaxyIframe')))

        iframe = driver.find_element(By.ID, 'galaxyIframe')

        driver.switch_to.frame(iframe)

        sleep(3)

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DROPDOWN['Dropdown'])))

        ov_dropdown = driver.find_element(By.XPATH, DROPDOWN['Dropdown'])

        ov_dropdown.click()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DROPDOWN['Item container'])))

        item_cont = driver.find_element(By.XPATH, DROPDOWN['Item container'])

        sleep(1)

        ov_items = item_cont.find_elements(By.TAG_NAME, DROPDOWN['Items'])

        i += 1

        if i == 8:

            break
        else:
            pass

def get_active_users():

    driver = start_driver()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Audience'])))

    audience = driver.find_element(By.XPATH, MAIN_BTNS['Audience']) 

    audience.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, METRICS['Active users'])))

    act_users = driver.find_element(By.XPATH, METRICS['Active users'])

    act_users.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'galaxyIframe')))

    iframe = driver.find_element(By.ID, 'galaxyIframe')

    driver.switch_to.frame(iframe)

    st_year = 2009
    end_year = 2011

    while True:

        sleep(3)

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Date selector'])))

        date_selector = driver.find_element(By.XPATH, DATE_RANGE['Date selector'])

        date_selector.click()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input start'])))

        start = driver.find_element(By.XPATH, DATE_RANGE['Input start'])

        sleep(1)

        start.clear()

        sleep(1)

        start.send_keys(f'Jan 4, {st_year}')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input end'])))

        end = driver.find_element(By.XPATH, DATE_RANGE['Input end'])

        sleep(1)

        end.clear()

        sleep(1)

        end.send_keys(f'Jan 3, {end_year}')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Apply'])))

        sleep(1)

        apply = driver.find_element(By.XPATH, DATE_RANGE['Apply'])

        apply.click()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, EXPORT['Export'])))

        sleep(2)

        WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

        export = driver.find_element(By.XPATH, EXPORT['Export'])

        sleep(1)

        driver.execute_script("arguments[0].click();", export)

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, EXPORT['GS'])))

        sleep(1)

        gsheets = driver.find_element(By.XPATH, EXPORT['GS'])

        sleep(1)

        driver.execute_script("arguments[0].click();", gsheets)

        current_handles = driver.window_handles

        google_sheets_save(driver, current_handles)

        driver.refresh()

        sleep(2.5)

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'galaxyIframe')))

        iframe = driver.find_element(By.ID, 'galaxyIframe')

        driver.switch_to.frame(iframe)

        sleep(10)

        st_year += 1
        end_year += 1

        if end_year > 2023:

            driver.quit()

            break

def get_event_metrics():

    driver = start_driver()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Behavior'])))

    behavior = driver.find_element(By.XPATH, MAIN_BTNS['Behavior']) 

    behavior.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Events'])))

    events = driver.find_element(By.XPATH, MAIN_BTNS['Events'])

    events.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Overview'])))

    overview = driver.find_elements(By.XPATH, MAIN_BTNS['Overview'])[1]

    overview.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'galaxyIframe')))

    iframe = driver.find_element(By.ID, 'galaxyIframe')

    driver.switch_to.frame(iframe)

    sleep(3)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, DROPDOWN['Dropdown'])))

    ov_dropdown = driver.find_element(By.CLASS_NAME, DROPDOWN['Dropdown'])

    ov_dropdown.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, DROPDOWN['Items'])))

    ov_items = driver.find_elements(By.CLASS_NAME, DROPDOWN['Items'])

    i = 2

    while True:

        ov_items[i].click()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, DATE_RANGE['Date selector'])))

        date_selector = driver.find_element(By.CLASS_NAME, DATE_RANGE['Date selector'])

        date_selector.click()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, DATE_RANGE['Input start'])))

        start = driver.find_element(By.CLASS_NAME, DATE_RANGE['Input start'])

        sleep(1)

        start.clear

        sleep(1)

        start.send_keys(f'Jan 4, 2009')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, DATE_RANGE['Input end'])))

        end = driver.find_element(By.CLASS_NAME, DATE_RANGE['Input end'])

        sleep(1)

        end.clear()

        sleep(1)

        end.send_keys(f'Jan 3, 2023')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, DATE_RANGE['Apply'])))

        apply = driver.find_element(By.CLASS_NAME, DATE_RANGE['Apply'])

        apply.click()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, EXPORT['Export'])))

        export = driver.find_element(By.CLASS_NAME, EXPORT['Export'])

        export.click()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, EXPORT['GS'])))

        gsheets = driver.find_element(By.CLASS_NAME, EXPORT['GS'])

        gsheets.click()

        current_handles = driver.window_handles

        google_sheets_save(driver, current_handles)

        driver.refresh()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'galaxyIframe')))

        iframe = driver.find_element(By.ID, 'galaxyIframe')

        driver.switch_to.frame(iframe)

        sleep(3)

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, DROPDOWN['Dropdown'])))

        ov_dropdown = driver.find_element(By.CLASS_NAME, DROPDOWN['Dropdown'])

        ov_dropdown.click()
        
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, DROPDOWN['Items'])))

        ov_items = driver.find_elements(By.CLASS_NAME, DROPDOWN['Items'])

        i += 1

        if i == 5:

            break
        else:
            pass

def get_audience_behavior_data():

    driver = start_driver()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Audience'])))

    audience = driver.find_element(By.XPATH, MAIN_BTNS['Audience']) 

    audience.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Behavior'])))

    behavior = driver.find_element(By.XPATH, MAIN_BTNS['Behavior']) 

    behavior.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, METRICS['New vs Returning'])))

    newvsret = driver.find_element(By.XPATH, METRICS['New vs Returning'])

    newvsret.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, OTHERS['Select metric'])))

    select = driver.find_element(By.XPATH, OTHERS['Select metric'])

    select.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, OTHERS['Metric input'])))

    metric_input = driver.find_element(By.CLASS_NAME, OTHERS['Metric input'])

    sleep(1)

    metric_input.clear()

    sleep(1)

    metric_input.send_keys('New users')

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, METRICS['New Users'])))

    new_users = driver.find_element(By.XPATH, METRICS['New Users'])

    new_users.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, DATE_RANGE['Date selector'])))

    date_selector = driver.find_element(By.CLASS_NAME, DATE_RANGE['Date selector'])

    date_selector.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, DATE_RANGE['Input start'])))

    start = driver.find_element(By.CLASS_NAME, DATE_RANGE['Input start'])

    sleep(1)

    start.clear()

    sleep(1)

    start.send_keys(f'Jan 4, 2009')

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, DATE_RANGE['Input end'])))

    end = driver.find_element(By.CLASS_NAME, DATE_RANGE['Input end'])

    sleep(1)

    end.clear()

    sleep(1)

    end.send_keys(f'Jan 3, 2023')

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, DATE_RANGE['Apply'])))

    apply = driver.find_element(By.CLASS_NAME, DATE_RANGE['Apply'])

    apply.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, EXPORT['Export'])))

    export = driver.find_element(By.CLASS_NAME, EXPORT['Export'])

    export.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, EXPORT['GS'])))

    gsheets = driver.find_element(By.CLASS_NAME, EXPORT['GS'])

    gsheets.click()

    current_handles = driver.window_handles

    google_sheets_save(driver, current_handles)


def get_new_vs_returning():

    # Set up dates from 01/04/2009 to 01/03/2023:

    start_date = date(day=16,month=2,year=2022)
    end_date = date(day=3,month=1,year=2023)

    # Automation:

    driver = start_driver()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Audience'])))

    audience = driver.find_element(By.XPATH, MAIN_BTNS['Audience']) 

    audience.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Behavior'])))

    behavior = driver.find_element(By.XPATH, MAIN_BTNS['Behavior']) 

    behavior.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, METRICS['New vs Returning'])))

    nvr = driver.find_element(By.XPATH, METRICS['New vs Returning'])

    nvr.click()

    sleep(3)

    WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'galaxyIframe')))

    iframe = driver.find_element(By.ID, 'galaxyIframe')

    driver.switch_to.frame(iframe)

    sleep(1.5)

    while True:

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Date selector'])))

        try:

            date_selector = driver.find_element(By.XPATH, DATE_RANGE['Date selector'])

            date_selector.click()
        
        except Exception as e:

            print(e)

            WebDriverWait(driver, 70).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[74]/div[2]/div/div[4]/div')))

            sleep(1)

            warning = driver.find_element(By.XPATH, '/html/body/div[74]/div[2]/div/div[4]/div')

            sleep(0.5)

            warning.click()

            sleep(1.5)

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Date selector'])))

            date_selector = driver.find_element(By.XPATH, DATE_RANGE['Date selector'])

            date_selector.click()


        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input end'])))

        end = driver.find_element(By.XPATH, DATE_RANGE['Input end'])

        sleep(0.5)

        end.clear()

        sleep(0.5)

        end.send_keys(f'{start_date.month}/{start_date.day}/{start_date.year}')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input start'])))

        start = driver.find_element(By.XPATH, DATE_RANGE['Input start'])

        sleep(0.5)

        start.clear()

        sleep(0.5)

        start.send_keys(f'{start_date.month}/{start_date.day}/{start_date.year}')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Apply'])))

        apply = driver.find_element(By.XPATH, DATE_RANGE['Apply'])

        sleep(1)

        apply.click()

        sleep(1)

        WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

        sleep(1.5)

        #SAVE INTO json:

        is_returning_users_first = 'Returning' in driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[1]/td[3]').text


        if is_returning_users_first:

            new_users = re.sub(r'\(.*\)', '' , driver.find_element(By.XPATH, METRICS['Returning users']).text)
            
            ret_users = re.sub(r'\(.*\)', '' , driver.find_element(By.XPATH, METRICS['New users']).text)

        else:

            new_users = re.sub(r'\(.*\)', '' , driver.find_element(By.XPATH, METRICS['New users']).text)

            ret_users = re.sub(r'\(.*\)', '' , driver.find_element(By.XPATH, METRICS['Returning users']).text)

        

        new_vs_ret = {
            'Date' : start_date.strftime('%m/%d/%Y'),
            'NewVisitors' : new_users,
            'ReturningVisitors': ret_users
        }   

        print('*' * 100 + f""" 
        
        {new_vs_ret}
        
        """ 
        + '*' * 100)

        with open('nvr_martes_financiero.json') as f:

            data = json.load(f)

        with open('nvr_martes_financiero.json', 'w') as f:

            data['Data'].append(new_vs_ret)
            
            json.dump(data, f, indent=4)

        start_date += timedelta(days=1)

        if start_date > end_date:

            break

        else:

            pass


def google_sheets_save(driver : webdriver.Chrome, current_handles):

    sleep(10)

    driver.switch_to.window(driver.window_handles[1])

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'confirmActionButton')))

    save_btn = driver.find_element(By.ID, 'confirmActionButton')

    save_btn.click()

    sleep(5)

    driver.close()

    driver.switch_to.window(driver.window_handles[0])


def get_login_entitlement():

    # Set up dates from 01/04/2009 to 01/03/2023:

    start_date = date(day=15,month=3,year=2022)
    end_date = date(day=3,month=1,year=2023)

    # Automation:

    driver = start_driver()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Behavior'])))

    behavior = driver.find_element(By.XPATH, MAIN_BTNS['Behavior']) 

    behavior.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Events'])))

    events = driver.find_element(By.XPATH, MAIN_BTNS['Events'])

    events.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Top events'])))

    top = driver.find_element(By.XPATH, MAIN_BTNS['Top events'])

    top.click()

    sleep(5)

    WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'galaxyIframe')))

    iframe = driver.find_element(By.ID, 'galaxyIframe')

    driver.switch_to.frame(iframe)

    sleep(3)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, OTHERS['Login entitlement'])))

    log_en = driver.find_element(By.XPATH, OTHERS['Login entitlement'])

    log_en.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Evolok'])))

    evolok = driver.find_element(By.XPATH, MAIN_BTNS['Evolok'])

    evolok.click()

    sleep(5)

    while True:

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Date selector'])))

        date_selector = driver.find_element(By.XPATH, DATE_RANGE['Date selector'])

        date_selector.click()
        
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input end'])))

        end = driver.find_element(By.XPATH, DATE_RANGE['Input end'])

        sleep(1)

        end.clear()

        sleep(1)

        end.send_keys(f'{start_date.month}/{start_date.day}/{start_date.year}')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input start'])))

        start = driver.find_element(By.XPATH, DATE_RANGE['Input start'])

        sleep(1)

        start.clear()

        sleep(1)

        start.send_keys(f'{start_date.month}/{start_date.day}/{start_date.year}')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Apply'])))

        apply = driver.find_element(By.XPATH, DATE_RANGE['Apply'])

        sleep(2)

        apply.click()

        sleep(3)

        WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

        sleep(3)

        #SAVE INTO json:

        try:

            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[1]/td[4]/div')))

            log_ent = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[1]/td[4]/div')

            log_ent_val = re.sub(r'\(.*\)', '', log_ent.text)

        except TimeoutException:

            log_ent_val = '0'

        try:

            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[2]/td[4]/div')))

            log = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[2]/td[4]/div')

            log_val = re.sub(r'\(.*\)', '', log.text)

        except TimeoutException:

            log_val = '0'

        try:

            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[3]/td[4]/div')))

            ent = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[3]/td[4]/div')

            ent_val = re.sub(r'\(.*\)', '', ent.text)

        except TimeoutException:

            ent_val = '0'
        

        LOGIN_ENTITLEMENT = {
            'Login' : log_ent_val,
            'Entitlement' : log_val,
            'Login/Entitlement' : ent_val,
        }   

        with open('login_entitlement_data.json') as f:

            data = json.load(f)

        with open('login_entitlement_data.json', 'w') as f:

            data['Data'].append(LOGIN_ENTITLEMENT)
            
            json.dump(data, f, indent=4)


        start_date += timedelta(days=1)

        if start_date > end_date:

            break

        else:

            pass

def get_traffic_by_page(driver: webdriver.Chrome):

    # Set up dates from 01/04/2009 to 01/03/2023:

    start_date = date(day=15,month=3,year=2022)
    end_date = date(day=3,month=1,year=2023)

    # Automation:

    driver = start_driver()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Behavior'])))

    behavior = driver.find_element(By.XPATH, MAIN_BTNS['Behavior']) 

    behavior.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Events'])))

    events = driver.find_element(By.XPATH, MAIN_BTNS['Events'])

    events.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Top events'])))

    top = driver.find_element(By.XPATH, MAIN_BTNS['Top events'])

    top.click()

    sleep(5)

    WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'galaxyIframe')))

    iframe = driver.find_element(By.ID, 'galaxyIframe')

    driver.switch_to.frame(iframe)

    sleep(3)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, OTHERS['Login entitlement'])))

    log_en = driver.find_element(By.XPATH, OTHERS['Login entitlement'])

    log_en.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Evolok'])))

    evolok = driver.find_element(By.XPATH, MAIN_BTNS['Evolok'])

    evolok.click()

    sleep(5)

    while True:

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Date selector'])))

        date_selector = driver.find_element(By.XPATH, DATE_RANGE['Date selector'])

        date_selector.click()
        
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input end'])))

        end = driver.find_element(By.XPATH, DATE_RANGE['Input end'])

        sleep(1)

        end.clear()

        sleep(1)

        end.send_keys(f'{start_date.month}/{start_date.day}/{start_date.year}')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input start'])))

        start = driver.find_element(By.XPATH, DATE_RANGE['Input start'])

        sleep(1)

        start.clear()

        sleep(1)

        start.send_keys(f'{start_date.month}/{start_date.day}/{start_date.year}')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Apply'])))

        apply = driver.find_element(By.XPATH, DATE_RANGE['Apply'])

        sleep(2)

        apply.click()

        sleep(3)

        WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

        sleep(3)

        #SAVE INTO json:

        try:

            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[1]/td[4]/div')))

            log_ent = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[1]/td[4]/div')

            log_ent_val = re.sub(r'\(.*\)', '', log_ent.text)

        except TimeoutException:

            log_ent_val = '0'

        try:

            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[2]/td[4]/div')))

            log = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[2]/td[4]/div')

            log_val = re.sub(r'\(.*\)', '', log.text)

        except TimeoutException:

            log_val = '0'

        try:

            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[3]/td[4]/div')))

            ent = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[3]/td[4]/div')

            ent_val = re.sub(r'\(.*\)', '', ent.text)

        except TimeoutException:

            ent_val = '0'
        

        LOGIN_ENTITLEMENT = {
            'Login' : log_ent_val,
            'Entitlement' : log_val,
            'Login/Entitlement' : ent_val,
        }   

        with open('login_entitlement_data.json') as f:

            data = json.load(f)

        with open('login_entitlement_data.json', 'w') as f:

            data['Data'].append(LOGIN_ENTITLEMENT)
            
            json.dump(data, f, indent=4)


        start_date += timedelta(days=1)

        if start_date > end_date:

            break

        else:

            pass

def get_traffic_by_device(
                          start_date : date,
                          end_date : date,
                          url : str,
                          op_file : str
                          ):
    # Automation:
    
    driver = start_driver(url)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Audience'])))

    audience = driver.find_element(By.XPATH, MAIN_BTNS['Audience']) 

    audience.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Mobile'])))

    mobile = driver.find_element(By.XPATH, MAIN_BTNS['Mobile'])

    mobile.click()

    sleep(3)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Overview'])))

    top = driver.find_elements(By.XPATH, MAIN_BTNS['Overview'])[1]

    top.click()

    sleep(4)

    WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'galaxyIframe')))

    iframe = driver.find_element(By.ID, 'galaxyIframe')

    driver.switch_to.frame(iframe)

    sleep(3)

    while True:

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Date selector'])))

        date_selector = driver.find_element(By.XPATH, DATE_RANGE['Date selector'])

        date_selector.click()
        
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input end'])))

        end = driver.find_element(By.XPATH, DATE_RANGE['Input end'])

        sleep(1)

        end.clear()

        sleep(1)

        end.send_keys(f'{start_date.month}/{start_date.day}/{start_date.year}')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input start'])))

        start = driver.find_element(By.XPATH, DATE_RANGE['Input start'])

        sleep(1)

        start.clear()

        sleep(1)

        start.send_keys(f'{start_date.month}/{start_date.day}/{start_date.year}')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Apply'])))

        apply = driver.find_element(By.XPATH, DATE_RANGE['Apply'])

        sleep(2)

        apply.click()

        sleep(3)

        WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

        sleep(3)

        #SAVE INTO json:

        order = {
                'desktop':1,
                'mobile':2,
                'tablet':3
                } 

        keys = order.keys()

        try:

            for i in range(1,4):

                path = f'/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[{i}]/td[3]/span'

                try:

                    device = driver.find_element(By.XPATH, path).text

                    order[device] = i

                except Exception:

                    order[keys[i-1]] = None
        
        except Exception:

            pass

        try:

            desktop = driver.find_element(By.XPATH, f'/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[{order["desktop"]}]/td[4]/div')
                                                        
            desktop = re.sub(r'\(.*\)', '', desktop.text)

        except :

            desktop = '0'

        try:

            mobile = driver.find_element(By.XPATH, f'/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[{order["mobile"]}]/td[4]/div')

            mobile = re.sub(r'\(.*\)', '', mobile.text)

        except Exception:

            mobile = '0'

        try:

            tablet = driver.find_element(By.XPATH, f'/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[{order["tablet"]}]/td[4]/div')

            tablet = re.sub(r'\(.*\)', '', tablet.text)

        except Exception:

            tablet = '0'
        

        tbd = {
            'Date' : start_date.strftime('%m/%d/%Y'),
            'Desktop' : desktop,
            'Mobile' : mobile,
            'Tablet' : tablet,
        }   

        with open(f'{op_file}') as f:

            data = json.load(f)

        with open(f'{op_file}', 'w') as f:

            data['Data'].append(tbd)
            
            json.dump(data, f, indent=4)


        start_date += timedelta(days=1)

        if start_date > end_date:

            break

        else:

            pass

def get_traffic_by_channel(
                          start_date : date,
                          end_date : date,
                          url : str,
                          op_file : str
                          ):
    # Automation:
    
    driver = start_driver(url)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Acquisition'])))

    audience = driver.find_element(By.XPATH, MAIN_BTNS['Acquisition']) 

    audience.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['All Traffic'])))

    mobile = driver.find_element(By.XPATH, MAIN_BTNS['All Traffic'])

    mobile.click()

    sleep(3)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, MAIN_BTNS['Channels'])))

    top = driver.find_element(By.XPATH, MAIN_BTNS['Channels'])

    top.click()

    sleep(4)

    WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'galaxyIframe')))

    iframe = driver.find_element(By.ID, 'galaxyIframe')

    driver.switch_to.frame(iframe)

    sleep(3)

    while True:

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Date selector'])))

        date_selector = driver.find_element(By.XPATH, DATE_RANGE['Date selector'])

        date_selector.click()
        
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input end'])))

        end = driver.find_element(By.XPATH, DATE_RANGE['Input end'])

        sleep(1)

        end.clear()

        sleep(1)

        end.send_keys(f'{start_date.month}/{start_date.day}/{start_date.year}')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input start'])))

        start = driver.find_element(By.XPATH, DATE_RANGE['Input start'])

        sleep(1)

        start.clear()

        sleep(1)

        start.send_keys(f'{start_date.month}/{start_date.day}/{start_date.year}')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Apply'])))

        apply = driver.find_element(By.XPATH, DATE_RANGE['Apply'])

        sleep(2)

        apply.click()

        sleep(3)

        WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

        sleep(3)

        #SAVE INTO json:

        order = {
                'Direct': None,
                'Organic Search': None,
                '(Other)': None,
                'Email' : None,
                'Social': None,
                'Referral': None,
                'Paid Search': None,
                'Display': None,
                } 

        keys = order.keys()

        try:

            for i in range(1,9):

                path = f'/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[{i}]/td[3]/span'
                        
                try:

                    device = driver.find_element(By.XPATH, path).text

                    order[device] = i

                except Exception:

                    order[keys[i-1]] = None
        
        except Exception:

            pass

        try:

            direct = driver.find_element(By.XPATH, f'/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[{order["Direct"]}]/td[4]/div')
                                                        
            direct = re.sub(r'\(.*\)', '', direct.text)

        except :

            direct = '0'

        try:

            org = driver.find_element(By.XPATH, f'/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[{order["Organic Search"]}]/td[4]/div')

            org = re.sub(r'\(.*\)', '', org.text)

        except Exception:

            org = '0'

        try:

            other = driver.find_element(By.XPATH, f'/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[{order["(Other)"]}]/td[4]/div')

            other = re.sub(r'\(.*\)', '', other.text)

        except Exception:

            other = '0'

        try:

            email = driver.find_element(By.XPATH, f'/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[{order["Email"]}]/td[4]/div')
                                                        
            email = re.sub(r'\(.*\)', '', email.text)

        except :

            email = '0'

        try:

            social = driver.find_element(By.XPATH, f'/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[{order["Social"]}]/td[4]/div')

            social = re.sub(r'\(.*\)', '', social.text)

        except Exception:

            social = '0'

        try:

            ref = driver.find_element(By.XPATH, f'/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[{order["Referral"]}]/td[4]/div')

            ref = re.sub(r'\(.*\)', '', ref.text)

        except Exception:

            ref = '0'

        try:

            ps = driver.find_element(By.XPATH, f'/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[{order["Paid Search"]}]/td[4]/div')
                                                        
            ps = re.sub(r'\(.*\)', '', ps.text)

        except :

            ps = '0'

        try:

            display = driver.find_element(By.XPATH, f'/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody/tr[{order["Display"]}]/td[4]/div')

            display = re.sub(r'\(.*\)', '', display.text)

        except Exception:

            display = '0'
        

        tbc = {
            'Date' : start_date.strftime('%m/%d/%Y'),
            'Direct': direct,
            'Organic Search': org,
            '(Other)': other,
            'Email' : email,
            'Social': social,
            'Referral': ref,
            'Paid Search': ps,
            'Display': display
            }  

        print(f'\n Retrieving: {tbc}')

        with open(f'{op_file}') as f:

            data = json.load(f)

        with open(f'{op_file}', 'w') as f:

            data['Data'].append(tbc)
            
            json.dump(data, f, indent=4)


        start_date += timedelta(days=1)

        if start_date > end_date:

            break

        else:

            pass

def get_historics_nv( start_date : date,
                    end_date : date,
                    url : str,
                    op_file : str, 
                    *btns):

    driver = start_driver(url)

    for btn in btns:

        if btn == 'Overview':

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, f"//span[text()='{btn}']")))

            button = driver.find_elements(By.XPATH, f"//span[text()='{btn}']")[1] 

            button.click()

        else:

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, f"//span[text()='{btn}']")))

            button = driver.find_element(By.XPATH, f"//span[text()='{btn}']") 

            button.click()

        sleep(1.5)

    WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'galaxyIframe')))

    iframe = driver.find_element(By.ID, 'galaxyIframe')

    driver.switch_to.frame(iframe)

    sleep(1.5)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Date selector'])))

    date_selector = driver.find_element(By.XPATH, DATE_RANGE['Date selector'])

    date_selector.click()
    
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input end'])))

    end = driver.find_element(By.XPATH, DATE_RANGE['Input end'])

    sleep(0.5)

    end.clear()

    sleep(0.5)

    end.send_keys(f'{end_date.month}/{end_date.day}/{end_date.year}')

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input start'])))

    start = driver.find_element(By.XPATH, DATE_RANGE['Input start'])

    sleep(0.5)

    start.clear()

    sleep(0.5)

    start.send_keys(f'{start_date.month}/{start_date.day}/{start_date.year}')

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Apply'])))

    apply = driver.find_element(By.XPATH, DATE_RANGE['Apply'])

    sleep(1)

    apply.click()

    sleep(1.5)

    WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

    sleep(1.5)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE['DD btn'])))

    dropdown = driver.find_element(By.XPATH, DATE['DD btn']) 

    dropdown.click()

    sleep(1.5)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE['DD input'])))

    input = driver.find_element(By.XPATH, DATE['DD input']) 

    input.clear()

    input.send_keys('Date')

    sleep(1)

    input.send_keys(Keys.ENTER)

    sleep(1.5)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE['Date'])))

    date = driver.find_element(By.XPATH, DATE['Date']) 

    date.click()

    sleep(1.5)

    WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

    date = driver.find_element(By.XPATH, DATE['Date']) 

    date.click()

    sleep(1.5)

    WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE['Select'])))

    select = driver.find_element(By.XPATH, DATE['Select']) 

    select.click()

    sleep(1.5)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE['Max rows'])))

    max_rows = driver.find_element(By.XPATH, DATE['Max rows']) 

    max_rows.click()

    sleep(1)

    data = []

    data_row = {}

    date = None

    file_number = 1

    while True:

        WebDriverWait(driver, 70).until(EC.invisibility_of_element_located((By.XPATH, OTHERS['Loading'])))

        table = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div[4]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/div[2]/div/table/tbody')

        rows = table.find_elements(By.TAG_NAME, 'tr')

        row_count = 0

        while True:

            try:

                elements = rows[row_count].find_elements(By.TAG_NAME, 'td')

                if date == elements[3].text:

                    date = elements[3].text
                    
                    data_row[elements[2].text] = re.sub(r'\(.*\)', '' , elements[4].text)

                else:

                    if len(data_row) > 0:

                        data.append(data_row)
                        print(f'Retrieved record: {data_row}')

                    data_row = {}

                    date = elements[3].text

                    f_date = f'{date[4:6]}/{date[6:8]}/{date[0:4]}'

                    data_row['Date'] = f_date

                    data_row[elements[2].text] = re.sub(r'\(.*\)', '' , elements[4].text)

                row_count += 1

            except IndexError:

                with open(f'{op_file}({file_number}).json') as f:

                    all_data = json.load(f)

                with open(f'{op_file}({file_number}).json', 'w') as f:

                    all_data['Data'].append(data)
                    
                    json.dump(data, f, indent=4)

                # df = pd.DataFrame(data, columns=list(max(data, key=len).keys()))

                # df.fillna('0', inplace=True)

                # values = df.values.tolist()

                # for value in values:

                #     sheet.append(value)

                # wb.save(op_file)

                # data = []

                # del df

                # del values                 

                # break

            if row_count > 4999:

                with open(f'{op_file}({file_number}).json') as f:

                    all_data = json.load(f)

                with open(f'{op_file}({file_number}).json', 'w') as f:

                    all_data['Data'].append(data)
                    
                    json.dump(data, f, indent=4)

                file_number += 1

                # df = pd.DataFrame(data, columns=list(max(data, key=len).keys()))

                # df.fillna('0', inplace=True)

                # values = df.values.tolist()

                # for value in values:

                #     sheet.append(value)

                # wb.save(op_file)

                data = []

                # del df

                # del values                 

                break

        #At the end...

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE['Pagination'])))

        next = driver.find_element(By.XPATH, DATE['Pagination']) 

        next.click()
        
        sleep(1)

def save_historic_to_excel(op_file):

    file_num = 1

    all_columns = []

    all_data = []

    while True:

        try:

            with open(f'{op_file}({file_num}).json') as f:

                data = json.load(f)['Data']

                for element in data:

                    all_data.append(element)

                    keys = element.keys()

                    for key in keys:

                        if key not in all_columns:

                            all_columns.append(key)

                        else:

                            continue

                file_num += 1

        except Exception:

            break

    df = pd.DataFrame(data, index =pd.RangeIndex(0,len(data)),columns = all_columns)
    
    df.to_excel(f"{op_file}.xlsx", index=False, columns= all_columns)

    print('Saved!')



def google_sheets_save(driver : webdriver.Chrome, current_handles):

    sleep(10)

    driver.switch_to.window(driver.window_handles[1])

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'confirmActionButton')))

    save_btn = driver.find_element(By.ID, 'confirmActionButton')

    save_btn.click()

    sleep(5)

    driver.close()

    driver.switch_to.window(driver.window_handles[0])

def save_to_excel(op_file, columns):

    with open(f'{op_file}') as f:

        data = json.load(f)['Data']

        print('Saving to excel...')

        df = pd.DataFrame(data, index =pd.RangeIndex(0,len(data)),columns = columns)
        
        df.to_excel(f"{op_file.replace('.json', '')}.xlsx", index=False, columns= columns)

        print('Saved!')

def clean_and_fix():

    new_data = []

    start_date = date(day=5,month=1,year=2009)
    end_date = date(day=4,month=1,year=2023)

    with open('new_vs_ret_data.json') as f:

        data = json.load(f)['Data']

        for row in data:

            row['Date'] = start_date.strftime('%m/%d/%Y')

            new_data.append(row)
        
            start_date += timedelta(days=1)

            if start_date > end_date:

                break

            else:

                pass
        
    with open('new_vs_ret_with_dates.json') as f:

            data = json.load(f)

    with open('new_vs_ret_with_dates.json', 'w') as f:

        for row in new_data:

            data['Data'].append(row)
        
        json.dump(data, f, indent=4)


def find_outlayers():

    with open('new_vs_ret_data.json') as f:

        data = json.load(f)['Data']

        last_row = 0

        for row in data:

            if int(row['ReturningVisitors'].replace(',','')) > last_row * 2 and last_row > 1000:

                print(f'Outlayer: {row["Date"]}')

            else:

                last_row = int(row['ReturningVisitors'].replace(',',''))

    with open('new_vs_ret_data.json') as f:

        data = json.load(f)['Data']

        last_date = ''

        for row in data:

            if row['Date'] == last_date:

                print(last_date)

            else:

                last_date = row['Date']        

start_date = date(day=1,month=1,year=2009)
end_date = date(day=24,month=4,year=2023)
url = 'https://analytics.google.com/analytics/web/?authuser=2#/report-home/a6897643w13263073p13942924'
tvd_file = 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/Prensa_historico/channel_traffic_prensa'


if __name__ == '__main__':

    get_historics_nv(start_date, end_date, url, tvd_file, 'Acquisition', 'All Traffic', 'Channels')