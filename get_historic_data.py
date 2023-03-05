from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from selenium.webdriver.common.by import By

MAIN_BTNS = {
    'Audience' : "//span[text()='Audience']",
    'Overview' : "//span[text()='Overview']",
    'Behavior' : "//span[text()='Behavior']",
    'Events' : "//span[text()='Events']",
    'Evolok' : "//span[text()='EVOLOK']",
}

METRICS = {
    'Active users' : "//span[text()='Active Users']",
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
    'Dropdown' : 'ID-buttonText._GAb-_GAci-_GAhi._GAPe',
    'Items' : '_GAMQ'
}

DATE_RANGE = {
    'Date selector' : 'ID-container._GAPd',
    'Input start' : 'ID-datecontrol-primary-start._GATg ACTION-daterange_input.TARGET-primary_start._GAfl',
    'Input end' : 'ID-datecontrol-primary-end _GATg.ACTION-daterange_input.TARGET-primary_end',
    'Apply' : 'ID-apply.ACTION-apply.TARGET-.C_DATECONTROL_APPLY',
    
}

EXPORT = {
    'Export' : 'ID-exportControlButton.C_REPORTTOOLBAR_BTN_SIMPLE.ACTION-exportMenu.TARGET-',
    'GS' : 'ACTION-export.TARGET-FOR_TRIX'
}

OTHERS = {
    'Login_entitlement' : "//span[text()='Login/Entitlement']"
}
 

def start_driver():

    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get("https://analytics.google.com/analytics/web/?authuser=2#/a6897643w13263073p13942924/")

    return driver

def get_audience_overview_metrics():

    driver = start_driver()

    audience = driver.find_element(By.XPATH, MAIN_BTNS['Audience']) 

    audience.click()

    overview = driver.find_element(By.XPATH, MAIN_BTNS['Overview']) 

    overview.click()

    ov_dropdown = driver.find_element(By.CLASS_NAME, DROPDOWN['Dropdown'])

    ov_dropdown.click()

    ov_items = driver.find_elements(By.CLASS_NAME, DROPDOWN['Items'])

    i = 0

    while True:

        ov_items[i].click()

        date_selector = driver.find_element(By.CLASS_NAME, DATE_RANGE['Date selector'])

        date_selector.click()

        st_year = 2009
        end_year = 2011

        while True:

            start = driver.find_element(By.CLASS_NAME, DATE_RANGE['Input start'])

            start.send_keys(f'Jan 4, {st_year}')

            end = driver.find_element(By.CLASS_NAME, DATE_RANGE['Input end'])

            end.send_keys(f'Jan 3, {end_year}')

            apply = driver.find_element(By.CLASS_NAME, DATE_RANGE['Apply'])

            apply.click()

            export = driver.find_element(By.CLASS_NAME, EXPORT['Export'])

            export.click()

            gsheets = driver.find_element(By.CLASS_NAME, EXPORT['GS'])

            gsheets.click()

            st_year += 2
            end_year += 2

            if end_year > 2023:

                break

        driver.refresh()

        ov_dropdown = driver.find_element(By.CLASS_NAME, DROPDOWN['Dropdown'])

        ov_dropdown.click()

        ov_items = driver.find_elements(By.CLASS_NAME, DROPDOWN['Items'])

        i += 1

        if i == 8:

            break
        else:
            pass

def get_active_users():

    pass

def get_event_metrics():

    pass

def get_audience_behavior_data():

    pass