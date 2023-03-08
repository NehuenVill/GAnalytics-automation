from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from selenium.webdriver.common.by import By

MAIN_BTNS = {
    'Audience' : "//span[text()='Audience']",
    'Overview' : "//span[text()='Overview']",
    'Behavior' : "//span[text()='Behavior']",
    'Events' : "//span[text()='Events']",
    'Top events' : "//span[text()='Top Events']",
    'Evolok' : "//span[text()='EVOLOK']"
}

METRICS = {
    'Active users' : "//span[text()='Active Users']",
    'New vs Returning' : "//span[text()='New vs Returning']",
    'New users' : "//div[text()='New Users']"
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
    'Login entitlement' : "//span[text()='Login/Entitlement']",
    'Select metric' : "//div[text()='Select a metric']",
    'Metric input' : "ID-searchBox"
}
 

def start_driver():

    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get("https://analytics.google.com/analytics/web/?authuser=2#/a6897643w13263073p13942924/")

    sleep(150)

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

            google_sheets_save(driver)

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

    driver = start_driver()

    audience = driver.find_element(By.XPATH, MAIN_BTNS['Audience']) 

    audience.click()

    act_users = driver.find_element(By.XPATH, METRICS['Active users'])

    act_users.click()

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

        google_sheets_save(driver)

        st_year += 2
        end_year += 2

        if end_year > 2023:

            break

def get_event_metrics():

    driver = start_driver()

    behavior = driver.find_element(By.XPATH, MAIN_BTNS['Behavior']) 

    behavior.click()

    events = driver.find_element(By.XPATH, MAIN_BTNS['Events'])

    events.click()

    overview = driver.find_elements(By.XPATH, MAIN_BTNS['Overview'])[1]

    overview.click()

    ov_dropdown = driver.find_element(By.CLASS_NAME, DROPDOWN['Dropdown'])

    ov_dropdown.click()

    ov_items = driver.find_elements(By.CLASS_NAME, DROPDOWN['Items'])

    i = 2

    while True:

        ov_items[i].click()

        date_selector = driver.find_element(By.CLASS_NAME, DATE_RANGE['Date selector'])

        date_selector.click()

        start = driver.find_element(By.CLASS_NAME, DATE_RANGE['Input start'])

        start.send_keys(f'Jan 4, 2009')

        end = driver.find_element(By.CLASS_NAME, DATE_RANGE['Input end'])

        end.send_keys(f'Jan 3, 2023')

        apply = driver.find_element(By.CLASS_NAME, DATE_RANGE['Apply'])

        apply.click()

        export = driver.find_element(By.CLASS_NAME, EXPORT['Export'])

        export.click()

        gsheets = driver.find_element(By.CLASS_NAME, EXPORT['GS'])

        gsheets.click()

        google_sheets_save(driver)

        driver.refresh()

        ov_dropdown = driver.find_element(By.CLASS_NAME, DROPDOWN['Dropdown'])

        ov_dropdown.click()

        ov_items = driver.find_elements(By.CLASS_NAME, DROPDOWN['Items'])

        i += 1

        if i == 5:

            break
        else:
            pass

def get_audience_behavior_data():

    driver = start_driver()

    audience = driver.find_element(By.XPATH, MAIN_BTNS['Audience']) 

    audience.click()

    behavior = driver.find_element(By.XPATH, MAIN_BTNS['Behavior']) 

    behavior.click()

    newvsret = driver.find_element(By.XPATH, METRICS['New vs Returning'])

    newvsret.click()

    select = driver.find_element(By.XPATH, OTHERS['Select metric'])

    select.click()

    metric_input = driver.find_element(By.CLASS_NAME, OTHERS['Metric input'])

    metric_input.send_keys('New users')

    new_users = driver.find_element(By.XPATH, METRICS['New Users'])

    new_users.click()

    start = driver.find_element(By.CLASS_NAME, DATE_RANGE['Input start'])

    start.send_keys(f'Jan 4, 2009')

    end = driver.find_element(By.CLASS_NAME, DATE_RANGE['Input end'])

    end.send_keys(f'Jan 3, 2023')

    apply = driver.find_element(By.CLASS_NAME, DATE_RANGE['Apply'])

    apply.click()

    export = driver.find_element(By.CLASS_NAME, EXPORT['Export'])

    export.click()

    gsheets = driver.find_element(By.CLASS_NAME, EXPORT['GS'])

    gsheets.click()

    google_sheets_save(driver)

def get_evolok_metrics():

    driver = start_driver()

    behavior = driver.find_element(By.XPATH, MAIN_BTNS['Behavior']) 

    behavior.click()

    events = driver.find_element(By.XPATH, MAIN_BTNS['Events'])

    events.click()

    top = driver.find_element(By.XPATH, MAIN_BTNS['Top Events'])

    top.click()

    log_en = driver.find_element(By.XPATH, OTHERS['Login entitlement'])

    log_en.click()

    date_selector = driver.find_element(By.CLASS_NAME, DATE_RANGE['Date selector'])

    date_selector.click()

    start = driver.find_element(By.CLASS_NAME, DATE_RANGE['Input start'])

    start.send_keys(f'Jan 4, 2009')

    end = driver.find_element(By.CLASS_NAME, DATE_RANGE['Input end'])

    end.send_keys(f'Jan 3, 2023')

    apply = driver.find_element(By.CLASS_NAME, DATE_RANGE['Apply'])

    apply.click()

    export = driver.find_element(By.CLASS_NAME, EXPORT['Export'])

    export.click()

    gsheets = driver.find_element(By.CLASS_NAME, EXPORT['GS'])

    gsheets.click()

    google_sheets_save(driver)


def google_sheets_save(driver : webdriver.Chrome):

    driver.switch_to.window(driver.window_handles[1])

    save_btn = driver.find_element(By.ID, 'confirmActionButton')

    save_btn.click()

    sleep(7.5)

    driver.close()

    driver.switch_to.window(driver.window_handles[0])


if __name__ == '__main__':

    start_driver()