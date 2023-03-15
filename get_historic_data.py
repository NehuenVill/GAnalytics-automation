from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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
    'Metric input' : "ID-searchBox"
}
 

def start_driver():

    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get("https://analytics.google.com/analytics/web/?authuser=2#/a6897643w13263073p13942924/")

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

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, DROPDOWN['Dropdown'])))

    ov_dropdown = driver.find_element(By.CLASS_NAME, DROPDOWN['Dropdown'])

    ov_dropdown.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, DROPDOWN['Items'])))

    ov_items = driver.find_elements(By.CLASS_NAME, DROPDOWN['Items'])

    i = 0

    while True:

        ov_items[i].click()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Date selector'])))

        date_selector = driver.find_element(By.XPATH, DATE_RANGE['Date selector'])

        date_selector.click()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input start'])))

        start = driver.find_element(By.CLASS_NAME, DATE_RANGE['Input start'])

        sleep(1)

        start.clear()

        sleep(1)

        start.send_keys(f'Jan 4, 2009')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Input end'])))

        end = driver.find_element(By.CLASS_NAME, DATE_RANGE['Input end'])

        sleep(1)

        end.clear()

        sleep(1)

        end.send_keys(f'Jan 3, 2023')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, DATE_RANGE['Apply'])))

        apply = driver.find_element(By.XPATH, DATE_RANGE['Apply'])

        apply.click()

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

    st_year = 2015
    end_year = 2017

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

        apply = driver.find_element(By.XPATH, DATE_RANGE['Apply'])

        apply.click()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, EXPORT['Export'])))

        sleep(10)

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

        st_year += 2
        end_year += 2

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

## later!!!

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



def google_sheets_save(driver : webdriver.Chrome, current_handles):

    sleep(10)

    driver.switch_to.window(driver.window_handles[1])

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'confirmActionButton')))

    save_btn = driver.find_element(By.ID, 'confirmActionButton')

    save_btn.click()

    sleep(5)

    driver.close()

    driver.switch_to.window(driver.window_handles[0])


if __name__ == '__main__':

    get_active_users()