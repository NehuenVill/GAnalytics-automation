from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep

MAIN_BTNS = {
    'Audience' : 'ga-nav-link.guided-help-visitors.has-icon.is-group.includes-active-state',
    'Overview' : 'ga-nav-link.guided-help-visitors-overview.includes-active-state.is-active-state'
}

METRICS = {
    'Active users' : 'ga-nav-link.guided-help-visitors-actives.includes-active-state.is-active-state',
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

driver = webdriver.Chrome(ChromeDriverManager().install())
driver.get("https://analytics.google.com/analytics/web/?authuser=2#/a6897643w13263073p13942924/")


