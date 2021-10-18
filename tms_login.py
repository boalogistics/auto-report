from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# TODO rename module to tms actions and include all direct interactions here?

def login(url, boolheadless=True):
    # activate headless mode
    options = Options()
    options.headless = boolheadless

    # activate chrome driver
    browser = webdriver.Chrome(options=options)
    browser.get(url)

    # page elements to login
    boa_user = browser.find_element_by_id("txb-username")
    boa_pw = browser.find_element_by_id("txb-password")
    login_button = browser.find_element_by_id("ctl00_ContentBody_butLogin")

    # login credentials
    boa_user.send_keys("daigo@boalogistics.com")
    boa_pw.send_keys("ship12345")
    login_button.click()
    return browser
