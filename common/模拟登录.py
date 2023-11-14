from utils import simulate_login
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions



browser  = simulate_login.create_chrome_driver(headless = True)
browser.get('https://ecig.cn')

browser.implicitly_wait(5)
# example
def get_cookies(username,password):
    enter_button = browser.find_element(By.CSS_SELECTOR,'#app > div > div > div.d-flex.login-content1 > div > div > div.login-top > a:nth-child(1) > span')
    enter_button.click()
    user = browser.find_element(By.CSS_SELECTOR,'#login > div > div > div.d-flex.login-content1 > div > div > div.d-flex.j-center > div > div > div:nth-child(4) > input')
    user.send_keys(username)
    pswd = browser.find_element(By.CSS_SELECTOR,'#login > div > div > div.d-flex.login-content1 > div > div > div.d-flex.j-center > div > div > div.el-input.el-input--suffix > input')
    pswd.send_keys(password)
    login_button = browser.find_element(By.CSS_SELECTOR,'#login > div > div > div.d-flex.login-content1 > div > div > div.d-flex.j-center > div > div > button')
    login_button.click()
    wait_obj = WebDriverWait(browser,5)
    wait_obj.until(expected_conditions.presence_of_element_located((By.CSS_SELECTOR,'#app > div > div.ec-header-box > div.ec-header-box-nav > div > div > div.ant-tabs-bar.ant-tabs-top-bar.ant-tabs-card-bar > div > div > div > div > div:nth-child(1) > div:nth-child(5) > span')))
    cookie_str = '; '.join([f"{cookie['name']}={cookie['value']}" for cookie in browser.get_cookies()])

    return cookie_str

cookies = get_cookies('','')
print(cookies)