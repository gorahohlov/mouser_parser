import pickle
import json

from fake_useragent import UserAgent
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.window import WindowTypes

# from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.chrome.service import Service


WEBDRIVER_SERVICE_DIR = r'c:/users/user/dev/mouser_parse_proj/wd_service_dir/'
EXTENSION_PATH = r'c:/users/user/dev/mouser_parse_proj/browsec_ext.crx'
URL = ('https://www.mouser.com/ProductDetail/ISSI/IS43DR16128B-25EBLI?qs='
       + 'abmNkq9no6AZ%252BuuB8Bpieg%3D%3D'
       )
# URL_1 = 'https://www.chipdip.ru/product/ecs-3953m-500-bn-tr-1'
FILE_NAME = r'c:/users/user/dev/mouser_parse_proj/mouser_cookies.json'


def save_cookies(file_name='cookies.pkl'):
    with open(file_name, 'wb') as file:
        pickle.dump(driver.get_cookies(), file)


def edit_cookies(file_name=None, imported_cookies=None):
    """preprocessing imported cookies to selenium webdriver add_cookies
    method format"""
    valid_cookies_keys = {'httpOnly', 'value', 'name', 'domain',
                          'secure', 'path', 'expiry', 'sameSite'}
    valid_sameSite_values = {'no_restriction': 'None', 'unspecified': 'Lax',
                             'strict': 'Strict'}
    if file_name:
        with open(file_name, 'r') as file:
            imported_cookies = json.load(file)
    elif not imported_cookies:
        return

    for i in range(len(imported_cookies)):
        cookie = imported_cookies.pop(i)
        if cookie.get('expirationDate'):
            cookie['expiry'] = int(cookie.pop('expirationDate'))
        cookie = {key: cookie[key] for key in valid_cookies_keys if key in cookie}
        if valid_sameSite_values.get(cookie['sameSite']):
            cookie['sameSite'] = valid_sameSite_values[cookie.pop('sameSite')]
        imported_cookies.insert(i, cookie)

    if file_name:
        with open(file_name, 'w') as file:
            json.dump(imported_cookies, file)
    return imported_cookies


def load_cookies(file_name='cookies.pkl'):
    with open(file_name, 'rb') as file:
        cookies = pickle.load(file)
        for cookie in cookies:
            driver.add_cookie(cookie)


def save_json_cookies(file_name=FILE_NAME):
    with open(file_name, 'w') as file:
        json.dump(cookies, file)


def load_json_cookies(file_name=FILE_NAME):
    with open(file_name, 'r') as file:
        cookies = json.load(file)
        for cookie in cookies:
            driver.add_cookie(cookie)


if __name__ == '__main__':

#     with open(r'c:/users/user/desktop/mouser_cookies.json', 'r') as file:
#         cookies = json.load(file)
#     edit_cookies(imported_cookies=cookies)
    service = webdriver.chrome.service.Service(ChromeDriverManager().install())

    ua = UserAgent()
    user_agent = ua.random

#     options = Options()
    options = webdriver.ChromeOptions()
    options.add_argument('--enable-javascript')
    options.add_argument(f'user-agent={user_agent}')
#     optoins.add_argument('--headless')
#     options.add_extension(EXTENSION_PATH)
    options.add_argument(f'--user-data-dir={WEBDRIVER_SERVICE_DIR}')

#    these options helps to hide automation browser functions
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(service=service, options=options)
    driver.delete_all_cookies()
    load_json_cookies(FILE_NAME)
#     driver.refresh()
    driver.get(URL)
#     save_cookies()
#     driver.switch_to.new_window(WindowTypes.TAB)
#     driver.get(URL_1)
#     save_cookies()
#     load_cookies())