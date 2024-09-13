# import pickle
import json

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.window import WindowTypes
from lxml import html
from fake_useragent import UserAgent

# from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.chrome.service import Service


WEBDRIVER_SERVICE_DIR = r'c:/users/user/dev/mouser_parse_proj/wd_service_dir/'
EXTENSION_PATH = r'c:/users/user/dev/mouser_parse_proj/browsec_ext.crx'
URL = ('https://eu.mouser.com/ProductDetail/Altera/EPCS1SI8N?qs='
       + 'jblrfmjbeiEDfo5ju%2FfMLw%3D%3D'
       )
URL_1 = 'https://www.chipdip.ru/product/ecs-3953m-500-bn-tr-1'
FILE_NAME = r'c:/users/user/dev/mouser_parse_proj/mouser_cookies.json'
PATH = r'c:/users/user/dev/mouser_parse_proj/'
USER_AGENT = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
              + ' (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'
              )


# def save_cookies(file_name='cookies.pkl'):
#     with open(file_name, 'wb') as file:
#         pickle.dump(driver.get_cookies(), file)


def edit_cookies(file_name=None, imported_cookies=None):
    """Preprocessing imported cookies to selenium webdriver add_cookies
    method format"""
    valid_cookies_keys = {'httpOnly', 'value', 'name', 'domain',
                          'secure', 'path', 'expiry', 'sameSite'}
    valid_samesite_values = {'no_restriction': 'None', 'unspecified': 'Lax',
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
        cookie = {key: cookie[key] for key in valid_cookies_keys
                  if key in cookie}
        if valid_samesite_values.get(cookie['sameSite']):
            cookie['sameSite'] = valid_samesite_values[cookie.pop('sameSite')]
        imported_cookies.insert(i, cookie)

    if file_name:
        with open(file_name, 'w') as file:
            json.dump(imported_cookies, file)

    return imported_cookies


def gen_cookies_file(json_file=None, json_cookies=None):
    """Generate cookies.txt file in Netscape/Mozilla cookie file format
    from valid_cookies_keys in json_cookies"""
#     cookie_format = ['domain', 'flag', 'path', 'secure', 'expiration',
#                      'name', 'value']
    if json_file:
        with open(json_file, 'r') as file:
            json_cookies = json.load(file)
    elif not json_cookies:
        return

    with open(PATH + 'cookies.txt', 'w') as file:
        file.write('# Netscape HTTP Cookie File')
        for cookie in json_cookies:
            if cookie['httpOnly']:
                line = '#HttpOnly_' + cookie['domain'] + '\t'
            else:
                line = cookie['domain'] + '\t'
            line += 'TRUE\t' if cookie['domain'][0] == '.' else 'FALSE\t'
            line += cookie['path'] + '\t'
            line += str(cookie['secure']).upper() + '\t'
            line += str(int(cookie['expiry'])
                       ) + '\t' if cookie.get('expiry') else '0\t'
            line += cookie['name'] + '\t'
            line += cookie['value'] + '\n'
#             for i in range(5):
#                 line += cookie[cookie_format[i]] + '\t'
#             line += cookie['value'] + '\n'
            file.write(line)


# def load_cookies(file_name='cookies.pkl'):
#     with open(file_name, 'rb') as file:
#         cookies = pickle.load(file)
#         for cookie in cookies:
#             driver.add_cookie(cookie)


def save_json_cookies(file_name=FILE_NAME):
    with open(file_name, 'w') as file:
        json.dump(cookies, file)


def load_json_cookies(file_name=FILE_NAME):
    with open(file_name, 'r') as file:
        cookies = json.load(file)
        for cookie in cookies:
            driver.add_cookie(cookie)


if __name__ == '__main__':

#     edit_cookies(file_name=FILE_NAME)
#     with open(r'c:/users/user/desktop/mouser_cookies.json', 'r') as file:
#         cookies = json.load(file)
#     edit_cookies(imported_cookies=cookies)
    service = webdriver.chrome.service.Service(ChromeDriverManager().install())

    ua = UserAgent()
    user_agent = ua.random

#     options = Options()
    options = webdriver.ChromeOptions()
    options.add_argument('--enable-javascript')
    options.add_argument(f'user-agent={USER_AGENT}')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
#     optoins.add_argument('--headless')
#     options.add_extension(EXTENSION_PATH)
    options.add_argument(f'--user-data-dir={WEBDRIVER_SERVICE_DIR}')

#    these options helps to hide automation browser functions
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)

    breakpoint()
#     load_json_cookies(FILE_NAME)  # only for debugging purpose

    driver = webdriver.Chrome(service=service, options=options)
#     driver.get(URL_1)
    driver.get('https://eu.mouser.com')
    driver.delete_all_cookies()
    load_json_cookies(FILE_NAME)
    driver.refresh()
#     driver.get(URL)
#     save_cookies()
#     driver.switch_to.new_window(WindowTypes.TAB)
#     save_cookies()
#     load_cook:es())
