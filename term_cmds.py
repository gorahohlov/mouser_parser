import os
import base64

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from pprint import pprint as pp

URL = ('https://eu.mouser.com/ProductDetail/Altera/'
       'EPCS1SI8N?qs=jblrfmjbeiEDfo5ju%2FfMLw%3D%3D')

ARTICLE_LIST = [
    'APHCM2012SYCK-F01',
    'APHCM2012ZGC-F01',
    'WSL2512R0100FEA18',
    'AD8051ARZ-REEL7',
    '09031966921',
    '09185106323',
    '09692815421',
    '09670254715',
    '09670009922',
    '09670254715',
    '09670094715',
    '0900BL15C050E',
    '502585-0470',
    '51065-0500',
    'M80-0030006',
    'VS-10CTQ150-M3',
    'LT5538IDD#TRPBF',
    'DLFCV-1600+',
    'GRM31CR71E106KA12L',
    'GRM32ER71H106KA12L',
    'BZV55-B10,115',
    '09062486823',
    'DW-10-09-G-D-300',
    'FTSH-113-01-F-MT',
]


def save_page_as_pdf(driver, datasheet_path):
    part_number_el = driver.find_element(
        By.XPATH, '//span[@id="spnManufacturerPartNumber"]')
    part_number = part_number_el.text.strip()
    filename = f"{part_number}.pdf"
    full_path = os.path.join(datasheet_path, filename)
    result = driver.execute_cdp_cmd("Page.printToPDF", {
            "printBackground": True,
            "landscape": False,
            "paperWidth": 8.27,
            "paperHeight": 11.69}
            )
    with open(full_path, "wb") as f:
            f.write(base64.b64decode(result["data"]))
    print(f" PDF сохранён как: {full_path}")


driver = uc.Chrome()
driver.set_window_size(1530, 864)  # 1910, 1030
driver.set_window_position(x=1, y=1)
driver.get(URL)

driver.find_element(By.ID, 'onetrust-pc-btn-handler').click()
# driver.find_element(By.XPATH, '//button[text()="Cookie Settings"]').click()
driver.find_element(
        By.XPATH, 
        (
            '//button[@class="save-preference-btn-handler '
            'onetrust-close-btn-handler"]'
        )
).click()
# driver.find_element(
#         By.XPATH, 
#         '//button[contains(@class, "save-preference-btn-handler")]'
# ).click()
# driver.find_element(By.XPATH, 
#                     '//button[text()="Confirm My Choices"]'
# ).click()

# driver.find_element(By.ID, '1_hylRightFlagText').click()
# driver.find_element(By.ID, '1_divRightFlagImg').click()
driver.find_element(By.XPATH, '//span[text()="Mouser Europe"]').click()


# driver.find_element(By.CSS_SELECTOR, 'button[lang="en"]').click()
driver.find_element(By.ID, 'ddlLanguageMenuButton').click()
driver.find_element(By.XPATH, '//button[text()="русский язык"]').click()

# driver.find_elements(
#         By.XPATH, 
#         '//ol[contains(@class, "breadcrumb")]/li[@class][a[@href]]'
# )
header = driver.find_element(By.XPATH, 
                             '//ol[@class="breadcrumb mBodyCompact"]')
header_list = header.text.split('\n')
pp(header_list)

elements = driver.find_elements(By.XPATH, 
                                '//tr[contains(@id, "pdp_specs_SpecList")]')
elem_list = [elem.text for elem in elements]
elem_dict = {key: val for key, val in 
             (elem.replace(':\n ', ': ').split(': ', 1) for elem in elem_list)
            }
# elem_dict = {}
# for elem in elem_list:
#     norm_elem = elem.replace(':\n ', ': ')
#     key, val = norm_elem.split(': ', 1)
#     elem_dict[key] = val
pp(elem_dict)

