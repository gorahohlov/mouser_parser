import os
import base64

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import StaleElementReferenceException


DATASHEET_DIR = 'c:/users/user/dev/mouser_parse_proj/datasheets/'
URL = (
    'https://eu.mouser.com/ProductDetail/Altera/EPCS1SI8N?qs='
    'jblrfmjbeiEDfo5ju%2FfMLw%3D%3D'
)


class ElementHelper:
    def __init__(self, driver, timeout=6):
        self.driver = driver
        self.wait = WebDriverWait(driver, timeout)

    def get_texts_safe(self, xpath, retries=3):
        """Returns elements text list by xpath w/o StaleException"""
        for _ in range(retries):
            try:
                elements = self.driver.find_elements(By.XPATH, xpath)
                return [el.text for el in elements]
            except StaleElementReferenceException:
                continue
        return []

    def click_when_clickable(self, by, value):
        """Waits clickable elem loading and click to it"""
        try:
            element = self.wait.until(EC.element_to_be_clickable((by, value)))
            element.click()
            return True
        except TimeoutException:
            return False

    def get_attribute_safe(self, xpath, attr, retries=3):
        """Returns attribute values list for whole elems by xpath"""
        for _ in range(retries):
            try:
                elements = self.driver.find_elements(By.XPATH, xpath)
                return [el.get_attribute(attr) for el in elements]
            except StaleElementReferenceException:
                continue
        return []

    def wait_for_presence(self, by, value):
        """Waits until the element become visible"""
        return self.wait.until(EC.presence_of_element_located((by, value)))

def initialize_mouser_ui(driver):
    """Initial mouser chrome session setup"""
    driver.get(URL)
    wait = WebDriverWait(driver, 6)
    try:
        wait.until(
            EC.element_to_be_clickable(
                (By.ID, 'onetrust-pc-btn-handler')
            )
        ).click()
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, 
                 '//button[@class="save-preference-btn-handler '
                 'onetrust-close-btn-handler"]'
                 )
            )
        ).click()
        wait.until(
            EC.element_to_be_clickable(
                (By.ID, '1_hylRightFlagText')
            )
        ).click()
        wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'button[lang="en"]')
            )
        ).click()
        wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'button[lang="ru"]')
            )
        ).click()
        print('Initial site setup mouser.com completed')
    except TimeoutException:
        print('Initial site setup mouser.com failed to complete normally')

#     driver.find_element(By.ID, 'onetrust-pc-btn-handler').click()
#     driver.find_element(By.XPATH, 
#                         '//button[text()="Cookie Settings"]').click()
#     driver.find_element(By.XPATH, 
#                         '//button[@class="save-preference-btn-handler '
#                         'onetrust-close-btn-handler"]').click()
#     driver.find_element(By.ID, '1_hylRightFlagText').click()
#     driver.find_element(By.ID, "1_divRightFlagImg").click()
#     driver.find_element(By.CSS_SELECTOR, 'button[lang="en"]').click()

def start_driver_with_position(width=788, height=864): # x=755, y=0,)
    options = uc.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument(f"--window-size={width}, {height}")
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = uc.Chrome(options=options)
    driver.set_window_position(x=755, y=0)
#     driver.set_window_size(width, height)
    return driver

def save_page_as_pdf(driver, datasheet_path):
    part_number_el = driver.find_element(
        By.XPATH, '//span[@id="spnManufacturerPartNumber"]')
    part_number = part_number_el.text.strip()

    # 2 Формируем полный путь
    filename = f"{part_number}.pdf"
    full_path = os.path.join(datasheet_path, filename)

    # 3 Используем DevTools Protocol (CDP) для сохранения PDF
    result = driver.execute_cdp_cmd("Page.printToPDF", {
            "printBackground": True,
            "landscape": False,
            "paperWidth": 8.27,  # A4 width (in inches)
            "paperHeight": 11.69}  # A4 height
            )

    # 4 Записываем PDF в файл
    with open(full_path, "wb") as f:
            f.write(base64.b64decode(result["data"]))

    print(f" PDF сохранён как: {full_path}")


if __name__ == "__main__":

    # options = uc.ChromeOptions()
    # driver = uc.Chrome(options=options)

    # driver.set_window_position(1190, 1)
    # driver.set_window_size(735, 1080)

    # Открываем страницу
    # URL = (
    #     'https://eu.mouser.com/ProductDetail/Altera/EPCS1SI8N?qs='
    #     'jblrfmjbeiEDfo5ju%2FfMLw%3D%3D'
    # )

    # driver.get(URL)

    driver = start_driver_with_position()
    driver.get(URL)
    driver.save_screenshot("./dev/mouser_parse_proj/screenshots/screen.png")
#     initialize_mouser_ui(driver)
    helper = ElementHelper(driver)
    helper.click_when_clickable(By.ID, 'onetrust-pc-btn-handler')
    helper.click_when_clickable(By.XPATH, 
                                '//button[@class="save-preference-btn-handler'
                                ' onetrust-close-btn-handler"]')
    helper.click_when_clickable(By.ID, '1_hylRightFlagText')
    WebDriverWait(driver, 10).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    helper.click_when_clickable(By.CSS_SELECTOR, 'button[lang="en"]')
    helper.click_when_clickable(By.CSS_SELECTOR, 'button[lang="ru"]')
    WebDriverWait(driver, 10).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    driver.save_screenshot("after_lang_click.png")
    WebDriverWait(driver, 6).until(
            lambda d: d.find_element(By.TAG_NAME, "html")
                      .get_attribute("lang") == "ru"
    )
    print('Initial site setup mouser.com completed')

    tr_list = helper.get_texts_safe('//tr[contains(@id, '
                                    '"pdp_specs_SpecList")]')
    bc_list = helper.get_texts_safe('//ol[@class="breadcrumb"]'
                                    '/li[@class][a[@href]]')

    save_page_as_pdf(driver, DATASHEET_DIR)

    # Получаем все строки таблицы спецификаций
    rows = driver.find_elements(
                    By.XPATH, '//tr[contains(@id, "pdp_specs_SpecList")]')

    spec_string = ''
    for row in rows:
        spec_string += row.text + ', '
    print(f'Specification characteristics:\n{spec_string}')

    # Получаем все <li> внутри ol.breadcrumb, у которых есть class и <a href>
    breadcrumb_items = driver.find_elements(
                         By.XPATH, 
                         '//ol[@class="breadcrumb"]/li[@class][a[@href]]')

    # Извлекаем текст
    breadcrumb_texts = []
    for li in breadcrumb_items:
        try:
            text = li.text.strip()
            if text:
                breadcrumb_texts.append(text)
        except Exception as e:
            print(f"Ошибка при извлечении breadcrumb: {e}")

    print("\nНавигационная цепочка (breadcrumb):")
    print(" > ".join(breadcrumb_texts))

    # Закрываем драйвер
    print("Модуль штатно завершил все процедуры обработки "
          "и завершает свою работу")
    driver.quit()
