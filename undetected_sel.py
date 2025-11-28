import os
import base64
import time
import random
import json

import undetected_chromedriver as uc
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import StaleElementReferenceException
from pprint import pprint as pp


DATASHEET_DIR = r'c:\users\user\dev\mouser_parse_proj\mouser_datasheets'
EXCEL_FILE = r'c:\users\user\dev\mouser_parse_proj\mouser_datasheets\mouser_data.xlsx'

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


def human_delay(min_sec=0.5, max_sec=2.0):
    """Random delay to simulate human behavior"""
    time.sleep(random.uniform(min_sec, max_sec))


class ElementHelper:
    def __init__(self, driver, timeout=10):
        self.driver = driver
        self.wait = WebDriverWait(driver, timeout)

    def get_texts_safe(self, xpath, retries=3):
        """Returns elements text list by xpath w/o StaleException"""
        for _ in range(retries):
            try:
                elements = self.driver.find_elements(By.XPATH, xpath)
                return [el.text for el in elements]
            except StaleElementReferenceException:
                human_delay(0.3, 0.7)
                continue
        return []

    def click_when_clickable(self, by, value, timeout=None):
        """Waits clickable elem loading and click it"""
        try:
            wait = WebDriverWait(self.driver, timeout) if timeout else self.wait
            element = wait.until(EC.element_to_be_clickable((by, value)))
            human_delay(0.3, 0.8)
            element.click()
            human_delay(0.5, 1.5)
            return True
        except TimeoutException:
            print(f"Timeout waiting for element: {by}={value}")
            return False

    def get_attribute_safe(self, xpath, attr, retries=3):
        """Returns attribute values list for whole elems by xpath"""
        for _ in range(retries):
            try:
                elements = self.driver.find_elements(By.XPATH, xpath)
                return [el.get_attribute(attr) for el in elements]
            except StaleElementReferenceException:
                human_delay(0.3, 0.7)
                continue
        return []

    def wait_for_presence(self, by, value, timeout=None):
        """Waits until the element become visible"""
        wait = WebDriverWait(self.driver, timeout) if timeout else self.wait
        return wait.until(EC.presence_of_element_located((by, value)))

#region commented "Initial mouser chrome session setup"
# def initialize_mouser_ui(driver):
#     """Initial mouser chrome session setup"""
#     driver.get(URL)
#     wait = WebDriverWait(driver, 6)
#     try:
#         wait.until(
#             EC.element_to_be_clickable(
#                 (By.ID, 'onetrust-pc-btn-handler')
#             )
#         ).click()
#         wait.until(
#             EC.element_to_be_clickable(
#                 (By.XPATH, 
#                  '//button[@class="save-preference-btn-handler '
#                  'onetrust-close-btn-handler"]'
#                  )
#             )
#         ).click()
#         wait.until(
#             EC.element_to_be_clickable(
#                 (By.ID, '1_hylRightFlagText')
#             )
#         ).click()
#         wait.until(
#             EC.presence_of_element_located(
#                 (By.CSS_SELECTOR, 'button[lang="en"]')
#             )
#         ).click()
#         wait.until(
#             EC.presence_of_element_located(
#                 (By.CSS_SELECTOR, 'button[lang="ru"]')
#             )
#         ).click()
#         print('Initial site setup mouser.com completed')
#     except TimeoutException:
#         print('Initial site setup mouser.com failed to complete normally')

#     driver.find_element(By.ID, 'onetrust-pc-btn-handler').click()
#     driver.find_element(By.XPATH, 
#                         '//button[text()="Cookie Settings"]').click()
#     driver.find_element(By.XPATH, 
#                         '//button[@class="save-preference-btn-handler '
#                         'onetrust-close-btn-handler"]').click()
#     driver.find_element(By.ID, '1_hylRightFlagText').click()
#     driver.find_element(By.ID, "1_divRightFlagImg").click()
#     driver.find_element(By.CSS_SELECTOR, 'button[lang="en"]').click()
#endregion

def start_driver_with_position(width=1910, height=1030, x=1, y=1, headless=False):
    """Initialize undetected Chrome driver with minimal settings"""
    options = uc.ChromeOptions()
    
    if headless:
        options.add_argument("--headless=new")
    
    # Минимальный набор опций как в term_cmds.py
    options.add_argument(f"--window-size={width},{height}")
    
    # Создаем драйвер с минимальными настройками (как в term_cmds.py)
    driver = uc.Chrome(options=options, version_main=None)
    
    if not headless:
        driver.set_window_size(width, height)
        driver.set_window_position(x=x, y=y)
    
    return driver

def save_page_as_pdf(driver, datasheet_path, article=None):
    """Save current page as PDF if not already exists"""
    try:
        part_number_el = driver.find_element(
            By.XPATH, '//span[@id="spnManufacturerPartNumber"]')
        part_number = part_number_el.text.strip()
    except Exception as e:
        print(f"Не удалось получить part number: {e}")
        part_number = article if article else "unknown"
    
    filename = f"{part_number}.pdf"
    full_path = os.path.join(datasheet_path, filename)
    
    # Проверяем существование файла
    if os.path.exists(full_path):
        print(f"PDF уже существует (пропускаем): {filename}")
        return {'saved': False, 'exists': True, 'filename': filename, 'path': full_path}
    
    try:
        result = driver.execute_cdp_cmd("Page.printToPDF", {
                "printBackground": True,
                "landscape": False,
                "paperWidth": 8.27,
                "paperHeight": 11.69}
                )
        with open(full_path, "wb") as f:
                f.write(base64.b64decode(result["data"]))
        print(f"PDF сохранён: {filename}")
        return {'saved': True, 'exists': False, 'filename': filename, 'path': full_path}
    except Exception as e:
        print(f"Ошибка сохранения PDF: {e}")
        return {'saved': False, 'exists': False, 'filename': filename, 'error': str(e)}


def initialize_mouser_session(driver, helper):
    """Initial mouser site setup: cookies, language, region"""
    print("Инициализация сессии Mouser...")
    
    # Ждем загрузки страницы
    human_delay(3, 5)
    
    # Cookie settings - увеличены таймауты
    if helper.click_when_clickable(By.ID, 'onetrust-pc-btn-handler', timeout=20):
        print("  [OK] Cookie Settings открыты")
    else:
        print("  [SKIP] Cookie Settings не найдены (возможно уже приняты)")
    
    if helper.click_when_clickable(By.XPATH, 
                                   '//button[@class="save-preference-btn-handler '
                                   'onetrust-close-btn-handler"]', timeout=15):
        print("  [OK] Cookie Settings сохранены")
    else:
        print("  [SKIP] Кнопка подтверждения cookies не найдена")
    
    # Выбор региона
    if helper.click_when_clickable(By.XPATH, '//span[text()="Mouser Europe"]', timeout=15):
        print("  [OK] Регион: Mouser Europe")
        human_delay(2, 3)
    else:
        print("  [SKIP] Выбор региона не требуется")
    
    # Выбор языка
    if helper.click_when_clickable(By.ID, 'ddlLanguageMenuButton', timeout=15):
        print("  [OK] Меню языка открыто")
        
        if helper.click_when_clickable(By.XPATH, '//button[text()="русский язык"]', timeout=10):
            print("  [OK] Язык: русский")
            # Ждем применения языка
            human_delay(3, 5)
        else:
            print("  [SKIP] Русский язык не найден в меню")
    else:
        print("  [SKIP] Меню языка не найдено (возможно уже выбран)")
    
    print("Инициализация завершена\n")
    return True


def get_article_url(article):
    """Generate Mouser search URL for article"""
    # Можно использовать прямой поиск
    return f"https://eu.mouser.com/c/?q={article}"


def extract_product_data(driver, helper, article):
    """Extract breadcrumb and specifications from product page"""
    data = {
        'article': article,
        'breadcrumb': [],
        'specifications': {},
        'part_number': None
    }
    
    try:
        # Получаем breadcrumb
        header = driver.find_element(By.XPATH, '//ol[@class="breadcrumb mBodyCompact"]')
        header_list = header.text.split('\n')
        data['breadcrumb'] = header_list
        print(f"Заголовок: {' > '.join(header_list)}")
    except Exception as e:
        print(f"Не удалось получить breadcrumb: {e}")
    
    try:
        # Получаем спецификации
        elements = driver.find_elements(By.XPATH, '//tr[contains(@id, "pdp_specs_SpecList")]')
        elem_list = [elem.text for elem in elements]
        elem_dict = {key: val for key, val in 
                     (elem.replace(':\n ', ': ').split(': ', 1) for elem in elem_list)
                     }
        data['specifications'] = elem_dict
        print(f"Спецификаций найдено: {len(elem_dict)}")
    except Exception as e:
        print(f"Не удалось получить спецификации: {e}")
    
    try:
        # Получаем part number
        part_number_el = driver.find_element(By.XPATH, '//span[@id="spnManufacturerPartNumber"]')
        data['part_number'] = part_number_el.text.strip()
        print(f"Part Number: {data['part_number']}")
    except Exception as e:
        print(f"Не удалось получить part number: {e}")
    
    return data


def save_to_excel(all_data, excel_file):
    """Save collected data to Excel file"""
    print(f"\nСохранение данных в Excel: {excel_file}")
    
    # Подготовка данных для DataFrame
    rows = []
    for data in all_data:
        # Базовая информация
        row = {
            'Артикул': data.get('article', ''),
            'Part Number': data.get('part_number', ''),
            'Заголовок': ' > '.join(data.get('breadcrumb', [])),
            'PDF сохранен': 'Да' if data.get('pdf_saved') or data.get('pdf_exists') else 'Нет',
            'PDF файл': data.get('pdf_filename', ''),
        }
        
        # Добавляем спецификации как отдельные колонки
        specs = data.get('specifications', {})
        for key, value in specs.items():
            row[f'Спец: {key}'] = value
        
        # Ошибки
        if 'error' in data:
            row['Ошибка'] = data['error']
        
        rows.append(row)
    
    # Создаем DataFrame
    df = pd.DataFrame(rows)
    
    # Сохраняем в Excel
    try:
        df.to_excel(excel_file, index=False, engine='openpyxl')
        print(f"[OK] Данные успешно сохранены в Excel")
        print(f"  Строк: {len(df)}")
        print(f"  Колонок: {len(df.columns)}")
        return True
    except Exception as e:
        print(f"Ошибка сохранения в Excel: {e}")
        # Пробуем сохранить в CSV как запасной вариант
        csv_file = excel_file.replace('.xlsx', '.csv')
        try:
            df.to_csv(csv_file, index=False, encoding='utf-8-sig')
            print(f"[OK] Данные сохранены в CSV: {csv_file}")
            return True
        except Exception as e2:
            print(f"Ошибка сохранения в CSV: {e2}")
            return False


def process_article(driver, helper, article, datasheet_dir, first_article=False):
    """Process single article: load page, extract data, save PDF"""
    print(f"\n{'='*60}")
    print(f"Обработка артикула: {article}")
    print(f"{'='*60}")
    
    # Формируем URL
    search_url = get_article_url(article)
    print(f"Загружаем: {search_url}")
    
    driver.get(search_url)
    human_delay(3, 5)
    
    # Первый раз нужна инициализация
    if first_article:
        initialize_mouser_session(driver, helper)
        # После инициализации может потребоваться перезагрузка страницы
        print(f"Перезагружаем страницу артикула...")
        driver.get(search_url)
        human_delay(2, 4)
    
    # Извлекаем данные
    data = extract_product_data(driver, helper, article)
    
    # Сохраняем PDF
    pdf_result = save_page_as_pdf(driver, datasheet_dir, article)
    data['pdf_saved'] = pdf_result.get('saved', False)
    data['pdf_exists'] = pdf_result.get('exists', False)
    data['pdf_filename'] = pdf_result.get('filename', '')
    
    print(f"Артикул {article} обработан")
    
    return data


if __name__ == "__main__":
    # Настройки запуска
    TEST_MODE = False  # True - тестовый режим (первые 2 артикула), False - полная обработка
    HEADLESS = False   # True - headless режим, False - видимое окно браузера
    SKIP_INIT = False  # True - пропустить инициализацию (cookies/язык), если уже настроено
    
    # Создаем директорию для PDF если не существует
    os.makedirs(DATASHEET_DIR, exist_ok=True)
    
    # Определяем список артикулов для обработки
    articles_to_process = ARTICLE_LIST[:6] if TEST_MODE else ARTICLE_LIST
    
    print("="*60)
    print("Запуск автоматизированного сбора данных с Mouser.com")
    print("="*60)
    print(f"Режим: {'ТЕСТОВЫЙ' if TEST_MODE else 'ПОЛНАЯ ОБРАБОТКА'}")
    print(f"Браузер: {'Headless' if HEADLESS else 'Видимый'}")
    print(f"Артикулов для обработки: {len(articles_to_process)}")
    print(f"Директория для PDF: {DATASHEET_DIR}")
    print(f"Файл Excel: {EXCEL_FILE}\n")
    
    # Инициализируем драйвер
    driver = start_driver_with_position(headless=HEADLESS)
    helper = ElementHelper(driver)
    
    # Инициализация сессии (только один раз, если требуется)
    if not SKIP_INIT:
        print("Загружаем главную страницу Mouser для инициализации...")
        driver.get('https://eu.mouser.com')
        human_delay(2, 4)
        initialize_mouser_session(driver, helper)
    else:
        print("Пропуск инициализации (SKIP_INIT=True)")
    
    # Обрабатываем все артикулы
    all_data = []
    for idx, article in enumerate(articles_to_process, 1):
        print(f"\nПрогресс: {idx}/{len(articles_to_process)}")
        try:
            data = process_article(
                driver, 
                helper, 
                article, 
                DATASHEET_DIR,
                first_article=False  # Инициализация уже выполнена
            )
            all_data.append(data)
            
            # Задержка между артикулами
            if idx < len(articles_to_process):
                delay = random.uniform(2, 5)
                print(f"Пауза {delay:.1f}с перед следующим артикулом...")
                time.sleep(delay)
                
        except Exception as e:
            print(f"ОШИБКА при обработке {article}: {e}")
            all_data.append({
                'article': article,
                'error': str(e),
                'breadcrumb': [],
                'specifications': {},
                'part_number': None
            })
    
    # Закрываем драйвер
    print("\n" + "="*60)
    print("Обработка завершена")
    print("="*60)
    print(f"Успешно обработано: {len([d for d in all_data if 'error' not in d])}")
    print(f"Ошибок: {len([d for d in all_data if 'error' in d])}")
    
    driver.quit()
    
    # Сохраняем результаты в Excel
    save_to_excel(all_data, EXCEL_FILE)
    
    # Дополнительно сохраняем в JSON для отладки
    results_file = os.path.join(DATASHEET_DIR, 'results.json')
    try:
        with open(results_file, 'w', encoding='utf-8') as f:
            json.dump(all_data, f, ensure_ascii=False, indent=2)
        print(f"Результаты также сохранены в JSON: {results_file}")
    except Exception as e:
        print(f"Не удалось сохранить JSON: {e}")
