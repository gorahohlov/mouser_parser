"""Простой тест для проверки работы undetected_chromedriver с Mouser"""
import undetected_chromedriver as uc
import time

# Тестовый URL
url = 'https://eu.mouser.com/c/?q=APHCM2012SYCK-F01'

print("Запуск браузера...")
driver = uc.Chrome()
driver.set_window_size(1530, 864)
driver.set_window_position(x=1, y=1)

print(f"Переход на {url}")
driver.get(url)

print("Ожидание 10 секунд для загрузки...")
time.sleep(10)

print(f"Текущий URL: {driver.current_url}")
print(f"Title: {driver.title}")

# Проверяем, не заблокирован ли доступ
if "access denied" in driver.title.lower() or "blocked" in driver.page_source.lower():
    print("ВНИМАНИЕ: Похоже, что доступ заблокирован!")
else:
    print("Доступ к сайту получен успешно!")

print("\nЗакрытие браузера через 5 секунд...")
time.sleep(5)
driver.quit()
print("Тест завершен")
