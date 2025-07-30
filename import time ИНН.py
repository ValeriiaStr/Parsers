import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Пути к файлам и драйверу
INN_FILE_PATH = r"C:\Users\valer\OneDrive\Рабочий стол\ИНН.txt"
CHROMEDRIVER_PATH = r"D:\Пайтон\chromedriver-win64\chromedriver-win64\chromedriver.exe"
OUTPUT_FILE = "результаты.xlsx"

# Настройки браузера
options = webdriver.ChromeOptions()
# options.add_argument("--headless")  # отключите, если хотите видеть браузер
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=options)

wait = WebDriverWait(driver, 15)  # таймаут 15 секунд

# Чтение списка ИНН из файла
with open(INN_FILE_PATH, 'r', encoding='utf-8') as f:
    inn_list = [line.strip() for line in f if line.strip()]

results = []

try:
    for inn in inn_list:
        try:
            print(f"\nОбработка ИНН: {inn}")
            url = f"https://spark-interfax.ru/search?Query={inn}"
            driver.get(url)

            # Ждем загрузки страницы и появления блока результатов
            result_block_selector = "#search-result-items > li"
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, result_block_selector)))
            except:
                print(f"Результат не найден или страница не загрузилась для ИНН {inn}")
                results.append({"ИНН": inn, "Данные": "Результат не найден"})
                continue

            # Проверяем первый результат (или нужный по порядку)
            first_result_selector = "#search-result-items > li:nth-child(1)"
            first_result_element = driver.find_element(By.CSS_SELECTOR, first_result_selector)

            # Внутри этого элемента ищем ИНН по селектору
            inn_selector_in_result = "div > div:nth-child(2) > div > div > div.code > span:nth-child(2) > span"
            inn_element_in_result = first_result_element.find_element(By.CSS_SELECTOR, inn_selector_in_result)
            found_inn = inn_element_in_result.text.strip()

            if found_inn != inn:
                print(f"ИНН не совпадает: искомый {inn}, найденный {found_inn}")
                results.append({"ИНН": inn, "Данные": "ИНН не совпадает"})
                continue

            # Если совпало, извлекаем нужные данные
            data_selector = "div > div:nth-child(2) > div > div > div:nth-child(4) > span"
            data_element = first_result_element.find_element(By.CSS_SELECTOR, data_selector)
            data_text = data_element.text.strip()

            print(f"ИНН совпадает. Полученные данные: {data_text}")
            results.append({"ИНН": inn, "Данные": data_text})

        except Exception as e:
            print(f"Ошибка при обработке ИНН {inn}: {e}")
            results.append({"ИНН": inn, "Данные": "Ошибка"})
        finally:
            time.sleep(1)  # задержка между итерациями

except KeyboardInterrupt:
    print("\nОбработка прервана пользователем. Сохраняю результаты...")

finally:
    driver.quit()
    # Сохраняем результаты в Excel-файл при любом исходе
    df = pd.DataFrame(results)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\nРезультаты сохранены в файл {OUTPUT_FILE}")