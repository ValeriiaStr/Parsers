import os
import re
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook

# Функция для очистки имени файла/папки
def clean_article_name(name):
    return re.sub(r'[^A-Za-z0-9]', '_', name)

# Чтение списка артикулов из файла
def read_articles_from_file(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        return [line.strip() for line in f if line.strip()]

# Создаем книгу Excel и добавляем заголовки
wb = Workbook()
ws = wb.active
ws.title = "Отзывы"
ws.append(['Артикул', 'Количество отзывов', 'Оценка', 'Отзыв', 'Достоинства', 'Недостатки', 'Фото', 'Видео'])

# Настройки Chrome для подавления логов и запуска в фоновом режиме (опционально)
options = Options()
options.add_argument("--log-level=3")
# options.add_argument("--headless")  # Раскомментируйте, если хотите запускать без окна браузера

service = Service(r"D:\Пайтон\chromedriver-win64\chromedriver-win64\chromedriver.exe")
service.log_path = os.devnull

driver = webdriver.Chrome(service=service, options=options)

try:
    # Укажите полный путь к файлу с артикулами
    articles_file = r"C:\Users\valer\OneDrive\Рабочий стол\articles.txt.txt"
    articles = read_articles_from_file(articles_file)

    for article_number in articles:
        print(f"\nОбработка артикула: {article_number}")
        # 1. Переход на страницу товара по артикулу
        product_url = f'https://www.wildberries.ru/catalog/{article_number}/detail.aspx'
        driver.get(product_url)
        wait = WebDriverWait(driver, 15)

        # 2. Попытка найти кнопку "Отзывы" по id и кликнуть её (если есть)
        try:
            feedback_button = wait.until(EC.element_to_be_clickable((By.ID, 'comments_reviews_link')))
            feedback_button.click()
            print("Кнопка отзывов найдена и нажата.")
        except:
            print("Кнопка отзывов не найдена или не нажата.")
            continue  # Переходим к следующему артикулу

        time.sleep(2)  # Ждем загрузки

        # Попытка найти количество отзывов внутри кнопки (если нужно)
        try:
            count_span = driver.find_element(By.CSS_SELECTOR, "button#comments_reviews_link span.product-review__count")
            review_count = count_span.text.strip()
        except:
            review_count = "0"

        # 3. На странице отзывов ищем отзывы и оценки
        try:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'li.comments__item.feedback')))
        except:
            print("Отзывы не найдены.")
            continue

        # Прокрутка страницы для загрузки всех отзывов (если есть подгрузка)
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        # Нажатие "Показать еще" (если есть)
        while True:
            try:
                load_more_btn = driver.find_element(By.CLASS_NAME, 'feedback__load-more')
                if load_more_btn.is_displayed():
                    load_more_btn.click()
                    print("Кликаю по 'Показать еще'...")
                    time.sleep(3)
                else:
                    break
            except:
                break

        review_elements = driver.find_elements(By.CSS_SELECTOR, 'li.comments__item.feedback')

        for elem in review_elements:
            # Извлечение текста отзыва
            try:
                review_text_elem = elem.find_element(By.CSS_SELECTOR, 'p.feedback__text.j-feedback__text')
                review_text = review_text_elem.text.strip()
            except:
                review_text = ""

            # Извлечение оценки (звезд)
            try:
                rating_span = elem.find_element(By.CSS_SELECTOR, 'span.feedback__rating')
                classes = rating_span.get_attribute('class').split()
                rating_class = next((c for c in classes if c.startswith('star')), None)
                if rating_class:
                    rating_value = int(rating_class.replace('star', ''))
                else:
                    rating_value= "Нет оценки"
            except:
                rating_value= "Нет оценки"

            # Извлечение достоинств (преимущества)
            try:
                pros_elem= elem.find_element(By.CSS_SELECTOR,'span.feedback__text--item-feedback--pro')
                pros_texts= [span.text.strip() for span in pros_elem.find_elements(By.CSS_SELECTOR,'span.feedback__text--item')]
                pros_texts_str= "; ".join(pros_texts)
            except:
                pros_texts_str= ""

            # Извлечение недостатков (минусы)
            try:
                cons_elem= elem.find_element(By.CSS_SELECTOR,'span.feedback__text--item-feedback--con')
                cons_texts= [span.text.strip() for span in cons_elem.find_elements(By.CSS_SELECTOR,'span.feedback__text--item')]
                cons_texts_str= "; ".join(cons_texts)
            except:
                cons_texts_str= ""

            # Фото отзывы (если есть)
            photos_list_str= ""
            try:
                photos_ul= elem.find_element(By.CSS_SELECTOR,'ul.feedback__photos')
                photo_imgs= photos_ul.find_elements(By.TAG_NAME,'img')
                photo_urls= [img.get_attribute('src') for img in photo_imgs]
                photos_list_str= "; ".join(photo_urls)
            except:
                photos_list_str= ""

            # Видео отзывы (если есть кнопка просмотра видео)
            video_link_str= ""
            try:
                video_btns= elem.find_elements(By.CSS_SELECTOR,'button.feedback__video-btn[aria-label="Посмотреть видеоотзыв"]')
                if video_btns and video_btns[0].is_displayed():
                    video_link_str= "Видео есть"
                    # Можно дополнительно кликнуть и получить ссылку или скриншот видео при необходимости.
                else:
                    video_link_str= ""
            except:
                video_link_str= ""

            ws.append([article_number, review_count, rating_value, review_text, pros_texts_str, cons_texts_str, photos_list_str, video_link_str])

        print(f"Обработано {len(review_elements)} отзывов для артикула {article_number}.")

finally:
    driver.quit()

# Сохраняем Excel файл после завершения обработки всех артикулов
wb.save('reviews.xlsx')
print("Все отзывы сохранены в reviews.xlsx.")