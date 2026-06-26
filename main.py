from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from time import sleep

options = Options()
options.add_argument("--start-maximized")

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)

parsed_data = []

try:
    print("Загрузка страницы...")
    try:
        driver.get("https://funpay.com/lots/436/")
    except Exception as e:
        print(f"Не удалось загрузить сайт: {e}")
        driver.quit()
        exit()

    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tc-item")))

    items = driver.find_elements(By.CLASS_NAME, "tc-item")
    print(f"Найдено лотов на странице: {len(items)}. Обрабатываем первые 50...")

    for item in items[0:50]:
        try:
            cena = item.find_element(By.CLASS_NAME, "tc-price").text
        except:
            cena = "Не указана"

        try:
            opis = item.find_element(By.CLASS_NAME, "tc-desc-text").text
        except:
            opis = "Без описания"

        try:
            nik = item.find_element(By.CLASS_NAME, "media-user-name").text
        except:
            nik = "Скрыт"

        stars_elements = item.find_elements(By.CLASS_NAME, "fas")
        rating_stars = len(stars_elements)
        
        try:
            reviews_count = item.find_element(By.CLASS_NAME, "rating-mini-count").text
        except:
            reviews_count = "0"

        try:
            time_on_site = item.find_element(By.CLASS_NAME, "media-user-info").text
        except:
            time_on_site = "Нет данных"

        print(f"Ник: {nik} | Цена: {cena} | Отзывов: {reviews_count} (Звезд: {rating_stars}/5) | На сайте: {time_on_site}")

        parsed_data.append({
            "Никнейм": nik,
            "Цена": cena,
            "Описание": opis,
            "Рейтинг (Звезд)": rating_stars,
            "Количество отзывов": reviews_count,
            "Время на сайте": time_on_site
        })

    print("\nСохранение данных в Excel...")
    df = pd.DataFrame(parsed_data)
    
    df.to_excel("funpay_lots.xlsx", index=False)
    print("Готово! Файл 'funpay_lots.xlsx' успешно сохранен.")

finally:
    driver.quit()