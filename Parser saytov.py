import requests
from bs4 import BeautifulSoup
import pandas as pd
import logging

# Настраиваем логирование, чтобы видеть процесс работы в консоли
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# ИСПРАВЛЕНО: Точный адрес магазина книг для тренировки парсинга
URL = "https://books.toscrape.com/" 
N = 5 #Сколько страниц нужно просмотреть


URLT = ""
def Downloading_HTML(URLL):
    global URLT
        # Имитируем реальный браузер, чтобы сайт нас не заблокировал (User-Agent)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
        # 1. Скачиваем HTML-код страницы
    try:
        response = requests.get(URLL, headers=headers, timeout=10)
        response.raise_for_status() # Проверяем, что сайт ответил успешно (код 200)
    except Exception as e:
        logging.error(f"Не удалось загрузить страницу: {e}")
        return

    # 2. Передаем код страницы парсеру BeautifulSoup
    soup = BeautifulSoup(response.text, "html.parser")
    
    # Создаем пустой список, куда будем складывать словари с данными книг
    books_data = []

    # 3. ИСПРАВЛЕНО: Используем class_ вместо class, чтобы не нарушать синтаксис Python
    books = soup.find_all("article", class_="product_pod")
    logging.info(f"Найдено {len(books)} товаров на странице. Начинаем сбор данных...")

    for book in books:
        try:
            # Извлекаем название книги (оно лежит внутри тега <h3> -> <a> в атрибуте title)
            title = book.h3.a["title"]
            
            # Извлекаем цену (тег <p class="price_color">)
            price = book.find("p", class_="price_color").text
            # Очищаем цену от лишних символов валюты (£), оставляя только цифры и точку
            price_cleaned = float(price.replace("£", "").replace("Â", ""))
            
            # Извлекаем статус наличия (тег <p class="instock availability">)
            availability = book.find("p", class_="instock availability").text.strip()

            # Формируем относительную ссылку на картинку товара
            img_src = book.find("img", class_="thumbnail")["src"]
            full_img_url = URL + img_src.replace("../", "")

            # Сохраняем собранные данные одной книги в словарь
            books_data.append({
                "Название книги": title,
                "Цена (£)": price_cleaned,
                "Статус": availability,
                "Ссылка на обложку": full_img_url
            })
        except Exception as item_error:
            logging.warning(f"Ошибка при разборе отдельного товара: {item_error}")
            continue
        # 1. Ищем кнопку "Next" на странице
    next_button = soup.find("li", class_="next")

    if next_button:
        # 2. Вытаскиваем значение атрибута href (в вашем случае это "page-2.html")
        next_page_url = next_button.a["href"]
        
        # 3. Собираем полную ссылку для следующего запроса
        # Получится: http://toscrape.com (или аналогично)
        from urllib.parse import urljoin
        full_next_url = urljoin(URLL, next_page_url)
        URLT = full_next_url
        print(f"Ссылка на следующую страницу найдена: {full_next_url}")
    else:
        print("Это последняя страница, кнопки Next больше нет.")

    return books_data

def parse_books():
    logging.info("Запуск парсера. Подключение к сайту...")

    # Главный плоский список для ВСЕХ книг со всех страниц
    all_books = []
    
    # Начинаем с базового URL первой страницы
    current_url = URL  
    I = 0
    # Автоматический цикл: работаем, пока есть куда переходить
    while current_url and I <= N :
        # Скачиваем текущую страницу и получаем список книг с неё
        page_books = Downloading_HTML(current_url)
        
        # Если что-то скачалось, добавляем элементы в общий список
        if page_books:
            # ВАЖНО: используем extend вместо append, чтобы список оставался плоским!
            all_books.extend(page_books)
        
        # Переключаем current_url на следующую страницу, которую нашла функция Downloading_HTML
        # Если кнопка Next пропадет, Downloading_HTML не обновит глобальную переменную URLT, 
        # но чтобы цикл гарантированно остановился на последней странице, мы управляем ссылкой:
        if current_url == URLT or not URLT:
            break
        else:
            current_url = URLT
            # Сбрасываем URLT перед следующим шагом для безопасности
            globals()['URLT'] = "" 
        I += 1
        

    # 4. Переносим объединенные данные в таблицу Pandas и сохраняем в Excel
    if all_books:
        df = pd.DataFrame(all_books)   
        
        # Имя готового файла
        output_file = "books_catalog.xlsx"
        
        # Сохраняем в Excel без вывода индексов строк
        df.to_excel(output_file, index=False)
        logging.info(f"🎉 Сбор успешно завершен! Всего собрано книг со всех страниц: {len(all_books)}")
        logging.info(f"Данные сохранены в файл: {output_file}")
    else:
        logging.warning("Данные не были собраны. Проверьте структуру HTML-кода сайта.")
        
if __name__ == "__main__":
    parse_books()

