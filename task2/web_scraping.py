from typing import List, Dict
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


def search_themes_and_get_links(themes: List[str], driver_path: str) -> Dict[str, List[str]]:
    """
    Поиск в интернете по заданным темам и получение первых трех уникальных ссылок для каждой темы.

    Параметры:
        themes (List[str]): Список строк, представляющих темы для поиска.
        driver_path (str): Путь к исполняемому файлу драйвера браузера Chrome.

    Возвращает:
        Dict[str, List[str]]: Словарь, где ключи - это темы поиска, а значения - списки с первыми тремя уникальными ссылками.
    """

    # Настройка опций для Chrome, включая имитацию User-Agent для имитации запросов от реального пользователя.
    chrome_options = Options()
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")

    # Инициализация драйвера с указанными опциями и путем к исполняемому файлу драйвера.
    service = Service(executable_path=driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    # Словарь для хранения тем и соответствующих ссылок.
    themes_links: Dict[str, List[str]] = {}

    for theme in themes:
        # Переход на страницу поиска.
        driver.get('https://ya.ru')

        # Ожидание появления поля поиска и ввод темы запроса.
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, 'text'))
        )
        search_box.send_keys(theme)
        search_box.send_keys(Keys.ENTER)

        # Ожидание загрузки результатов поиска.
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'li.serp-item div.organic a.link'))
        )

        # Получение элементов со ссылками в результатах поиска.
        links_elements = driver.find_elements(By.CSS_SELECTOR, 'li.serp-item div.organic a.link')

        # Сбор первых трех уникальных ссылок.
        unique_links: List[str] = []
        for element in links_elements:
            if len(unique_links) < 3:
                link = element.get_attribute('href')
                if link not in unique_links:  # Проверка на уникальность ссылки.
                    unique_links.append(link)

                if len(unique_links) == 3:  # Прекращение сбора после нахождения трех уникальных ссылок.
                    break

        themes_links[theme] = unique_links

    # Завершение работы драйвера и закрытие браузера.
    driver.quit()
    # Информативное сообщение о завершении процесса.
    print('3/5 Необходимые ссылки получены')
    return themes_links
