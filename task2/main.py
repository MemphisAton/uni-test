

import os
from typing import NoReturn

from config import load_config, Config
from email_sender import send_email
from excel_utils import read_themes_from_xlwings_with_table_search, insert_links_into_excel
from web_scraping import search_themes_and_get_links

# Загрузка конфигурации из файла .env
config: Config = load_config('.env')


def main() -> NoReturn:
    """
    Главная функция, координирующая чтение тем из Excel-файла, поиск ссылок по этим темам,
    обновление файла с найденными ссылками и отправку файла по электронной почте.
    """
    # Получение пути к файлу из конфигурации
    path_to_file = config.db.path_to_file

    # Проверка существования файла
    if not os.path.exists(path_to_file):
        print(f"Файл {path_to_file} не найден.")
        return

    print(f"1/5 Файл {path_to_file} найден.")

    # Чтение тем из Excel-файла
    themes = read_themes_from_xlwings_with_table_search(path_to_file)

    if themes:
        # Поиск ссылок по темам
        links_for_themes = search_themes_and_get_links(themes, config.db.driver_path)

        # Вставка найденных ссылок в Excel-файл
        insert_links_into_excel(path_to_file, links_for_themes)

        try:
            # Отправка файла по электронной почте
            send_email(config.db.subject, config.db.body, path_to_file)
            print("5/5 Отправка файла выполнена успешно.")
        except Exception as e:
            print(f"Произошла ошибка при отправке электронной почты: {e}")


if __name__ == "__main__":
    main()
