from contextlib import contextmanager
from typing import List, Dict, Iterator

import xlwings as xw


@contextmanager
def open_xl_book(path_to_file: str) -> Iterator[xw.main.Book]:
    """
    Контекстный менеджер для открытия книги Excel.

    Аргументы:
        path_to_file (str): Путь к файлу Excel, который нужно открыть.

    Возвращает:
        Iterator[xw.main.Book]: Открытая книга.

    Вызывает:
        Exception: Передает любое исключение, возникшее в контексте, после обязательного закрытия приложения Excel.
    """
    app = xw.App(visible=False)
    try:
        book = app.books.open(path_to_file)
        yield book
    except Exception as e:
        print(f"Не удалось открыть книгу: {e}")
        app.quit()
        raise
    finally:
        book.close()
        app.quit()


def read_themes_from_xlwings_with_table_search(path_to_file: str) -> List[str]:
    """
    Читает темы из листа Excel, начиная с ячейки A2 вниз.

    Аргументы:
        path_to_file (str): Путь к рабочей книге Excel.

    Возвращает:
        List[str]: Список тем, прочитанных из листа Excel. Возвращает пустой список, если темы не найдены.
    """
    try:
        with open_xl_book(path_to_file) as book:
            sheet = book.sheets['Sheet1']
            # Начинаем с ячейки 'A2', так как предполагается, что 'A1' - это заголовок
            first_cell = sheet.range('A2')
            # Идём вниз до первой пустой ячейки
            last_cell = first_cell.end('down')
            # Если в A2 уже пусто, значит нет тем для чтения
            if last_cell == first_cell and first_cell.value is None:
                return []
            # Собираем все темы в диапазоне от A2 до последней заполненной ячейки
            themes_range = sheet.range(f'A2:A{last_cell.row}')
            themes = [cell.value for cell in themes_range if cell.value is not None]
            print('2/5 Список тем сформирован')
            return themes
    except FileNotFoundError:
        print("Файл не найден. Проверьте путь к файлу.")
        return []
    except Exception as e:
        print(f"Произошла ошибка: {e.__class__.__name__}: {e}")
        return []


def insert_links_into_excel(path_to_file: str, themes_links: Dict[str, List[str]]) -> None:
    """
    Вставляет темы и соответствующие ссылки в лист Excel.

    Аргументы:
        path_to_file (str): Путь к рабочей книге Excel.
        themes_links (Dict[str, List[str]]): Словарь, где ключи - это темы, а значения - списки ссылок, ассоциированных с каждой темой.
    """
    try:
        with open_xl_book(path_to_file) as book:
            sheet = book.sheets['Sheet1']
            # Первый столбец с данными это "Theme", второй "Sources"
            theme_column_letter = 'A'
            sources_column_letter = 'B'
            # Найдем все существующие темы и их последний ряд
            existing_themes = sheet.range(
                f"{theme_column_letter}2:{theme_column_letter}{sheet.cells.last_cell.row}").value
            last_row = len(existing_themes) + 1 if existing_themes else 1
            # Очищаем существующие данные (исключая заголовки)
            if last_row > 1:
                sheet.range(f"{theme_column_letter}2:{sources_column_letter}{last_row}").clear_contents()
            # Заполнение данными начиная со второй строки
            row = 2
            for theme, links in themes_links.items():
                for link in links:
                    sheet.range(f"{theme_column_letter}{row}").value = theme
                    cell = sheet.range(f"{sources_column_letter}{row}")
                    cell.value = link
                    cell.api.WrapText = True
                    row += 1
            # Определение диапазона для автофильтра
            full_range = f"{theme_column_letter}1:{sources_column_letter}{row - 1}"
            if row > 2:  # Применение автофильтра
                sheet.range(full_range).api.AutoFilter(Field=1)
            print('4/5 Файл подготовлен к отправке')
            book.save()
    except Exception as e:
        print(f"Произошла ошибка: {e}")
