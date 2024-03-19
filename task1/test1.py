from contextlib import contextmanager
from typing import Generator

import xlwings as xw
from xlwings.main import Book


# Создаем контекстный менеджер для работы с книгой Excel
@contextmanager
def open_xl_book(path: str) -> Generator[Book, None, None]:
    """
    Контекстный менеджер для открытия и автоматического закрытия книги Excel.

    :param path: Путь к файлу Excel.
    :yield: Объект книги Excel, с которым можно взаимодействовать.
    """
    book = xw.Book(path)  # Открываем книгу по указанному пути
    try:
        yield book
    finally:
        book.close()  # Закрываем книгу по завершению работы в контексте


# Путь к файлу
path = 'ts1.xlsx'

# Определяем цвета
done_color = xw.utils.rgb_to_int((0, 255, 0))  # Зеленый
progress_color = xw.utils.rgb_to_int((255, 0, 0))  # Красный

# Использование контекстного менеджера для работы с книгой Excel
with open_xl_book(path) as wb:
    sheet = wb.sheets['Sheet1']  # Выбираем лист по имени
    try:
        table = sheet.tables[0]  # Предполагаем, что таблица - первый объект на листе
        status_col_index = None

        # Ищем столбец "Status" в заголовках таблицы
        for col in range(table.range.columns.count):
            if table.range(1, col + 1).value == "Status":
                status_col_index = col + 1
                break

        # Если столбец "Status" найден
        if status_col_index:
            # Получаем диапазон данных в столбце "Status", исключая заголовок
            data_range = table.range.offset(1, 0).resize(table.range.rows.count - 1, table.range.columns.count)

            # Проходимся по строкам в диапазоне данных
            for row in data_range.rows:
                status = row(status_col_index).value
                if status == "Done":
                    row.color = done_color
                elif status == "In progress":
                    row.color = progress_color

            # Сохраняем изменения в файле
            wb.save(path)
        else:
            print('Столбец "Status" не найден в таблице.')
    except IndexError:
        print('Таблица на листе не найдена.')
