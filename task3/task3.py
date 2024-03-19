from typing import Set

import matplotlib.pyplot as plt
import pandas as pd
import regex

# Путь к файлу с текстом
file_path: str = 'ts.txt'


def is_palindrome(word: str) -> bool:
    """
    Проверяет, является ли слово палиндромом.

    :param word: Строка, слово для проверки.
    :return: Возвращает True, если слово является палиндромом, иначе False.
    """
    clean_word = ''.join(char for char in word if char.isalnum()).lower()
    return clean_word == clean_word[::-1]


# Чтение файла и поиск палиндромов
with open(file_path, 'r', encoding='utf-8') as file:
    text = file.read().lower()

# Используем regex для разделения текста на слова
words: Set[str] = regex.findall(r'\w+', text)

# Находим уникальные палиндромы
unique_palindromes: Set[str] = {word for word in words if is_palindrome(word)}

# Подсчитываем уникальные палиндромы по количеству символов
palindrome_count: dict = {}
for palindrome in unique_palindromes:
    length = len(palindrome)
    palindrome_count[length] = palindrome_count.get(length, 0) + 1

# Создаем DataFrame из результатов
palindrome_table: pd.DataFrame = pd.DataFrame({
    'Количество символов': list(palindrome_count.keys()),
    'Количество слов': list(palindrome_count.values())
}).sort_values(by='Количество символов')

# Визуализация результатов круговой диаграммой
fig, ax = plt.subplots(figsize=(12, 8))
ax.pie(palindrome_table['Количество слов'], labels=palindrome_table['Количество символов'], autopct='%1.1f%%',
       startangle=90)
ax.axis('equal')  # Гарантирует, что круговая диаграмма будет нарисована как круг.
plt.title('Распределение палиндромов по количеству символов')

# Добавляем таблицу с результатами под диаграммой
columns = ['Количество символов', 'Количество слов']
table = plt.table(cellText=palindrome_table.values, colLabels=columns, loc='bottom', cellLoc='center',
                  bbox=[0.25, -0.5, 0.5, 0.3])
plt.subplots_adjust(left=0.2, bottom=0.3)

plt.show()
