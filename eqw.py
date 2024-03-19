from contextlib import contextmanager

import xlwings as xw


@contextmanager
def open_xl_book(path):
    app = xw.App(visible=False)
    book = xw.Book(path)
    try:
        yield book
    finally:
        book.close()
        app.quit()


def read_themes_from_xlwings_with_table_search(path):
    try:
        with open_xl_book(path) as book:
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

            return themes
    except FileNotFoundError:
        print("Файл не найден. Проверьте путь к файлу.")
        return []
    except Exception as e:
        print(f"Произошла ошибка: {e}")
        return []


path = 'ts2.xlsx'

themes = read_themes_from_xlwings_with_table_search(path)
print(themes)

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options


def search_themes_and_get_links(themes, driver_path='chromedriver.exe'):
    chrome_options = Options()
    # Задаем User-Agent
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")

    service = Service(executable_path=driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    themes_links = {}  # Словарь для хранения тем и соответствующих ссылок

    for theme in themes:
        driver.get('https://ya.ru')

        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, 'text'))
        )
        search_box.send_keys(theme)
        search_box.send_keys(Keys.ENTER)

        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'li.serp-item div.organic a.link'))
        )

        links_elements = driver.find_elements(By.CSS_SELECTOR, 'li.serp-item div.organic a.link')

        # Собираем первые три уникальные ссылки в список
        unique_links = []
        for element in links_elements:
            if len(unique_links) < 3:
                link = element.get_attribute('href')
                if link not in unique_links:  # Проверка на уникальность
                    unique_links.append(link)

            if len(unique_links) == 3:
                break

        themes_links[theme] = unique_links

    driver.quit()
    return themes_links


# Вызов функции и получение ссылок
links_for_themes = search_themes_and_get_links(themes)
print(links_for_themes)


# links_for_themes = {'Искусственный интеллект': ['https://www.nur.kz/family/school/1817736-iskusstvennyj-intellekt-sovremennye-vozmoznosti-i-perspektivy/', 'https://medium.com/nuances-of-programming/%D0%B8%D1%81%D0%BA%D1%83%D1%81%D1%81%D1%82%D0%B2%D0%B5%D0%BD%D0%BD%D1%8B%D0%B9-%D0%B8%D0%BD%D1%82%D0%B5%D0%BB%D0%BB%D0%B5%D0%BA%D1%82-%D0%B8-%D0%BD%D0%B0%D1%88%D0%B5-%D0%B1%D1%83%D0%B4%D1%83%D1%89%D0%B5%D0%B5-a5eacc7a9a41', 'https://ru.wikipedia.org/wiki/%D0%98%D1%81%D0%BA%D1%83%D1%81%D1%81%D1%82%D0%B2%D0%B5%D0%BD%D0%BD%D1%8B%D0%B9_%D0%B8%D0%BD%D1%82%D0%B5%D0%BB%D0%BB%D0%B5%D0%BA%D1%82'], 'Глобальное потепление': ['https://medium.com/%D0%B0%D0%BD%D0%B0%D0%BB%D1%96%D1%82%D0%B8%D1%87%D0%BD%D0%B8%D0%B9-%D1%86%D0%B5%D0%BD%D1%82%D1%80-%D1%81%D1%82%D1%80%D0%B0%D1%82%D0%B5%D0%B3%D1%96%D1%8F/%D0%B3%D0%BB%D0%BE%D0%B1%D0%B0%D0%BB%D1%8C%D0%BD%D0%BE%D0%B5-%D0%BF%D0%BE%D1%82%D0%B5%D0%BF%D0%BB%D0%B5%D0%BD%D0%B8%D0%B5-%D1%87%D1%82%D0%BE-%D1%8D%D1%82%D0%BE-%D0%B8-%D0%BF%D0%BE%D1%87%D0%B5%D0%BC%D1%83-%D0%BE-%D0%BD%D0%B5%D0%BC-%D1%82%D0%B0%D0%BA-%D0%BC%D0%BD%D0%BE%D0%B3%D0%BE-%D0%B3%D0%BE%D0%B2%D0%BE%D1%80%D1%8F%D1%82-da820bef00c8', 'https://ru.wikipedia.org/wiki/%D0%93%D0%BB%D0%BE%D0%B1%D0%B0%D0%BB%D1%8C%D0%BD%D0%BE%D0%B5_%D0%BF%D0%BE%D1%82%D0%B5%D0%BF%D0%BB%D0%B5%D0%BD%D0%B8%D0%B5', 'https://ru.wikipedia.org/wiki/%D0%93%D0%BB%D0%BE%D0%B1%D0%B0%D0%BB%D1%8C%D0%BD%D0%BE%D0%B5_%D0%BF%D0%BE%D1%82%D0%B5%D0%BF%D0%BB%D0%B5%D0%BD%D0%B8%D0%B5#%D0%9E%D0%B1%D1%89%D0%B8%D0%B5_%D1%81%D0%B2%D0%B5%D0%B4%D0%B5%D0%BD%D0%B8%D1%8F'], 'Квантовые компьютеры': ['https://ru.wikipedia.org/wiki/%D0%9A%D0%B2%D0%B0%D0%BD%D1%82%D0%BE%D0%B2%D1%8B%D0%B9_%D0%BA%D0%BE%D0%BC%D0%BF%D1%8C%D1%8E%D1%82%D0%B5%D1%80', 'https://ru.wikipedia.org/wiki/%D0%9A%D0%B2%D0%B0%D0%BD%D1%82%D0%BE%D0%B2%D1%8B%D0%B9_%D0%BA%D0%BE%D0%BC%D0%BF%D1%8C%D1%8E%D1%82%D0%B5%D1%80#%D0%92%D0%B2%D0%B5%D0%B4%D0%B5%D0%BD%D0%B8%D0%B5', 'https://ru.wikipedia.org/wiki/%D0%9A%D0%B2%D0%B0%D0%BD%D1%82%D0%BE%D0%B2%D1%8B%D0%B9_%D0%BA%D0%BE%D0%BC%D0%BF%D1%8C%D1%8E%D1%82%D0%B5%D1%80#%D0%A2%D0%B5%D0%BE%D1%80%D0%B8%D1%8F'], 'Электромобили': ['https://yabs.yandex.ru/count/WWOejI_zOoVX2Ld60SKF01DTRYOQbKgbKga4mGHzFfSxUxRVkVE6Er-_u_M6ErmRHoEIaAX1GrGYC6MJHajKePCVIcyZGen2CmW67KS6eC8Aq3801LQ0bW6ehmAq343j1MWNw16b0Eq2FUuAq4eRdNA-9XwxeY-k7ksSxf6uoP4FOV_1wprmMVJurD3SlgFGvrGAlMzy2dKwx-SzhQFIsyQTUGwDvbAdBRLNu61n_DZheY9s9ZkxyMe9g1-b6D1176JEnbi1a7m06XfPE58DF4IhOM-gWOTeDMpfgPfoJLWgN2Dsm4dNzHfz0w4NnRuXK3x8D5m6eC82s7girbj403BV71vK4AyF2rjcjib3ZyAsmjDNNGovgMF1J3fsX_bck761mX2SZGYOhW15R4G7a8m5QEhv8UNcjm8WkScj04XkxXEoLJ-imqJxmHIW7FbVvflPzpBVp3xtaodvUm9MWryRm90aFnl0a2HCCnFlZ5bcufqnnnU-h6mfuH1r9qdljDcNFvprGNgoprEoeysFQGy-m8mB4YZ_hQ4aK4_pMrhNTKWbWNgQrtgwDN8YdSzwh_ipD2UjRhRHo-g8fgtcB4wUyXlmmmMNn6oD8nJyX0aYS2276fJZBfe9Q6oG5dD92T0KQpVFzUeBDpW-EbBNwBFJuNqqb9m45Lk2GocB8grT40mlMfdbK000~2?etext=2202.xLedaxAv2ihW1cDAnu0VhIaXWp_qMk0daScACoctlpNubGRib3N6eW9neGZmcnpl.32bfea6507a8d2bcbd6482b22a550707e11d34fa&from=ya.ru%3Bsearch%26%23x2F%3B%3Bweb%3B%3B0%3B&q=%D1%8D%D0%BB%D0%B5%D0%BA%D1%82%D1%80%D0%BE%D0%BC%D0%BE%D0%B1%D0%B8%D0%BB%D0%B8', 'https://yabs.yandex.ru/count/WXiejI_zOoVX2LdP0MqG05EUSoOQbKgbKga4mGHzFfSxUxRVkVE6Er-_u_M6EzmX8gKP8suS6I7LiDOuX6eRH40aR1oDI4AY1WrHGUo0APiqMg8AdVnGUXiHOXIQGJ3eE3805LQ0bW4ei0Aq341z5Q1b06elGBj0ZoW5Q1VeSLU0LjhebFCryTWLVN7rQETqZyHDZdm8-mzUxu7BeiUdXkRs7OK-frBeVUDJgDDvF-zf7PNUDkxCSsWqbpfjgRq21ulZnruN5R4psTcDLqj0_IZ5W0xY87CstWg0v0FGqCZ2aMhW8LeDUrCDF4IhOKjFrPHhmbBX6h43Jhgkr-WR2BqezWs1ya6cuJ805HR0rcErtY82aFdcyA22U7rOs3ApJHvw5BOLdhxgOiXD7Gjcqh4xp3V3ZGiKXk1iHC1q1IXY8pg0P2n0KyyFAJU_5G39JMu5G79pdv2j-c4T9jeFfW3boFyotyo-b_bczhcVJCdV4x0Q_De0XYJvsG26966QcNXdp38Jxumvll1bPKi9XwWxINgdpRxyuQmFqfDzdf8TRNvSmOCFCEE21Cf_QoW9rDFybjPr7L99e9xcjPvkZPn8vxFUw_uCpKchcstqiZgYQQivY_D4-IMHC4ebPciS0HJyH4M80ZdOe1ztKQP4OwUq64HAmAPVktdkLryumF7HZRf6dvqExAEZv2IesG8SIbaKQUs2O7ZHoYo70W00~2?etext=2202.xLedaxAv2ihW1cDAnu0VhIaXWp_qMk0daScACoctlpNubGRib3N6eW9neGZmcnpl.32bfea6507a8d2bcbd6482b22a550707e11d34fa&from=ya.ru%3Bsearch%26%23x2F%3B%3Bweb%3B%3B0%3B&q=%D1%8D%D0%BB%D0%B5%D0%BA%D1%82%D1%80%D0%BE%D0%BC%D0%BE%D0%B1%D0%B8%D0%BB%D0%B8', 'https://yabs.yandex.ru/count/WYaejI_zOoVX2LdW0TqG08FVToOQbKgbKga4mGHzFfSxUxRVkVE6Er-_u_M6EzmX8gKP8suS6I7LiDOuXEeYgmogZIGQGqOrIDWu6f64H0qQeehO0LCsQRH45JhveVGs8iGeD8DWq75a02gi02q3K605Q1c0-Yf0om3KNe1sWHvH2j0kqEEk0AsqqIddQ-AnA_hYwj7EwHw9cnpv4FOVlDu3bqMFJmtDxJiAVKwbqFl6fr2dytxUqpeglMtScUVGQ2vrsb9x1GuMnu-zBYfYPxAp6w-MWFfHYW4Tn47cRBmL0Ca7eA6HXIFLm4Eq6lQc6dY8LiEMdgefruIbmZLY1vnqNQ_HDn1wKUmR0kM3JCDb02eiWAt7Qhr5123ppU511V7wiB1bPfiyz2XiApnzrSMGcpeMpAHZTvXlXXiNA0p1sOY0wGfGn4Pq0CbOWAQU7r9kVYi0aflS2e3avZuXM_N3Eaoq7qm1of7_PRwPVI_ppUnpFvcIloTWDVYr0Gn9yhC134d2D3FnpfXb9juPStpXoygM4WvHTvBqJfjz-SDP7wGd-pmbEzhy68C77s361GcK_zPG4gYd-Isjwpgaaa0zpMiztHevaSvdlTVz6PgJLZVRw6LrH3DpfLAfj9gpEBRNTHq3JLR-y0aA94ZCrZW24ln4HOW2ETYW7tTHfaHZfxGOH4h0fb-xUUvNNpZ0yT6DkaQVdGxiewFa9AZP0XnAMHJY8YjP2AeB~2?etext=2202.xLedaxAv2ihW1cDAnu0VhIaXWp_qMk0daScACoctlpNubGRib3N6eW9neGZmcnpl.32bfea6507a8d2bcbd6482b22a550707e11d34fa&from=ya.ru%3Bsearch%26%23x2F%3B%3Bweb%3B%3B0%3B&q=%D1%8D%D0%BB%D0%B5%D0%BA%D1%82%D1%80%D0%BE%D0%BC%D0%BE%D0%B1%D0%B8%D0%BB%D0%B8'], 'Космический туризм': ['https://medium.com/space-review/%D0%BA%D0%B0%D0%BA-%D1%81%D1%82%D0%B0%D1%82%D1%8C-%D0%BA%D0%BE%D1%81%D0%BC%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%B8%D0%BC-%D1%82%D1%83%D1%80%D0%B8%D1%81%D1%82%D0%BE%D0%BC-e1c225bc67bd', 'https://ru.wikipedia.org/wiki/%D0%9A%D0%BE%D1%81%D0%BC%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%B8%D0%B9_%D1%82%D1%83%D1%80%D0%B8%D0%B7%D0%BC', 'https://trends.rbc.ru/trends/industry/5f22cf589a794765d3c449b9'], 'Виртуальная реальность': ['https://medium.com/vision-dti/%D0%B2%D0%B8%D1%80%D1%82%D1%83%D0%B0%D0%BB%D1%8C%D0%BD%D0%B0%D1%8F-%D1%80%D0%B5%D0%B0%D0%BB%D1%8C%D0%BD%D0%BE%D1%81%D1%82%D1%8C-eec4c277110', 'https://ru.wikipedia.org/wiki/%D0%92%D0%B8%D1%80%D1%82%D1%83%D0%B0%D0%BB%D1%8C%D0%BD%D0%B0%D1%8F_%D1%80%D0%B5%D0%B0%D0%BB%D1%8C%D0%BD%D0%BE%D1%81%D1%82%D1%8C', 'https://ru.wikipedia.org/wiki/%D0%92%D0%B8%D1%80%D1%82%D1%83%D0%B0%D0%BB%D1%8C%D0%BD%D0%B0%D1%8F_%D1%80%D0%B5%D0%B0%D0%BB%D1%8C%D0%BD%D0%BE%D1%81%D1%82%D1%8C#%D0%A0%D0%B5%D0%B0%D0%BB%D0%B8%D0%B7%D0%B0%D1%86%D0%B8%D1%8F'], 'Блокчейн технологии': ['https://medium.com/chainlink-community/%D1%87%D1%82%D0%BE-%D1%82%D0%B0%D0%BA%D0%BE%D0%B5-%D1%82%D0%B5%D1%85%D0%BD%D0%BE%D0%BB%D0%BE%D0%B3%D0%B8%D1%8F-%D0%B1%D0%BB%D0%BE%D0%BA%D1%87%D0%B5%D0%B9%D0%BD-91c23c89c108', 'https://practicum.yandex.ru/blog/chto-takoe-blokchain-i-kak-eto-rabotaet/', 'https://habr.com/ru/companies/iticapital/articles/340992/'], 'Генная инженерия': ['https://trends.rbc.ru/trends/futurology/612f77ad9a7947ce386b68ba', 'https://ru.wikipedia.org/wiki/%D0%93%D0%B5%D0%BD%D0%B5%D1%82%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%B0%D1%8F_%D0%B8%D0%BD%D0%B6%D0%B5%D0%BD%D0%B5%D1%80%D0%B8%D1%8F', 'https://www.youtube.com/watch?v=W6gpicnWXXg'], 'Умные города': ['https://shakirhodler.medium.com/%D0%BA%D0%BE%D0%BD%D1%86%D0%B5%D0%BF%D1%86%D0%B8%D1%8F-smart-city-77277e4a742a', 'https://te-st.org/2015/08/05/social-theory-of-the-smart-city/', 'https://ru.wikipedia.org/wiki/%D0%A3%D0%BC%D0%BD%D1%8B%D0%B9_%D0%B3%D0%BE%D1%80%D0%BE%D0%B4'], 'Возобновляемые источники энергии': ['https://www.un.org/ru/climatechange/what-is-renewable-energy', 'https://ru.wikipedia.org/wiki/%D0%92%D0%BE%D0%B7%D0%BE%D0%B1%D0%BD%D0%BE%D0%B2%D0%BB%D1%8F%D0%B5%D0%BC%D0%B0%D1%8F_%D1%8D%D0%BD%D0%B5%D1%80%D0%B3%D0%B8%D1%8F', 'https://admiralmarkets.com/ru/education/articles/shares/vozobnovljaemaja-energia']}


def insert_links_into_excel(path, themes_links):
    try:
        with open_xl_book(path) as book:
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

            # Применение автофильтра
            if row > 2:  # Фильтр возможен только если есть более одной строки данных
                sheet.range(full_range).api.AutoFilter(Field=1)

            book.save()
    except Exception as e:
        print(f"Произошла ошибка: {e}")


insert_links_into_excel(path, links_for_themes)

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


# import config  # Предполагается, что у вас есть модуль config с данными для входа


def send_email(subject: str, message: str, attachment_path: str = path) -> bool:
    """
    Функция отправляет сообщение с одной почты на другую с возможностью добавления вложения.
    :param subject: тема письма
    :param message: тело письма
    :param attachment_path: путь к файлу вложения (опционально)
    :return: булевый эквивалент итога отправки
    """
    sender_email = 'testmem.email@yandex.ru'
    sender_password = 'izlnpvwjwnaugcyi'
    recipient_email = input('Введите почту для отправки файла:   ')  # 'memphisaton@gmail.com'

    # Создаем экземпляр MIMEMultipart
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email

    # Прикрепляем текст письма
    msg.attach(MIMEText(message, 'plain'))

    attachment_filename = "ts.xlsx"  # Указываем имя файла для вложения
    if attachment_path is None:
        attachment_path = attachment_filename  # Если путь к файлу не предоставлен, используем имя файла

    try:
        with open(attachment_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            f'attachment; filename="{attachment_filename}"')  # Используем только имя файла
            msg.attach(part)
    except IOError:
        print(f'Could not read attachment {attachment_path}.')
        return False

    try:
        server = smtplib.SMTP('smtp.yandex.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        print('Файл отправлен')
        return True
    except Exception as error:
        print("Error sending email:", str(error))
        return False


send_email('Список тем', 'Файл во вложении')