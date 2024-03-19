import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from config import load_config, Config
config: Config = load_config('.env')


def send_email(subject: str, message: str, attachment_path) -> bool:
    """
    Функция отправляет сообщение с одной почты на другую с возможностью добавления вложения.
    :param subject: тема письма
    :param message: тело письма
    :param attachment_path: путь к файлу вложения (опционально)
    :return: булевый эквивалент итога отправки
    """
    sender_email = config.db.sender_email
    sender_password = config.db.sender_password
    recipient_email = config.db.recipient_email

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
        return True
    except Exception as error:
        print("Error sending email:", str(error))
        return False
