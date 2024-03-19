from dataclasses import dataclass
from typing import Optional

from environs import Env


@dataclass
class DatabaseConfig:
    sender_email: str  # Почта с которой отправляются репорты
    sender_password: str  # Пароль от почты
    recipient_email: str  # Почта куда отправлять
    path_to_file: str  # путь до отправляемого файла
    subject: str  # тема письма
    body: str  # тело письма
    driver_path: str   # путь к драйверам поисковика
@dataclass
class Config:
    db: DatabaseConfig


def load_config(path: Optional[str]) -> Config:
    """

    :rtype: object
    """
    env: Env = Env()  # Создаем экземпляр класса Env
    env.read_env(path)  # Добавляем в переменные окружения данные, прочитанные из файла .env
    return Config(db=DatabaseConfig(sender_email=env('sender_email'),
                                    sender_password=env('sender_password'),
                                    recipient_email=env('recipient_email'),
                                    subject=env('subject'),
                                    body=env('body'),
                                    driver_path=env('driver_path'),
                                    path_to_file=env('path_to_file')))
