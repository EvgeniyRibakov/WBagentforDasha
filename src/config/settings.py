"""
Настройки проекта для работы с Wildberries Seller
Адаптировано под Wildberries на основе Ozon проекта
"""

from pydantic_settings import BaseSettings
from pathlib import Path
from typing import Optional


class Settings(BaseSettings):
    """Настройки приложения"""
    
    # Wildberries Seller
    wildberries_start_url: str = "https://seller.wildberries.ru/analytics-reports/sales"
    # Данные для авторизации (если сессия истекла)
    phone_number: Optional[str] = None  # Телефон для авторизации (опционально)
    email: Optional[str] = None  # Email для авторизации (опционально)
    password: Optional[str] = None  # Пароль для авторизации (опционально)
    
    # Задержки (увеличены для Wildberries)
    delay_before_click: float = 1.5
    delay_after_click: float = 1.5
    delay_before_type: float = 1.0
    delay_after_type: float = 1.5
    delay_between_keys: float = 0.12
    delay_page_load: float = 4.0
    delay_between_actions: float = 1.5  # Задержка между действиями
    
    # Профиль Chrome (ОБЯЗАТЕЛЬНО - содержит сохранённую авторизацию)
    chrome_user_data_dir: Optional[str] = None
    chrome_profile_name: str = "Profile 2"
    
    # Папки
    downloads_dir: Path = Path("downloads")
    logs_dir: Path = Path("logs")
    pages_code_dir: Path = Path("pages_code")
    
    # Telegram бот (опционально)
    telegram_bot_token: Optional[str] = None
    telegram_bot_password: Optional[str] = None
    
    # Google Sheets (опционально)
    upload_to_google_sheets: bool = False
    google_sheets_url: Optional[str] = None
    google_sheets_credentials_path: Optional[str] = None
    
    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"
        extra = "ignore"  # Игнорировать дополнительные поля из .env


