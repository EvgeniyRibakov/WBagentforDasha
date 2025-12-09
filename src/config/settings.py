"""Настройки приложения."""
import os
from pathlib import Path
from typing import Optional

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Настройки приложения."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )

    # URL страницы отчётов Wildberries
    wildberries_start_url: str = Field(
        default="https://seller.wildberries.ru/analytics-reports/sales",
        description="URL страницы отчётов о продажах",
    )

    # Данные для авторизации
    phone_number: Optional[str] = Field(
        default=None,
        description="Номер телефона для авторизации",
    )

    # Профиль Yandex Browser
    yandex_browser_path: Optional[str] = Field(
        default=None,
        description="Путь к исполняемому файлу Yandex Browser",
    )
    yandex_user_data_dir: Optional[str] = Field(
        default=None,
        description="Путь к папке User Data Yandex Browser",
    )
    yandex_profile_name: Optional[str] = Field(
        default=None,
        description="Имя профиля Yandex Browser (например, 'Default' или 'Profile 1')",
    )
    yandex_browser_version: Optional[int] = Field(
        default=None,
        description="Версия Yandex Browser (например, 138). Если не указана, будет определена автоматически",
    )

    # Задержки между действиями (в секундах)
    delay_before_click: float = Field(default=1.5, description="Задержка перед кликом")
    delay_after_click: float = Field(default=1.5, description="Задержка после клика")
    delay_before_type: float = Field(default=1.0, description="Задержка перед вводом текста")
    delay_after_type: float = Field(default=1.5, description="Задержка после ввода текста")
    delay_between_keys: float = Field(default=0.12, description="Задержка между символами")
    delay_page_load: float = Field(default=4.0, description="Задержка загрузки страницы")
    delay_between_actions: float = Field(default=1.5, description="Задержка между действиями")

    # Папки проекта
    downloads_dir: str = Field(default="downloads", description="Папка для скачанных файлов")
    logs_dir: str = Field(default="logs", description="Папка для логов")
    data_dir: str = Field(default="data", description="Папка для обработанных данных")

    # Путь к файлу с примером первой строки
    example_first_stroke_file: str = Field(
        default="example_first_stroke.XLSX",
        description="Путь к файлу с примером первой строки для замены",
    )

    # Таймауты ожидания элементов (в секундах)
    element_wait_timeout: int = Field(default=20, description="Таймаут ожидания элемента")

    @property
    def downloads_path(self) -> Path:
        """Возвращает путь к папке downloads."""
        return Path(self.downloads_dir).resolve()

    @property
    def logs_path(self) -> Path:
        """Возвращает путь к папке logs."""
        return Path(self.logs_dir).resolve()

    @property
    def data_path(self) -> Path:
        """Возвращает путь к папке data."""
        return Path(self.data_dir).resolve()

    @property
    def example_first_stroke_path(self) -> Path:
        """Возвращает путь к файлу с примером первой строки."""
        return Path(self.example_first_stroke_file).resolve()
