"""
Логирование для проекта Wildberries
Переиспользовано из Ozon проекта
"""

from loguru import logger
from pathlib import Path
import sys


def setup_logger(logs_dir: Path = Path("logs"), level: str = "INFO"):
    """
    Настройка логирования
    
    Args:
        logs_dir: Папка для сохранения логов
        level: Уровень логирования
    """
    logs_dir.mkdir(parents=True, exist_ok=True)
    
    # Удаляем стандартный обработчик
    logger.remove()
    
    # Добавляем обработчик для консоли
    logger.add(
        sys.stdout,
        format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>",
        level=level,
        colorize=True
    )
    
    # Добавляем обработчик для файла
    logger.add(
        logs_dir / "app_{time:YYYY-MM-DD}.log",
        format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
        level=level,
        rotation="00:00",  # Новая лог-файл каждый день в полночь
        retention="30 days",  # Хранить логи 30 дней
        compression="zip",  # Сжимать старые логи
        encoding="utf-8"
    )
    
    # Добавляем отдельный файл для ошибок
    logger.add(
        logs_dir / "errors_{time:YYYY-MM-DD}.log",
        format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
        level="ERROR",
        rotation="00:00",
        retention="90 days",
        compression="zip",
        encoding="utf-8"
    )
    
    return logger
