"""Настройка логирования."""
import sys
from pathlib import Path
from datetime import datetime

from loguru import logger


def setup_logger(logs_dir: Path) -> None:
    """Настройка логирования.

    Args:
        logs_dir: Путь к папке для логов
    """
    # Создаём папку для логов, если её нет
    logs_dir.mkdir(parents=True, exist_ok=True)

    # Формат даты для имени файла
    date_str = datetime.now().strftime("%Y-%m-%d")

    # Удаляем стандартный обработчик
    logger.remove()

    # Добавляем обработчик для консоли
    logger.add(
        sys.stdout,
        format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>",
        level="INFO",
        colorize=True,
    )

    # Добавляем обработчик для файла с общими логами
    logger.add(
        logs_dir / f"app_{date_str}.log",
        format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
        level="DEBUG",
        rotation="00:00",
        retention="30 days",
        encoding="utf-8",
    )

    # Добавляем обработчик для файла с ошибками
    logger.add(
        logs_dir / f"errors_{date_str}.log",
        format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
        level="ERROR",
        rotation="00:00",
        retention="30 days",
        encoding="utf-8",
    )
