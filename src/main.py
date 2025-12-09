"""Точка входа в приложение."""
import sys
from pathlib import Path

from loguru import logger

from src.config import Settings
from src.utils import setup_logger
from src.agents import BrowserAgent


def main() -> int:
    """Основная функция приложения.

    Returns:
        Код возврата (0 - успех, 1 - ошибка)
    """
    try:
        # Загрузка настроек
        settings = Settings()

        # Настройка логирования
        setup_logger(settings.logs_path)
        logger.info("=" * 60)
        logger.info("Запуск WBagentforDasha")
        logger.info("=" * 60)

        # Проверка наличия файла с примером первой строки
        if not settings.example_first_stroke_path.exists():
            logger.error(f"Файл с примером первой строки не найден: {settings.example_first_stroke_path}")
            logger.error("Убедитесь, что файл example_first_stroke.XLSX существует в корне проекта")
            return 1

        # Информация о Yandex Browser
        logger.info("Используется Yandex Browser")
        if settings.yandex_browser_path:
            logger.info(f"Путь к браузеру из .env: {settings.yandex_browser_path}")
        if settings.yandex_user_data_dir:
            logger.info(f"Путь к User Data из .env: {settings.yandex_user_data_dir}")
        if settings.yandex_profile_name:
            logger.info(f"Профиль из .env: {settings.yandex_profile_name}")

        # Создание агента
        agent = BrowserAgent(settings)

        # Выполнение основного потока
        agent.execute_flow()

        logger.success("=" * 60)
        logger.success("Работа завершена успешно")
        logger.success("=" * 60)
        return 0

    except KeyboardInterrupt:
        logger.warning("Прервано пользователем")
        return 1

    except Exception as e:
        logger.error(f"Критическая ошибка: {e}")
        logger.exception("Детали ошибки:")
        return 1


if __name__ == "__main__":
    sys.exit(main())
