"""
Точка входа в приложение парсера Wildberries
"""

import sys
from pathlib import Path

# Добавляем корневую директорию в путь для импортов
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.config.settings import Settings
from src.utils.logger import setup_logger
from src.agents.browser_agent import BrowserAgent

logger = setup_logger()


def main():
    """Основная функция приложения"""
    logger.info("="*60)
    logger.info("Запуск парсера Wildberries")
    logger.info("="*60)
    
    try:
        # Загружаем настройки
        settings = Settings()
        
        logger.info(f"Стартовая страница: {settings.wildberries_start_url}")
        logger.info(f"Профиль Chrome: {settings.chrome_profile_name}")
        logger.info(f"Папка загрузок: {settings.downloads_dir}")
        
        # Проверка обязательных настроек
        if not settings.chrome_user_data_dir:
            logger.error("⚠ CHROME_USER_DATA_DIR не указан в настройках!")
            logger.error("Укажите путь к профилю Chrome в .env файле")
            return False
        
        # Создаём агент
        agent = BrowserAgent(settings)
        
        # Выполняем основной поток
        success = agent.execute_flow()
        
        if success:
            logger.success("="*60)
            logger.success("✓ Процесс завершён успешно!")
            logger.success("="*60)
        else:
            logger.error("="*60)
            logger.error("✗ Процесс завершён с ошибками")
            logger.error("="*60)
        
        return success
        
    except KeyboardInterrupt:
        logger.warning("Прервано пользователем")
        return False
    except Exception as e:
        logger.error(f"Критическая ошибка: {e}")
        logger.exception("Детали ошибки:")
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)


