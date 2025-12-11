"""
Скрипт для РУЧНОЙ авторизации в Yandex Browser с сохранением профиля.

ИНСТРУКЦИЯ:
1. Запустите этот скрипт: python manual_auth.py
2. Откроется Yandex Browser с чистым профилем для автоматизации
3. ВРУЧНУЮ авторизуйтесь на странице WB (введите номер и код из SMS)
4. После успешной авторизации ЗАКРОЙТЕ браузер
5. Сессия сохранится в профиле yandex_automation_profile/
6. При следующих запусках основного скрипта авторизация НЕ потребуется!
"""

import time
from pathlib import Path
import undetected_chromedriver as uc
from loguru import logger

def manual_authorization():
    """Запуск браузера для ручной авторизации."""
    
    logger.info("=" * 60)
    logger.info("РЕЖИМ РУЧНОЙ АВТОРИЗАЦИИ")
    logger.info("=" * 60)
    logger.info("")
    logger.info("Сейчас откроется Yandex Browser")
    logger.info("АВТОРИЗУЙТЕСЬ ВРУЧНУЮ на странице Wildberries")
    logger.info("После авторизации закройте браузер")
    logger.info("Сессия сохранится для последующих запусков")
    logger.info("")
    logger.info("=" * 60)
    
    # Настройки браузера
    options = uc.ChromeOptions()
    
    # Путь к Yandex Browser
    browser_path = Path(r"C:\Users\fisher\AppData\Local\Yandex\YandexBrowser\Application\browser.exe")
    options.binary_location = str(browser_path)
    
    # Используем изолированный профиль в папке проекта
    automation_user_data = Path("./yandex_automation_profile").resolve()
    automation_user_data.mkdir(parents=True, exist_ok=True)
    
    options.add_argument(f'--user-data-dir={str(automation_user_data.absolute())}')
    options.add_argument('--profile-directory=Default')
    
    # Папка для скачивания
    downloads_dir = Path("./downloads").absolute()
    downloads_dir.mkdir(exist_ok=True)
    
    # Антидетект
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    
    logger.info(f"Профиль: {automation_user_data.absolute()}")
    logger.info(f"Папка скачивания: {downloads_dir}")
    logger.info("")
    logger.info("Запуск браузера...")
    
    try:
        # Запуск браузера
        driver = uc.Chrome(
            options=options,
            browser_executable_path=str(browser_path),
            version_main=140,
            use_subprocess=False,
        )
        
        # Настройка папки скачивания через CDP
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": str(downloads_dir)
        })
        
        logger.success("✓ Браузер запущен")
        logger.info("")
        logger.info("=" * 60)
        logger.info("Открываем страницу Wildberries...")
        
        # Открываем страницу авторизации WB
        driver.get("https://seller.wildberries.ru/analytics-reports/sales")
        
        logger.info("")
        logger.info("=" * 60)
        logger.success("АВТОРИЗУЙТЕСЬ ВРУЧНУЮ В БРАУЗЕРЕ")
        logger.info("После авторизации нажмите Enter в терминале...")
        logger.info("=" * 60)
        
        input("Нажмите Enter после авторизации: ")
        
        logger.info("")
        logger.success("✓ Закрытие браузера...")
        logger.success("✓ Сессия сохранена в профиле")
        logger.info("")
        logger.info("=" * 60)
        logger.success("ГОТОВО! Теперь можно запускать основной скрипт:")
        logger.info("  python -m src.main")
        logger.success("Авторизация НЕ потребуется!")
        logger.info("=" * 60)
        
        driver.quit()
        
    except Exception as e:
        logger.error(f"Ошибка: {e}")
        raise


if __name__ == "__main__":
    manual_authorization()

