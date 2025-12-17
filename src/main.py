"""Точка входа в приложение."""
import sys
import argparse
from pathlib import Path
from datetime import datetime, timedelta

from loguru import logger
import psutil

from src.config import Settings
from src.utils import setup_logger
from src.agents import BrowserAgent


def kill_yandex_processes() -> int:
    """Закрывает все процессы Yandex Browser.
    
    Returns:
        Количество закрытых процессов
    """
    killed_count = 0
    processes_to_kill = ['browser.exe', 'YandexBrowser.exe']
    
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.info['name'] and proc.info['name'].lower() in [p.lower() for p in processes_to_kill]:
                logger.info(f"Закрытие процесса: {proc.info['name']} (PID: {proc.info['pid']})")
                proc.kill()
                killed_count += 1
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    
    if killed_count > 0:
        logger.info(f"✓ Закрыто процессов Yandex Browser: {killed_count}")
        # Даём время процессам завершиться
        import time
        time.sleep(2)
    else:
        logger.info("✓ Процессы Yandex Browser не найдены")
    
    return killed_count


def show_startup_warning() -> bool:
    """Показывает предупреждение перед запуском и запрашивает подтверждение.
    
    Returns:
        True если пользователь подтвердил запуск, False если отменил
    """
    print("\n" + "=" * 70)
    print("⚠️  ВАЖНОЕ ПРЕДУПРЕЖДЕНИЕ ПЕРЕД ЗАПУСКОМ СКРИПТА")
    print("=" * 70)
    print("\nПеред запуском скрипта необходимо:")
    print("1. ✅ Закрыть ВСЕ окна Yandex Browser вручную")
    print("2. ✅ Выключить антивирус (или добавить папку проекта в исключения)")
    print("3. ✅ Убедиться, что у вас есть доступ к телефону и почте для кодов авторизации")
    print("\n" + "-" * 70)
    print("Скрипт автоматически закроет все процессы Yandex Browser перед запуском.")
    print("-" * 70 + "\n")
    
    while True:
        response = input("Вы выполнили все пункты и готовы продолжить? (да/нет): ").strip().lower()
        if response in ['да', 'yes', 'y', 'д']:
            return True
        elif response in ['нет', 'no', 'n', 'н']:
            print("\n❌ Запуск отменён пользователем.")
            return False
        else:
            print("Пожалуйста, введите 'да' или 'нет'")


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
        
        # Показываем предупреждение и запрашиваем подтверждение
        if not show_startup_warning():
            return 1
        
        # Закрываем все процессы Yandex Browser
        logger.info("Закрытие процессов Yandex Browser...")
        kill_yandex_processes()

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

        # Парсинг аргументов командной строки
        parser = argparse.ArgumentParser(description='WBagentforDasha - Автоматизация выгрузки отчётов Wildberries')
        parser.add_argument(
            '--date',
            type=str,
            help='Дата для скачивания отчётов в формате DD.MM.YYYY (например: 10.12.2025). По умолчанию - вчерашний день',
            default=None
        )
        args = parser.parse_args()
        
        # Обработка даты
        target_date = None
        if args.date:
            try:
                target_date = datetime.strptime(args.date, "%d.%m.%Y").date()
                logger.info(f"✓ Указана дата для скачивания: {target_date.strftime('%d.%m.%Y')}")
            except ValueError:
                logger.error(f"❌ Неверный формат даты: {args.date}. Используйте формат DD.MM.YYYY (например: 10.12.2025)")
                return 1
        else:
            # По умолчанию - вчерашний день
            target_date = (datetime.now() - timedelta(days=1)).date()
            logger.info(f"✓ Используется дата по умолчанию (вчера): {target_date.strftime('%d.%m.%Y')}")

        # Создание агента
        agent = BrowserAgent(settings)

        # Выполнение основного потока
        agent.execute_flow(target_date=target_date)

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
