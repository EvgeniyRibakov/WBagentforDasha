"""
Автоматическое скачивание детализированных отчётов Wildberries
Аналог кнопки "Выгрузить в Excel" на странице аналитики
Поддерживает работу с несколькими кабинетами
"""

import sys
import io
from datetime import datetime, timedelta
from pathlib import Path
from wb_sales_parser import WBSalesParser
import os
from dotenv import load_dotenv

# Устанавливаем UTF-8 для вывода в Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# Загружаем переменные из .env файла (на уровень выше, в корне проекта)
env_path = Path(__file__).parent.parent / ".env"
load_dotenv(dotenv_path=env_path)


def get_cabinets_from_env():
    """
    Получить список кабинетов и их токенов из .env файла
    
    Returns:
        Словарь {название_кабинета: токен}
    """
    cabinets = {}
    
    # Путь к .env файлу в корне проекта
    env_file_path = Path(__file__).parent.parent / ".env"
    
    # Читаем .env файл напрямую для получения всех переменных
    if env_file_path.exists():
        try:
            encodings = ['utf-8', 'utf-8-sig', 'cp1251', 'latin-1']
            env_content = None
            
            for encoding in encodings:
                try:
                    with open(env_file_path, 'r', encoding=encoding) as f:
                        env_content = f.read()
                        break
                except UnicodeDecodeError:
                    continue
            
            if env_content:
                for line in env_content.split('\n'):
                    line = line.strip()
                    if line and not line.startswith('#'):
                        if '=' in line:
                            key, value = line.split('=', 1)
                            key = key.strip()
                            value = value.strip().strip('"').strip("'")
                            
                            # Если значение не пустое, добавляем кабинет
                            if value:
                                cabinets[key] = value
        except Exception as e:
            print(f"⚠ Ошибка при чтении .env файла: {e}")
    
    return cabinets


def get_data_folder_path():
    """
    Получить путь к папке data с сегодняшней датой
    
    Returns:
        Путь к папке в формате data/ДД.ММ.ГГГГ
    """
    project_root = Path(__file__).parent.parent  # На уровень выше от src
    today_formatted = datetime.now().strftime("%d.%m.%Y")
    data_folder = project_root / "data" / today_formatted
    return str(data_folder)


def download_yesterday_report_all_cabinets():
    """
    Скачать отчёты за вчерашний день для всех кабинетов из .env
    
    Returns:
        True если все отчёты скачаны успешно, False при ошибке
    """
    cabinets = get_cabinets_from_env()
    
    if not cabinets:
        print("❌ Не найдено ни одного кабинета с токеном в .env файле")
        print("Проверьте, что в .env указаны токены для кабинетов")
        return False
    
    print(f"✓ Найдено кабинетов: {len(cabinets)}")
    print(f"Кабинеты: {', '.join(cabinets.keys())}\n")
    
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    today_formatted = datetime.now().strftime("%d.%m.%Y")
    
    # Получаем путь к папке data с сегодняшней датой
    data_folder = get_data_folder_path()
    print(f"📁 Папка для сохранения: {data_folder}\n")
    
    # Обрабатываем каждый кабинет по очереди
    import time
    
    for idx, (cabinet_name, api_token) in enumerate(cabinets.items(), 1):
        print(f"{'='*60}")
        print(f"Обработка кабинета: {cabinet_name} ({idx}/{len(cabinets)})")
        print(f"{'='*60}")
        
        # Проверяем токен
        if not api_token or len(api_token.strip()) == 0:
            print(f"❌ Ошибка: Токен для кабинета '{cabinet_name}' пустой!")
            print("Остановка выполнения. Исправьте .env файл и запустите снова.")
            return False
        
        try:
            parser = WBSalesParser(api_token.strip())
            
            # Формируем имя файла: Название_переменной_01.01.2020.xlsx
            filename = f"{cabinet_name}_{today_formatted}.xlsx"
            
            # Скачиваем отчёт в папку data с датой
            success = parser.download_report_to_excel(
                date_from=yesterday,
                date_to=yesterday,
                filename=filename,
                data_folder=data_folder
            )
            
            if not success:
                print(f"⚠ Не удалось скачать отчёт для кабинета '{cabinet_name}'")
                print("Продолжаем обработку следующих кабинетов...")
                continue
            
            filepath = Path(data_folder) / filename
            print(f"✓ Отчёт для кабинета '{cabinet_name}' успешно сохранён: {filepath}\n")
            
            # Добавляем задержку между кабинетами, чтобы избежать ошибки 429
            if idx < len(cabinets):
                print("⏳ Задержка 3 секунды перед следующим кабинетом...\n")
                time.sleep(3)
            
        except Exception as e:
            print(f"❌ Критическая ошибка при обработке кабинета '{cabinet_name}': {e}")
            print("Остановка выполнения.")
            import traceback
            traceback.print_exc()
            return False
    
    print(f"{'='*60}")
    print(f"✓ Все отчёты успешно скачаны! Обработано кабинетов: {len(cabinets)}")
    print(f"{'='*60}")
    return True


def download_custom_period_all_cabinets(date_from: str, date_to: str):
    """
    Скачать отчёты за указанный период для всех кабинетов
    
    Args:
        date_from: Дата начала (YYYY-MM-DD)
        date_to: Дата окончания (YYYY-MM-DD)
    
    Returns:
        True если все отчёты скачаны успешно, False при ошибке
    """
    cabinets = get_cabinets_from_env()
    
    if not cabinets:
        print("❌ Не найдено ни одного кабинета с токеном в .env файле")
        return False
    
    print(f"✓ Найдено кабинетов: {len(cabinets)}")
    print(f"Период: {date_from} - {date_to}\n")
    
    today_formatted = datetime.now().strftime("%d.%m.%Y")
    
    # Получаем путь к папке data с сегодняшней датой
    data_folder = get_data_folder_path()
    print(f"📁 Папка для сохранения: {data_folder}\n")
    
    for cabinet_name, api_token in cabinets.items():
        print(f"{'='*60}")
        print(f"Обработка кабинета: {cabinet_name}")
        print(f"{'='*60}")
        
        if not api_token or len(api_token.strip()) == 0:
            print(f"❌ Ошибка: Токен для кабинета '{cabinet_name}' пустой!")
            return False
        
        try:
            parser = WBSalesParser(api_token.strip())
            # Формат: Название_переменной_01.01.2020.xlsx
            filename = f"{cabinet_name}_{today_formatted}.xlsx"
            
            success = parser.download_report_to_excel(
                date_from=date_from,
                date_to=date_to,
                filename=filename,
                data_folder=data_folder
            )
            
            if not success:
                print(f"❌ Ошибка при скачивании отчёта для кабинета '{cabinet_name}'")
                return False
            
            filepath = Path(data_folder) / filename
            print(f"✓ Отчёт для кабинета '{cabinet_name}' сохранён: {filepath}\n")
            
        except Exception as e:
            print(f"❌ Ошибка при обработке кабинета '{cabinet_name}': {e}")
            return False
    
    print(f"✓ Все отчёты успешно скачаны!")
    return True


def main():
    """Основная функция"""
    if len(sys.argv) > 1:
        # Если переданы аргументы командной строки
        if sys.argv[1] == "--yesterday":
            download_yesterday_report_all_cabinets()
        elif sys.argv[1] == "--period" and len(sys.argv) == 4:
            date_from = sys.argv[2]
            date_to = sys.argv[3]
            download_custom_period_all_cabinets(date_from, date_to)
        elif sys.argv[1] == "--help":
            print("Использование:")
            print("  python wb_report_downloader.py --yesterday")
            print("    Скачать отчёты за вчерашний день для всех кабинетов из .env")
            print("")
            print("  python wb_report_downloader.py --period YYYY-MM-DD YYYY-MM-DD")
            print("    Скачать отчёты за указанный период для всех кабинетов")
            print("    Пример: python wb_report_downloader.py --period 2024-12-01 2024-12-03")
            print("")
            print("Формат .env файла:")
            print("  COSMO=токен_кабинета")
            print("  MMA=токен_кабинета")
            print("  MAB=токен_кабинета")
            print("  ...")
        else:
            print("Неизвестная команда. Используйте --help для справки")
    else:
        # По умолчанию скачиваем за вчера для всех кабинетов
        print("Скачивание отчётов за вчерашний день для всех кабинетов...")
        download_yesterday_report_all_cabinets()


if __name__ == "__main__":
    main()

