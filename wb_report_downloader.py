"""
Автоматическое скачивание детализированных отчётов Wildberries
Аналог кнопки "Выгрузить в Excel" на странице аналитики
"""

import sys
from datetime import datetime, timedelta
from pathlib import Path
from wb_sales_parser import WBSalesParser, main as get_token_main
import os
from dotenv import load_dotenv

# Загружаем переменные из .env файла
env_path = Path(__file__).parent / ".env"
load_dotenv(dotenv_path=env_path)


def download_yesterday_report():
    """Скачать отчёт за вчерашний день"""
    api_token = os.getenv("WB_API_TOKEN")
    
    if not api_token:
        print("❌ API токен не найден в .env файле")
        return False
    
    parser = WBSalesParser(api_token)
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    
    filename = f"wb_report_{yesterday}.xlsx"
    return parser.download_report_to_excel(
        date_from=yesterday,
        date_to=yesterday,
        filename=filename
    )


def download_custom_period(date_from: str, date_to: str, filename: str = None):
    """
    Скачать отчёт за указанный период
    
    Args:
        date_from: Дата начала (YYYY-MM-DD)
        date_to: Дата окончания (YYYY-MM-DD)
        filename: Имя файла (опционально)
    """
    api_token = os.getenv("WB_API_TOKEN")
    
    if not api_token:
        print("❌ API токен не найден в .env файле")
        return False
    
    parser = WBSalesParser(api_token)
    return parser.download_report_to_excel(
        date_from=date_from,
        date_to=date_to,
        filename=filename
    )


def main():
    """Основная функция"""
    if len(sys.argv) > 1:
        # Если переданы аргументы командной строки
        if sys.argv[1] == "--yesterday":
            download_yesterday_report()
        elif sys.argv[1] == "--period" and len(sys.argv) == 4:
            date_from = sys.argv[2]
            date_to = sys.argv[3]
            download_custom_period(date_from, date_to)
        elif sys.argv[1] == "--help":
            print("Использование:")
            print("  python wb_report_downloader.py --yesterday")
            print("    Скачать отчёт за вчерашний день")
            print("")
            print("  python wb_report_downloader.py --period YYYY-MM-DD YYYY-MM-DD")
            print("    Скачать отчёт за указанный период")
            print("    Пример: python wb_report_downloader.py --period 2024-12-01 2024-12-03")
        else:
            print("Неизвестная команда. Используйте --help для справки")
    else:
        # По умолчанию скачиваем за вчера
        print("Скачивание отчёта за вчерашний день...")
        download_yesterday_report()


if __name__ == "__main__":
    main()

