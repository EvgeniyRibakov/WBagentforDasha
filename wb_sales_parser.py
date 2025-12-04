"""
Парсер эндпоинта продаж Wildberries API
https://statistics-api.wildberries.ru/api/v1/supplier/sales
"""

import requests
import json
from datetime import datetime, timedelta
from typing import Optional, Dict, List
import os
from pathlib import Path
from dotenv import load_dotenv

# Загружаем переменные из .env файла
# Явно указываем путь к файлу .env в корне проекта
env_path = Path(__file__).parent / ".env"
load_dotenv(dotenv_path=env_path)


class WBSalesParser:
    """Класс для парсинга данных о продажах из Wildberries API"""
    
    BASE_URL = "https://statistics-api.wildberries.ru/api/v1/supplier/sales"
    REPORT_URL = "https://statistics-api.wildberries.ru/api/v1/supplier/reportDetailByPeriod"
    
    def __init__(self, api_token: str):
        """
        Инициализация парсера
        
        Args:
            api_token: API токен от Wildberries
        """
        self.api_token = api_token
        self.headers = {
            "Authorization": f"Bearer {api_token}",
            "Content-Type": "application/json"
        }
    
    def get_sales(
        self, 
        date_from: Optional[str] = None,
        date_to: Optional[str] = None,
        flag: Optional[int] = None
    ) -> Dict:
        """
        Получить данные о продажах
        
        Args:
            date_from: Дата начала периода в формате YYYY-MM-DD (опционально)
            date_to: Дата окончания периода в формате YYYY-MM-DD (опционально)
            flag: Флаг для фильтрации (0 - все продажи, 1 - только новые) (опционально)
        
        Returns:
            Словарь с данными о продажах или ошибкой
        """
        params = {}
        
        if date_from:
            params["dateFrom"] = date_from
        if date_to:
            params["dateTo"] = date_to
        if flag is not None:
            params["flag"] = flag
        
        try:
            response = requests.get(
                self.BASE_URL,
                headers=self.headers,
                params=params if params else None,
                timeout=30
            )
            
            response.raise_for_status()
            return {
                "success": True,
                "data": response.json(),
                "status_code": response.status_code
            }
            
        except requests.exceptions.HTTPError as e:
            return {
                "success": False,
                "error": f"HTTP ошибка: {e}",
                "status_code": response.status_code if 'response' in locals() else None,
                "response_text": response.text if 'response' in locals() else None
            }
        except requests.exceptions.RequestException as e:
            return {
                "success": False,
                "error": f"Ошибка запроса: {e}"
            }
        except json.JSONDecodeError as e:
            return {
                "success": False,
                "error": f"Ошибка парсинга JSON: {e}",
                "response_text": response.text if 'response' in locals() else None
            }
    
    def get_sales_last_days(self, days: int = 7) -> Dict:
        """
        Получить продажи за последние N дней
        
        Args:
            days: Количество дней назад
        
        Returns:
            Словарь с данными о продажах
        """
        date_to = datetime.now().strftime("%Y-%m-%d")
        date_from = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
        
        return self.get_sales(date_from=date_from, date_to=date_to)
    
    def get_sales_yesterday(self) -> Dict:
        """
        Получить продажи только за вчерашний день
        
        Returns:
            Словарь с данными о продажах за вчера
        """
        yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        return self.get_sales(date_from=yesterday, date_to=yesterday)
    
    def get_report_detail(
        self,
        date_from: str,
        date_to: str,
        rrdid: Optional[int] = None,
        limit: int = 100000
    ) -> Dict:
        """
        Получить детализированный отчёт о продажах (аналог "Выгрузить в Excel")
        
        Args:
            date_from: Дата начала периода в формате YYYY-MM-DD
            date_to: Дата окончания периода в формате YYYY-MM-DD
            rrdid: Идентификатор отчёта (опционально)
            limit: Лимит записей (по умолчанию 100000)
        
        Returns:
            Словарь с данными отчёта или ошибкой
        """
        params = {
            "dateFrom": date_from,
            "dateTo": date_to,
            "limit": limit
        }
        
        if rrdid:
            params["rrdid"] = rrdid
        
        try:
            response = requests.get(
                self.REPORT_URL,
                headers=self.headers,
                params=params,
                timeout=60
            )
            
            response.raise_for_status()
            return {
                "success": True,
                "data": response.json(),
                "status_code": response.status_code
            }
            
        except requests.exceptions.HTTPError as e:
            return {
                "success": False,
                "error": f"HTTP ошибка: {e}",
                "status_code": response.status_code if 'response' in locals() else None,
                "response_text": response.text if 'response' in locals() else None
            }
        except requests.exceptions.RequestException as e:
            return {
                "success": False,
                "error": f"Ошибка запроса: {e}"
            }
        except json.JSONDecodeError as e:
            return {
                "success": False,
                "error": f"Ошибка парсинга JSON: {e}",
                "response_text": response.text if 'response' in locals() else None
            }
    
    def get_report_yesterday(self) -> Dict:
        """
        Получить детализированный отчёт за вчерашний день
        
        Returns:
            Словарь с данными отчёта за вчера
        """
        yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        return self.get_report_detail(date_from=yesterday, date_to=yesterday)
    
    def save_to_json(self, data: Dict, filename: str = "wb_sales.json"):
        """
        Сохранить данные в JSON файл
        
        Args:
            data: Данные для сохранения
            filename: Имя файла
        """
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"Данные сохранены в файл: {filename}")
    
    def save_to_excel(self, data: List[Dict], filename: str = "wb_report.xlsx"):
        """
        Сохранить данные в Excel файл (аналог "Выгрузить в Excel")
        
        Args:
            data: Список словарей с данными для сохранения
            filename: Имя файла Excel
        """
        try:
            import pandas as pd
            
            if not data:
                print("⚠ Нет данных для сохранения в Excel")
                return
            
            # Создаём DataFrame из данных
            df = pd.DataFrame(data)
            
            # Сохраняем в Excel
            df.to_excel(filename, index=False, engine='openpyxl')
            print(f"✓ Данные сохранены в Excel файл: {filename}")
            print(f"  Всего строк: {len(df)}")
            
        except ImportError:
            print("❌ Для сохранения в Excel необходимо установить библиотеки:")
            print("   pip install pandas openpyxl")
        except Exception as e:
            print(f"❌ Ошибка при сохранении в Excel: {e}")
    
    def download_report_to_excel(
        self,
        date_from: str,
        date_to: str,
        filename: Optional[str] = None,
        use_detailed_api: bool = True
    ) -> bool:
        """
        Скачать отчёт и сохранить в Excel
        
        Args:
            date_from: Дата начала периода в формате YYYY-MM-DD
            date_to: Дата окончания периода в формате YYYY-MM-DD
            filename: Имя файла Excel (если не указано, будет сгенерировано автоматически)
            use_detailed_api: Попытаться использовать детализированный API (если доступен)
        
        Returns:
            True если успешно, False если ошибка
        """
        print(f"Получение отчёта за период {date_from} - {date_to}...")
        
        data = None
        
        # Сначала пробуем получить детализированный отчёт через API
        if use_detailed_api:
            report_data = self.get_report_detail(date_from=date_from, date_to=date_to)
            if report_data.get("success"):
                data = report_data.get("data", [])
                if isinstance(data, list) and data:
                    print("✓ Получены данные через детализированный API")
        
        # Если детализированный API не сработал, используем базовый API
        if not data:
            print("⚠ Детализированный API недоступен, используем базовый API...")
            sales_data = self.get_sales(date_from=date_from, date_to=date_to)
            
            if not sales_data.get("success"):
                print(f"❌ Ошибка получения данных: {sales_data.get('error')}")
                return False
            
            data = sales_data.get("data", [])
            if not isinstance(data, list):
                print("❌ Неожиданный формат данных")
                return False
            
            if not data:
                print("⚠ Нет данных за указанный период")
                return False
            
            print(f"✓ Получены данные через базовый API ({len(data)} записей)")
        
        if not filename:
            filename = f"wb_report_{date_from}_to_{date_to}.xlsx"
        
        self.save_to_excel(data, filename)
        return True
    
    def print_sales_summary(self, sales_data: Dict):
        """
        Вывести краткую сводку по продажам
        
        Args:
            sales_data: Данные о продажах
        """
        if not sales_data.get("success"):
            print(f"Ошибка: {sales_data.get('error')}")
            return
        
        data = sales_data.get("data", [])
        if not isinstance(data, list):
            print("Неожиданный формат данных")
            return
        
        print(f"\n=== Сводка по продажам ===")
        print(f"Всего записей: {len(data)}")
        
        if data:
            total_sum = sum(item.get("totalPrice", 0) for item in data)
            print(f"Общая сумма продаж: {total_sum:.2f} руб.")
            
            # Группировка по артикулам
            articles = {}
            for item in data:
                article = item.get("nmId", "Неизвестно")
                if article not in articles:
                    articles[article] = {"count": 0, "sum": 0}
                articles[article]["count"] += item.get("quantity", 0)
                articles[article]["sum"] += item.get("totalPrice", 0)
            
            print(f"\nТоп-5 артикулов по количеству:")
            sorted_articles = sorted(articles.items(), key=lambda x: x[1]["count"], reverse=True)
            for i, (article, stats) in enumerate(sorted_articles[:5], 1):
                print(f"{i}. Артикул {article}: {stats['count']} шт., {stats['sum']:.2f} руб.")


def main():
    """Основная функция для запуска парсера"""
    
    # Проверяем наличие .env файла
    env_file = Path(__file__).parent / ".env"
    api_token = None
    
    if env_file.exists():
        print(f"✓ Файл .env найден: {env_file}")
        
        # Пытаемся прочитать файл напрямую
        try:
            # Пробуем разные кодировки
            encodings = ['utf-8', 'utf-8-sig', 'cp1251', 'latin-1']
            env_content = None
            
            for encoding in encodings:
                try:
                    with open(env_file, 'r', encoding=encoding) as f:
                        env_content = f.read()
                        break
                except UnicodeDecodeError:
                    continue
            
            if env_content is None:
                print("❌ Не удалось прочитать файл .env (проблема с кодировкой)")
            else:
                print(f"Размер файла: {len(env_content)} символов")
                print(f"Содержимое файла (первые 100 символов): {repr(env_content[:100])}")
                
                # Парсим файл вручную
                lines_found = []
                for line_num, line in enumerate(env_content.split('\n'), 1):
                    original_line = line
                    line = line.strip()
                    if line and not line.startswith('#'):
                        if '=' in line:
                            key, value = line.split('=', 1)
                            key = key.strip()
                            value = value.strip().strip('"').strip("'")
                            lines_found.append(f"Строка {line_num}: ключ='{key}', значение длина={len(value)}")
                            
                            if key == 'WB_API_TOKEN' or key == 'WB_API_TOKEN ' or 'WB_API_TOKEN' in key:
                                if value:
                                    api_token = value
                                    print(f"✓ Токен найден в файле")
                                else:
                                    print(f"❌ Строка {line_num}: ключ найден, но значение пустое")
                                    print(f"   Содержимое строки: {repr(original_line)}")
                
                if not api_token:
                    if lines_found:
                        print(f"\nНайдено строк в .env: {len(lines_found)}")
                        for info in lines_found:
                            print(f"  {info}")
                        print(f"\nИщем ключ: 'WB_API_TOKEN'")
                        print(f"Проверьте, что ключ написан точно так же (без пробелов, регистр важен)")
                    else:
                        print(f"\n⚠ В файле .env не найдено ни одной строки с ключом=значение")
                        print(f"Проверьте формат файла")
                    
        except Exception as e:
            print(f"⚠ Ошибка при чтении .env файла: {e}")
            import traceback
            traceback.print_exc()
    else:
        print(f"⚠ Файл .env не найден в: {env_file}")
    
    # Если токен не найден в файле, пытаемся получить из переменных окружения
    if not api_token:
        api_token = os.getenv("WB_API_TOKEN")
        if api_token:
            api_token = api_token.strip().strip('"').strip("'")
            print(f"✓ Токен найден в переменных окружения")
    
    # Финальная проверка
    if not api_token:
        print("\n❌ Ошибка: API токен не указан!")
        print("Проверьте файл .env или переменные окружения.")
        return
    
    print(f"✓ Токен загружен")
    
    # Создаем экземпляр парсера
    parser = WBSalesParser(api_token)
    
    # Получаем детализированный отчёт за вчерашний день и сохраняем в Excel
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    print(f"Скачивание детализированного отчёта за {yesterday}...")
    
    success = parser.download_report_to_excel(
        date_from=yesterday,
        date_to=yesterday,
        filename=f"wb_report_{yesterday}.xlsx"
    )
    
    if success:
        print(f"\n✓ Отчёт успешно скачан и сохранён в Excel!")
    else:
        print(f"\n⚠ Не удалось скачать отчёт. Пробуем получить базовые данные о продажах...")
        sales_data = parser.get_sales_yesterday()
        parser.print_sales_summary(sales_data)
        
        if sales_data.get("success"):
            parser.save_to_json(sales_data["data"], "wb_sales.json")
    
    # Другие примеры использования:
    # Скачать отчёт за конкретный период в Excel:
    # parser.download_report_to_excel(date_from="2024-01-01", date_to="2024-01-31")
    # 
    # Получить продажи за последние 7 дней:
    # sales_data = parser.get_sales_last_days(days=7)


if __name__ == "__main__":
    main()

