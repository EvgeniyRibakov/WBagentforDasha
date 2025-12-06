"""
Парсер детализированных отчётов о продажах Wildberries API
Использует API для получения детализированных отчётов (аналог "Выгрузить в Excel")
"""

import requests
import json
from datetime import datetime, timedelta
from typing import Optional, Dict, List
import os
from pathlib import Path
from dotenv import load_dotenv

# Загружаем переменные из .env файла
# Явно указываем путь к файлу .env в корне проекта (на уровень выше от src)
env_path = Path(__file__).parent.parent / ".env"
load_dotenv(dotenv_path=env_path)


class WBSalesParser:
    """Класс для парсинга данных о продажах из Wildberries API"""
    
    # Основной API для детализированных отчётов (v5 - правильная версия)
    DETAILED_REPORT_API = "https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod"
    # Альтернативные варианты эндпоинта
    DETAILED_REPORT_API_V1 = "https://statistics-api.wildberries.ru/api/v1/supplier/reportDetailByPeriod"
    DETAILED_REPORT_API_V2 = "https://statistics-api.wildberries.ru/api/v2/supplier/reportDetailByPeriod"
    
    # Новый API для создания и получения отчётов
    ANALYTICS_API_BASE = "https://seller-analytics-api.wildberries.ru/api/v2"
    REPORT_CREATE_URL = f"{ANALYTICS_API_BASE}/nm-report/downloads"
    REPORT_GET_URL = f"{ANALYTICS_API_BASE}/nm-report/downloads/file"
    
    # Старый API (оставлен для совместимости)
    REPORT_API_BASE = "https://seller-weekly-report.wildberries.ru/ns/reportsviewer/analytics-back/api/report"
    REPORT_DOWNLOAD_URL = f"{REPORT_API_BASE}/supplier-goods/xlsx"
    
    def __init__(self, api_token: str):
        """
        Инициализация парсера
        
        Args:
            api_token: API токен от Wildberries (JWT токен для authorizev3)
        """
        self.api_token = api_token
        self.headers = {
            "Authorization": f"Bearer {api_token}",
            "Content-Type": "application/json"
        }
        # Заголовки для нового API отчётов (seller-analytics-api использует HeaderApiKey)
        self.analytics_headers = {
            "Authorization": f"Bearer {api_token}",
            "HeaderApiKey": api_token,  # Также пробуем HeaderApiKey на случай если нужен он
            "Content-Type": "application/json"
        }
        # Заголовки для старого API отчётов (использует authorizev3)
        self.report_headers = {
            "authorizev3": api_token,
            "Content-Type": "application/json",
            "accept": "*/*",
            "accept-language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
            "origin": "https://seller.wildberries.ru",
            "referer": "https://seller.wildberries.ru/",
            "sec-ch-ua": '"Google Chrome";v="143", "Chromium";v="143", "Not A(Brand";v="24"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"'
        }
        
        # Заголовки для скачивания Excel файла
        self.excel_headers = {
            "authorizev3": api_token,
            "accept": "*/*",
            "accept-encoding": "gzip, deflate, br, zstd",
            "accept-language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
            "origin": "https://seller.wildberries.ru",
            "referer": "https://seller.wildberries.ru/",
            "sec-ch-ua": '"Google Chrome";v="143", "Chromium";v="143", "Not A(Brand";v="24"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"'
        }
    
    def get_report_detail(
        self,
        date_from: str,
        date_to: str,
        rrdid: Optional[int] = None,
        limit: int = 100000
    ) -> Dict:
        """
        Получить детализированный отчёт о продажах через API reportDetailByPeriod
        Это правильный API для получения детализированных отчётов (аналог "Выгрузить в Excel")
        
        Args:
            date_from: Дата начала периода в формате YYYY-MM-DD
            date_to: Дата окончания периода в формате YYYY-MM-DD
            rrdid: Идентификатор отчёта (опционально)
            limit: Лимит записей (по умолчанию 100000)
        
        Returns:
            Словарь с данными отчёта или ошибкой
        """
        # API v5 требует формат RFC3339 для дат (YYYY-MM-DDTHH:MM:SSZ)
        # ВАЖНО: API v5 возвращает еженедельные детализации, которые формируются только на следующей неделе
        # Для ежедневных данных нужно использовать другие API
        date_from_rfc = f"{date_from}T00:00:00Z"
        date_to_rfc = f"{date_to}T23:59:59Z"
        
        # Пробуем разные варианты эндпоинтов
        # v1 и v2 могут возвращать ежедневные данные, v5 - только еженедельные
        endpoints_to_try = [
            (self.DETAILED_REPORT_API_V1, {"dateFrom": date_from, "dateTo": date_to, "limit": limit, "rrdid": rrdid if rrdid else None}),  # v1 - может быть ежедневный
            (self.DETAILED_REPORT_API_V2, {"dateFrom": date_from, "dateTo": date_to, "limit": limit, "rrdid": rrdid if rrdid else None}),  # v2 - может быть ежедневный
            (self.DETAILED_REPORT_API, {"dateFrom": date_from_rfc, "dateTo": date_to_rfc, "rrdid": rrdid if rrdid is not None else 0}),  # v5 - еженедельный (пробуем в последнюю очередь)
        ]
        
        for endpoint, params_dict in endpoints_to_try:
            try:
                # Убираем None значения из параметров
                params = {k: v for k, v in params_dict.items() if v is not None}
                
                # Определяем версию API для логирования
                api_version = "v1" if "/v1/" in endpoint else ("v2" if "/v2/" in endpoint else "v5")
                print(f"  Пробуем API {api_version}...")
                
                response = requests.get(
                    endpoint,
                    headers=self.headers,
                    params=params,
                    timeout=120
                )
                
                if response.status_code == 200:
                    try:
                        data = response.json()
                        print(f"✓ Успешно получены данные через API {api_version}")
                        print(f"  Тип данных: {type(data)}, длина: {len(data) if isinstance(data, list) else 'N/A'}")
                        
                        # Проверяем структуру данных
                        if isinstance(data, list):
                            if len(data) > 0:
                                first_item = data[0]
                                if isinstance(first_item, dict):
                                    print(f"  Колонки: {', '.join(list(first_item.keys())[:15])}...")
                                    has_brand = 'brand' in first_item
                                    has_subject = 'subject' in first_item
                                    has_warehouse = 'warehouseName' in first_item
                                    has_supplier_article = 'supplierArticle' in first_item
                                    
                                    if has_brand or has_subject or has_warehouse or has_supplier_article:
                                        print(f"✓ Получена правильная структура данных")
                                    else:
                                        print(f"⚠ Структура данных может быть неполной")
                            else:
                                # Пустой список
                                if api_version == "v5":
                                    print(f"⚠ API v5 вернул пустой список")
                                    print(f"  ВАЖНО: API v5 возвращает еженедельные детализации, которые формируются только на следующей неделе")
                                    print(f"  Для ежедневных данных используйте API v1 или v2")
                                else:
                                    print(f"⚠ API {api_version} вернул пустой список (возможно, для этого периода нет данных)")
                                # Продолжаем пробовать другие эндпоинты
                                continue
                        else:
                            print(f"⚠ API вернул не список: {type(data)}")
                            if isinstance(data, dict):
                                print(f"  Ключи: {list(data.keys())[:10]}")
                        
                        return {
                            "success": True,
                            "data": data,
                            "status_code": response.status_code
                        }
                    except json.JSONDecodeError:
                        # Возможно это не JSON, а другой формат
                        print(f"⚠ Ответ от {endpoint} не является JSON")
                        continue
                elif response.status_code == 404:
                    # Эндпоинт не найден, пробуем следующий
                    print(f"⚠ API {api_version} вернул 404 (эндпоинт не найден)")
                    continue
                else:
                    # Другая ошибка, пробуем следующий эндпоинт
                    print(f"⚠ API {api_version} вернул HTTP {response.status_code}")
                    if hasattr(response, 'text'):
                        print(f"  Ответ: {response.text[:200]}")
                    continue
            except requests.exceptions.RequestException as e:
                print(f"⚠ Ошибка запроса к {endpoint}: {e}")
                continue
        
        # Если ни один эндпоинт не сработал
        return {
            "success": False,
            "error": "Не удалось получить отчёт ни через один из доступных эндпоинтов",
            "status_code": None,
            "response_text": None
        }
    
    def create_analytics_report(
        self,
        date_from: str,
        date_to: str,
        report_type: str = "DETAIL_HISTORY_REPORT"
    ) -> Dict:
        """
        Создать задание на генерацию отчёта через новый API
        
        Args:
            date_from: Дата начала периода в формате YYYY-MM-DD
            date_to: Дата окончания периода в формате YYYY-MM-DD
            report_type: Тип отчёта
                - DETAIL_HISTORY_REPORT - воронка продаж (возвращает dt, openCardCount и т.д.)
                - STOCK_HISTORY_REPORT_CSV - история остатков
                Нужен отчёт с полями: Бренд, Предмет, Сезон, Коллекция, Наименование, Артикул поставщика...
        
        Returns:
            Словарь с downloadId или ошибкой
        """
        import uuid
        
        report_id = str(uuid.uuid4())
        
        # Параметры для отчёта зависят от типа отчёта
        if report_type == "DETAIL_HISTORY_REPORT":
            params = {
                "startDate": date_from,
                "endDate": date_to,
                "groupBy": "nmId",  # Группировка по артикулам WB
                "timezone": "Europe/Moscow"
            }
        elif report_type == "STOCK_HISTORY_REPORT_CSV":
            # История остатков - может содержать нужные поля
            params = {
                "startDate": date_from,
                "endDate": date_to,
                "timezone": "Europe/Moscow"
            }
        else:
            # Для других типов отчётов параметры могут отличаться
            params = {
                "startDate": date_from,
                "endDate": date_to,
                "timezone": "Europe/Moscow"
            }
        
        request_body = {
            "id": report_id,
            "reportType": report_type,
            "params": params
        }
        
        try:
            print(f"Создание задания на генерацию отчёта (ID: {report_id})...")
            response = requests.post(
                self.REPORT_CREATE_URL,
                headers=self.analytics_headers,
                json=request_body,
                timeout=60
            )
            
            print(f"Ответ сервера: HTTP {response.status_code}")
            if response.status_code == 200 or response.status_code == 201:
                try:
                    data = response.json()
                    print(f"Ответ JSON: {json.dumps(data, ensure_ascii=False)[:200]}")
                    download_id = data.get("downloadId") or data.get("id") or report_id
                    print(f"✓ Задание создано, downloadId: {download_id}")
                    return {
                        "success": True,
                        "downloadId": download_id,
                        "reportId": report_id,
                        "data": data,
                        "status_code": response.status_code
                    }
                except json.JSONDecodeError:
                    # Возможно ответ не JSON
                    print(f"⚠ Ответ не является JSON: {response.text[:200]}")
                    return {
                        "success": True,
                        "downloadId": report_id,
                        "reportId": report_id,
                        "data": response.text,
                        "status_code": response.status_code
                    }
            else:
                error_text = response.text[:500] if hasattr(response, 'text') else None
                if response.status_code == 403:
                    print(f"⚠ HTTP 403: Отчёт недоступен для этого кабинета (возможно, нужна подписка)")
                else:
                    print(f"❌ HTTP {response.status_code}: {error_text}")
                return {
                    "success": False,
                    "error": f"HTTP ошибка {response.status_code} при создании отчёта",
                    "status_code": response.status_code,
                    "response_text": error_text
                }
        except requests.exceptions.RequestException as e:
            return {
                "success": False,
                "error": f"Ошибка запроса при создании отчёта: {e}"
            }
    
    def get_analytics_report_file(
        self,
        download_id: str
    ) -> Dict:
        """
        Получить отчёт по downloadId
        
        Args:
            download_id: ID задания на генерацию отчёта
        
        Returns:
            Словарь с бинарными данными ZIP архива или ошибкой
        """
        try:
            url = f"{self.REPORT_GET_URL}/{download_id}"
            print(f"Получение отчёта по downloadId: {download_id}...")
            response = requests.get(
                url,
                headers=self.analytics_headers,
                timeout=120,
                stream=True
            )
            
            if response.status_code == 200:
                content_type = response.headers.get('Content-Type', '')
                if 'zip' in content_type.lower() or 'application/zip' in content_type.lower():
                    print(f"✓ Получен ZIP архив ({len(response.content)} байт)")
                    return {
                        "success": True,
                        "data": response.content,
                        "format": "zip",
                        "status_code": response.status_code
                    }
                else:
                    # Возможно это не ZIP
                    return {
                        "success": True,
                        "data": response.content,
                        "format": "unknown",
                        "status_code": response.status_code,
                        "content_type": content_type
                    }
            else:
                error_text = response.text[:500] if hasattr(response, 'text') else None
                return {
                    "success": False,
                    "error": f"HTTP ошибка {response.status_code} при получении отчёта",
                    "status_code": response.status_code,
                    "response_text": error_text
                }
        except requests.exceptions.RequestException as e:
            return {
                "success": False,
                "error": f"Ошибка запроса при получении отчёта: {e}"
            }
    
    def get_report_detail_by_period(
        self,
        date_from: str,
        date_to: str,
        supplier_id: Optional[str] = None
    ) -> Dict:
        """
        Получить детализированный отчёт о продажах через новый API (аналог "Выгрузить в Excel")
        Использует правильный эндпоинт со страницы analytics-reports/sales
        
        Args:
            date_from: Дата начала периода в формате YYYY-MM-DD
            date_to: Дата окончания периода в формате YYYY-MM-DD
            supplier_id: ID поставщика (опционально, будет извлечён из токена если не указан)
        
        Returns:
            Словарь с бинарными данными Excel файла или ошибкой
        """
        import base64
        
        # Пытаемся извлечь supplier_id из токена, если не указан
        if not supplier_id:
            try:
                # JWT токен состоит из трёх частей, разделённых точками
                parts = self.api_token.split('.')
                if len(parts) >= 2:
                    # Декодируем payload (вторая часть)
                    payload = parts[1]
                    # Добавляем padding если нужно
                    padding = 4 - len(payload) % 4
                    if padding != 4:
                        payload += '=' * padding
                    decoded = base64.urlsafe_b64decode(payload)
                    token_data = json.loads(decoded)
                    # Пытаемся найти supplier_id в токене (пробуем разные варианты полей)
                    supplier_id = (
                        token_data.get('user') or 
                        token_data.get('supplier_id') or 
                        token_data.get('supplierId') or
                        token_data.get('userId') or
                        token_data.get('id')
                    )
                    if supplier_id:
                        print(f"✓ Извлечён supplier_id из токена: {supplier_id}")
                    else:
                        # Выводим все ключи токена для отладки
                        print(f"⚠ Поля в токене: {list(token_data.keys())}")
            except Exception as e:
                print(f"⚠ Не удалось извлечь supplier_id из токена: {e}")
        
        # Если supplier_id не найден, продолжаем без него
        if not supplier_id:
            print("⚠ supplier_id не найден, будет использован 'unknown' в URL")
        
        # Пробуем разные варианты эндпоинтов для создания/получения отчёта
        create_endpoints = [
            f"{self.REPORT_API_BASE}/supplier-goods/create",
            f"{self.REPORT_API_BASE}/supplier-goods/generate",
            f"{self.REPORT_API_BASE}/supplier-goods",
            f"{self.REPORT_API_BASE}/supplier-goods/request",
        ]
        
        report_hash = None
        download_url = None
        
        # Пробуем создать задание на генерацию отчёта
        for create_url in create_endpoints:
            try:
                create_params = {
                    "dateFrom": date_from,
                    "dateTo": date_to
                }
                
                # Пробуем POST запрос
                create_response = requests.post(
                    create_url,
                    headers=self.report_headers,
                    json=create_params,
                    timeout=60
                )
                
                if create_response.status_code == 200:
                    try:
                        report_data = create_response.json()
                        report_hash = (
                            report_data.get("reportId") or 
                            report_data.get("id") or 
                            report_data.get("hash") or 
                            report_data.get("report_hash") or
                            report_data.get("reportHash")
                        )
                        download_url = (
                            report_data.get("downloadUrl") or 
                            report_data.get("url") or 
                            report_data.get("download_url") or
                            report_data.get("downloadURL")
                        )
                        if report_hash or download_url:
                            print(f"✓ Получен hash/URL через POST {create_url}")
                            if report_hash:
                                print(f"  Hash: {report_hash}")
                            break
                    except Exception as e:
                        print(f"⚠ Ошибка парсинга JSON ответа от {create_url}: {e}")
                        print(f"  Ответ: {create_response.text[:200]}")
                
                # Пробуем GET запрос с параметрами
                create_response = requests.get(
                    create_url,
                    headers=self.report_headers,
                    params=create_params,
                    timeout=60
                )
                
                if create_response.status_code == 200:
                    try:
                        report_data = create_response.json()
                        report_hash = (
                            report_data.get("reportId") or 
                            report_data.get("id") or 
                            report_data.get("hash") or 
                            report_data.get("report_hash") or
                            report_data.get("reportHash")
                        )
                        download_url = (
                            report_data.get("downloadUrl") or 
                            report_data.get("url") or 
                            report_data.get("download_url") or
                            report_data.get("downloadURL")
                        )
                        if report_hash or download_url:
                            print(f"✓ Получен hash/URL через GET {create_url}")
                            if report_hash:
                                print(f"  Hash: {report_hash}")
                            break
                    except Exception as e:
                        print(f"⚠ Ошибка парсинга JSON ответа от {create_url}: {e}")
                        print(f"  Ответ: {create_response.text[:200]}")
                elif create_response.status_code == 404:
                    # Эндпоинт не существует, пропускаем
                    continue
                else:
                    print(f"⚠ HTTP {create_response.status_code} от {create_url}")
            except requests.exceptions.Timeout:
                print(f"⚠ Таймаут при запросе к {create_url}")
                continue
            except Exception as e:
                print(f"⚠ Ошибка при запросе к {create_url}: {e}")
                continue
        
        # Если не удалось получить hash, пробуем скачать напрямую
        # Структура URL: /supplier-goods/xlsx/supplier-goods-{supplier_id}-{date_from}-{date_to}-{hash}
        if not download_url:
            if report_hash:
                download_url = f"{self.REPORT_DOWNLOAD_URL}/supplier-goods-{supplier_id or 'unknown'}-{date_from}-{date_to}-{report_hash}"
            else:
                # Пробуем использовать эндпоинт без hash (может быть работает)
                download_url = f"{self.REPORT_DOWNLOAD_URL}/supplier-goods-{supplier_id or 'unknown'}-{date_from}-{date_to}"
                print(f"⚠ Hash не получен, пробуем скачать без hash")
        
        # Пробуем скачать отчёт
        try:
            print(f"Попытка скачать отчёт: {download_url}")
            response = requests.get(
                download_url,
                headers=self.excel_headers,  # Используем заголовки для Excel
                timeout=120,
                stream=True  # Для больших файлов
            )
            
            if response.status_code == 200:
                # Проверяем Content-Type
                content_type = response.headers.get('Content-Type', '')
                if 'excel' in content_type.lower() or 'spreadsheet' in content_type.lower() or 'xlsx' in content_type.lower() or 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in content_type:
                    # Это Excel файл, возвращаем как бинарные данные
                    print(f"✓ Получен Excel файл ({len(response.content)} байт)")
                    return {
                        "success": True,
                        "data": response.content,  # Бинарные данные Excel
                        "status_code": response.status_code,
                        "format": "xlsx",
                        "url_used": download_url
                    }
                else:
                    # Возможно это JSON или другой формат
                    try:
                        data = response.json()
                        # Если это JSON с информацией об отчёте, пробуем извлечь URL
                        if isinstance(data, dict):
                            new_url = data.get("downloadUrl") or data.get("url") or data.get("download_url")
                            if new_url:
                                print(f"✓ Получен URL для скачивания из JSON ответа")
                                # Рекурсивно вызываем себя с новым URL
                                return self._download_from_url(new_url)
                        return {
                            "success": True,
                            "data": data,
                            "status_code": response.status_code,
                            "format": "json",
                            "url_used": download_url
                        }
                    except:
                        # Текстовый формат
                        return {
                            "success": True,
                            "data": response.text,
                            "status_code": response.status_code,
                            "format": "text",
                            "url_used": download_url
                        }
            else:
                error_text = response.text[:500] if hasattr(response, 'text') else None
                return {
                    "success": False,
                    "error": f"HTTP ошибка {response.status_code} при скачивании отчёта",
                    "status_code": response.status_code,
                    "response_text": error_text,
                    "url_used": download_url
                }
        except requests.exceptions.RequestException as e:
            return {
                "success": False,
                "error": f"Ошибка запроса при скачивании отчёта: {e}",
                "url_used": download_url
            }
    
    def _download_from_url(self, url: str) -> Dict:
        """Вспомогательный метод для скачивания по прямому URL"""
        try:
            response = requests.get(
                url,
                headers=self.report_headers,
                timeout=120,
                stream=True
            )
            
            if response.status_code == 200:
                return {
                    "success": True,
                    "data": response.content,
                    "status_code": response.status_code,
                    "format": "xlsx",
                    "url_used": url
                }
            else:
                return {
                    "success": False,
                    "error": f"HTTP ошибка {response.status_code}",
                    "status_code": response.status_code,
                    "url_used": url
                }
        except Exception as e:
            return {
                "success": False,
                "error": f"Ошибка при скачивании: {e}",
                "url_used": url
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
    
    def save_to_excel(self, data: List[Dict], filename: str = "wb_report.xlsx", data_folder: Optional[str] = None):
        """
        Сохранить данные в Excel файл (аналог "Выгрузить в Excel")
        
        Args:
            data: Список словарей с данными для сохранения
            filename: Имя файла Excel
            data_folder: Папка для сохранения (если None, сохраняет в текущую директорию)
        """
        try:
            import pandas as pd
            
            if not data:
                print("⚠ Нет данных для сохранения в Excel")
                return
            
            # Создаём DataFrame из данных
            df = pd.DataFrame(data)
            
            # Определяем полный путь к файлу
            if data_folder:
                # Создаём папку если её нет
                Path(data_folder).mkdir(parents=True, exist_ok=True)
                filepath = Path(data_folder) / filename
            else:
                filepath = Path(filename)
            
            # Сохраняем в Excel (заменяем существующий файл если есть)
            df.to_excel(filepath, index=False, engine='openpyxl')
            print(f"✓ Данные сохранены в Excel файл: {filepath}")
            print(f"  Всего строк: {len(df)}")
            
        except ImportError:
            print("❌ Для сохранения в Excel необходимо установить библиотеки:")
            print("   pip install pandas openpyxl")
        except Exception as e:
            print(f"❌ Ошибка при сохранении в Excel: {e}")
    
    def get_sales_data(self, date_from: str, date_to: str) -> Dict:
        """
        Получить данные о продажах из /api/v1/supplier/sales
        
        Args:
            date_from: Дата начала периода в формате YYYY-MM-DD
            date_to: Дата окончания периода в формате YYYY-MM-DD
        
        Returns:
            Словарь с данными о продажах
        """
        url = "https://statistics-api.wildberries.ru/api/v1/supplier/sales"
        params = {
            "dateFrom": date_from,
            "dateTo": date_to,
            "flag": 0  # 0 - все продажи
        }
        
        try:
            print(f"📊 Получение данных о продажах из /api/v1/supplier/sales...")
            response = requests.get(url, headers=self.headers, params=params, timeout=60)
            
            if response.status_code == 200:
                data = response.json()
                print(f"✓ Получено {len(data)} записей о продажах")
                return {"success": True, "data": data}
            else:
                print(f"⚠ HTTP {response.status_code} при получении продаж: {response.text[:200]}")
                return {"success": False, "error": f"HTTP {response.status_code}", "data": []}
        except Exception as e:
            print(f"❌ Ошибка при получении продаж: {e}")
            return {"success": False, "error": str(e), "data": []}
    
    def get_product_cards(self) -> Dict:
        """
        Получить информацию о товарах из /api/v1/supplier/cards/list
        
        Returns:
            Словарь с данными о товарах (nmId -> информация о товаре)
        """
        url = "https://suppliers-api.wildberries.ru/content/v1/cards/cursor/list"
        
        # Для получения всех карточек используем cursor-based pagination
        all_cards = []
        cursor = None
        cursor_nm_id = None
        
        try:
            print(f"📦 Получение информации о товарах из /api/v1/supplier/cards...")
            
            while True:
                request_body = {
                    "sort": {
                        "cursor": {
                            "limit": 1000
                        },
                        "filter": {
                            "withPhoto": -1
                        }
                    }
                }
                
                if cursor and cursor_nm_id is not None:
                    request_body["sort"]["cursor"]["updatedAt"] = cursor
                    request_body["sort"]["cursor"]["nmID"] = cursor_nm_id
                
                response = requests.post(url, headers=self.headers, json=request_body, timeout=60)
                
                if response.status_code == 200:
                    data = response.json()
                    cards = data.get("data", {}).get("cards", [])
                    if not cards:
                        break
                    
                    all_cards.extend(cards)
                    
                    # Проверяем есть ли ещё данные
                    cursor_data = data.get("data", {}).get("cursor", {})
                    if not cursor_data or not cursor_data.get("updatedAt"):
                        break
                    
                    cursor = cursor_data.get("updatedAt")
                    cursor_nm_id = cursor_data.get("nmID", 0)
                    print(f"  Загружено {len(all_cards)} карточек...")
                else:
                    print(f"⚠ HTTP {response.status_code} при получении карточек: {response.text[:200]}")
                    break
            
            # Создаём словарь nmId -> карточка
            cards_dict = {}
            for card in all_cards:
                nm_id = card.get("nmID")
                if nm_id:
                    cards_dict[nm_id] = card
            
            print(f"✓ Получено {len(cards_dict)} карточек товаров")
            return {"success": True, "data": cards_dict}
            
        except Exception as e:
            print(f"❌ Ошибка при получении карточек: {e}")
            return {"success": False, "error": str(e), "data": {}}
    
    def get_stocks_data(self) -> Dict:
        """
        Получить данные об остатках из /api/v1/supplier/stocks
        
        Returns:
            Словарь с данными об остатках по складам
        """
        url = "https://statistics-api.wildberries.ru/api/v1/supplier/stocks"
        params = {
            "dateFrom": (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        }
        
        try:
            print(f"📦 Получение данных об остатках из /api/v1/supplier/stocks...")
            response = requests.get(url, headers=self.headers, params=params, timeout=60)
            
            if response.status_code == 200:
                data = response.json()
                print(f"✓ Получено {len(data)} записей об остатках")
                return {"success": True, "data": data}
            else:
                print(f"⚠ HTTP {response.status_code} при получении остатков: {response.text[:200]}")
                return {"success": False, "error": f"HTTP {response.status_code}", "data": []}
        except Exception as e:
            print(f"❌ Ошибка при получении остатков: {e}")
            return {"success": False, "error": str(e), "data": []}
    
    def build_combined_report(
        self,
        date_from: str,
        date_to: str,
        sales_data: List[Dict],
        product_cards: Dict,
        stocks_data: List[Dict]
    ) -> List[Dict]:
        """
        Объединить данные из разных источников в один отчёт
        
        Args:
            date_from: Дата начала периода
            date_to: Дата окончания периода
            sales_data: Данные о продажах
            product_cards: Словарь карточек товаров (nmId -> карточка)
            stocks_data: Данные об остатках
        
        Returns:
            Список словарей с объединёнными данными
        """
        print(f"🔗 Объединение данных из разных источников...")
        
        # Группируем продажи по nmId, складу и размеру
        sales_by_key = {}  # (nmId, warehouse, size) -> {ordered: 0, buyouts: 0, ordered_cost: 0, buyouts_sum: 0}
        
        for sale in sales_data:
            nm_id = sale.get("nmId") or sale.get("nm_id")
            # В /api/v1/supplier/sales может не быть warehouseName и techSize
            # Используем пустые строки если нет
            warehouse = sale.get("warehouseName", "") or sale.get("warehouse_name", "")
            size = sale.get("techSize", "") or sale.get("tech_size", "") or sale.get("size", "")
            quantity = sale.get("quantity", 0)
            total_price = sale.get("totalPrice", 0) or sale.get("total_price", 0)
            # В /api/v1/supplier/sales может не быть isRealization, используем другие поля
            is_realization = sale.get("isRealization", False) or sale.get("is_realization", False)
            # Если нет явного флага, считаем что все продажи - это выкупы
            if not any(key in sale for key in ["isRealization", "is_realization"]):
                is_realization = True
            
            key = (nm_id, warehouse, size)
            
            if key not in sales_by_key:
                sales_by_key[key] = {
                    "ordered": 0,
                    "buyouts": 0,
                    "ordered_cost": 0.0,
                    "buyouts_sum": 0.0
                }
            
            # Заказано - это все продажи
            sales_by_key[key]["ordered"] += quantity
            sales_by_key[key]["ordered_cost"] += total_price
            
            # Выкупили - это только isRealization=True (или все, если флаг не указан)
            if is_realization:
                sales_by_key[key]["buyouts"] += quantity
                sales_by_key[key]["buyouts_sum"] += total_price
        
        # Группируем остатки по nmId, складу и размеру
        stocks_by_key = {}  # (nmId, warehouse, size) -> quantity
        
        for stock in stocks_data:
            nm_id = stock.get("nmId")
            
            # Остатки могут быть в массиве warehouses
            warehouses = stock.get("warehouses", [])
            if warehouses:
                for wh in warehouses:
                    warehouse = wh.get("warehouseName", "")
                    quantity = wh.get("quantity", 0)
                    size = stock.get("techSize", "")
                    
                    key = (nm_id, warehouse, size)
                    stocks_by_key[key] = stocks_by_key.get(key, 0) + quantity
            else:
                # Старый формат - остатки напрямую в объекте
                warehouse = stock.get("warehouseName", "")
                size = stock.get("techSize", "")
                quantity = stock.get("quantity", 0)
                
                key = (nm_id, warehouse, size)
                stocks_by_key[key] = stocks_by_key.get(key, 0) + quantity
        
        # Собираем все уникальные ключи (из продаж и остатков)
        all_keys = set(sales_by_key.keys()) | set(stocks_by_key.keys())
        
        # Собираем финальный отчёт
        report_rows = []
        
        for (nm_id, warehouse, size) in all_keys:
            sales_info = sales_by_key.get((nm_id, warehouse, size), {
                "ordered": 0,
                "buyouts": 0,
                "ordered_cost": 0.0,
                "buyouts_sum": 0.0
            })
            # Получаем информацию о товаре
            card = product_cards.get(nm_id, {})
            
            # Извлекаем данные из карточки
            # Пробуем разные варианты названий полей
            brand = card.get("brand", "") or card.get("Бренд", "")
            subject = card.get("subject", "") or card.get("Предмет", "") or card.get("category", "")
            season = card.get("season", "") or card.get("Сезон", "")
            collection = card.get("collection", "") or card.get("Коллекция", "")
            name = card.get("imtName", "") or card.get("imt_name", "") or card.get("Наименование", "") or card.get("title", "")
            supplier_article = card.get("supplierArticle", "") or card.get("supplier_article", "") or card.get("Артикул поставщика", "")
            barcode = ""
            
            # Ищем баркод для данного размера
            sizes = card.get("sizes", []) or card.get("Размеры", [])
            for size_info in sizes:
                if isinstance(size_info, dict):
                    tech_size = size_info.get("techSize") or size_info.get("tech_size") or size_info.get("Размер", "")
                    if tech_size == size or str(tech_size) == str(size):
                        barcode = size_info.get("barcode", "") or size_info.get("Баркод", "")
                        break
            
            # Получаем остаток
            stock_quantity = stocks_by_key.get((nm_id, warehouse, size), 0)
            
            # Создаём строку отчёта
            row = {
                "Бренд": brand,
                "Предмет": subject,
                "Сезон": season,
                "Коллекция": collection,
                "Наименование": name,
                "Артикул поставщика": supplier_article,
                "Номенклатура": nm_id,
                "Баркод": barcode,
                "Размер": size,
                "Контракт": "",  # Не доступно через API
                "Склад": warehouse,
                "Заказано шт": sales_info["ordered"],
                "Заказано себестоимость": sales_info["ordered_cost"],
                "Выкупили шт": sales_info["buyouts"],
                "Выкупили руб": sales_info["buyouts_sum"],
                "Текущий остаток": stock_quantity
            }
            
            report_rows.append(row)
        
        print(f"✓ Собрано {len(report_rows)} строк отчёта")
        return report_rows
    
    def download_report_to_excel(
        self,
        date_from: str,
        date_to: str,
        filename: Optional[str] = None,
        use_detailed_api: bool = True,
        data_folder: Optional[str] = None
    ) -> bool:
        """
        Скачать отчёт и сохранить в Excel
        
        Args:
            date_from: Дата начала периода в формате YYYY-MM-DD
            date_to: Дата окончания периода в формате YYYY-MM-DD
            filename: Имя файла Excel (если не указано, будет сгенерировано автоматически)
            use_detailed_api: Попытаться использовать детализированный API (если доступен)
            data_folder: Папка для сохранения (если None, сохраняет в текущую директорию)
        
        Returns:
            True если успешно, False если ошибка
        """
        print(f"Получение отчёта за период {date_from} - {date_to}...")
        
        # Сначала пробуем собрать данные из разных эндпоинтов
        print("🔧 Пробуем собрать данные из разных эндпоинтов...")
        try:
            # 1. Получаем данные о продажах
            sales_result = self.get_sales_data(date_from, date_to)
            if not sales_result.get("success") or not sales_result.get("data"):
                print("⚠ Не удалось получить данные о продажах, пробуем другие методы...")
                raise Exception("Нет данных о продажах")
            
            sales_data = sales_result.get("data", [])
            if not sales_data:
                print("⚠ Нет данных о продажах за указанный период")
                raise Exception("Нет данных о продажах")
            
            # 2. Получаем информацию о товарах
            cards_result = self.get_product_cards()
            if not cards_result.get("success"):
                print("⚠ Не удалось получить информацию о товарах, пробуем другие методы...")
                raise Exception("Нет данных о товарах")
            
            product_cards = cards_result.get("data", {})
            
            # 3. Получаем данные об остатках
            stocks_result = self.get_stocks_data()
            stocks_data = stocks_result.get("data", []) if stocks_result.get("success") else []
            
            # 4. Объединяем данные
            combined_report = self.build_combined_report(
                date_from=date_from,
                date_to=date_to,
                sales_data=sales_data,
                product_cards=product_cards,
                stocks_data=stocks_data
            )
            
            if combined_report:
                # Сохраняем в Excel
                if not filename:
                    filename = f"wb_report_{date_from}_to_{date_to}.xlsx"
                
                self.save_to_excel(combined_report, filename=filename, data_folder=data_folder)
                print(f"✓ Отчёт успешно собран из разных эндпоинтов и сохранён")
                return True
            else:
                print("⚠ Не удалось собрать отчёт из разных эндпоинтов")
                raise Exception("Не удалось собрать отчёт")
                
        except Exception as e:
            print(f"⚠ Ошибка при сборе данных из разных эндпоинтов: {e}")
            print("Пробуем другие методы...")
        
        # Сначала пробуем получить отчёт через новый API (seller-analytics-api)
        # Пробуем разные типы отчётов
        if use_detailed_api:
            report_types_to_try = [
                "STOCK_HISTORY_REPORT_CSV",  # История остатков - может содержать нужные поля
                "DETAIL_HISTORY_REPORT",  # Воронка продаж (не подходит, но пробуем)
            ]
            
            for report_type in report_types_to_try:
                try:
                    print(f"Пробуем создать отчёт типа '{report_type}' через новый API (seller-analytics-api)...")
                    create_result = self.create_analytics_report(date_from=date_from, date_to=date_to, report_type=report_type)
                    if create_result.get("success"):
                        print(f"✓ Отчёт типа '{report_type}' успешно создан")
                        download_id = create_result.get("downloadId")
                        if download_id:
                            # Ждём немного, чтобы отчёт успел сгенерироваться
                            import time
                            print("⏳ Ожидание генерации отчёта (5 секунд)...")
                            time.sleep(5)
                            
                            # Получаем отчёт
                            report_result = self.get_analytics_report_file(download_id)
                            if report_result.get("success"):
                                zip_data = report_result.get("data")
                                if zip_data and report_result.get("format") == "zip":
                                    # Распаковываем ZIP и конвертируем CSV в Excel
                                    try:
                                        import zipfile
                                        import io
                                        import pandas as pd
                                        
                                        with zipfile.ZipFile(io.BytesIO(zip_data)) as zip_file:
                                            # Ищем CSV файлы в архиве
                                            csv_files = [f for f in zip_file.namelist() if f.endswith('.csv')]
                                            if csv_files:
                                                # Читаем первый CSV файл
                                                csv_file = csv_files[0]
                                                with zip_file.open(csv_file) as f:
                                                    df = pd.read_csv(f, encoding='utf-8-sig')
                                                
                                                # Определяем полный путь к файлу
                                                if not filename:
                                                    filename = f"wb_report_{date_from}_to_{date_to}.xlsx"
                                                
                                                if data_folder:
                                                    Path(data_folder).mkdir(parents=True, exist_ok=True)
                                                    filepath = Path(data_folder) / filename
                                                else:
                                                    filepath = Path(filename)
                                                
                                                # Проверяем структуру данных
                                                # Если это воронка продаж (неправильная структура), пробуем следующий тип отчёта
                                                if 'dt' in df.columns and 'openCardCount' in df.columns:
                                                    print("⚠ Получен отчёт воронки продаж, а не детализация продаж")
                                                    print(f"  Тип отчёта: {report_type}")
                                                    print(f"  Текущие колонки: {', '.join(df.columns.tolist()[:5])}...")
                                                    print("  Нужны колонки: Бренд, Предмет, Сезон, Коллекция, Наименование, Артикул поставщика...")
                                                    print("  Пробуем следующий тип отчёта...")
                                                    # Выходим из вложенных блоков, чтобы перейти к следующему типу отчёта
                                                    raise StopIteration("Пробуем следующий тип отчёта")
                                                elif 'brand' in df.columns or 'subject' in df.columns or 'supplierArticle' in df.columns or 'warehouseName' in df.columns:
                                                    # Правильная структура - сохраняем
                                                    df.to_excel(filepath, index=False, engine='openpyxl')
                                                    print(f"✓ Отчёт сохранён в Excel файл: {filepath}")
                                                    print(f"  Тип отчёта: {report_type}")
                                                    print(f"  Всего строк: {len(df)}")
                                                    print(f"  Колонки: {', '.join(df.columns.tolist()[:10])}...")
                                                    return True
                                                else:
                                                    # Неизвестная структура - сохраняем для проверки
                                                    print(f"⚠ Неизвестная структура данных")
                                                    print(f"  Тип отчёта: {report_type}")
                                                    print(f"  Колонки: {', '.join(df.columns.tolist())}")
                                                    df.to_excel(filepath, index=False, engine='openpyxl')
                                                    print(f"✓ Отчёт сохранён в Excel файл: {filepath}")
                                                    print(f"  Всего строк: {len(df)}")
                                                    return True
                                            else:
                                                print("⚠ В ZIP архиве не найдено CSV файлов")
                                                raise StopIteration("Пробуем следующий тип отчёта")
                                    except StopIteration:
                                        # Пробуем следующий тип отчёта
                                        raise
                                    except Exception as e:
                                        print(f"⚠ Ошибка при обработке ZIP архива: {e}")
                                        raise StopIteration("Пробуем следующий тип отчёта")
                                else:
                                    print(f"⚠ Отчёт получен, но формат не ZIP: {report_result.get('format')}")
                                    raise StopIteration("Пробуем следующий тип отчёта")
                            else:
                                error_msg = report_result.get("error", "Неизвестная ошибка")
                                print(f"⚠ Ошибка при получении отчёта: {error_msg}")
                                raise StopIteration("Пробуем следующий тип отчёта")
                        else:
                            print("⚠ downloadId не получен из ответа")
                            raise StopIteration("Пробуем следующий тип отчёта")
                    else:
                        error_msg = create_result.get("error", "Неизвестная ошибка")
                        print(f"⚠ Ошибка при создании отчёта типа '{report_type}': {error_msg}")
                        if create_result.get("response_text"):
                            print(f"  Ответ сервера: {create_result.get('response_text')[:200]}")
                        # Пробуем следующий тип отчёта
                        continue
                except StopIteration:
                    # Пробуем следующий тип отчёта
                    continue
                except Exception as e:
                    print(f"⚠ Неожиданная ошибка при обработке отчёта типа '{report_type}': {e}")
                    continue
            
            print("⚠ Новый API не вернул правильную структуру данных")
            print("Пробуем другие методы...")
            
            print("⚠ Новый API не вернул правильную структуру данных")
            print("Пробуем другие методы...")
        
        # Пробуем получить отчёт через API reportDetailByPeriod (v1, v2, v5)
        if use_detailed_api:
            print("Пробуем получить отчёт через API reportDetailByPeriod (v1, v2, v5)...")
            report_data = self.get_report_detail(date_from=date_from, date_to=date_to)
            if report_data.get("success"):
                # Если получены бинарные данные Excel
                if report_data.get("format") == "xlsx" and isinstance(report_data.get("data"), bytes):
                    excel_data = report_data.get("data")
                    
                    # Определяем полный путь к файлу
                    if not filename:
                        filename = f"wb_report_{date_from}_to_{date_to}.xlsx"
                    
                    if data_folder:
                        Path(data_folder).mkdir(parents=True, exist_ok=True)
                        filepath = Path(data_folder) / filename
                    else:
                        filepath = Path(filename)
                    
                    # Сохраняем бинарные данные напрямую в файл
                    try:
                        with open(filepath, 'wb') as f:
                            f.write(excel_data)
                        print(f"✓ Отчёт сохранён в Excel файл: {filepath}")
                        print(f"  Размер файла: {len(excel_data)} байт")
                        return True
                    except Exception as e:
                        print(f"❌ Ошибка при сохранении Excel файла: {e}")
                        return False
                else:
                    # Если получены JSON данные, обрабатываем
                    data = report_data.get("data", [])
                    if isinstance(data, list) and data:
                        print(f"✓ Получены данные через детализированный API reportDetailByPeriod ({len(data)} записей)")
                        
                        # Проверяем структуру данных
                        first_item = data[0] if data else {}
                        if isinstance(first_item, dict):
                            # Проверяем наличие нужных полей
                            has_brand = 'brand' in first_item
                            has_subject = 'subject' in first_item
                            has_warehouse = 'warehouseName' in first_item
                            has_supplier_article = 'supplierArticle' in first_item
                            
                            if has_brand or has_subject or has_warehouse or has_supplier_article:
                                print(f"✓ Структура данных соответствует требуемой")
                                print(f"  Колонки: {', '.join(list(first_item.keys())[:10])}...")
                                
                                if not filename:
                                    filename = f"wb_report_{date_from}_to_{date_to}.xlsx"
                                self.save_to_excel(data, filename, data_folder=data_folder)
                                return True
                            else:
                                print(f"⚠ Структура данных не соответствует требуемой")
                                print(f"  Колонки: {', '.join(list(first_item.keys())[:10])}...")
                                print("  Нужны: brand, subject, warehouseName, supplierArticle...")
                        else:
                            print(f"⚠ API вернул данные, но формат неожиданный: {type(data)}")
        
        # Убрано дублирование - уже пробовали выше
            if report_data.get("success"):
                data = report_data.get("data", [])
                if isinstance(data, list):
                    if len(data) > 0:
                        print(f"✓ Получены данные через детализированный API reportDetailByPeriod ({len(data)} записей)")
                        
                        # Проверяем структуру данных
                        first_item = data[0]
                        if isinstance(first_item, dict):
                            # Проверяем наличие нужных полей
                            has_brand = 'brand' in first_item
                            has_subject = 'subject' in first_item
                            has_warehouse = 'warehouseName' in first_item
                            has_supplier_article = 'supplierArticle' in first_item
                            
                            print(f"  Колонки: {', '.join(list(first_item.keys())[:15])}...")
                            
                            if has_brand or has_subject or has_warehouse or has_supplier_article:
                                print(f"✓ Структура данных соответствует требуемой")
                                if not filename:
                                    filename = f"wb_report_{date_from}_to_{date_to}.xlsx"
                                self.save_to_excel(data, filename, data_folder=data_folder)
                                return True
                            else:
                                print(f"⚠ Структура данных не соответствует требуемой")
                                print("  Нужны: brand, subject, warehouseName, supplierArticle...")
                                # Но всё равно сохраняем, возможно данные правильные, просто названия полей другие
                                print("  Сохраняем данные для проверки...")
                                if not filename:
                                    filename = f"wb_report_{date_from}_to_{date_to}.xlsx"
                                self.save_to_excel(data, filename, data_folder=data_folder)
                                return True
                        else:
                            print(f"⚠ Первый элемент не является словарём: {type(first_item)}")
                    else:
                        print(f"⚠ API вернул пустой список")
                else:
                    print(f"⚠ API вернул данные, но формат неожиданный: {type(data)}")
                    if isinstance(data, dict):
                        print(f"  Ключи в ответе: {list(data.keys())}")
            else:
                error_msg = report_data.get("error", "Неизвестная ошибка")
                print(f"⚠ Ошибка при получении детализированного отчёта: {error_msg}")
                if report_data.get("response_text"):
                    print(f"  Ответ сервера: {report_data.get('response_text')[:200]}")
        
        # Если детализированный API не сработал, возвращаем ошибку
        print("❌ Не удалось получить детализированный отчёт ни через один из доступных API")
        print("Проверьте токен и доступность API")
        return False
    
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
        print(f"\n❌ Не удалось скачать отчёт")
        print("Проверьте токен и доступность API")
    
    # Другие примеры использования:
    # Скачать отчёт за конкретный период в Excel:
    # parser.download_report_to_excel(date_from="2024-01-01", date_to="2024-01-31")


if __name__ == "__main__":
    main()

