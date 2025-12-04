# Парсер продаж Wildberries API

Скрипты для автоматизации получения данных о продажах из API Wildberries и скачивания детализированных отчётов (аналог кнопки "Выгрузить в Excel").

## Установка

1. Установите зависимости:
```bash
pip install -r requirements.txt
```

## Получение API токена

1. Войдите в личный кабинет продавца на Wildberries
2. Перейдите в раздел «Настройки» → «Доступ к API»
3. Создайте новый токен с доступом к «Статистика»
4. Сохраните токен в безопасном месте

## Использование

### Вариант 1: Через .env файл (рекомендуется)

1. Создайте файл `.env` в корне проекта (можно скопировать `.env_sample`)
2. Добавьте в него ваш токен:
```
WB_API_TOKEN=ваш_токен_здесь
```
3. Запустите скрипт:
```bash
# Скачать детализированный отчёт за вчера в Excel
python wb_report_downloader.py

# Или получить базовые данные о продажах
python wb_sales_parser.py
```

**Важно:** Файл `.env` уже добавлен в .gitignore, чтобы не попасть в репозиторий.

### Вариант 2: Через переменную окружения

```bash
# Windows PowerShell
$env:WB_API_TOKEN="ваш_токен"
python wb_sales_parser.py

# Windows CMD
set WB_API_TOKEN=ваш_токен
python wb_sales_parser.py

# Linux/Mac
export WB_API_TOKEN="ваш_токен"
python wb_sales_parser.py
```

### Вариант 3: Ввод токена при запуске

Просто запустите скрипт, и он запросит токен:
```bash
python wb_sales_parser.py
```

## Автоматизация скачивания отчётов

### Скачать отчёт за вчерашний день

```bash
python wb_report_downloader.py --yesterday
```

### Скачать отчёт за указанный период

```bash
python wb_report_downloader.py --period 2024-12-01 2024-12-03
```

### Настройка автоматического запуска (Windows)

1. Откройте "Планировщик заданий" (Task Scheduler)
2. Создайте новое задание
3. Установите триггер (например, ежедневно в 9:00)
4. В действии укажите:
   - Программа: `python`
   - Аргументы: `C:\путь\к\проекту\wb_report_downloader.py --yesterday`
   - Рабочая папка: `C:\путь\к\проекту`

## Использование в коде

```python
from wb_sales_parser import WBSalesParser

# Создаем парсер
parser = WBSalesParser(api_token="ваш_токен")

# Скачать детализированный отчёт за вчера в Excel (аналог "Выгрузить в Excel")
parser.download_report_to_excel(
    date_from="2024-12-03",
    date_to="2024-12-03",
    filename="report.xlsx"
)

# Получить детализированный отчёт за период
report_data = parser.get_report_detail(
    date_from="2024-12-01",
    date_to="2024-12-03"
)

# Получить продажи за последние 7 дней
sales_data = parser.get_sales_last_days(days=7)

# Получить продажи за конкретный период
sales_data = parser.get_sales(
    date_from="2024-01-01",
    date_to="2024-01-31"
)

# Вывести сводку
parser.print_sales_summary(sales_data)

# Сохранить в JSON
parser.save_to_json(sales_data["data"], "sales.json")

# Сохранить в Excel
parser.save_to_excel(sales_data["data"], "sales.xlsx")
```

## Доступные методы API

### Базовые данные о продажах (`/api/v1/supplier/sales`)
- `dateFrom` - Дата начала периода (формат: YYYY-MM-DD)
- `dateTo` - Дата окончания периода (формат: YYYY-MM-DD)
- `flag` - Флаг фильтрации (0 - все продажи, 1 - только новые)

### Детализированный отчёт (`/api/v1/supplier/reportDetailByPeriod`)
Аналог кнопки "Выгрузить в Excel" на странице аналитики продаж.
- `dateFrom` - Дата начала периода (формат: YYYY-MM-DD)
- `dateTo` - Дата окончания периода (формат: YYYY-MM-DD)
- `limit` - Лимит записей (по умолчанию 100000)
- `rrdid` - Идентификатор отчёта (опционально)

## Формат данных

API возвращает массив объектов с информацией о продажах, включая:
- Артикул (nmId)
- Количество (quantity)
- Общая цена (totalPrice)
- Дата продажи
- И другие поля

Подробнее: https://dev.wildberries.ru/

