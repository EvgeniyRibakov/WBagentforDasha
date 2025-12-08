"""
Основной браузерный агент для работы с Wildberries Seller
Аналог Ozon проекта, адаптированный под Wildberries
"""

import os
import time
import random
from pathlib import Path
from typing import Optional, List, Dict
from datetime import datetime, timedelta

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from loguru import logger

from src.config.settings import Settings


class BrowserAgent:
    """Агент для автоматизации работы с Wildberries Seller через браузер"""
    
    # Список кабинетов для обработки
    CABINETS = [
        {"name": "MAU", "id": "53607"},
        {"name": "MAB", "id": "121614"},
        {"name": "MMA", "id": "174711"},
        {"name": "cosmo", "id": "224650"},
        {"name": "dreamlab", "id": "1140223"},
        {"name": "beautylab", "id": "4428365"},
    ]
    
    def __init__(self, settings: Settings):
        """
        Инициализация агента
        
        Args:
            settings: Настройки приложения
        """
        self.settings = settings
        self.driver: Optional[webdriver.Chrome] = None
        self.downloads_dir = Path(settings.downloads_dir).absolute()
        self.downloads_dir.mkdir(parents=True, exist_ok=True)
        self.data_dir = Path("data")
        self.data_dir.mkdir(parents=True, exist_ok=True)
        
    def start_browser(self) -> None:
        """Запуск браузера с профилем Chrome"""
        options = Options()
        
        # НЕ используем headless - Wildberries может блокировать
        # options.add_argument("--headless=new")  # НЕ включать!
        
        # Базовые настройки
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-blink-features=AutomationControlled")
        
        # КРИТИЧЕСКИ ВАЖНО: Использование профиля Chrome
        if self.settings.chrome_user_data_dir:
            user_data_path = Path(self.settings.chrome_user_data_dir)
            path_str = str(user_data_path.absolute())
            
            if os.path.exists(path_str):
                # Указываем путь к папке User Data
                options.add_argument(f'--user-data-dir={path_str}')
                # Указываем имя профиля
                options.add_argument(f"--profile-directory={self.settings.chrome_profile_name}")
                
                logger.info(f"Используется профиль Chrome: {path_str} / {self.settings.chrome_profile_name}")
                
                # Проверяем, что папка профиля существует
                profile_path = user_data_path / self.settings.chrome_profile_name
                if os.path.exists(str(profile_path.absolute())):
                    logger.success(f"✓ Папка профиля найдена: {profile_path}")
                else:
                    logger.warning(f"⚠ Папка профиля не найдена: {profile_path}")
            else:
                logger.warning(f"⚠ Путь к профилю Chrome не существует: {path_str}")
        
        # Убираем признаки автоматизации
        options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        options.add_experimental_option('useAutomationExtension', False)
        
        # Настройка скачивания файлов
        downloads_path = str(self.downloads_dir.absolute())
        prefs = {
            "download.default_directory": downloads_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_setting_values.automatic_downloads": 1,
        }
        options.add_experimental_option("prefs", prefs)
        
        # User-Agent
        options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        
        # Запускаем браузер
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        
        # Убираем признаки автоматизации через JavaScript
        self.driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
            'source': '''
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                })
            '''
        })
        
        logger.success("✓ Браузер запущен с профилем Chrome")
    
    def wait_for_element(self, by: By, value: str, timeout: int = 20) -> Optional[object]:
        """
        Ожидание появления элемента
        
        Args:
            by: Способ поиска (By.ID, By.CSS_SELECTOR и т.д.)
            value: Значение для поиска
            timeout: Таймаут ожидания в секундах
            
        Returns:
            WebElement или None если не найден
        """
        try:
            element = WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((by, value))
            )
            return element
        except TimeoutException:
            logger.error(f"Элемент не найден: {by}={value} (таймаут {timeout}с)")
            return None
    
    def wait_for_clickable(self, by: By, value: str, timeout: int = 20) -> Optional[object]:
        """
        Ожидание кликабельности элемента
        
        Args:
            by: Способ поиска
            value: Значение для поиска
            timeout: Таймаут ожидания в секундах
            
        Returns:
            WebElement или None если не найден
        """
        try:
            element = WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((by, value))
            )
            return element
        except TimeoutException:
            logger.error(f"Элемент не кликабелен: {by}={value} (таймаут {timeout}с)")
            return None
    
    def wait_for_dynamic_content(self, timeout: int = 20) -> None:
        """
        Ожидание полной загрузки динамического контента
        
        Args:
            timeout: Таймаут ожидания в секундах
        """
        try:
            # Ожидаем готовности документа
            WebDriverWait(self.driver, timeout).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            
            # Дополнительная проверка на отсутствие загрузчиков
            WebDriverWait(self.driver, timeout).until(
                lambda d: len(d.find_elements(By.CSS_SELECTOR, ".loader, .spinner, [class*='loading']")) == 0
            )
            
            logger.debug("Динамический контент загружен")
        except TimeoutException:
            logger.warning("Таймаут ожидания загрузки динамического контента")
    
    def click_button(self, by: By, value: str, description: str = "") -> bool:
        """
        Клик по кнопке с ожиданием и задержками
        
        Args:
            by: Способ поиска
            value: Значение для поиска
            description: Описание действия для логов
            
        Returns:
            True если успешно, False при ошибке
        """
        try:
            # Ожидание перед кликом
            time.sleep(self.settings.delay_before_click)
            
            element = self.wait_for_clickable(by, value)
            if not element:
                logger.error(f"Не удалось найти кликабельный элемент: {description or value}")
                return False
            
            # Прокрутка к элементу
            self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
            time.sleep(0.5)
            
            # Клик
            element.click()
            logger.info(f"✓ Клик выполнен: {description or value}")
            
            # Ожидание после клика
            time.sleep(self.settings.delay_after_click)
            
            return True
            
        except Exception as e:
            logger.error(f"Ошибка при клике на {description or value}: {e}")
            return False
    
    def fill_input(self, by: By, value: str, text: str, description: str = "", clear_first: bool = True) -> bool:
        """
        Заполнение поля ввода с имитацией человеческого ввода
        
        Args:
            by: Способ поиска
            value: Значение для поиска
            text: Текст для ввода
            description: Описание действия для логов
            clear_first: Очистить поле перед вводом
            
        Returns:
            True если успешно, False при ошибке
        """
        try:
            # Ожидание перед вводом
            time.sleep(self.settings.delay_before_type)
            
            element = self.wait_for_element(by, value)
            if not element:
                logger.error(f"Не удалось найти поле ввода: {description or value}")
                return False
            
            # Прокрутка к элементу
            self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
            time.sleep(0.5)
            
            # Очистка поля если нужно
            if clear_first:
                element.clear()
                time.sleep(0.3)
            
            # Имитация человеческого ввода (посимвольный ввод)
            for char in text:
                element.send_keys(char)
                time.sleep(random.uniform(0.05, self.settings.delay_between_keys))
            
            logger.info(f"✓ Поле заполнено: {description or value} = {text}")
            
            # Ожидание после ввода
            time.sleep(self.settings.delay_after_type)
            
            return True
            
        except Exception as e:
            logger.error(f"Ошибка при заполнении поля {description or value}: {e}")
            return False
    
    def navigate_to_url(self, url: str) -> bool:
        """
        Переход на URL с ожиданием загрузки
        
        Args:
            url: URL для перехода
            
        Returns:
            True если успешно, False при ошибке
        """
        try:
            logger.info(f"Переход на страницу: {url}")
            self.driver.get(url)
            
            # Ожидание загрузки страницы
            time.sleep(self.settings.delay_page_load)
            self.wait_for_dynamic_content()
            
            logger.success(f"✓ Страница загружена: {url}")
            return True
            
        except Exception as e:
            logger.error(f"Ошибка при переходе на {url}: {e}")
            return False
    
    def check_session(self) -> bool:
        """
        Проверка активной сессии (нет формы авторизации)
        
        Returns:
            True если сессия активна, False если требуется авторизация
        """
        try:
            # Проверяем наличие элементов авторизации
            auth_elements = self.driver.find_elements(
                By.CSS_SELECTOR, 
                "input[type='password'], input[name*='password'], input[id*='password'], "
                "button[type='submit'][class*='login'], form[class*='auth']"
            )
            
            if auth_elements:
                logger.error("⚠ Обнаружена форма авторизации - сессия истекла!")
                return False
            
            logger.debug("✓ Сессия активна")
            return True
            
        except Exception as e:
            logger.warning(f"Не удалось проверить сессию: {e}")
            return True  # Предполагаем что сессия активна
    
    def process_cabinet(self, cabinet: Dict[str, str]) -> bool:
        """
        Обработка одного кабинета
        
        Args:
            cabinet: Словарь с данными кабинета {"name": "...", "id": "..."}
            
        Returns:
            True если успешно, False при ошибке
        """
        cabinet_name = cabinet["name"]
        cabinet_id = cabinet["id"]
        
        logger.info(f"{'='*60}")
        logger.info(f"Обработка кабинета: {cabinet_name} (ID: {cabinet_id})")
        logger.info(f"{'='*60}")
        
        try:
            # Шаг 2.1: Ввод ID кабинета
            logger.info("Шаг 2.1: Ввод ID кабинета")
            time.sleep(self.settings.delay_between_actions)
            
            if not self.fill_input(
                By.ID, 
                "suppliers-search", 
                cabinet_id,
                description=f"Поле поиска кабинета ({cabinet_name})"
            ):
                logger.error(f"Не удалось ввести ID кабинета {cabinet_name}")
                return False
            
            # Задержка для поиска и загрузки результатов
            time.sleep(self.settings.delay_between_actions)
            
            # Шаг 2.2: Настройка периода отчёта
            logger.info("Шаг 2.2: Настройка периода отчёта")
            time.sleep(self.settings.delay_between_actions)
            
            # Клик по кнопке календаря
            calendar_button = self.wait_for_clickable(
                By.CSS_SELECTOR,
                "button.Date-input__icon-button__WnbzIWQzsq",
                timeout=10
            )
            if not calendar_button:
                logger.error("Не удалось найти кнопку календаря")
                return False
            
            time.sleep(self.settings.delay_before_click)
            calendar_button.click()
            time.sleep(self.settings.delay_after_click)
            time.sleep(self.settings.delay_between_actions)
            
            # Вчерашняя дата в формате DD.MM.YYYY
            yesterday = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
            
            # Заполнение поля начала периода
            start_date_element = self.wait_for_element(By.ID, "startDate", timeout=10)
            if not start_date_element:
                logger.error("Не удалось найти поле startDate")
                return False
            
            time.sleep(self.settings.delay_before_type)
            start_date_element.clear()
            time.sleep(0.3)
            for char in yesterday:
                start_date_element.send_keys(char)
                time.sleep(random.uniform(0.05, self.settings.delay_between_keys))
            time.sleep(self.settings.delay_after_type)
            time.sleep(self.settings.delay_between_actions)
            
            # Заполнение поля окончания периода
            end_date_element = self.wait_for_element(By.ID, "endDate", timeout=10)
            if not end_date_element:
                logger.error("Не удалось найти поле endDate")
                return False
            
            time.sleep(self.settings.delay_before_type)
            end_date_element.clear()
            time.sleep(0.3)
            for char in yesterday:
                end_date_element.send_keys(char)
                time.sleep(random.uniform(0.05, self.settings.delay_between_keys))
            time.sleep(self.settings.delay_after_type)
            
            # Клик по кнопке "Сохранить"
            save_button = self.wait_for_clickable(
                By.CSS_SELECTOR,
                "button.Button-link--main__bEAy5pip1O[type='submit']",
                timeout=10
            )
            if not save_button:
                logger.error("Не удалось найти кнопку 'Сохранить'")
                return False
            
            time.sleep(self.settings.delay_before_click)
            save_button.click()
            logger.info("✓ Период сохранён")
            time.sleep(self.settings.delay_after_click)
            time.sleep(self.settings.delay_between_actions)
            
            # Шаг 2.3: Выгрузка отчёта
            logger.info("Шаг 2.3: Выгрузка отчёта")
            time.sleep(3)  # Ожидание после сохранения периода (как указано в алгоритме)
            time.sleep(self.settings.delay_between_actions)
            
            # Запоминаем время перед скачиванием для поиска нового файла
            files_before = set(self.downloads_dir.glob("*.xlsx"))
            files_before_time = {f: f.stat().st_mtime for f in files_before}
            
            # Клик по кнопке "Выгрузить в Excel"
            # Используем XPath так как CSS :has() может не поддерживаться
            download_button = self.wait_for_clickable(
                By.XPATH,
                "//button[.//span[contains(text(), 'Выгрузить в Excel')]]",
                timeout=10
            )
            
            # Альтернативный селектор по классу
            if not download_button:
                download_button = self.wait_for_clickable(
                    By.CSS_SELECTOR,
                    "button.Button-link__1abzU3JUeb.Button-link--button-big__Bi4mHiOkNS",
                    timeout=10
                )
            
            if not download_button:
                logger.error("Не удалось найти кнопку 'Выгрузить в Excel'")
                return False
            
            time.sleep(self.settings.delay_before_click)
            download_button.click()
            logger.info("✓ Клик по кнопке скачивания выполнен")
            time.sleep(self.settings.delay_after_click)
            
            # Шаг 2.4: Ожидание скачивания файла
            logger.info("Шаг 2.4: Ожидание скачивания файла")
            downloaded_file = self._wait_for_download(timeout=60)
            
            if not downloaded_file:
                logger.error(f"Файл не был скачан для кабинета {cabinet_name}")
                return False
            
            logger.success(f"✓ Файл скачан: {downloaded_file.name}")
            
            # Обработка файла
            return self._process_downloaded_file(downloaded_file, cabinet_name, yesterday)
            
        except Exception as e:
            logger.error(f"Ошибка при обработке кабинета {cabinet_name}: {e}")
            logger.exception("Детали ошибки:")
            return False
    
    def _wait_for_download(self, timeout: int = 60) -> Optional[Path]:
        """
        Ожидание скачивания нового файла
        
        Args:
            timeout: Таймаут ожидания в секундах
            
        Returns:
            Path к скачанному файлу или None
        """
        start_time = time.time()
        initial_files = {f.name: f.stat().st_mtime for f in self.downloads_dir.glob("*.xlsx")}
        
        logger.info("Ожидание скачивания файла...")
        
        while time.time() - start_time < timeout:
            # Получаем текущие файлы
            current_files = list(self.downloads_dir.glob("*.xlsx"))
            
            # Ищем новые файлы или файлы с изменённым временем
            for file in current_files:
                try:
                    file_name = file.name
                    file_time = file.stat().st_mtime
                    file_size = file.stat().st_size
                    
                    # Проверяем что файл новый или изменён
                    is_new = file_name not in initial_files
                    is_modified = file_name in initial_files and file_time > initial_files[file_name]
                    
                    if (is_new or is_modified) and file_size > 0:
                        # Проверяем что файл не заблокирован (ещё скачивается)
                        # Делаем несколько проверок размера с задержкой
                        size1 = file.stat().st_size
                        time.sleep(1)
                        size2 = file.stat().st_size
                        time.sleep(1)
                        size3 = file.stat().st_size
                        
                        # Если размер не меняется - файл скачан
                        if size1 == size2 == size3 and size1 > 0:
                            logger.info(f"Найден новый файл: {file.name} (размер: {size1} байт)")
                            return file
                except (OSError, PermissionError) as e:
                    # Файл может быть заблокирован, продолжаем ждать
                    logger.debug(f"Файл {file.name} заблокирован, ожидание...")
                    continue
                except Exception as e:
                    logger.debug(f"Ошибка при проверке файла {file.name}: {e}")
                    continue
            
            time.sleep(1)
        
        logger.error(f"Таймаут ожидания скачивания файла ({timeout}с)")
        return None
    
    def _process_downloaded_file(
        self, 
        file_path: Path, 
        cabinet_name: str, 
        date_str: str
    ) -> bool:
        """
        Обработка скачанного файла: переименование, замена первой строки, копирование
        
        Args:
            file_path: Путь к скачанному файлу
            cabinet_name: Название кабинета
            date_str: Дата в формате DD.MM.YYYY
            
        Returns:
            True если успешно, False при ошибке
        """
        try:
            import pandas as pd
            import openpyxl
            import shutil
            
            logger.info(f"Обработка файла: {file_path.name}")
            
            # Шаг 1: Переименование файла
            new_name = f"{cabinet_name.lower()}_{date_str}.xlsx"
            new_path = self.downloads_dir / new_name
            
            if file_path != new_path:
                file_path.rename(new_path)
                logger.info(f"✓ Файл переименован: {new_name}")
                file_path = new_path
            
            # Шаг 2: Замена первой строки
            logger.info("Замена первой строки из example_first_stroke.XLSX")
            
            # Загружаем шаблон первой строки
            template_path = Path("example_first_stroke.XLSX")
            if not template_path.exists():
                logger.error(f"Файл шаблона не найден: {template_path}")
                return False
            
            # Читаем первую строку из шаблона
            template_df = pd.read_excel(template_path, nrows=1)
            template_row = template_df.iloc[0].to_dict()
            
            # Читаем скачанный файл
            df = pd.read_excel(file_path)
            
            # Удаляем первую строку
            df = df.iloc[1:].reset_index(drop=True)
            
            # Создаём новую первую строку из шаблона
            new_first_row = pd.DataFrame([template_row])
            
            # Объединяем: новая первая строка + остальные данные
            df = pd.concat([new_first_row, df], ignore_index=True)
            
            # Сохраняем изменения
            df.to_excel(file_path, index=False, engine='openpyxl')
            logger.success("✓ Первая строка заменена")
            
            # Шаг 3: Создание резервной копии в data/
            data_folder = self.data_dir / date_str
            data_folder.mkdir(parents=True, exist_ok=True)
            
            backup_path = data_folder / new_name
            shutil.copy2(file_path, backup_path)
            logger.success(f"✓ Резервная копия создана: {backup_path}")
            
            return True
            
        except Exception as e:
            logger.error(f"Ошибка при обработке файла: {e}")
            logger.exception("Детали ошибки:")
            return False
    
    def execute_flow(self) -> bool:
        """
        Выполнение основного потока работы для всех кабинетов
        
        Returns:
            True если все кабинеты обработаны успешно, False при ошибке
        """
        try:
            logger.info("Начало выполнения основного потока")
            
            # Шаг 1: Запуск браузера
            self.start_browser()
            
            # Шаг 2: Переход на стартовую страницу
            if not self.navigate_to_url(self.settings.wildberries_start_url):
                logger.error("Не удалось перейти на стартовую страницу")
                return False
            
            # Шаг 3: Проверка сессии
            if not self.check_session():
                logger.error("Сессия истекла или требуется авторизация!")
                logger.error("Пожалуйста, авторизуйтесь вручную в профиле Chrome и запустите скрипт снова")
                return False
            
            # Шаг 4: Обработка каждого кабинета
            for cabinet in self.CABINETS:
                if not self.process_cabinet(cabinet):
                    logger.error(f"Ошибка при обработке кабинета {cabinet['name']}")
                    logger.error("Остановка выполнения")
                    return False
                
                # Возврат на стартовую страницу перед следующим кабинетом
                if cabinet != self.CABINETS[-1]:  # Не для последнего кабинета
                    logger.info("Возврат на стартовую страницу для следующего кабинета")
                    if not self.navigate_to_url(self.settings.wildberries_start_url):
                        logger.error("Не удалось вернуться на стартовую страницу")
                        return False
            
            logger.success("✓ Все кабинеты успешно обработаны!")
            return True
            
        except Exception as e:
            logger.error(f"Критическая ошибка в основном потоке: {e}")
            logger.exception("Детали ошибки:")
            return False
        
        finally:
            # Закрытие браузера
            if self.driver:
                logger.info("Закрытие браузера")
                self.driver.quit()
