"""
Основной браузерный агент для работы с Wildberries Seller
Аналог Ozon проекта, адаптированный под Wildberries
"""

import os
import time
import random
import subprocess
import psutil
from pathlib import Path
from typing import Optional, List, Dict
from datetime import datetime, timedelta

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
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
        
    def _kill_chrome_processes(self) -> None:
        """Принудительное закрытие всех процессов Chrome и ChromeDriver"""
        logger.info("Закрытие всех процессов Chrome и ChromeDriver...")
        
        processes_killed = 0
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    proc_name = proc.info['name'].lower()
                    if 'chrome' in proc_name or 'chromedriver' in proc_name:
                        logger.info(f"Завершение процесса: {proc.info['name']} (PID: {proc.info['pid']})")
                        proc.kill()
                        processes_killed += 1
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
            
            if processes_killed > 0:
                logger.info(f"✓ Завершено процессов: {processes_killed}")
                time.sleep(2)  # Даём время процессам завершиться
            else:
                logger.info("Процессы Chrome не найдены")
                
        except Exception as e:
            logger.warning(f"Ошибка при закрытии процессов Chrome: {e}")
            # Альтернативный способ через taskkill (Windows)
            try:
                subprocess.run(['taskkill', '/F', '/IM', 'chrome.exe'], 
                             capture_output=True, timeout=5)
                subprocess.run(['taskkill', '/F', '/IM', 'chromedriver.exe'], 
                             capture_output=True, timeout=5)
                logger.info("Попытка закрыть Chrome через taskkill выполнена")
                time.sleep(2)
            except Exception as e2:
                logger.warning(f"Не удалось закрыть через taskkill: {e2}")
    
    def start_browser(self) -> None:
        """Запуск браузера с профилем Chrome"""
        logger.info("Запуск Chrome...")
        
        options = Options()
        
        # Базовые аргументы
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-blink-features=AutomationControlled")
        
        # Профиль Chrome
        if self.settings.chrome_user_data_dir:
            user_data_dir = os.path.expandvars(self.settings.chrome_user_data_dir)
            path_str = str(Path(user_data_dir).absolute())
            
            if os.path.exists(path_str):
                # Используем абсолютный путь
                options.add_argument(f'--user-data-dir={path_str}')
                options.add_argument(f'--profile-directory={self.settings.chrome_profile_name}')
                logger.info(f"Профиль: {path_str} / {self.settings.chrome_profile_name}")
            else:
                logger.error(f"❌ Путь к профилю не существует: {path_str}")
                raise FileNotFoundError(f"Профиль Chrome не найден: {path_str}")
        
        # Настройка скачивания
        downloads_path = str(self.downloads_dir.absolute())
        prefs = {
            "download.default_directory": downloads_path,
            "download.prompt_for_download": False,
        }
        options.add_experimental_option("prefs", prefs)
        
        # Убираем признаки автоматизации
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        
        # Запуск браузера
        logger.info("Инициализация ChromeDriver...")
        service = Service(ChromeDriverManager().install())
        
        logger.info("Запуск Chrome...")
        try:
            self.driver = webdriver.Chrome(service=service, options=options)
        except Exception as e:
            error_msg = str(e)
            logger.error(f"❌ Ошибка запуска Chrome: {error_msg}")
            
            # Если ошибка DevToolsActivePort - пробуем альтернативный способ
            if "DevToolsActivePort" in error_msg:
                logger.info("Пробуем альтернативный способ запуска...")
                # Пробуем без некоторых аргументов
                options_alt = Options()
                options_alt.add_argument("--no-sandbox")
                if self.settings.chrome_user_data_dir:
                    user_data_dir = os.path.expandvars(self.settings.chrome_user_data_dir)
                    path_str = str(Path(user_data_dir).absolute())
                    options_alt.add_argument(f'--user-data-dir={path_str}')
                    options_alt.add_argument(f'--profile-directory={self.settings.chrome_profile_name}')
                downloads_path = str(self.downloads_dir.absolute())
                prefs = {
                    "download.default_directory": downloads_path,
                    "download.prompt_for_download": False,
                }
                options_alt.add_experimental_option("prefs", prefs)
                self.driver = webdriver.Chrome(service=service, options=options_alt)
            else:
                raise
        
        logger.success("✓ Браузер запущен")
        
        # Максимизируем окно
        try:
            self.driver.maximize_window()
        except:
            pass
        
        time.sleep(2)
        
        # СРАЗУ открываем страницу Wildberries (прямо в коде)
        wb_url = "https://seller.wildberries.ru/analytics-reports/sales"
        logger.info(f"Открытие страницы: {wb_url}")
        self.driver.get(wb_url)
        
        # Ждём загрузки
        time.sleep(5)
        logger.success(f"✓ Страница открыта: {self.driver.current_url}")
    
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
            if not self.driver:
                logger.error("Браузер не запущен!")
                return False
            
            logger.info(f"Открытие страницы в адресной строке: {url}")
            
            # Принудительный переход на URL
            self.driver.get(url)
            
            # Ожидание загрузки страницы
            logger.info("Ожидание загрузки страницы...")
            time.sleep(self.settings.delay_page_load)
            
            # Ожидание готовности документа
            WebDriverWait(self.driver, 20).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            
            self.wait_for_dynamic_content()
            
            # Проверяем текущий URL
            current_url = self.driver.current_url
            logger.info(f"Текущий URL после перехода: {current_url}")
            
            # Проверяем, что мы на правильной странице
            if url in current_url or "wildberries.ru" in current_url:
                logger.success(f"✓ Страница загружена: {current_url}")
                return True
            else:
                logger.warning(f"⚠ Перешли на другую страницу: {current_url}")
                logger.warning(f"Ожидалось: {url}")
                # Пробуем перейти ещё раз
                logger.info("Повторная попытка перехода...")
                self.driver.get(url)
                time.sleep(self.settings.delay_page_load)
                current_url = self.driver.current_url
                if url in current_url or "wildberries.ru" in current_url:
                    logger.success(f"✓ Страница загружена после повторной попытки: {current_url}")
                    return True
                else:
                    logger.error(f"Не удалось перейти на нужную страницу. Текущий URL: {current_url}")
                    return False
            
        except Exception as e:
            logger.error(f"Ошибка при переходе на {url}: {e}")
            logger.exception("Детали ошибки:")
            return False
    
    def check_session(self) -> bool:
        """
        Проверка активной сессии (нет формы авторизации)
        
        Returns:
            True если сессия активна, False если требуется авторизация
        """
        try:
            current_url = self.driver.current_url
            logger.info(f"Проверка сессии. Текущий URL: {current_url}")
            
            # Проверяем наличие элементов авторизации
            auth_selectors = [
                "input[type='password']",
                "input[name*='password']",
                "input[id*='password']",
                "input[placeholder*='парол']",
                "input[placeholder*='Парол']",
                "button[type='submit'][class*='login']",
                "button[class*='auth']",
                "button:contains('Войти')",
                "button:contains('Вход')",
                "form[class*='auth']",
                "form[class*='login']",
            ]
            
            auth_found = False
            for selector in auth_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if elements:
                        logger.warning(f"⚠ Найден элемент авторизации: {selector}")
                        auth_found = True
                        break
                except:
                    continue
            
            # Проверяем наличие элементов страницы отчётов (признак успешной авторизации)
            report_selectors = [
                "input[id='suppliers-search']",
                "input[name='suppliers-search']",
                "input[placeholder*='ИНН']",
                "input[placeholder*='ID']",
                "button:contains('Выгрузить в Excel')",
                "button[class*='Date-input__icon-button']",
            ]
            
            report_found = False
            for selector in report_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if elements:
                        logger.debug(f"✓ Найден элемент страницы отчётов: {selector}")
                        report_found = True
                        break
                except:
                    continue
            
            # Если найдены элементы авторизации И нет элементов отчётов - требуется авторизация
            if auth_found and not report_found:
                logger.error("⚠ Обнаружена форма авторизации - сессия истекла или требуется вход!")
                return False
            
            # Если найдены элементы отчётов - авторизация успешна
            if report_found:
                logger.success("✓ Сессия активна, страница отчётов загружена")
                return True
            
            # Если ничего не найдено - проверяем URL
            if "login" in current_url.lower() or "auth" in current_url.lower():
                logger.error("⚠ URL указывает на страницу авторизации")
                return False
            
            logger.warning("⚠ Не удалось определить статус авторизации, предполагаем что требуется авторизация")
            return False
            
        except Exception as e:
            logger.warning(f"Ошибка при проверке сессии: {e}")
            return False  # В случае ошибки предполагаем что требуется авторизация
    
    def handle_authorization(self) -> bool:
        """
        Автоматическая авторизация на Wildberries
        
        Returns:
            True если авторизация успешна, False при ошибке
        """
        try:
            if not self.settings.password:
                logger.error("Пароль не указан в настройках")
                return False
            
            logger.info("Поиск формы авторизации...")
            time.sleep(2)
            
            # Ищем поле для телефона/email
            phone_email_selectors = [
                "input[type='tel']",
                "input[name*='phone']",
                "input[id*='phone']",
                "input[name*='email']",
                "input[id*='email']",
                "input[placeholder*='Телефон']",
                "input[placeholder*='телефон']",
                "input[placeholder*='Email']",
                "input[placeholder*='email']",
            ]
            
            phone_email_field = None
            for selector in phone_email_selectors:
                try:
                    field = self.driver.find_element(By.CSS_SELECTOR, selector)
                    if field.is_displayed():
                        phone_email_field = field
                        logger.info(f"✓ Найдено поле для телефона/email: {selector}")
                        break
                except:
                    continue
            
            if not phone_email_field:
                logger.error("❌ Не найдено поле для телефона/email")
                return False
            
            # Вводим телефон или email
            login_value = self.settings.phone_number or self.settings.email
            if not login_value:
                logger.error("❌ Не указан телефон или email в настройках")
                return False
            
            logger.info("Ввод телефона/email...")
            phone_email_field.clear()
            time.sleep(self.settings.delay_before_type)
            
            for char in login_value:
                phone_email_field.send_keys(char)
                time.sleep(self.settings.delay_between_keys)
            
            time.sleep(self.settings.delay_after_type)
            
            # Ищем поле для пароля
            password_selectors = [
                "input[type='password']",
                "input[name*='password']",
                "input[id*='password']",
                "input[placeholder*='парол']",
                "input[placeholder*='Парол']",
            ]
            
            password_field = None
            for selector in password_selectors:
                try:
                    field = self.driver.find_element(By.CSS_SELECTOR, selector)
                    if field.is_displayed():
                        password_field = field
                        logger.info(f"✓ Найдено поле для пароля: {selector}")
                        break
                except:
                    continue
            
            if not password_field:
                logger.error("❌ Не найдено поле для пароля")
                return False
            
            # Вводим пароль
            logger.info("Ввод пароля...")
            password_field.clear()
            time.sleep(self.settings.delay_before_type)
            
            for char in self.settings.password:
                password_field.send_keys(char)
                time.sleep(self.settings.delay_between_keys)
            
            time.sleep(self.settings.delay_after_type)
            
            # Ищем кнопку входа
            login_button_selectors = [
                "button[type='submit']",
                "button[class*='login']",
                "button[class*='auth']",
                "button[class*='enter']",
                "button:contains('Войти')",
                "button:contains('Вход')",
                "input[type='submit']",
            ]
            
            login_button = None
            for selector in login_button_selectors:
                try:
                    button = self.driver.find_element(By.CSS_SELECTOR, selector)
                    if button.is_displayed() and button.is_enabled():
                        login_button = button
                        logger.info(f"✓ Найдена кнопка входа: {selector}")
                        break
                except:
                    continue
            
            if not login_button:
                logger.error("❌ Не найдена кнопка входа")
                return False
            
            # Нажимаем кнопку входа
            logger.info("Нажатие кнопки входа...")
            time.sleep(self.settings.delay_before_click)
            login_button.click()
            time.sleep(self.settings.delay_after_click)
            
            # Ждём завершения авторизации
            logger.info("Ожидание завершения авторизации...")
            time.sleep(5)
            
            # Проверяем успешность авторизации
            if self.check_session():
                logger.success("✓ Авторизация успешна!")
                return True
            else:
                logger.warning("⚠ Авторизация может быть не завершена, проверьте вручную")
                return False
                
        except Exception as e:
            logger.error(f"Ошибка при авторизации: {e}")
            logger.exception("Детали ошибки:")
            return False
    
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
            
            # Шаг 1: Запуск браузера (БЕЗ закрытия существующих процессов Chrome)
            logger.info("Шаг 1: Запуск браузера Chrome...")
            logger.warning("⚠ ВАЖНО: Убедитесь, что Chrome НЕ запущен с профилем Profile 2!")
            logger.warning("⚠ Если Chrome запущен - закройте его вручную перед запуском скрипта")
            try:
                self.start_browser()
            except Exception as e:
                logger.error(f"Критическая ошибка при запуске браузера: {e}")
                logger.error("Возможные причины:")
                logger.error("1. Chrome уже запущен с профилем Profile 2 - закройте его вручную")
                logger.error("2. Профиль заблокирован - закройте все окна Chrome")
                logger.error("3. Недостаточно прав доступа к профилю")
                return False
            
            if not self.driver:
                logger.error("Браузер не был запущен!")
                return False
            
            # Шаг 2: Проверка что страница открыта (уже открыта в start_browser)
            logger.info("Шаг 2: Проверка страницы...")
            current_url = self.driver.current_url
            logger.info(f"Текущий URL: {current_url}")
            
            # Если не на нужной странице - переходим
            if "wildberries.ru" not in current_url or "analytics-reports/sales" not in current_url:
                logger.warning("Не на странице отчётов, переходим...")
                self.driver.get(self.settings.wildberries_start_url)
                time.sleep(5)
                WebDriverWait(self.driver, 20).until(
                    lambda d: d.execute_script("return document.readyState") == "complete"
                )
                logger.success("✓ Страница загружена")
            
            # Шаг 3: Проверка и авторизация (если требуется)
            if not self.check_session():
                logger.warning("⚠ Требуется авторизация")
                
                # Пробуем автоматическую авторизацию
                if self.settings.phone_number or self.settings.email:
                    logger.info("Попытка автоматической авторизации...")
                    if self.handle_authorization():
                        logger.success("✓ Авторизация успешна")
                        # Проверяем снова после авторизации
                        time.sleep(3)
                        if not self.check_session():
                            logger.error("❌ Авторизация не прошла проверку")
                            return False
                    else:
                        logger.error("❌ Авторизация не удалась")
                        logger.error("Пожалуйста, авторизуйтесь вручную в браузере и запустите скрипт снова")
                        logger.info("Ожидание 60 секунд для ручной авторизации...")
                        time.sleep(60)
                        # Проверяем снова после ожидания
                        if not self.check_session():
                            return False
                else:
                    logger.error("❌ Данные для авторизации не указаны в .env")
                    logger.error("Добавьте PHONE_NUMBER или EMAIL и PASSWORD в .env файл")
                    logger.info("Ожидание 60 секунд для ручной авторизации...")
                    time.sleep(60)
                    # Проверяем снова после ожидания
                    if not self.check_session():
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


