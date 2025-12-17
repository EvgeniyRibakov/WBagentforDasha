"""Скрипт установки WBagentforDasha на новом устройстве.

Выполняет все необходимые шаги для настройки проекта:
1. Проверка версии Python
2. Установка зависимостей
3. Создание .env файла из .env_sample
4. Проверка наличия Yandex Browser
"""
import sys
import subprocess
import shutil
from pathlib import Path


def check_python_version() -> bool:
    """Проверяет версию Python (должна быть 3.12+)."""
    print("=" * 70)
    print("Шаг 1: Проверка версии Python")
    print("=" * 70)
    
    version = sys.version_info
    print(f"Текущая версия Python: {version.major}.{version.minor}.{version.micro}")
    
    if version.major < 3 or (version.major == 3 and version.minor < 12):
        print("❌ ОШИБКА: Требуется Python 3.12 или выше!")
        print("Скачайте Python с https://www.python.org/downloads/")
        print("При установке обязательно отметьте 'Add Python to PATH'")
        return False
    
    print("✓ Версия Python подходит")
    return True


def install_dependencies() -> bool:
    """Устанавливает зависимости из requirements.txt."""
    print("\n" + "=" * 70)
    print("Шаг 2: Установка зависимостей")
    print("=" * 70)
    
    requirements_file = Path("requirements.txt")
    if not requirements_file.exists():
        print("❌ ОШИБКА: Файл requirements.txt не найден!")
        return False
    
    print("Установка пакетов из requirements.txt...")
    try:
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "-r", str(requirements_file)
        ])
        print("✓ Зависимости установлены успешно")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ ОШИБКА при установке зависимостей: {e}")
        return False


def create_env_file() -> bool:
    """Создаёт .env файл из .env_sample."""
    print("\n" + "=" * 70)
    print("Шаг 3: Создание файла .env")
    print("=" * 70)
    
    env_sample = Path(".env_sample")
    env_file = Path(".env")
    
    if not env_sample.exists():
        print("❌ ОШИБКА: Файл .env_sample не найден!")
        return False
    
    if env_file.exists():
        response = input("Файл .env уже существует. Перезаписать? (да/нет): ").strip().lower()
        if response not in ['да', 'yes', 'y', 'д']:
            print("✓ Файл .env оставлен без изменений")
            return True
    
    try:
        shutil.copy(env_sample, env_file)
        print("✓ Файл .env создан из .env_sample")
        print("\n⚠️  ВАЖНО: Откройте файл .env и заполните:")
        print("   - PHONE_NUMBER (номер телефона для авторизации)")
        print("   - Остальные настройки можно оставить по умолчанию")
        return True
    except Exception as e:
        print(f"❌ ОШИБКА при создании .env: {e}")
        return False


def check_yandex_browser() -> bool:
    """Проверяет наличие Yandex Browser."""
    print("\n" + "=" * 70)
    print("Шаг 4: Проверка Yandex Browser")
    print("=" * 70)
    
    import os
    from pathlib import Path
    
    # Стандартный путь к Yandex Browser
    default_paths = [
        Path(os.path.expandvars("%LOCALAPPDATA%")) / "Yandex" / "YandexBrowser" / "Application" / "browser.exe",
        Path(os.path.expandvars("%PROGRAMFILES%")) / "Yandex" / "YandexBrowser" / "Application" / "browser.exe",
        Path(os.path.expandvars("%PROGRAMFILES(X86)%")) / "Yandex" / "YandexBrowser" / "Application" / "browser.exe",
    ]
    
    for browser_path in default_paths:
        if browser_path.exists():
            print(f"✓ Yandex Browser найден: {browser_path}")
            return True
    
    print("⚠️  Yandex Browser не найден в стандартных местах")
    print("   Если браузер установлен, укажите путь в файле .env (YANDEX_BROWSER_PATH)")
    print("   Скачать Yandex Browser: https://browser.yandex.ru/")
    return True  # Не критично, можно указать путь вручную


def main() -> int:
    """Основная функция установки."""
    print("\n" + "=" * 70)
    print("УСТАНОВКА WBagentforDasha")
    print("=" * 70)
    print("\nЭтот скрипт выполнит все необходимые шаги для настройки проекта.\n")
    
    # Проверка версии Python
    if not check_python_version():
        return 1
    
    # Установка зависимостей
    if not install_dependencies():
        return 1
    
    # Создание .env файла
    if not create_env_file():
        return 1
    
    # Проверка Yandex Browser
    check_yandex_browser()
    
    # Финальные инструкции
    print("\n" + "=" * 70)
    print("✓ УСТАНОВКА ЗАВЕРШЕНА")
    print("=" * 70)
    print("\nСледующие шаги:")
    print("1. Откройте файл .env и заполните PHONE_NUMBER")
    print("2. Запустите первую авторизацию:")
    print("   python manual_auth.py")
    print("3. После авторизации можно запускать основной скрипт:")
    print("   python -m src.main")
    print("\nДля скачивания отчётов за конкретную дату:")
    print("   python -m src.main --date 10.12.2025")
    print("\n" + "=" * 70)
    
    return 0


if __name__ == "__main__":
    sys.exit(main())


