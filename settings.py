"""
HRMS System Configuration
Конфигурация системы управления персоналом
"""
import os
from pathlib import Path

# Base paths
BASE_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
EXCEL_FILE = BASE_DIR / "otdel_kadrov.xlsm"
TEMPLATES_DIR = BASE_DIR / "templates"
REPORTS_DIR = BASE_DIR / "приказы"  # RUSSIAN: orders
LOGS_DIR = BASE_DIR / "logs"
PERSONAL_FILES_DIR = BASE_DIR / "личные_дела"  # RUSSIAN: personal files


def ensure_directories():
    """Create required directories if they don't exist"""
    REPORTS_DIR.mkdir(exist_ok=True)
    LOGS_DIR.mkdir(exist_ok=True)
    PERSONAL_FILES_DIR.mkdir(exist_ok=True)
    TEMPLATES_DIR.mkdir(exist_ok=True)

# Alert thresholds
CONTRACT_ALERT_MONTHS_AHEAD = 3  # Show all contracts expiring in next 2-3 months
CONTRACT_ALERT_HIGH_PRIORITY_DAYS = 7
BIRTHDAY_ALERT_DAYS_AHEAD = 30  # Show birthdays for next 30 days (1 month)

# Vacation settings
ANNUAL_VACATION_DAYS = 28

# Logging settings
LOG_ROTATION_SIZE = "10 MB"
LOG_RETENTION_DAYS = 30
DEBUG_MODE = False

# UI settings
WINDOW_WIDTH = 1200
WINDOW_HEIGHT = 800
WINDOW_TITLE = "HRMS - Система управления персоналом"

# Excel sheet names
SHEET_EMPLOYEES = "Сотрудники"
SHEET_ARCHIVE = "Архив"
SHEET_VACATIONS = "Отпуска"
SHEET_ORDER_LOG = "Журнал событий"
SHEET_SETTINGS = "Настройки"
SHEET_STAFF_TABLE = "Штатка"  # NOT PRIORITY

# Column mappings (БЕЗ подчеркиваний, с пробелами!)
EMPLOYEE_COLUMNS = [
    "Таб. №", "ФИО", "Подразделение", "Должность", "Дата принятия", 
    "Дата рождения", "Пол", "Гражданин РБ", "Резидент РБ", "Пенсионер",
    "Форма оплаты", "Ставка", "Начало контракта", "Конец контракта",
    "Личный №", "Страховой №", "№ паспорта", "Путь к личному делу"
]

VACATION_COLUMNS = [
    "ID записи", "Таб. №", "ФИО", "Дата начала", 
    "Дата окончания", "Тип отпуска", "Количество дней"
]

STAFF_TABLE_COLUMNS = [
    "Подразделение", "Должность", "Оклад мин", "Оклад макс", "Количество ставок"
]

REFERENCE_COLUMNS = [
    "Должность", "События", "Подразделение", "Форма оплаты", "Ставка"
]

# Valid values
VALID_DEPARTMENTS = ["Завод КТМ", "Основное"]
VALID_GENDERS = ["М", "Ж"]
VALID_YES_NO = ["Да", "Нет"]

# Order numbering
ORDER_NUMBER_STORAGE = "Настройки"  # Sheet name for storing last order numbers
ORDER_NUMBER_FORMAT = "{number:03d}-{code}"  # e.g., "001-П"

# Event types (for order generation)
EVENT_TYPES = [
    "Больничный",
    "Отпуск за свой счет",
    "Отпуск трудовой",
    "Перевод",
    "Прием на работу",
    "Продление контракта",
    "Увольнение"
]

# Order log columns
ORDER_LOG_COLUMNS = [
    "Номер приказа",
    "Тип события",
    "Дата создания",
    "ФИО",
    "Таб. №",
    "Путь к файлу"
]

# Order type codes for numbering (e.g., "001-П")
ORDER_TYPE_CODES = {
    "Прием на работу": "П",
    "Увольнение": "У",
    "Отпуск трудовой": "ОТ",
    "Отпуск за свой счет": "ОТБ",
    "Больничный": "Б",
    "Перевод": "ПЕР",
    "Продление контракта": "ПК"
}
