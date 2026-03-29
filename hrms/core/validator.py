# core/validator.py

from datetime import datetime, date
from typing import Tuple, Dict, List, Optional
import pandas as pd
from core.db_engine import ExcelDatabase
from core.logger import logger


class DataValidator:
    """
    Валидация данных и проверка бизнес-правил.
    """

    def __init__(self, db: ExcelDatabase):
        """
        Initialize validator with database connection.

        Args:
            db: ExcelDatabase instance
        """
        self.db = db

    def validate_vacation_data(self, vacation_data: dict) -> Tuple[bool, str]:
        """
        Валидирует данные отпуска.

        Args:
            vacation_data: Словарь с данными отпуска

        Returns:
            Кортеж (is_valid, error_message)
        """
        required_fields = ["Таб. №", "Дата начала", "Дата окончания", "Тип отпуска"]

        for field in required_fields:
            if field not in vacation_data or vacation_data[field] is None:
                return False, f"Отсутствует обязательное поле: {field}"

        tab_number = vacation_data.get("Таб. №")
        if not tab_number:
            return False, "Не указан табельный номер"

        try:
            employees = self.db.get_employees()
            if employees.empty or tab_number not in employees["Таб. №"].values:
                return False, f"Сотрудник с табельным номером {tab_number} не найден"
        except Exception as e:
            logger.warning(f"Could not verify employee: {e}")

        start_date = vacation_data.get("Дата начала")
        end_date = vacation_data.get("Дата окончания")

        is_valid, error = self.validate_date_logic(start_date, end_date)
        if not is_valid:
            return False, error

        vacation_type = vacation_data.get("Тип отпуска")
        valid_types = ["Трудовой отпуск", "Отпуск за свой счет", "Учебный отпуск", "Декретный"]
        if vacation_type not in valid_types:
            return False, f"Неверный тип отпуска: {vacation_type}"

        return True, ""

    def check_vacation_overlap(
        self,
        tab_number: int,
        start_date: datetime,
        end_date: datetime,
        vacation_id: Optional[int] = None
    ) -> Tuple[bool, List[Dict]]:
        """
        Проверяет пересечение дат отпусков для одного сотрудника.

        Args:
            tab_number: Табельный номер сотрудника
            start_date: Дата начала отпуска
            end_date: Дата окончания отпуска
            vacation_id: ID отпуска (для исключения при редактировании)

        Returns:
            Кортеж (has_overlap, list_of_overlapping_vacations)
        """
        try:
            vacations = self.db.get_vacations(tab_number)
        except Exception as e:
            logger.error(f"Error checking vacation overlap: {e}")
            return False, []

        if vacations.empty:
            return False, []

        if vacation_id is not None:
            vacations = vacations[vacations.get("ID записи") != vacation_id]

        overlaps = []
        for _, vac in vacations.iterrows():
            vac_start = vac.get("Дата начала")
            vac_end = vac.get("Дата окончания")

            if vac_start is None or vac_end is None:
                continue

            if not (end_date < vac_start or start_date > vac_end):
                overlaps.append({
                    "id": vac.get("ID записи"),
                    "start": vac_start,
                    "end": vac_end,
                    "type": vac.get("Тип отпуска")
                })

        return len(overlaps) > 0, overlaps

    def validate_reference_value(self, field: str, value: str) -> Tuple[bool, str]:
        """
        Проверяет, что значение поля существует в справочнике.

        Args:
            field: Название поля
            value: Значение для проверки

        Returns:
            Кортеж (is_valid, error_message)
        """
        reference_fields = {
            "Подразделение": ["Завод КТМ", "Основное"],
            "Пол": ["М", "Ж"],
            "Гражданин РБ": ["Да", "Нет"],
            "Резидент РБ": ["Да", "Нет"],
            "Пенсионер": ["Да", "Нет"],
            "Форма оплаты": ["Почасовая", "Оклад"],
        }

        if field in reference_fields:
            valid_values = reference_fields[field]
            if value not in valid_values:
                return False, f"Неверное значение '{value}' для поля '{field}'. Допустимые: {', '.join(valid_values)}"

        return True, ""

    def validate_date_logic(
        self,
        start_date: datetime,
        end_date: datetime
    ) -> Tuple[bool, str]:
        """
        Проверяет логику дат (начало раньше конца).

        Args:
            start_date: Дата начала
            end_date: Дата окончания

        Returns:
            Кортеж (is_valid, error_message)
        """
        if start_date is None or end_date is None:
            return False, "Даты начала и окончания обязательны"

        if isinstance(start_date, str):
            try:
                start_date = pd.to_datetime(start_date, dayfirst=True)
            except:
                return False, "Неверный формат даты начала"

        if isinstance(end_date, str):
            try:
                end_date = pd.to_datetime(end_date, dayfirst=True)
            except:
                return False, "Неверный формат даты окончания"

        if start_date > end_date:
            return False, "Дата начала не может быть позже даты окончания"

        return True, ""

    def validate_employee_data(self, employee_data: dict) -> Tuple[bool, str]:
        """
        Валидирует данные сотрудника.

        Args:
            employee_data: Словарь с данными сотрудника

        Returns:
            Кортеж (is_valid, error_message)
        """
        required_fields = ["Таб. №", "ФИО"]
        for field in required_fields:
            if field not in employee_data or not employee_data[field]:
                return False, f"Отсутствует обязательное поле: {field}"

        if "Дата принятия" in employee_data and employee_data["Дата принятия"]:
            is_valid, error = self.validate_date_logic(
                employee_data["Дата принятия"],
                datetime.now()
            )
            if not is_valid:
                return False, f"Дата принятия: {error}"

        if "Дата рождения" in employee_data and employee_data["Дата рождения"]:
            birth_date = employee_data["Дата рождения"]
            if isinstance(birth_date, str):
                try:
                    birth_date = pd.to_datetime(birth_date, dayfirst=True)
                except:
                    return False, "Неверный формат даты рождения"
            
            if birth_date > datetime.now():
                return False, "Дата рождения не может быть в будущем"

        for field in ["Подразделение", "Пол", "Гражданин РБ", "Резидент РБ", "Пенсионер"]:
            if field in employee_data and employee_data[field]:
                is_valid, error = self.validate_reference_value(field, employee_data[field])
                if not is_valid:
                    return False, error

        return True, ""
