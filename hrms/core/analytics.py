# core/analytics.py

from datetime import date, datetime
from typing import List, Dict
import pandas as pd

class AnalyticsEngine:
    """
    Класс для выполнения аналитических расчетов, связанных с данными сотрудников.
    """

    def calculate_age(self, birth_date: date) -> int:
        """
        Рассчитывает возраст сотрудника на сегодняшний день.

        Args:
            birth_date: Дата рождения сотрудника.

        Returns:
            Полных лет.
        """
        if not isinstance(birth_date, (date, datetime)):
            return 0
        
        today = date.today()
        return today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))

    def calculate_tenure(self, hire_date: date) -> tuple[int, int]:
        """
        Рассчитывает стаж работы сотрудника в годах и месяцах.

        Args:
            hire_date: Дата приема на работу.

        Returns:
            Кортеж (годы, месяцы).
        """
        if not isinstance(hire_date, (date, datetime)):
            return 0, 0
            
        today = date.today()
        
        years = today.year - hire_date.year
        months = today.month - hire_date.month
        
        if months < 0:
            years -= 1
            months += 12
            
        return years, months

    def calculate_contract_days_remaining(self, contract_end: date) -> int:
        """
        Рассчитывает количество дней до окончания контракта.

        Args:
            contract_end: Дата окончания контракта.

        Returns:
            Количество дней. Возвращает отрицательное число, если контракт истек.
        """
        if not isinstance(contract_end, (date, datetime)):
            return 0

        today = date.today()
        return (contract_end - today).days

    def get_contract_alerts(self, employees_df: pd.DataFrame) -> List[Dict]:
        """
        Получает и обрабатывает предупреждения по контрактам.

        - Получает ВСЕ контракты на 2-3 месяца вперед.
        - Группирует по месяцам (YYYY-MM).
        - Приоритет HIGH если <= 7 дней, иначе MEDIUM.
        - ВАЖНО: Красные алерты (HIGH) всегда в самом верху.
        - Сортировка: сначала HIGH по дате, потом MEDIUM по месяцам.

        Args:
            employees_df: DataFrame с данными сотрудников.

        Returns:
            Список словарей с данными по алертам.
        """
        if 'Конец контракта' not in employees_df.columns:
            return []

        today = pd.to_datetime(date.today())
        future_limit = today + pd.DateOffset(months=3)

        # 1. Фильтрация контрактов
        alerts_df = employees_df[
            pd.to_datetime(employees_df['Конец контракта'], errors='coerce', dayfirst=True).notna()
        ].copy()
        
        alerts_df['Конец контракта'] = pd.to_datetime(alerts_df['Конец контракта'], dayfirst=True)
        
        alerts_df = alerts_df[
            (alerts_df['Конец контракта'] >= today) &
            (alerts_df['Конец контракта'] <= future_limit)
        ]

        if alerts_df.empty:
            return []

        # 2. Расчет дней и приоритета
        alerts_df['days_remaining'] = (alerts_df['Конец контракта'] - today).dt.days
        alerts_df['priority'] = alerts_df['days_remaining'].apply(
            lambda x: 'HIGH' if x <= 7 else 'MEDIUM'
        )
        alerts_df['month_group'] = alerts_df['Конец контракта'].dt.strftime('%Y-%m')

        # 3. Разделение на HIGH и MEDIUM
        high_priority = alerts_df[alerts_df['priority'] == 'HIGH'].sort_values(by='Конец контракта')
        medium_priority = alerts_df[alerts_df['priority'] == 'MEDIUM'].sort_values(by='Конец контракта')

        # 4. Сортировка и объединение
        sorted_alerts = pd.concat([high_priority, medium_priority])

        return sorted_alerts.to_dict('records')

    def calculate_vacation_stats(self, employees_df: pd.DataFrame, vacations_df: pd.DataFrame) -> pd.DataFrame:
        """
        Рассчитывает статистику по отпускам для каждого сотрудника.

        Args:
            employees_df: DataFrame с данными сотрудников.
            vacations_df: DataFrame с данными об отпусках.

        Returns:
            DataFrame с колонками: Таб. №, ФИО, Использовано дней, Доступно дней, Осталось дней
        """
        if employees_df.empty:
            return pd.DataFrame(columns=["Таб. №", "ФИО", "Использовано дней", "Доступно дней", "Осталось дней"])

        result = employees_df[["Таб. №", "ФИО", "Дата принятия"]].copy()

        if not vacations_df.empty:
            used_days = vacations_df.groupby("Таб. №")["Количество дней"].sum().reset_index()
            used_days.columns = ["Таб. №", "Использовано дней"]
            result = result.merge(used_days, on="Таб. №", how="left")
            result["Использовано дней"] = result["Использовано дней"].fillna(0).astype(int)
        else:
            result["Использовано дней"] = 0

        result["Доступно дней"] = 28
        result["Осталось дней"] = result["Доступно дней"] - result["Использовано дней"]
        result = result[["Таб. №", "ФИО", "Использовано дней", "Доступно дней", "Осталось дней"]]

        return result

    def get_upcoming_birthdays(self, employees_df: pd.DataFrame, days_ahead: int = 30) -> List[Dict]:
        """
        Получает список сотрудников с приближающимися днями рождениями.

        Args:
            employees_df: DataFrame с данными сотрудников.
            days_ahead: Количество дней для поиска вперед (по умолчанию 30).

        Returns:
            Список словарей с данными по дням рождения.
        """
        if "Дата рождения" not in employees_df.columns:
            return []

        today = date.today()
        employees = employees_df.copy()
        employees["Дата рождения"] = pd.to_datetime(employees["Дата рождения"], errors="coerce", dayfirst=True)

        if employees["Дата рождения"].empty:
            return []

        employees = employees[employees["Дата рождения"].notna()].copy()

        birthdays = []
        for _, row in employees.iterrows():
            birth_date = row["Дата рождения"].date()
            bday_this_year = birth_date.replace(year=today.year)

            if bday_this_year < today:
                bday_next_year = birth_date.replace(year=today.year + 1)
                days_until = (bday_next_year - today).days
            else:
                days_until = (bday_this_year - today).days

            if 0 <= days_until <= days_ahead:
                birthdays.append({
                    "Таб. №": row.get("Таб. №"),
                    "ФИО": row.get("ФИО"),
                    "Дата рождения": birth_date.strftime("%d.%m.%Y"),
                    "Дней до дня рождения": days_until,
                    "Возраст": self.calculate_age(birth_date)
                })

        return sorted(birthdays, key=lambda x: x["Дней до дня рождения"])

    def get_dashboard_stats(self, employees_df: pd.DataFrame) -> Dict:
        """
        Рассчитывает общую статистику для дашборда.

        Args:
            employees_df: DataFrame с данными сотрудников.

        Returns:
            Словарь с различными статистиками.
        """
        if employees_df.empty:
            return {
                "total_employees": 0,
                "by_department": {},
                "by_position": {},
                "avg_age": 0,
                "avg_tenure_months": 0,
                "contract_status": {"active": 0, "expiring_soon": 0, "expired": 0}
            }

        total = len(employees_df)

        dept_counts = employees_df["Подразделение"].value_counts().to_dict() if "Подразделение" in employees_df.columns else {}
        position_counts = employees_df["Должность"].value_counts().to_dict() if "Должность" in employees_df.columns else {}

        ages = []
        for _, row in employees_df.iterrows():
            if pd.notna(row.get("Дата рождения")):
                ages.append(self.calculate_age(row["Дата рождения"]))
        avg_age = sum(ages) / len(ages) if ages else 0

        tenure_months_list = []
        for _, row in employees_df.iterrows():
            if pd.notna(row.get("Дата принятия")):
                years, months = self.calculate_tenure(row["Дата принятия"])
                tenure_months_list.append(years * 12 + months)
        avg_tenure_months = sum(tenure_months_list) / len(tenure_months_list) if tenure_months_list else 0

        contract_status = {"active": 0, "expiring_soon": 0, "expired": 0}
        if "Конец контракта" in employees_df.columns:
            today = pd.to_datetime(date.today(), dayfirst=True)
            contracts = pd.to_datetime(employees_df["Конец контракта"], errors="coerce", dayfirst=True)
            contract_status["expired"] = int((contracts < today).sum())
            contract_status["expiring_soon"] = int(((contracts >= today) & (contracts <= today + pd.DateOffset(months=1))).sum())
            contract_status["active"] = int(total - contract_status["expired"] - contract_status["expiring_soon"])

        return {
            "total_employees": total,
            "by_department": dept_counts,
            "by_position": position_counts,
            "avg_age": round(avg_age, 1),
            "avg_tenure_months": round(avg_tenure_months, 1),
            "contract_status": contract_status
        }
