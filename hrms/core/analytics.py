# core/analytics.py

from datetime import date, datetime
from typing import List, Dict
import pandas as pd

class AnalyticsEngine:
    """
    Класс для выполнения аналитических расчетов, связанных с данными сотрудников.
    """

    def calculate_age(self, birth_date) -> int:
        """
        Рассчитывает возраст сотрудника на сегодняшний день.
        """
        if not isinstance(birth_date, (date, datetime, pd.Timestamp)):
            return 0
        
        today = date.today()
        # Приводим к date для сравнения
        if hasattr(birth_date, 'date'):
            birth_date = birth_date.date()
            
        return today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))

    def calculate_tenure(self, hire_date) -> tuple[int, int]:
        """
        Рассчитывает стаж работы сотрудника в годах и месяцах.
        """
        if not isinstance(hire_date, (date, datetime, pd.Timestamp)):
            return 0, 0
            
        today = date.today()
        if hasattr(hire_date, 'date'):
            hire_date = hire_date.date()
        
        years = today.year - hire_date.year
        months = today.month - hire_date.month
        
        if months < 0:
            years -= 1
            months += 12
            
        return years, months

    def calculate_contract_days_remaining(self, contract_end) -> int:
        """
        Рассчитывает количество дней до окончания контракта.
        """
        if not isinstance(contract_end, (date, datetime, pd.Timestamp)):
            return 0

        # Приводим оба значения к pd.Timestamp для корректного вычитания
        today = pd.Timestamp(date.today())
        try:
            target = pd.Timestamp(contract_end)
            # Убираем часовой пояс если есть для сравнения
            if target.tzinfo is not None:
                target = target.tz_convert(None)
            return (target - today).days
        except:
            return 0

    def get_contract_alerts(self, employees_df: pd.DataFrame) -> List[Dict]:
        """
        Получает и обрабатывает предупреждения по контрактам.
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
        
        # Убираем TZ для сравнения если есть
        if alerts_df['Конец контракта'].dt.tz is not None:
            alerts_df['Конец контракта'] = alerts_df['Конец контракта'].dt.tz_localize(None)

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
            birth_dt = row["Дата рождения"]
            if birth_dt.tzinfo is not None:
                birth_dt = birth_dt.tz_convert(None)
            
            birth_date = birth_dt.date()
            try:
                bday_this_year = birth_date.replace(year=today.year)
            except ValueError: # 29 февраля
                bday_this_year = birth_date.replace(year=today.year, month=3, day=1)

            if bday_this_year < today:
                try:
                    bday_next_year = birth_date.replace(year=today.year + 1)
                except ValueError:
                    bday_next_year = birth_date.replace(year=today.year + 1, month=3, day=1)
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
            dob = row.get("Дата рождения")
            if pd.notna(dob):
                ages.append(self.calculate_age(dob))
        avg_age = sum(ages) / len(ages) if ages else 0

        tenure_months_list = []
        for _, row in employees_df.iterrows():
            hire = row.get("Дата принятия")
            if pd.notna(hire):
                years, months = self.calculate_tenure(hire)
                tenure_months_list.append(years * 12 + months)
        avg_tenure_months = sum(tenure_months_list) / len(tenure_months_list) if tenure_months_list else 0

        contract_status = {"active": 0, "expiring_soon": 0, "expired": 0}
        if "Конец контракта" in employees_df.columns:
            today = pd.Timestamp(date.today())
            contracts = pd.to_datetime(employees_df["Конец контракта"], errors="coerce", dayfirst=True)
            # Снимаем TZ
            if contracts.dt.tz is not None:
                contracts = contracts.dt.tz_localize(None)
                
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
