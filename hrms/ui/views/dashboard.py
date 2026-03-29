"""
HRMS Dashboard - Панель аналитики
Запуск: python dashboard.py (из Excel)
"""
import os
import sys

# Добавляем путь к корневой папке
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import xlwings as xw
from core.db_engine import ExcelDatabase
from core.analytics import AnalyticsEngine


def show_dashboard():
    """Показывает дашборд с статистикой"""
    
    try:
        # Подключаемся к Excel
        try:
            wb = xw.Book.caller()
        except:
            import settings
            wb = xw.Book(settings.EXCEL_FILE)
        
        db = ExcelDatabase()
        db.connect()
        
        # Читаем данные
        employees = db.get_employees()
        vacations = db.get_vacations()
        
        analytics = AnalyticsEngine()
        
        # Получаем статистику
        stats = analytics.get_dashboard_stats(employees)
        birthdays = analytics.get_upcoming_birthdays(employees, days_ahead=30)
        vacation_stats = analytics.calculate_vacation_stats(employees, vacations)
        
        # Находим или создаём лист "Дашборд"
        dashboard_sheet = None
        for sheet in wb.sheets:
            if sheet.name == "Дашборд":
                dashboard_sheet = sheet
                break
        
        if not dashboard_sheet:
            dashboard_sheet = wb.sheets.add("Дашборд")
        
        # Очищаем лист
        dashboard_sheet.range("A:Z").clear()
        
        # Заголовок
        dashboard_sheet.range("A1").value = "HRMS ДАШБОРД"
        dashboard_sheet.range("A1").font.bold = True
        dashboard_sheet.range("A1").font.size = 16
        
        # Общая статистика
        row = 3
        dashboard_sheet.range(f"A{row}").value = "ОБЩАЯ СТАТИСТИКА"
        dashboard_sheet.range(f"A{row}").font.bold = True
        row += 1
        
        dashboard_sheet.range(f"A{row}").value = "Всего сотрудников:"
        dashboard_sheet.range(f"B{row}").value = stats["total_employees"]
        row += 1
        
        dashboard_sheet.range(f"A{row}").value = "Средний возраст:"
        dashboard_sheet.range(f"B{row}").value = stats["avg_age"]
        row += 1
        
        dashboard_sheet.range(f"A{row}").value = "Средний стаж (мес.):"
        dashboard_sheet.range(f"B{row}").value = stats["avg_tenure_months"]
        row += 2
        
        # Статус контрактов
        dashboard_sheet.range(f"A{row}").value = "КОНТРАКТЫ"
        dashboard_sheet.range(f"A{row}").font.bold = True
        row += 1
        
        contract = stats["contract_status"]
        dashboard_sheet.range(f"A{row}").value = "Активных:"
        dashboard_sheet.range(f"B{row}").value = contract["active"]
        dashboard_sheet.range(f"B{row}").font.color = (0, 176, 80)  # Зелёный
        row += 1
        
        dashboard_sheet.range(f"A{row}").value = "Истекают скоро:"
        dashboard_sheet.range(f"B{row}").value = contract["expiring_soon"]
        dashboard_sheet.range(f"B{row}").font.color = (255, 192, 0)  # Оранжевый
        row += 1
        
        dashboard_sheet.range(f"A{row}").value = "Истекли:"
        dashboard_sheet.range(f"B{row}").value = contract["expired"]
        dashboard_sheet.range(f"B{row}").font.color = (192, 0, 0)  # Красный
        row += 2
        
        # По подразделениям
        if stats["by_department"]:
            dashboard_sheet.range(f"A{row}").value = "ПОДРАЗДЕЛЕНИЯ"
            dashboard_sheet.range(f"A{row}").font.bold = True
            row += 1
            
            for dept, count in stats["by_department"].items():
                dashboard_sheet.range(f"A{row}").value = dept
                dashboard_sheet.range(f"B{row}").value = count
                row += 1
            row += 1
        
        # Дни рождения
        dashboard_sheet.range(f"A{row}").value = "ДНИ РОЖДЕНИЯ (ближайшие 30 дней)"
        dashboard_sheet.range(f"A{row}").font.bold = True
        row += 1
        
        if birthdays:
            dashboard_sheet.range(f"A{row}").value = "ФИО"
            dashboard_sheet.range(f"B{row}").value = "Дата"
            dashboard_sheet.range(f"C{row}").value = "Дней до"
            for b in birthdays:
                row += 1
                dashboard_sheet.range(f"A{row}").value = b.get("ФИО", "")
                dashboard_sheet.range(f"B{row}").value = b.get("Дата рождения", "")
                dashboard_sheet.range(f"C{row}").value = b.get("Дней до дня рождения", "")
        else:
            dashboard_sheet.range(f"A{row}").value = "Нет ближайших дней рождения"
            row += 1
        
        row += 2
        
        # Отпуска
        dashboard_sheet.range(f"A{row}").value = "ОТПУСКА (остаток дней)"
        dashboard_sheet.range(f"A{row}").font.bold = True
        row += 1
        
        if not vacation_stats.empty:
            dashboard_sheet.range(f"A{row}").value = "ФИО"
            dashboard_sheet.range(f"B{row}").value = "Использовано"
            dashboard_sheet.range(f"C{row}").value = "Осталось"
            
            for _, vac in vacation_stats.iterrows():
                row += 1
                dashboard_sheet.range(f"A{row}").value = vac.get("ФИО", "")
                dashboard_sheet.range(f"B{row}").value = vac.get("Использовано дней", 0)
                dashboard_sheet.range(f"C{row}").value = vac.get("Осталось дней", 0)
        
        # Автоширина
        dashboard_sheet.autofit("c")
        
        db.disconnect()
        
        print("Дашборд создан на листе 'Дашборд'")
        
    except Exception as e:
        print(f"Ошибка: {e}")
        import traceback
        traceback.print_exc()
        input("Нажми Enter для выхода...")


if __name__ == "__main__":
    show_dashboard()
