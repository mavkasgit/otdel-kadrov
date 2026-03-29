"""
HRMS - Главное меню
"""
import os
import sys
import ctypes

# Важно! Установить app ID до создания окна
myappid = 'hrms.otdelkadrov'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import xlwings as xw
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox

from core.db_engine import ExcelDatabase
from core.analytics import AnalyticsEngine
import settings


def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")


def get_icon_path():
    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "hrms.ico")


class DashboardDialog:
    """Дашборд в отдельном окне"""
    
    def __init__(self, parent):
        self.db = None
        self.analytics = AnalyticsEngine()
        
        self.win = ttk.Toplevel(parent)
        self.win.title("Дашборд - Статистика")
        center_window(self.win, 500, 600)
        
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "..", "..", "icon.ico")
            icon_path = os.path.abspath(icon_path)
            self.win.iconbitmap(icon_path)
        except Exception as e:
            print(f"Icon error: {e}")
        
        self.setup_ui()
        self.load_data()
    
    def setup_ui(self):
        main = ttk.Frame(self.win, padding=15)
        main.pack(fill=BOTH, expand=True)
        
        ttk.Label(main, text="Дашборд", font=("Segoe UI", 18, "bold")).pack(pady=10)
        
        self.stats_frame = ttk.LabelFrame(main, text="Общая статистика")
        self.stats_frame.pack(fill="x", pady=10)
        
        self.bday_frame = ttk.LabelFrame(main, text="Дни рождения (30 дней)")
        self.bday_frame.pack(fill="both", expand=True, pady=10)
        
        self.vac_frame = ttk.LabelFrame(main, text="Отпуска")
        self.vac_frame.pack(fill="x", pady=10)
        
        ttk.Button(main, text="Закрыть", command=self.win.destroy, padding=10).pack(pady=10)
    
    def load_data(self):
        try:
            try:
                wb = xw.Book.caller()
            except:
                wb = xw.Book(settings.EXCEL_FILE)
            
            self.db = ExcelDatabase()
            self.db.connect()
            
            employees = self.db.get_employees()
            vacations = self.db.get_vacations()
            
            stats = self.analytics.get_dashboard_stats(employees)
            birthdays = self.analytics.get_upcoming_birthdays(employees, days_ahead=30)
            vac_stats = self.analytics.calculate_vacation_stats(employees, vacations)
            
            # Общая статистика
            for label, value in [
                ("Всего сотрудников:", stats.get("total_employees", 0)),
                ("Мужчин:", stats.get("male_count", 0)),
                ("Женщин:", stats.get("female_count", 0)),
                ("Контракты истекают:", stats.get("expiring_contracts_30d", 0)),
            ]:
                row = ttk.Frame(self.stats_frame)
                row.pack(fill="x", padx=10, pady=3)
                ttk.Label(row, text=label, width=25, anchor="w").pack(side=LEFT)
                ttk.Label(row, text=str(value), font=("Segoe UI", 11, "bold")).pack(side=LEFT)
            
            # Дни рождения
            if birthdays:
                for bday in birthdays[:10]:
                    name = bday.get("FIO", "")
                    date = bday.get("Data rozhdeniya", "")
                    if hasattr(date, 'strftime'):
                        date = date.strftime("%d.%m")
                    ttk.Label(self.bday_frame, text=f"{name} - {date}").pack(anchor="w", padx=10, pady=2)
            else:
                ttk.Label(self.bday_frame, text="Нет ближайших дней рождения").pack(pady=10)
            
            # Отпуска
            ttk.Label(self.vac_frame, text=f"В отпуске сейчас: {vac_stats.get('currently_on_vacation', 0)}").pack(anchor="w", padx=10, pady=3)
            ttk.Label(self.vac_frame, text=f"Запланировано: {vac_stats.get('planned_vacations', 0)}").pack(anchor="w", padx=10, pady=3)
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные:\n{e}")


class MainMenu:
    """Главное меню HRMS"""
    
    def __init__(self):
        self.root = ttk.Window(themename="yeti")
        self.root.title("HRMS - Отдел кадров")
        center_window(self.root, 450, 650)
        
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "..", "..", "icon.ico")
            icon_path = os.path.abspath(icon_path)
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Icon error: {e}")
        
        self.db = None
        self.analytics = AnalyticsEngine()
        
        self.setup_ui()
        self.root.mainloop()
    
    def setup_ui(self):
        header = ttk.Frame(self.root, bootstyle="primary", height=80)
        header.pack(fill="x")
        
        ttk.Label(header, text="HRMS", font=("Segoe UI", 24, "bold"),
                  bootstyle="inverse-primary").pack(pady=15)
        
        ttk.Label(header, text="Система управления персоналом",
                  bootstyle="inverse-primary", font=("Segoe UI", 10)).pack()
        
        menu_frame = ttk.Frame(self.root)
        menu_frame.pack(pady=25, padx=20, fill="both", expand=True)
        
        menu_buttons = [
            ("📊 Дашборд", self.open_dashboard, "success"),
            ("👥 Сотрудники", self.open_employees, "info"),
            ("🏖️ Отпуска", self.open_vacations, "info"),
            ("📄 Приказы", self.open_orders, "secondary"),
            ("⚙️ Настройки", self.open_settings, "secondary"),
        ]
        
        for text, cmd, style in menu_buttons:
            btn = ttk.Button(
                menu_frame, text=text, command=cmd,
                bootstyle=style, width=25
            )
            btn.pack(pady=6, fill="x")
        
        self.status_label = ttk.Label(self.root, text="", bootstyle="secondary")
        self.status_label.pack(side="bottom", pady=10)
        
        self.update_status()
    
    def update_status(self):
        try:
            try:
                wb = xw.Book.caller()
            except:
                wb = xw.Book(settings.EXCEL_FILE)
            
            if self.db is None:
                self.db = ExcelDatabase()
                self.db.connect()
            
            employees = self.db.get_employees()
            vacations = self.db.get_vacations()
            
            self.status_label.config(
                text=f"Сотрудников: {len(employees)} | Отпусков: {len(vacations)}"
            )
        except Exception as e:
            self.status_label.config(text=f"Статус: {e}")
    
    def open_dashboard(self):
        DashboardDialog(self.root)
    
    def open_employees(self):
        self.root.withdraw()
        from ui.views import employee_card
        employee_card.EmployeeCardDialog()
        self.root.deiconify()
    
    def open_vacations(self):
        self.root.withdraw()
        from ui.views import vacation_mgr
        vacation_mgr.VacationManagerDialog()
        self.root.deiconify()
    
    def open_orders(self):
        self.root.withdraw()
        from ui.views import order_generator
        order_generator.OrderGeneratorDialog()
        self.root.deiconify()
    
    def open_settings(self):
        info = f"Excel файл: {settings.EXCEL_FILE}\nПапка отчётов: {settings.REPORTS_DIR}\nПапка логов: {settings.LOGS_DIR}"
        messagebox.showinfo("Настройки", info)


def main():
    try:
        app = MainMenu()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("Нажми Enter для выхода...")


if __name__ == "__main__":
    main()
