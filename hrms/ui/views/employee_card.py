"""
HRMS - Card employee
"""
import os
import sys
import ctypes

myappid = 'hrms.otdelkadrov'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import xlwings as xw
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox, StringVar, Listbox, END
from PIL import Image, ImageTk
from datetime import datetime

from core.db_engine import ExcelDatabase
from core.analytics import AnalyticsEngine
from core.logger import logger
import settings
import pandas as pd


def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")


def get_icon_path():
    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "hrms.ico")


class EmployeeCardDialog:
    
    def __init__(self, parent=None, tab_number=None):
        if parent:
            self.root = ttk.Toplevel(parent)
        else:
            self.root = ttk.Window(themename="yeti")
            
        self.root.title("Карточка сотрудника")
        center_window(self.root, 500, 700)
        
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "..", "..", "icon.png")
            icon_path = os.path.abspath(icon_path)
            icon_img = Image.open(icon_path)
            self._icon = ImageTk.PhotoImage(icon_img)
            self.root.iconphoto(True, self._icon)
        except Exception as e:
            print(f"Icon error: {e}")
        
        self.db = None
        self.analytics = AnalyticsEngine()
        self.tab_number = tab_number
        self.employee_data = None
        self.full_employees_list = []
        
        self.setup_ui()
        
        if tab_number:
            self.load_employee(tab_number)
        
        if not parent:
            self.root.mainloop()
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        self.header_label = ttk.Label(main_frame, text="Карточка сотрудника", 
                                   font=("Segoe UI", 14, "bold"))
        self.header_label.pack(pady=10)
        
        if self.tab_number is None:
            # Поиск сотрудника
            ttk.Label(main_frame, text="Поиск сотрудника (имя или таб. №):").pack(pady=(10, 0))
            self.search_query = StringVar()
            self.search_query.trace_add("write", self.filter_employees)
            self.search_entry = ttk.Entry(main_frame, textvariable=self.search_query, width=50)
            self.search_entry.pack(pady=5)
            
            # Список результатов (вместо выпадающего списка)
            self.employee_listbox = Listbox(main_frame, height=6, width=60)
            self.employee_listbox.pack(pady=5)
            self.employee_listbox.bind("<<ListboxSelect>>", self.on_employee_selected)
            
            self.load_employee_list()
        
        self.info_frame = ttk.LabelFrame(main_frame, text="Основная информация")
        self.info_frame.pack(fill="x", pady=10)
        
        self.info_labels = {}
        # Список полей на русском языке для соответствия settings.py
        fields = [
            "Таб. №", "ФИО", "Подразделение", "Должность", "Дата принятия", 
            "Дата рождения", "Пол", "Гражданин РБ", "Резидент РБ", "Пенсионер"
        ]
        
        for i, field in enumerate(fields):
            ttk.Label(self.info_frame, text=f"{field}:", font=("Segoe UI", 10, "bold")).grid(
                row=i, column=0, sticky="w", pady=3
            )
            self.info_labels[field] = ttk.Label(self.info_frame, text="-", font=("Segoe UI", 10))
            self.info_labels[field].grid(row=i, column=1, sticky="w", pady=3, padx=10)
        
        self.contract_frame = ttk.LabelFrame(main_frame, text="Контракт")
        self.contract_frame.pack(fill="x", pady=10)
        
        ttk.Label(self.contract_frame, text="Начало:", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, sticky="w", pady=3
        )
        self.contract_start = ttk.Label(self.contract_frame, text="-")
        self.contract_start.grid(row=0, column=1, sticky="w", pady=3, padx=10)
        
        ttk.Label(self.contract_frame, text="Конец:", font=("Segoe UI", 10, "bold")).grid(
            row=1, column=0, sticky="w", pady=3
        )
        self.contract_end = ttk.Label(self.contract_frame, text="-")
        self.contract_end.grid(row=1, column=1, sticky="w", pady=3, padx=10)
        
        ttk.Label(self.contract_frame, text="Осталось:", font=("Segoe UI", 10, "bold")).grid(
            row=2, column=0, sticky="w", pady=3
        )
        self.contract_days = ttk.Label(self.contract_frame, text="-", font=("Segoe UI", 10, "bold"))
        self.contract_days.grid(row=2, column=1, sticky="w", pady=3, padx=10)
        
        self.stats_frame = ttk.LabelFrame(main_frame, text="Статистика")
        self.stats_frame.pack(fill="x", pady=10)
        
        ttk.Label(self.stats_frame, text="Возраст:", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, sticky="w", pady=3
        )
        self.age_label = ttk.Label(self.stats_frame, text="-")
        self.age_label.grid(row=0, column=1, sticky="w", pady=3, padx=10)
        
        ttk.Label(self.stats_frame, text="Стаж:", font=("Segoe UI", 10, "bold")).grid(
            row=1, column=0, sticky="w", pady=3
        )
        self.tenure_label = ttk.Label(self.stats_frame, text="-")
        self.tenure_label.grid(row=1, column=1, sticky="w", pady=3, padx=10)
        
        ttk.Label(self.stats_frame, text="Отпусков использовано:", font=("Segoe UI", 10, "bold")).grid(
            row=2, column=0, sticky="w", pady=3
        )
        self.vacation_used = ttk.Label(self.stats_frame, text="-")
        self.vacation_used.grid(row=2, column=1, sticky="w", pady=3, padx=10)
        
        ttk.Label(self.stats_frame, text="Отпусков осталось:", font=("Segoe UI", 10, "bold")).grid(
            row=3, column=0, sticky="w", pady=3
        )
        self.vacation_remain = ttk.Label(self.stats_frame, text="-", font=("Segoe UI", 10, "bold"))
        self.vacation_remain.grid(row=3, column=1, sticky="w", pady=3, padx=10)
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15)
        
        ttk.Button(btn_frame, text="Редактировать", command=self.edit_employee,
               bootstyle=WARNING, padding=10).pack(side=LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Создать приказ", command=self.create_order,
               bootstyle=SECONDARY, padding=10).pack(side=LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Закрыть", command=self.root.destroy,
               padding=10).pack(side=LEFT, padx=5)
    
    def load_employee_list(self):
        try:
            if self.db is None:
                try:
                    wb = xw.Book.caller()
                except:
                    wb = xw.Book(settings.EXCEL_FILE)
                
                self.db = ExcelDatabase()
                self.db.connect()
            
            employees = self.db.get_employees()
            
            self.full_employees_list = []
            for _, row in employees.iterrows():
                name = row.get("ФИО")
                if pd.notna(name) and str(name).strip():
                    tab_num = row.get("Таб. №")
                    tab_str = ""
                    if pd.notna(tab_num):
                        try:
                            tab_str = str(int(float(tab_num)))
                        except:
                            pass
                    self.full_employees_list.append((name, tab_str))
            
            self.full_employees_list.sort(key=lambda x: x[0])
            self.refresh_employee_list(self.full_employees_list)
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить сотрудников:\n{e}")
    
    def refresh_employee_list(self, employees):
        self.employee_listbox.delete(0, END)
        self.current_display_list = employees
        for name, tab in employees:
            display = f"{name}" + (f" (т.н. {tab})" if tab else "")
            self.employee_listbox.insert(END, display)
    
    def filter_employees(self, *args):
        query = self.search_query.get().lower()
        if not query:
            self.refresh_employee_list(self.full_employees_list)
            return
            
        filtered = [
            (name, tab) for name, tab in self.full_employees_list
            if query in name.lower() or (tab and query in tab)
        ]
        self.refresh_employee_list(filtered)
    
    def on_employee_selected(self, event):
        selection = self.employee_listbox.curselection()
        if not selection:
            return
        
        selected_name = self.current_display_list[selection[0]][0]
        self.load_employee(selected_name)
    
    def load_selected_employee(self):
        selection = self.employee_listbox.curselection()
        if not selection:
            return
        
        selected_name = self.current_display_list[selection[0]][0]
        self.load_employee(selected_name)
    
    def load_employee(self, search_value):
        try:
            if self.db is None:
                try:
                    wb = xw.Book.caller()
                except:
                    wb = xw.Book(settings.EXCEL_FILE)
                
                self.db = ExcelDatabase()
                self.db.connect()
            
            employee = self.db.find_employee(str(search_value))
            
            if not employee:
                messagebox.showerror("Ошибка", "Сотрудник не найден")
                return
            
            self.employee_data = employee
            self.tab_number = employee.get("Таб. №")
            if pd.notna(self.tab_number):
                self.tab_number = int(float(self.tab_number))
            else:
                self.tab_number = None
            
            self.header_label.config(text=f"Карточка: {self.employee_data.get('ФИО', '')}")
            
            # Поля из settings.EMPLOYEE_COLUMNS
            fields = [
                "Таб. №", "ФИО", "Подразделение", "Должность", "Дата принятия", 
                "Дата рождения", "Пол", "Гражданин РБ", "Резидент РБ", "Пенсионер"
            ]
            
            for field in fields:
                value = self.employee_data.get(field, "-")
                if pd.notna(value):
                    if hasattr(value, 'strftime'):
                        value = value.strftime("%d.%m.%Y")
                else:
                    value = "-"
                self.info_labels[field].config(text=str(value))
            
            contract_start = self.employee_data.get("Начало контракта")
            contract_end = self.employee_data.get("Конец контракта")
            
            if pd.notna(contract_start):
                self.contract_start.config(text=contract_start.strftime("%d.%m.%Y"))
            else:
                self.contract_start.config(text="-")
            
            if pd.notna(contract_end):
                self.contract_end.config(text=contract_end.strftime("%d.%m.%Y"))
                days_left = self.analytics.calculate_contract_days_remaining(contract_end)
                
                if days_left < 0:
                    self.contract_days.config(text=f"Истек {abs(days_left)} дн. назад", foreground="red")
                elif days_left <= 30:
                    self.contract_days.config(text=f"{days_left} дн.", foreground="orange")
                else:
                    self.contract_days.config(text=f"{days_left} дн.", foreground="green")
            else:
                self.contract_end.config(text="-")
                self.contract_days.config(text="-")
            
            birth_date = self.employee_data.get("Дата рождения")
            hire_date = self.employee_data.get("Дата принятия")
            
            if pd.notna(birth_date):
                self.age_label.config(text=f"{self.analytics.calculate_age(birth_date)} лет")
            else:
                self.age_label.config(text="-")
            
            if pd.notna(hire_date):
                years, months = self.analytics.calculate_tenure(hire_date)
                self.tenure_label.config(text=f"{years} лет, {months} мес.")
            else:
                self.tenure_label.config(text="-")
            
            vacations = self.db.get_vacations(self.tab_number)
            used = 0
            if not vacations.empty:
                used = vacations["Количество дней"].sum()
                if float(used).is_integer():
                    used = int(used)
            
            self.vacation_used.config(text=f"{used} дн.")
            remaining = 28 - used
            if float(remaining).is_integer():
                remaining = int(remaining)
            self.vacation_remain.config(text=f"{remaining} дн.", 
                                       foreground="green" if remaining > 0 else "red")
            
        except Exception as e:
            logger.exception(f"Failed to load employee data: {e}")
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные:\n{e}")
    
    def edit_employee(self):
        messagebox.showinfo("Информация", "Функция в разработке")
    
    def create_order(self):
        messagebox.showinfo("Информация", "Функция в разработке")


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--tab", type=int, help="Tabelnyj nomer sotrudnika")
    args = parser.parse_args()
    
    try:
        app = EmployeeCardDialog(tab_number=args.tab)
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("Nazmi Enter dlya vykhoda...")


if __name__ == "__main__":
    main()
