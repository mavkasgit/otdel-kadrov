"""
HRMS - Card employee
"""
import os
import sys
import ctypes

myappid = 'hrms.otdelkadrov'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import xlwings as xw
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox, StringVar
from datetime import datetime

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


class EmployeeCardDialog:
    
    def __init__(self, parent=None, tab_number=None):
        self.root = ttk.Window(themename="yeti")
        self.root.title("Карточка сотрудника")
        center_window(self.root, 500, 700)
        
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "..", "..", "icon.ico")
            icon_path = os.path.abspath(icon_path)
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Icon error: {e}")
        
        self.db = None
        self.analytics = AnalyticsEngine()
        self.tab_number = tab_number
        self.employee_data = None
        
        self.setup_ui()
        
        if tab_number:
            self.load_employee(tab_number)
        
        self.root.mainloop()
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        self.header_label = ttk.Label(main_frame, text="Карточка сотрудника", 
                                   font=("Segoe UI", 14, "bold"))
        self.header_label.pack(pady=10)
        
        if self.tab_number is None:
            ttk.Label(main_frame, text="Выберите сотрудника:").pack(pady=5)
            self.employee_var = StringVar()
            self.employee_combo = ttk.Combobox(main_frame, textvariable=self.employee_var,
                                               state="readonly", width=40)
            self.employee_combo.pack(pady=5)
            self.employee_combo.bind("<<ComboboxSelected>>", self.on_employee_selected)
            ttk.Button(main_frame, text="Загрузить", command=self.load_selected_employee,
                      bootstyle=INFO).pack(pady=5)
            self.load_employee_list()
        
        self.info_frame = ttk.LabelFrame(main_frame, text="Основная информация")
        self.info_frame.pack(fill="x", pady=10)
        
        self.info_labels = {}
        fields = ["Tab. #", "FIO", "Podrazdelenie", "Dolzhnost", "Data priyatiya", 
                  "Data rozhdeniya", "Pol", "Grazhdanin RB", "Rezident RB", "Pensioner"]
        
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
            try:
                wb = xw.Book.caller()
            except:
                wb = xw.Book(settings.EXCEL_FILE)
            
            self.db = ExcelDatabase()
            self.db.connect()
            
            employees = self.db.get_employees()
            self.employees_list = [(row["Tab. #"], row["FIO"]) for _, row in employees.iterrows()]
            
            self.employee_combo["values"] = [f"{tab} - {name}" for tab, name in self.employees_list]
            
        except Exception as e:
            messagebox.showerror("Oshibka", f"Ne udalos zagruzit sotrudnikov:\n{e}")
    
    def on_employee_selected(self, event):
        self.load_selected_employee()
    
    def load_selected_employee(self):
        if not self.employee_var.get():
            return
        
        tab_str = self.employee_var.get().split(" - ")[0]
        self.load_employee(int(tab_str))
    
    def load_employee(self, tab_number):
        try:
            if self.db is None:
                try:
                    wb = xw.Book.caller()
                except:
                    wb = xw.Book(settings.EXCEL_FILE)
                
                self.db = ExcelDatabase()
                self.db.connect()
            
            employees = self.db.get_employees()
            emp = employees[employees["Tab. #"] == tab_number]
            
            if emp.empty:
                messagebox.showerror("Oshibka", "Sotrudnik ne najden")
                return
            
            self.employee_data = emp.iloc[0]
            self.tab_number = tab_number
            
            self.header_label.config(text=f"Karto4ka: {self.employee_data.get('FIO', '')}")
            
            fields = ["Tab. #", "FIO", "Podrazdelenie", "Dolzhnost", "Data priyatiya", 
                      "Data rozhdeniya", "Pol", "Grazhdanin RB", "Rezident RB", "Pensioner"]
            
            for field in fields:
                value = self.employee_data.get(field, "-")
                if isinstance(value, datetime):
                    value = value.strftime("%d.%m.%Y")
                self.info_labels[field].config(text=str(value))
            
            contract_start = self.employee_data.get("Na4alo kontrakta")
            contract_end = self.employee_data.get("Konec kontrakta")
            
            if isinstance(contract_start, datetime):
                self.contract_start.config(text=contract_start.strftime("%d.%m.%Y"))
            else:
                self.contract_start.config(text="-")
            
            if isinstance(contract_end, datetime):
                self.contract_end.config(text=contract_end.strftime("%d.%m.%Y"))
                days_left = self.analytics.calculate_contract_days_remaining(contract_end)
                
                if days_left < 0:
                    self.contract_days.config(text=f"Istek {abs(days_left)} dn. nazad", fg="red")
                elif days_left <= 30:
                    self.contract_days.config(text=f"{days_left} dn.", fg="orange")
                else:
                    self.contract_days.config(text=f"{days_left} dn.", fg="green")
            else:
                self.contract_end.config(text="-")
                self.contract_days.config(text="-")
            
            birth_date = self.employee_data.get("Data rozhdeniya")
            hire_date = self.employee_data.get("Data priyatiya")
            
            if isinstance(birth_date, datetime):
                age = self.analytics.calculate_age(birth_date)
                self.age_label.config(text=f"{age} let")
            else:
                self.age_label.config(text="-")
            
            if isinstance(hire_date, datetime):
                years, months = self.analytics.calculate_tenure(hire_date)
                self.tenure_label.config(text=f"{years} let, {months} mes.")
            else:
                self.tenure_label.config(text="-")
            
            vacations = self.db.get_vacations(tab_number)
            if not vacations.empty:
                used = vacations["Kolichestvo dnej"].sum()
            else:
                used = 0
            
            self.vacation_used.config(text=f"{used} dn.")
            remaining = 28 - used
            self.vacation_remain.config(text=f"{remaining} dn.", 
                                       fg="green" if remaining > 0 else "red")
            
        except Exception as e:
            messagebox.showerror("Oshibka", f"Ne udalos zagruzit dannye:\n{e}")
            import traceback
            traceback.print_exc()
    
    def edit_employee(self):
        messagebox.showinfo("Informaciya", "Funkciya v razrabotke")
    
    def create_order(self):
        messagebox.showinfo("Informaciya", "Funkciya v razrabotke")


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
