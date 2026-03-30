"""
HRMS - Upravlenie otpuskami
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
from ui.components.date_picker import CustomDateEntry
from ttkbootstrap.constants import *
from tkinter import messagebox, StringVar, Toplevel, LEFT, RIGHT, Y, END, Listbox
from PIL import Image, ImageTk
from datetime import datetime

from core.db_engine import ExcelDatabase
from core.validator import DataValidator
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


class VacationManagerDialog:
    
    def __init__(self, parent=None):
        if parent:
            self.root = ttk.Toplevel(parent)
        else:
            self.root = ttk.Window(themename="yeti")
            
        self.root.title("Управление отпусками")
        center_window(self.root, 800, 700)
        
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
        self.validator = None
        self.full_employees_list = []
        
        self.employees = []
        self.selected_employee = StringVar()
        
        self.setup_ui()
        self.load_employees()
        
        if not parent:
            self.root.mainloop()
    
    def setup_ui(self):
        ttk.Label(self.root, text="Отпуска сотрудников", font=("Segoe UI", 14, "bold")).pack(pady=10)
        
        # Поиск сотрудника
        ttk.Label(self.root, text="Поиск сотрудника (имя или таб. №):").pack(pady=(5, 0))
        self.search_query = StringVar()
        self.search_query.trace_add("write", self.filter_employees)
        self.search_entry = ttk.Entry(self.root, textvariable=self.search_query, width=50)
        self.search_entry.pack(pady=5)
        
        # Список результатов
        self.employee_listbox = Listbox(self.root, height=6, width=60)
        self.employee_listbox.pack(pady=5)
        self.employee_listbox.bind("<<ListboxSelect>>", self.on_employee_selected)
        
        ttk.Label(self.root, text="Список отпусков:").pack(pady=(10, 5))
        
        scroll = ttk.Scrollbar(self.root)
        scroll.pack(side=RIGHT, fill=Y)
        self.vacations_listbox = Listbox(self.root, height=15, width=50, yscrollcommand=scroll.set)
        self.vacations_listbox.pack(pady=5)
        scroll.config(command=self.vacations_listbox.yview)
        
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Добавить отпуск", command=self.add_vacation,
               bootstyle=SUCCESS, padding=10).pack(side=LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Удалить", command=self.delete_vacation,
               bootstyle=DANGER, padding=10).pack(side=LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Закрыть", command=self.root.destroy,
               padding=10).pack(side=LEFT, padx=5)
        
        self.status_label = ttk.Label(self.root, text="", bootstyle=INFO)
        self.status_label.pack(pady=5)
    
    def load_employees(self):
        try:
            if self.db is None:
                try:
                    wb = xw.Book.caller()
                except:
                    wb = xw.Book(settings.EXCEL_FILE)
                
                self.db = ExcelDatabase()
                self.db.connect()
                self.validator = DataValidator(self.db)
            
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
            logger.exception(f"Failed to load employees in vacation manager")
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
        self.show_vacations()
    
    def show_vacations(self):
        self.vacations_listbox.delete(0, END)
        
        selection = self.employee_listbox.curselection()
        if not selection:
            return
        
        selected_name = self.current_display_list[selection[0]][0]
        employee = self.db.find_employee(selected_name)
        
        if not employee:
            self.vacations_listbox.insert(0, "Сотрудник не найден")
            return
        
        tab_number = employee.get("Таб. №")
        if pd.isna(tab_number):
            tab_number = None
        
        try:
            vacations = self.db.get_vacations(tab_number)
            
            if vacations.empty:
                self.vacations_listbox.insert(0, "Нет отпусков")
                return
            
            for _, vac in vacations.iterrows():
                start = vac.get("Дата начала")
                end = vac.get("Дата окончания")
                days = vac.get("Количество дней", 0)
                vtype = vac.get("Тип отпуска", "")
                
                # Форматируем даты если они есть
                start_str = start.strftime("%d.%m.%Y") if pd.notna(start) else "-"
                end_str = end.strftime("%d.%m.%Y") if pd.notna(end) else "-"
                
                display_text = f"{start_str} - {end_str} ({days} дн.) [{vtype}]"
                self.vacations_listbox.insert(END, display_text)
            
            # Статистика
            all_vac = self.db.get_vacations()
            if tab_number is not None:
                emp_vac = all_vac[all_vac["Таб. №"] == tab_number]
            else:
                emp_vac = all_vac[all_vac["ФИО"] == selected_name]
            total_days = int(emp_vac["Количество дней"].sum()) if not emp_vac.empty else 0
            remaining = 28 - total_days
            
            self.status_label.config(text=f"Использовано: {total_days} дн. Осталось: {remaining} дн.")
            
        except Exception as e:
            logger.exception(f"Failed to show vacations for tab number {tab_number}")
            messagebox.showerror("Ошибка", f"Не удалось загрузить отпуска:\n{e}")
    
    def add_vacation(self):
        selection = self.employee_listbox.curselection()
        if not selection:
            messagebox.showwarning("Внимание", "Выберите сотрудника")
            return
        
        selected_name = self.current_display_list[selection[0]][0]
        employee = self.db.find_employee(selected_name)
        
        if not employee:
            messagebox.showerror("Ошибка", "Сотрудник не найден")
            return
        
        tab_number = employee.get("Таб. №")
        if pd.notna(tab_number):
            tab_number = int(float(tab_number))
        else:
            tab_number = None
        
        dialog = AddVacationDialog(self.root, tab_number, selected_name)
        if dialog.result:
            vac_data = dialog.result
            
            if self.validator:
                is_valid, error = self.validator.validate_vacation_data(vac_data)
                if not is_valid:
                    messagebox.showerror("Ошибка валидации", error)
                    return
                
                start = vac_data["Дата начала"]
                end = vac_data["Дата окончания"]
                if tab_number:
                    has_overlap, overlaps = self.validator.check_vacation_overlap(
                        tab_number, start, end
                    )
                    
                    if has_overlap:
                        msg = "Внимание! Пересечение с существующими отпусками:\n"
                        for ov in overlaps:
                            msg += f"- {ov['start']} - {ov['end']}\n"
                        msg += "\nВсё равно добавить?"
                        if not messagebox.askyesno("Пересечение", msg):
                            return
            
            try:
                db_vac_data = {
                    "search_value": selected_name,
                    "Дата начала": vac_data["Дата начала"],
                    "Дата окончания": vac_data["Дата окончания"],
                    "Тип отпуска": vac_data["Тип отпуска"],
                    "Количество дней": vac_data["Количество дней"]
                }
                self.db.add_vacation(db_vac_data)
                messagebox.showinfo("Успешно", "Отпуск добавлен!")
                self.show_vacations()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось добавить отпуск:\n{e}")
    
    def delete_vacation(self):
        selection = self.employee_listbox.curselection()
        if not selection:
            return
        
        selection = self.vacations_listbox.curselection()
        if not selection:
            messagebox.showwarning("Внимание", "Выберите отпуск для удаления")
            return
        
        if messagebox.askyesno("Подтверждение", "Удалить выбранный отпуск?"):
            messagebox.showinfo("Инфо", "Функция в разработке")


class AddVacationDialog:
    
    def __init__(self, parent, tab_number, fio=None):
        self.result = None
        self.tab_number = tab_number
        self.fio = fio
        
        self.dialog = Toplevel(parent)
        self.dialog.title("Добавить отпуск")
        center_window(self.dialog, 400, 300)
        
        self.setup_ui()
        self.dialog.grab_set()
        self.dialog.wait_window()
    
    def setup_ui(self):
        if self.fio:
            ttk.Label(self.dialog, text=f"Сотрудник: {self.fio}").pack(pady=5)
        if self.tab_number:
            ttk.Label(self.dialog, text=f"Таб. №: {self.tab_number}").pack(pady=5)
        
        ttk.Label(self.dialog, text="Дата начала:").pack()
        self.start_entry = CustomDateEntry(self.dialog)
        self.start_entry.pack(pady=5)
        
        ttk.Label(self.dialog, text="Дата окончания:").pack()
        self.end_entry = CustomDateEntry(self.dialog)
        self.end_entry.pack(pady=5)
        
        ttk.Label(self.dialog, text="Тип отпуска:").pack()
        self.type_var = StringVar(value="Ежегодный отпуск")
        types = ["Ежегодный отпуск", "Отпуск без сохранения з/п", "Учебный отпуск", "Больничный"]
        
        for t in types:
            ttk.Radiobutton(self.dialog, text=t, variable=self.type_var, value=t).pack()
        
        btn_frame = ttk.Frame(self.dialog)
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="Сохранить", command=self.save, 
               bootstyle=SUCCESS, padding=10).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="Отмена", command=self.dialog.destroy,
               padding=10).pack(side=LEFT, padx=5)
    
    def save(self):
        try:
            start_str = self.start_entry.get()
            end_str = self.end_entry.get()
            
            start_date = datetime.strptime(start_str, "%d.%m.%Y")
            end_date = datetime.strptime(end_str, "%d.%m.%Y")
            
            days = (end_date - start_date).days + 1
            
            self.result = {
                "Таб. №": self.tab_number,
                "Дата начала": start_date,
                "Дата окончания": end_date,
                "Тип отпуска": self.type_var.get(),
                "Количество дней": days
            }
            
            self.dialog.destroy()
            
        except ValueError as e:
            messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ДД.ММ.ГГГГ")


def main():
    try:
        app = VacationManagerDialog()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("Nazmi Enter dlya vykhoda...")


if __name__ == "__main__":
    main()
