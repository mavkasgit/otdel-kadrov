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
from ttkbootstrap.constants import *
from tkinter import messagebox, StringVar, Listbox
from datetime import datetime

from core.db_engine import ExcelDatabase
from core.validator import DataValidator
from core.analytics import AnalyticsEngine
import settings


def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")


class VacationManagerDialog:
    
    def __init__(self, parent=None):
        self.root = ttk.Window(themename="yeti")
        self.root.title("Управление отпусками")
        center_window(self.root, 500, 600)
        
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "..", "..", "icon.ico")
            icon_path = os.path.abspath(icon_path)
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Icon error: {e}")
        
        self.db = None
        self.validator = None
        self.analytics = AnalyticsEngine()
        
        self.employees = []
        self.selected_employee = StringVar()
        
        self.setup_ui()
        self.load_employees()
        
        self.root.mainloop()
    
    def setup_ui(self):
        ttk.Label(self.root, text="Отпуска сотрудников", font=("Segoe UI", 14, "bold")).pack(pady=10)
        
        ttk.Label(self.root, text="Выберите сотрудника:").pack(pady=5)
        self.employee_combo = ttk.Combobox(self.root, textvariable=self.selected_employee, 
                                           state="readonly", width=40)
        self.employee_combo.pack(pady=5)
        self.employee_combo.bind("<<ComboboxSelected>>", self.on_employee_selected)
        
        ttk.Button(self.root, text="Показать отпуска", command=self.show_vacations,
               bootstyle=INFO, padding=10).pack(pady=10)
        
        ttk.Label(self.root, text="Список отпусков:").pack(pady=5)
        
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
            try:
                wb = xw.Book.caller()
            except:
                wb = xw.Book(settings.EXCEL_FILE)
            
            self.db = ExcelDatabase()
            self.db.connect()
            self.validator = DataValidator(self.db)
            
            employees = self.db.get_employees()
            self.employees = [(row["Tab. #"], row["FIO"]) for _, row in employees.iterrows()]
            
            self.employee_combo["values"] = [f"{tab} - {name}" for tab, name in self.employees]
            
        except Exception as e:
            messagebox.showerror("Oshibka", f"Ne udalos zagruzit sotrudnikov:\n{e}")
    
    def on_employee_selected(self, event):
        self.show_vacations()
    
    def show_vacations(self):
        self.vacations_listbox.delete(0, END)
        
        if not self.selected_employee.get():
            return
        
        tab_str = self.selected_employee.get().split(" - ")[0]
        tab_number = int(tab_str)
        
        try:
            vacations = self.db.get_vacations(tab_number)
            
            if vacations.empty:
                self.vacations_listbox.insert(0, "Net otpuskov")
                return
            
            for _, vac in vacations.iterrows():
                start = vac.get("Data nachala", "")
                end = vac.get("Data okonchaniya", "")
                days = vac.get("Kolichestvo dnej", 0)
                vtype = vac.get("Tip otpuska", "")
                
                if isinstance(start, datetime):
                    start = start.strftime("%d.%m.%Y")
                if isinstance(end, datetime):
                    end = end.strftime("%d.%m.%Y")
                
                text = f"{start} - {end} ({days} dn.) - {vtype}"
                self.vacations_listbox.insert(END, text)
            
            all_vac = self.db.get_vacations()
            emp_vac = all_vac[all_vac["Tab. #"] == tab_number]
            total_days = emp_vac["Kolichestvo dnej"].sum() if not emp_vac.empty else 0
            remaining = 28 - total_days
            
            self.status_label.config(text=f"Ispolzovano: {total_days} dn. Ostalos: {remaining} dn.")
            
        except Exception as e:
            messagebox.showerror("Oshibka", f"Ne udalos zagruzit otpusta:\n{e}")
    
    def add_vacation(self):
        if not self.selected_employee.get():
            messagebox.showwarning("Vnimanie", "Viberite sotrudnika")
            return
        
        tab_str = self.selected_employee.get().split(" - ")[0]
        tab_number = int(tab_str)
        
        dialog = AddVacationDialog(self.root, tab_number)
        if dialog.result:
            vac_data = dialog.result
            
            if self.validator:
                is_valid, error = self.validator.validate_vacation_data(vac_data)
                if not is_valid:
                    messagebox.showerror("Oshibka validacii", error)
                    return
                
                start = vac_data["Data nachala"]
                end = vac_data["Data okonchaniya"]
                has_overlap, overlaps = self.validator.check_vacation_overlap(
                    tab_number, start, end
                )
                
                if has_overlap:
                    msg = "Vnimanie! Peresechenie s sushestvuyushchimi otpuskami:\n"
                    for ov in overlaps:
                        msg += f"- {ov['start']} - {ov['end']}\n"
                    msg += "\nVsyo ravno dobavit?"
                    if not messagebox.askyesno("Peresechenie", msg):
                        return
            
            try:
                self.db.add_vacation(vac_data)
                messagebox.showinfo("Uspeshno", "Otpust dobavlen!")
                self.show_vacations()
            except Exception as e:
                messagebox.showerror("Oshibka", f"Ne udalos dobavit otpusk:\n{e}")
    
    def delete_vacation(self):
        if not self.selected_employee.get():
            return
        
        selection = self.vacations_listbox.curselection()
        if not selection:
            messagebox.showwarning("Vnimanie", "Viberite otpusk dlya udaleniya")
            return
        
        if messagebox.askyesno("Podtverzhdenie", "Udalit vbrannyj otpusk?"):
            messagebox.showinfo("Info", "Funkciya v razrabotke")


class AddVacationDialog:
    
    def __init__(self, parent, tab_number):
        self.result = None
        self.tab_number = tab_number
        
        self.dialog = Toplevel(parent)
        self.dialog.title("Добавить отпуск")
        center_window(self.dialog, 400, 300)
        
        self.setup_ui()
        self.dialog.grab_set()
        self.dialog.wait_window()
    
    def setup_ui(self):
        ttk.Label(self.dialog, text=f"Tab. #: {self.tab_number}").pack(pady=5)
        
        ttk.Label(self.dialog, text="Дата начала:").pack()
        self.start_entry = ttk.Entry(self.dialog)
        self.start_entry.pack(pady=5)
        self.start_entry.insert(0, "01.07.2025")
        
        ttk.Label(self.dialog, text="Дата окончания:").pack()
        self.end_entry = ttk.Entry(self.dialog)
        self.end_entry.pack(pady=5)
        self.end_entry.insert(0, "14.07.2025")
        
        ttk.Label(self.dialog, text="Тип отпуска:").pack()
        self.type_var = StringVar(value="Trudovoj otpusk")
        types = ["Trudovoj otpusk", "Otpusk za svoj schyot", "Uchebnyj otpusk", "Dekretnyj"]
        
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
                "Tab. #": self.tab_number,
                "Data nachala": start_date,
                "Data okonchaniya": end_date,
                "Tip otpuska": self.type_var.get(),
                "Kolichestvo dnej": days
            }
            
            self.dialog.destroy()
            
        except ValueError as e:
            messagebox.showerror("Oshibka", "Nevernyj format daty. Ispolzujte DD.MM.GGGG")


def main():
    try:
        app = VacationManagerDialog()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("Nazmi Enter dlya vykhoda...")


if __name__ == "__main__":
    main()
