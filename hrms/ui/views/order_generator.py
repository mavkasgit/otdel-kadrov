"""
HRMS - Генератор приказов
Создание приказов из шаблонов
"""
import os
import sys
import ctypes
import pandas as pd

myappid = 'hrms.otdelkadrov'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import xlwings as xw
import ttkbootstrap as ttk
import tkinter as tk
from ui.components.date_picker import CustomDateEntry
from ttkbootstrap.constants import *
from tkinter import messagebox, StringVar, Text, END
from tkinter import Listbox as TkListbox
from datetime import datetime
import locale

from core.db_engine import ExcelDatabase
from core.doc_generator import DocumentGenerator
import settings
import json


def format_date_with_weekday(date_obj):
    """Формат даты: понедельник, 30 марта 2026 г."""
    weekdays = {
        0: "понедельник",
        1: "вторник",
        2: "среда",
        3: "четверг",
        4: "пятница",
        5: "суббота",
        6: "воскресенье"
    }
    months = {
        1: "января", 2: "февраля", 3: "марта", 4: "апреля",
        5: "мая", 6: "июня", 7: "июля", 8: "августа",
        9: "сентября", 10: "октября", 11: "ноября", 12: "декабря"
    }
    wd = weekdays[date_obj.weekday()]
    month = months[date_obj.month]
    return f"{wd}, {date_obj.day} {month} {date_obj.year} г."


def parse_date_with_weekday(date_str):
    """Парсинг даты из формата с днём недели"""
    import re
    match = re.search(r'(\d+)\s+(\w+)\s+(\d+)', date_str)
    if match:
        day = int(match.group(1))
        year = int(match.group(3))
        month_map = {
            "января": 1, "февраля": 2, "марта": 3, "апреля": 4,
            "мая": 5, "июня": 6, "июля": 7, "августа": 8,
            "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12
        }
        month_name = match.group(2)
        month = month_map.get(month_name, 1)
        return datetime(year, month, day)
    return datetime.now()


def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")


class OrderGeneratorDialog:
    """Диалог генерации приказов"""
    
    def __init__(self, parent=None):
        if parent:
            self.root = ttk.Toplevel(parent)
        else:
            self.root = ttk.Window(themename="yeti")
            
        self.root.title("Генератор приказов")
        center_window(self.root, 1000, 750)
        
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "..", "..", "icon.png")
            icon_path = os.path.abspath(icon_path)
            icon_img = Image.open(icon_path)
            self._icon = ImageTk.PhotoImage(icon_img)
            self.root.iconphoto(True, self._icon)
        except Exception as e:
            print(f"Icon error: {e}")
        
        self.db = None
        self.order_type = StringVar(value="") # Сброс по умолчанию
        self.order_number_var = StringVar()
        self.employees_list = []
        self.employee_buttons = {}
        self.selected_employee = None
        self.selected_button = None
        
        self._load_pane_pos()
        self.setup_ui()
        self.load_employees()
        
        if not parent:
            self.root.mainloop()
    
    def _save_pane_pos(self, paned):
        """Сохранение позиции разделителя"""
        try:
            self.root.after(100, lambda: self._do_save_pane(paned))
        except:
            pass
    
    def _do_save_pane(self, paned):
        """Сохранение позиции после задержки"""
        try:
            pos = paned.sash_coord(0)[0]
            settings.ORDER_GEN_PANE_POS = int(pos)
            config_path = os.path.join(os.path.dirname(__file__), "..", "..", "settings.json")
            with open(config_path, "w") as f:
                json.dump({"order_gen_pane_pos": int(pos)}, f)
        except:
            pass
    
    def _load_pane_pos(self):
        """Загрузка позиции разделителя"""
        try:
            config_path = os.path.join(os.path.dirname(__file__), "..", "..", "settings.json")
            if os.path.exists(config_path):
                with open(config_path, "r") as f:
                    data = json.load(f)
                    settings.ORDER_GEN_PANE_POS = data.get("order_gen_pane_pos")
        except:
            pass
    
    def setup_ui(self):
        header = ttk.Frame(self.root, bootstyle=PRIMARY)
        header.pack(fill="x")
        
        ttk.Label(header, text="Генератор приказов", font=("Segoe UI", 16, "bold"),
                 bootstyle="inverse-primary").pack(pady=15)
        
        paned = tk.PanedWindow(self.root, orient="horizontal", sashwidth=10, bg="#CCCCCC")
        paned.pack(fill="both", expand=True, padx=10, pady=10)
        
        left_frame = ttk.LabelFrame(paned, text="Сотрудники")
        paned.add(left_frame, minsize=350, stretch="always")
        
        right_frame = ttk.LabelFrame(paned, text="Тип приказа")
        paned.add(right_frame, minsize=350, stretch="always")
        
        def set_initial_sash():
            try:
                if settings.ORDER_GEN_PANE_POS:
                    paned.sash_place(0, int(settings.ORDER_GEN_PANE_POS), 0)
                else:
                    paned.sash_place(0, 480, 0)
            except:
                pass
                
        self.root.after(100, set_initial_sash)
        
        paned.bind("<ButtonRelease-1>", lambda e: self._save_pane_pos(paned))
        
        self.paned = paned
        
        self.search_var = StringVar()
        search_entry = ttk.Entry(left_frame, textvariable=self.search_var)
        search_entry.pack(pady=5, padx=10, fill="x")
        search_entry.focus()
        
        self.search_entry = search_entry
        search_entry.bind("<KeyRelease>", self.on_search_change)
        
        # Сначала пакуем нижние элементы (dock to bottom)
        bottom_frame = ttk.Frame(left_frame)
        bottom_frame.pack(side="bottom", fill="x", pady=10)
        
        self.status_label = ttk.Label(bottom_frame, text="", bootstyle=INFO)
        self.status_label.pack(side="bottom", pady=2)
        
        btn_frame = ttk.Frame(bottom_frame)
        btn_frame.pack(side="bottom", fill="x")
        
        ttk.Button(btn_frame, text="Создать приказ", command=self.generate_order,
                  bootstyle=SUCCESS, padding=(15, 8)).pack(side="left", padx=5, expand=True, fill="x")
        ttk.Button(btn_frame, text="Закрыть", command=self.root.destroy,
                  padding=(15, 8)).pack(side="left", padx=5, expand=True, fill="x")
        
        self.selected_info = ttk.Label(left_frame, text="Выбран: None", 
                                       font=("Segoe UI", 10, "bold"), foreground="red")
        self.selected_info.pack(side="bottom", pady=5)
        
        # И только теперь пакуем список, чтобы он занимал оставшееся место (dock to top/fill)
        listbox_frame = ttk.Frame(left_frame)
        listbox_frame.pack(side="top", fill="both", expand=True, pady=5)
        
        scroll = ttk.Scrollbar(listbox_frame)
        scroll.pack(side="right", fill="y")
        
        self.employee_listbox = TkListbox(listbox_frame, yscrollcommand=scroll.set, 
                                        font=("Segoe UI", 10))
        self.employee_listbox.pack(side="left", fill="both", expand=True)
        scroll.config(command=self.employee_listbox.yview)
        
        self.employee_listbox.bind("<<ListboxSelect>>", self.on_employee_select)
        
        order_types = [
            "Прием на работу",
            "Увольнение", 
            "Отпуск трудовой",
            "Отпуск за свой счет",
            "Больничный",
            "Перевод",
            "Продление контракта"
        ]
        
        for ot in order_types:
            btn = ttk.Radiobutton(right_frame, text=ot, variable=self.order_type, value=ot,
                                  bootstyle="success-toolbutton")
            btn.pack(pady=2, fill="x", padx=10)
            self.employee_buttons[ot] = btn
        
        self.order_type.set("") # Сброс по умолчанию
        self.order_type.trace("w", self.on_type_click)
        
        form_frame = ttk.Frame(right_frame)
        form_frame.pack(fill="x", pady=10)
        
        try:
            locale.setlocale(locale.LC_TIME, 'ru_RU')
        except locale.Error:
            pass
        
        ttk.Label(form_frame, text="Дата приказа:").grid(row=0, column=0, sticky="w", pady=5, padx=10)
        self.date_entry = CustomDateEntry(form_frame, width=20, popup_coords=(1200 , 600), on_change=self.update_order_number)
        self.date_entry.grid(row=0, column=1, sticky="w", pady=5, padx=10)
        
        ttk.Label(form_frame, text="Номер приказа:").grid(row=1, column=0, sticky="w", pady=5, padx=10)
        
        # Контейнер для № и самого поля ввода
        num_entry_frame = ttk.Frame(form_frame)
        num_entry_frame.grid(row=1, column=1, sticky="w", pady=5, padx=10)
        
        ttk.Label(num_entry_frame, text="№", font=("Segoe UI", 11, "bold")).pack(side="left", padx=(0, 2))
        self.order_number_entry = ttk.Entry(num_entry_frame, textvariable=self.order_number_var, width=15, font=("Segoe UI", 11, "bold"), foreground="blue")
        self.order_number_entry.pack(side="left")
        
        recent_frame = ttk.LabelFrame(right_frame, text="Последние 5 приказов")
        recent_frame.pack(fill="both", expand=True, pady=10, padx=10)
        
        columns = ("number", "fio", "date", "type")
        self.recent_tree = ttk.Treeview(recent_frame, columns=columns, show="headings", height=5)
        self.recent_tree.heading("number", text="Номер приказа")
        self.recent_tree.column("number", width=90)
        self.recent_tree.heading("fio", text="ФИО")
        self.recent_tree.column("fio", width=120)
        self.recent_tree.heading("date", text="Дата")
        self.recent_tree.column("date", width=80)
        self.recent_tree.heading("type", text="Тип")
        self.recent_tree.column("type", width=100)
        self.recent_tree.pack(fill="both", expand=True, padx=5, pady=5)
        
        extra_frame = ttk.LabelFrame(right_frame, text="Дополнительно")
        extra_frame.pack(fill="x", pady=10, padx=10)
        
        self.extra_text = Text(extra_frame, height=4)
        self.extra_text.pack(fill="both", expand=True, padx=5, pady=5)
    
    def on_search_change(self, event=None):
        """Поиск сотрудников"""
        search_text = self.search_var.get().lower().strip()
        
        self.employee_listbox.delete(0, "end")
        
        if not search_text:
            self.refresh_employee_list()
            return
        
        for name, tab in self.employees_list:
            search_str = f"{name}" + (f" {tab}" if tab else "")
            if search_text in search_str.lower():
                display = name + (f" (т.н. {tab})" if tab else "")
                self.employee_listbox.insert("end", display)
    
    def on_employee_select(self, event):
        """Выбор сотрудника из списка"""
        selection = self.employee_listbox.curselection()
        if selection:
            idx = selection[0]
            name = self.employees_list[idx][0]
            tab = self.employees_list[idx][1]
            self.selected_employee = name
            
            display = name + (f" (т.н. {tab})" if tab else "")
            self.selected_info.config(text=f"Выбран: {display}", foreground="green")
    
    def on_type_click(self, *args):
        """Выбор типа приказа"""
        order_type = self.order_type.get()
        self.update_order_number()
    
    def load_employees(self):
        """Загрузка списка сотрудников"""
        try:
            try:
                wb = xw.Book.caller()
            except:
                wb = xw.Book(settings.EXCEL_FILE)
            
            self.db = ExcelDatabase()
            self.db.connect()
            
            employees = self.db.get_employees()
            
            self.employees_list = []
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
                    self.employees_list.append((name, tab_str))
            
            self.employees_list.sort(key=lambda x: x[0])
            
            self.refresh_employee_list()
            self.update_order_number()
            self.refresh_recent_orders()
            
        except Exception as e:
            logger.exception(f"Failed to load employees in order generator: {e}")
            messagebox.showerror("Ошибка", f"Не удалось загрузить:\n{e}")
    
    def refresh_employee_list(self):
        """Обновить список сотрудников"""
        self.employee_listbox.delete(0, "end")
        for name, tab in self.employees_list:
            display = f"{name}" + (f" (т.н. {tab})" if tab else "")
            self.employee_listbox.insert("end", display)
    
    def update_order_number(self):
        """Обновить номер приказа"""
        if self.db is None:
            self.order_number_var.set("Загрузка...")
            return
        try:
            order_type = self.order_type.get()
            if not order_type:
                self.order_number_var.set("")
                return
            
            # Получаем год из поля даты
            date_str = self.date_entry.get()
            try:
                order_year = datetime.strptime(date_str, "%d.%m.%Y").year
            except:
                order_year = datetime.now().year
                
            full_num = self.db.get_next_order_number(order_type, year=order_year)
            self.order_number_var.set(full_num)
        except Exception as e:
            self.order_number_var.set("")
    
    def generate_order(self):
        """Создать приказ"""
        if self.selected_employee is None:
            messagebox.showwarning("Внимание", "Выберите сотрудника")
            return
        
        if not self.date_entry.entry.get():
            messagebox.showwarning("Внимание", "Укажите дату")
            return
        
        try:
            search_value = self.selected_employee
            
            order_date = self.date_entry.entry.get()
            order_date = datetime.strptime(order_date, "%d.%m.%Y")
            order_type = self.order_type.get()
            
            if not order_type:
                messagebox.showwarning("Внимание", "Выберите тип приказа")
                return
            
            order_number = self.order_number_var.get()
            
            employee = self.db.find_employee(str(search_value))
            
            if not employee:
                messagebox.showerror("Ошибка", f"Сотрудник '{search_value}' не найден в базе.")
                return
            
            emp_dict = employee
            
            doc_gen = DocumentGenerator()
            file_path = doc_gen.generate_order(order_type, emp_dict, order_number, order_date)
            
            order_data = {
                "order_type": order_type,
                "search_value": search_value,
                "Номер приказа": order_number,
                "Дата создания": order_date,
                "Путь к файлу": file_path
            }
            
            self.db.add_order_log(order_data)
            
            self.status_label.config(text=f"Приказ № {order_number} создан")
            self.update_order_number()
            self.refresh_recent_orders()
            
            if messagebox.askyesno("Успех", f"Приказ № {order_number} успешно создан!\n\nФайл: {file_path}\n\nОткрыть файл сейчас?"):
                import os
                os.startfile(file_path)
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать приказ:\n{e}")
            import traceback
            traceback.print_exc()

    def refresh_recent_orders(self):
        """Обновление списка последних 5 приказов"""
        if self.db is None or not hasattr(self, 'recent_tree'):
            return
            
        for item in self.recent_tree.get_children():
            self.recent_tree.delete(item)
            
        try:
            df = self.db.get_order_log()
            if df.empty:
                return
                
            # Get last 5 reverse order
            recent = df.tail(5).iloc[::-1]
            
            for _, row in recent.iterrows():
                # Извлекаем значения с защитой от изменения названий колонок
                num_val = row.get("Номер приказа") if "Номер приказа" in row.index else row.iloc[0]
                type_val = row.get("Тип события") if "Тип события" in row.index else row.iloc[1]
                date_val = row.get("Дата создания") if "Дата создания" in row.index else row.iloc[2]
                fio_val = row.get("ФИО") if "ФИО" in row.index else row.iloc[3]

                # Форматируем значения
                num_str = str(num_val) if pd.notna(num_val) else ""
                type_str = str(type_val) if pd.notna(type_val) else ""
                fio_str = str(fio_val) if pd.notna(fio_val) else ""
                
                # Форматируем дату
                date_str = date_val.strftime("%d.%m.%Y") if (pd.notna(date_val) and hasattr(date_val, "strftime")) else ""
                
                # Форматируем ФИО (Иванов Иван Иванович -> Иванов И.И.)
                fio_clean = fio_str.strip()
                if fio_clean and fio_clean != "None":
                    fio_parts = fio_clean.split()
                    if len(fio_parts) >= 3:
                        fio_str = f"{fio_parts[0]} {fio_parts[1][0]}.{fio_parts[2][0]}."
                    elif len(fio_parts) == 2:
                        fio_str = f"{fio_parts[0]} {fio_parts[1][0]}."
                
                self.recent_tree.insert("", "end", values=(num_str, fio_str, date_str, type_str))
        except Exception as e:
            pass


def main():
    try:
        app = OrderGeneratorDialog()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("Нажми Enter для выхода...")


if __name__ == "__main__":
    main()
