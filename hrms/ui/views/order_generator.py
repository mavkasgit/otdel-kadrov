"""
HRMS - Генератор приказов
Создание приказов из шаблонов
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
from tkinter import messagebox, StringVar, Text, Frame, Label, Entry, Radiobutton, Button, LabelFrame
from datetime import datetime

from core.db_engine import ExcelDatabase
import settings


def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")


class OrderGeneratorDialog:
    """Диалог генерации приказов"""
    
    def __init__(self, parent=None):
        self.root = ttk.Window(themename="yeti") if parent is None else ttk.Window(parent, themename="yeti")
        self.root.title("Генератор приказов")
        center_window(self.root, 600, 700)
        
        try:
            from PIL import Image, ImageTk
            icon_path = os.path.join(os.path.dirname(__file__), "..", "..", "icon.ico")
            icon_path = os.path.abspath(icon_path)
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Icon error: {e}")
        
        self.db = None
        self.order_type = StringVar(value="Прием на работу")
        self.order_number_var = StringVar()
        self.employees_list = []
        
        self.setup_ui()
        self.load_employees()
        
        self.root.mainloop()
    
    def setup_ui(self):
        """Настройка интерфейса"""
        main_frame = Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Заголовок
        Label(main_frame, text="Создание приказа", font=("Arial", 14, "bold")).pack(pady=10)
        
        # Выбор типа приказа
        Label(main_frame, text="Тип приказа:", font=("Arial", 11, "bold")).pack(pady=5)
        
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
            Radiobutton(main_frame, text=ot, variable=self.order_type, 
                       value=ot, command=self.on_type_changed).pack(anchor="w", padx=20)
        
        # Выбор сотрудника
        Label(main_frame, text="\nВыберите сотрудника:", font=("Arial", 11, "bold")).pack(pady=5)
        
        self.employee_var = StringVar()
        self.employee_combo = ttk.Combobox(main_frame, textvariable=self.employee_var,
                                           state="readonly", width=40)
        self.employee_combo.pack(pady=5)
        
        # Дата приказа
        Label(main_frame, text="\nДата приказа:", font=("Arial", 11, "bold")).pack(pady=5)
        
        self.date_entry = Entry(main_frame, width=20)
        self.date_entry.pack(pady=5)
        self.date_entry.insert(0, datetime.now().strftime("%d.%m.%Y"))
        
        # Номер приказа (авто)
        Label(main_frame, text="\nНомер приказа:", font=("Arial", 11, "bold")).pack(pady=5)
        
        Label(main_frame, textvariable=self.order_number_var, 
              font=("Arial", 12), fg="blue").pack(pady=5)
        
        # Дополнительные поля
        self.extra_frame = LabelFrame(main_frame, text="Дополнительно", padx=10, pady=10)
        self.extra_frame.pack(fill="x", pady=10)
        
        self.extra_text = Text(self.extra_frame, height=5, width=50)
        self.extra_text.pack()
        
        # Кнопки
        btn_frame = Frame(main_frame)
        btn_frame.pack(pady=15)
        
        Button(btn_frame, text="Создать приказ", command=self.generate_order,
               bg="#4CAF50", fg="white", padx=20).pack(side="left", padx=5)
        
        Button(btn_frame, text="Закрыть", command=self.root.destroy,
               padx=20).pack(side="left", padx=5)
        
        # Статус
        self.status_label = Label(main_frame, text="", fg="blue")
        self.status_label.pack(pady=5)
    
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
            self.employees_list = [(row["Таб. №"], row["ФИО"]) for _, row in employees.iterrows()]
            
            self.employee_combo["values"] = [f"{tab} - {name}" for tab, name in self.employees_list]
            
            # Загружаем номер приказа
            self.update_order_number()
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить:\n{e}")
    
    def on_type_changed(self):
        """При изменении типа приказа"""
        self.update_order_number()
    
    def update_order_number(self):
        """Обновить номер приказа"""
        if self.db is None:
            self.order_number_var.set("Загрузка...")
            return
        try:
            order_type = self.order_type.get()
            next_num = self.db.get_next_order_number(order_type)
            self.order_number_var.set(next_num)
        except Exception as e:
            self.order_number_var.set(f"Ошибка: {e}")
    
    def generate_order(self):
        """Создать приказ"""
        if not self.employee_var.get():
            messagebox.showwarning("Внимание", "Выберите сотрудника")
            return
        
        if not self.date_entry.get():
            messagebox.showwarning("Внимание", "Укажите дату")
            return
        
        try:
            tab_str = self.employee_var.get().split(" - ")[0]
            tab_number = int(tab_str)
            
            order_date = datetime.strptime(self.date_entry.get(), "%d.%m.%Y")
            order_type = self.order_type.get()
            order_number = self.order_number_var.get()
            
            # Получить данные сотрудника
            employees = self.db.get_employees()
            emp = employees[employees["Таб. №"] == tab_number].iloc[0]
            
            # Записать в журнал приказов
            order_data = {
                "Номер приказа": order_number,
                "Тип события": order_type,
                "Дата создания": order_date,
                "ФИО": emp.get("ФИО", ""),
                "Таб. №": tab_number,
                "Путь к файлу": ""
            }
            
            self.db.add_order_log(order_data)
            
            # Обновить счётчик
            self.db.save_order_number(order_type)
            
            messagebox.showinfo("Успех", f"Приказ {order_number} создан!\n\nЗаписан в журнал приказов.")
            
            self.status_label.config(text=f"Приказ {order_number} создан")
            
            # Обновить номер для следующего приказа
            self.update_order_number()
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать приказ:\n{e}")
            import traceback
            traceback.print_exc()


def main():
    """Точка входа"""
    try:
        app = OrderGeneratorDialog()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("Нажми Enter для выхода...")


if __name__ == "__main__":
    main()
