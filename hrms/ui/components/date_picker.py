"""
HRMS - Стабильный локализованный выбор даты
Компонент на базе Frame + Entry + DatePickerDialog для предотвращения TclError
"""
import locale
from datetime import datetime
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import DatePickerDialog

import calendar

# Жестко прописываем русские названия (так как locale может сбоить на разных версиях Windows)
calendar.day_abbr = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
calendar.day_name = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]
calendar.month_name = ["", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
calendar.month_abbr = ["", "Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]

# Жесткий фикс для дней недели (ttkbootstrap использует внутренний переводчик, игнорируя системный Питона)
def _russian_headers(self):
    weekdays = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
    return weekdays[self.firstweekday:] + weekdays[:self.firstweekday]

DatePickerDialog._header_columns = _russian_headers

class CustomDateEntry(ttk.Frame):
    """
    Стабильный аналог DateEntry, использующий обычный Entry и кнопку
    вызова DatePickerDialog. Решает проблему TclError в сложной иерархии окон.
    """
    def __init__(self, master=None, width=15, popup_coords=None, **kwargs):
        # Очищаем kwargs от параметров, не относящихся к Frame
        self.bootstyle = kwargs.pop('bootstyle', DEFAULT)
        self.firstweekday = kwargs.pop('firstweekday', 0)
        self.dateformat = kwargs.pop('dateformat', "%d.%m.%Y")
        
        self.popup_coords = popup_coords
        
        super().__init__(master, **kwargs)
        
        # Общий стиль для компонентов
        self.entry = ttk.Entry(self, width=width, bootstyle=self.bootstyle)
        self.entry.pack(side=LEFT, fill=X, expand=True)
        
        # Кнопка вызова календаря (используем текст или символ, чтобы не зависеть от ресурсов)
        self.button = ttk.Button(
            self, 
            text="📅", 
            command=self._show_calendar,
            bootstyle=f"{self.bootstyle}-outline",
            width=3
        )
        self.button.pack(side=LEFT, padx=(2, 0))
        
        # Устанавливаем текущую дату по умолчанию
        self.entry.insert(0, datetime.now().strftime(self.dateformat))

    def _show_calendar(self):
        """Прямой вызов диалога выбора даты"""
        # Пытаемся распарсить текущую дату из поля ввода
        try:
            start_date = datetime.strptime(self.entry.get(), self.dateformat)
        except:
            start_date = datetime.now()
            
        import tkinter as tk
        
        # Определяем функцию калибровки
        def move_calendar():
            try:
                # Ищем всплывшее диалоговое окно
                for child in self.winfo_toplevel().winfo_children():
                    if isinstance(child, tk.Toplevel) and child.title() == "Выберите дату":
                        
                        # =========================================================
                        # КООРДИНАТЫ ПЕРЕДАЮТСЯ ИЗ ФАЙЛА ГДЕ ВЫЗЫВАЕТСЯ CustomDateEntry
                        # =========================================================
                        
                        if self.popup_coords:
                            MANUAL_X, MANUAL_Y = self.popup_coords
                        else:
                            # Значения по умолчанию, если окно не передало свои координаты
                            MANUAL_X = 500  
                            MANUAL_Y = 300  
                        
                        # Применяем новую позицию
                        child.geometry(f"+{MANUAL_X}+{MANUAL_Y}")
                        break
            except Exception:
                pass

        # Даем диалогу 50 мс на рендер, а потом перебиваем его координаты
        self.after(50, move_calendar)
        
        picker = DatePickerDialog(
            parent=self.winfo_toplevel(),
            title="Выберите дату",
            firstweekday=self.firstweekday,
            bootstyle=self.bootstyle,
            startdate=start_date
        )
        
        if picker.date_selected:
            selected_date = picker.date_selected
            if isinstance(selected_date, (datetime,)):
                date_str = selected_date.strftime(self.dateformat)
            else:
                # В некоторых версиях возвращается объект date
                date_str = selected_date.strftime(self.dateformat)
                
            self.entry.delete(0, 'end')
            self.entry.insert(0, date_str)

    def get(self):
        return self.entry.get()

    def delete(self, *args, **kwargs):
        self.entry.delete(*args, **kwargs)

    def insert(self, *args, **kwargs):
        self.entry.insert(*args, **kwargs)

    @property
    def value(self):
        try:
            return datetime.strptime(self.entry.get(), self.dateformat)
        except:
            return datetime.now()
