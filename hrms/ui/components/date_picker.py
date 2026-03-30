"""
HRMS - Стабильный локализованный выбор даты
Компонент на базе Frame + Entry + DatePickerDialog для предотвращения TclError
"""
import locale
from datetime import datetime
from typing import Any
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

# --- ОПТИМИЗАЦИЯ СКОРОСТИ КАЛЕНДАРЯ (ФИНАЛЬНАЯ ВЕРСИЯ) ---

def _optimized_draw_calendar(self):
    """Оптимизированная отрисовка: два пула виджетов (кнопки и текст) без лишних удалений"""
    self._update_widget_bootstyle()
    self._set_title()
    self._current_month_days()
    
    # Создаем контейнер и пулы только один раз
    if not hasattr(self, 'frm_dates') or not self.frm_dates.winfo_exists():
        self.frm_dates = ttk.Frame(self.frm_calendar)
        self.frm_dates.pack(fill=BOTH, expand=YES)
        self._pool_btns = [] # Пул для активных дней (Radiobuttons)
        self._pool_lbls = [] # Пул для неактивных дней (Labels)
        
        # Фиксируем сетку, чтобы она не "прыгала" по ширине
        for i in range(7):
            self.frm_dates.columnconfigure(i, weight=1, minsize=30)

    # Скрываем всё из текущей сетки (быстрее чем уничтожать/создавать)
    for b in self._pool_btns: b.grid_remove()
    for l in self._pool_lbls: l.grid_remove()

    b_idx = 0 # Индекс в пуле кнопок
    l_idx = 0 # Индекс в пуле меток

    # Синхронизируем внутреннюю переменную даты ttkbootstrap
    if all([self.date.month == self.date_selected.month,
            self.date.year == self.date_selected.year]):
        self.datevar.set(self.date_selected.day)
    else:
        self.datevar.set(0)

    for row, weekday_list in enumerate(self.monthdays):
        for col, day in enumerate(weekday_list):
            if day == 0:
                # Текст для дней соседнего месяца (без кружочков)
                text = self.monthdates[row][col].day
                if l_idx < len(self._pool_lbls):
                    lbl = self._pool_lbls[l_idx]
                    lbl.configure(text=text, foreground='gray70')
                    lbl.grid(row=row, column=col, sticky=NSEW)
                else:
                    lbl = ttk.Label(self.frm_dates, text=text, anchor=CENTER, padding=5, foreground='gray70')
                    lbl.grid(row=row, column=col, sticky=NSEW)
                    self._pool_lbls.append(lbl)
                l_idx += 1
            else:
                # Кнопка для активного дня
                text = day
                today = datetime.now().date()
                day_date = self.monthdates[row][col]

                # Определяем стиль ячейки
                if day_date == today:
                    day_style = "info-calendar"
                else:
                    day_style = f"{self.bootstyle}-calendar"

                # Замыкание для выбора
                def selected(r=row, c=col):
                    self._on_date_selected(r, c)

                if b_idx < len(self._pool_btns):
                    btn = self._pool_btns[b_idx]
                    btn.configure(text=text, bootstyle=day_style, command=selected, value=day)
                    btn.grid(row=row, column=col, sticky=NSEW)
                else:
                    btn = ttk.Radiobutton(
                        master=self.frm_dates,
                        variable=self.datevar,
                        value=day, # Значением должен быть номер дня
                        text=text,
                        bootstyle=day_style,
                        padding=5,
                        command=selected
                    )
                    btn.grid(row=row, column=col, sticky=NSEW)
                    self._pool_btns.append(btn)
                b_idx += 1

def _optimized_selection_callback(func):
    """Декоратор без удаления frm_dates"""
    def inner(self, *args):
        # Сохраняем оригинальную функцию (она может быть уже обернута)
        # Но нам нужно вызвать базовую логику изменения даты
        func(self, *args)
        self._draw_calendar()
    return inner

# Применяем патчи
DatePickerDialog._draw_calendar = _optimized_draw_calendar
DatePickerDialog._selection_callback = _optimized_selection_callback

# Переопределяем методы навигации с новым декоратором
# Мы берем базовую логику из исходного кода ttkbootstrap

@_optimized_selection_callback
def on_next_month(self) -> None:
    year, month = self._nextmonth(self.date.year, self.date.month)
    self.date = datetime(year=year, month=month, day=1).date()

@_optimized_selection_callback
def on_next_year(self, *_: Any) -> None:
    year = self.date.year + 1
    month = self.date.month
    self.date = datetime(year=year, month=month, day=1).date()

@_optimized_selection_callback
def on_prev_month(self) -> None:
    year, month = self._prevmonth(self.date.year, self.date.month)
    self.date = datetime(year=year, month=month, day=1).date()

@_optimized_selection_callback
def on_prev_year(self, *_: Any) -> None:
    year = self.date.year - 1
    month = self.date.month
    self.date = datetime(year=year, month=month, day=1).date()

@_optimized_selection_callback
def on_reset_date(self, *_: Any) -> None:
    self.date = self.startdate

DatePickerDialog.on_next_month = on_next_month
DatePickerDialog.on_next_year = on_next_year
DatePickerDialog.on_prev_month = on_prev_month
DatePickerDialog.on_prev_year = on_prev_year
DatePickerDialog.on_reset_date = on_reset_date

class CustomDateEntry(ttk.Frame):
    """
    Стабильный аналог DateEntry, использующий обычный Entry и кнопку
    вызова DatePickerDialog. Решает проблему TclError в сложной иерархии окон.
    """
    def __init__(self, master=None, width=15, popup_coords=None, on_change=None, **kwargs):
        # Очищаем kwargs от параметров, не относящихся к Frame
        self.bootstyle = kwargs.pop('bootstyle', DEFAULT)
        self.firstweekday = kwargs.pop('firstweekday', 0)
        self.dateformat = kwargs.pop('dateformat', "%d.%m.%Y")
        
        self.popup_coords = popup_coords
        self.on_change = on_change
        
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
                        if self.popup_coords:
                            MANUAL_X, MANUAL_Y = self.popup_coords
                        else:
                            MANUAL_X = 500  
                            MANUAL_Y = 300  
                        
                        child.geometry(f"+{MANUAL_X}+{MANUAL_Y}")
                        break
            except Exception:
                pass

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
                date_str = selected_date.strftime(self.dateformat)
                
            self.entry.delete(0, 'end')
            self.entry.insert(0, date_str)
            
            # Вызываем callback при изменении
            if self.on_change:
                self.on_change()

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
