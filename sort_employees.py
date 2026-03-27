import xlwings as xw
from tkinter import Tk, Label, Button, ttk, StringVar, messagebox
import traceback

def show_employee_selector():
    '''Показывает окно с выпадающим списком сотрудников'''
    
    try:
        # Пытаемся получить книгу через caller (из VBA)
        try:
            wb = xw.Book.caller()
        except:
            # Если не получилось, открываем файл напрямую
            wb = xw.Book('otdel-kadrov.xlsm')
        
        # Пытаемся найти лист "Сотрудники"
        employee_sheet = None
        
        for sheet in wb.sheets:
            if 'сотрудник' in sheet.name.lower():
                employee_sheet = sheet
                break
        
        if not employee_sheet:
            employee_sheet = wb.sheets[0]
        
        # Читаем список сотрудников из колонки C
        employees_range = employee_sheet.range('C2:C100').value
        employees = [emp for emp in employees_range if emp is not None and str(emp).strip()]
        
        if not employees:
            messagebox.showwarning("Предупреждение", "Не найдено сотрудников в колонке C!")
            return
        
        # Создаем GUI окно
        root = Tk()
        root.title("Выбор сотрудника")
        root.geometry("400x200")
        
        selected_employee = StringVar()
        
        Label(root, text="Выберите сотрудника:", font=("Arial", 12)).pack(pady=20)
        
        combo = ttk.Combobox(root, textvariable=selected_employee, 
                             values=employees, state="readonly", 
                             font=("Arial", 11), width=30)
        combo.pack(pady=10)
        
        if employees:
            combo.current(0)
        
        def on_sort():
            selected = selected_employee.get()
            if selected:
                sort_by_employee(employee_sheet, selected)
                root.destroy()
        
        Button(root, text="Сортировать", command=on_sort, 
               font=("Arial", 11), bg="#4CAF50", fg="white", 
               padx=20, pady=10).pack(pady=20)
        
        root.mainloop()
    
    except Exception as e:
        error_msg = f"Ошибка: {str(e)}\n\n{traceback.format_exc()}"
        print(error_msg)
        try:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n\n{str(e)}")
        except:
            pass
        input("Нажмите Enter для закрытия...")

def sort_by_employee(sheet, employee_name):
    '''Сортирует данные по выбранному сотруднику'''
    
    try:
        last_row = sheet.range('C' + str(sheet.cells.last_cell.row)).end('up').row
        
        data_range = sheet.range(f'A1:Z{last_row}')
        data = data_range.value
        
        header = data[0]
        rows = data[1:]
        
        def sort_key(row):
            name = str(row[2]) if row[2] else ""
            if name == employee_name:
                return (0, name)
            return (1, name)
        
        sorted_rows = sorted(rows, key=sort_key)
        
        sorted_data = [header] + sorted_rows
        data_range.value = sorted_data
        
        sheet.range('A2:Z2').color = (255, 255, 0)
        
        messagebox.showinfo("Успех", f"Сотрудник '{employee_name}' отсортирован!")
        
    except Exception as e:
        error_msg = f"Ошибка при сортировке: {str(e)}\n\n{traceback.format_exc()}"
        print(error_msg)
        messagebox.showerror("Ошибка сортировки", str(e))

if __name__ == '__main__':
    try:
        show_employee_selector()
    except Exception as e:
        print(f"КРИТИЧЕСКАЯ ОШИБКА: {e}")
        print(traceback.format_exc())
        input("Нажмите Enter для закрытия...")
