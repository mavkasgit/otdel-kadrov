import xlwings as xw
import os
import sys
import tkinter as tk
from tkinter import messagebox
import traceback

# Добавляем путь к папке проекта, чтобы Python находил папку 'core'
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# Импортируем наш основной модуль после настройки пути
from core.db_engine import ExcelDatabase
from core.logger import logger

def main():
    """Основная функция, вызываемая из Excel."""
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    
    try:
        logger.info("Python script started from Excel.")
        
        # Очищаем старые данные для наглядности
        sheet.range('A1:A10').clear_contents()
        sheet.autofit(axis="columns") # Расширяем колонки

        sheet.range('A1').value = "Статус: Подключено"
        
        db = ExcelDatabase()
        db.connect()
        logger.info(f"Connected to workbook: {db.workbook.name}")
        
        sheet.range('A2').value = f"База: {db.workbook.name}"
        
        df = db.get_employees()
        sheet.range('A3').value = f"Найдено сотрудников: {len(df)}"
        
        db.disconnect()
        
        sheet.range('A5').value = "✅ УСПЕХ"
        logger.info("Script finished successfully.")
        messagebox.showinfo("Готово", "Скрипт успешно выполнен!")

    except Exception as e:
        # Если что-то пошло не так, логируем и выводим ошибку
        logger.error(f"An error occurred: {e}")
        error_details = traceback.format_exc()
        logger.error(error_details)
        
        # Пытаемся вывести ошибку в Excel
        try:
            sheet.range('A5').value = "❌ ОШИБКА"
            sheet.range('A6').value = str(error_details)
            sheet.autofit(axis="columns")
            messagebox.showerror("Ошибка выполнения", str(error_details))
        except Exception as inner_e:
            logger.error(f"Could not report error to Excel: {inner_e}")
            # Последний шанс - записать в файл
            with open(os.path.join(project_root, "critical_error.txt"), "w", encoding="utf-8") as f:
                f.write(error_details)

if __name__ == "__main__":
    # Эта часть позволяет запускать скрипт напрямую из Python для теста,
    # но xlwings не будет работать без запущенного Excel.
    # Для отладки используйте 'Run Python file in Terminal' в VS Code
    # и убедитесь, что Excel с файлом otdel_kadrov.xlsm открыт.
    main()
