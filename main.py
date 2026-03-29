"""
HRMS System - Main Entry Point
Запускается из Excel через xlwings VBA макрос
"""
import os
import sys
import traceback

project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from core.db_engine import ExcelDatabase
from core.analytics import AnalyticsEngine
from core.logger import logger
from settings import ensure_directories


def run():
    """Основная функция, вызываемая из Excel."""
    import xlwings as xw
    
    ensure_directories()
    
    try:
        logger.info("HRMS started from Excel")
        
        # Получаем книгу из которой вызвали
        wb = xw.Book.caller()
        sheet = wb.sheets[0]
        
        # Подключаемся к базе
        db = ExcelDatabase()
        db.connect()
        
        logger.info(f"Connected to: {db.workbook.name}")
        
        # Тест: читаем сотрудников
        employees = db.get_employees()
        
        # Тест: читаем справочники
        references = db.get_references()
        
        # Тест: автонумерация приказов
        try:
            next_order = db.get_next_order_number("Прием на работу")
        except:
            next_order = "001-П (ошибка)"
        
        # Выводим результат в Excel
        sheet.range('A1').value = "HRMS - Тест"
        sheet.range('A2').value = f"Сотрудников: {len(employees)}"
        sheet.range('A3').value = f"Справочники: {list(references.keys())}"
        sheet.range('A4').value = f"Следующий приказ: {next_order}"
        sheet.range('A5').value = "✅ РАБОТАЕТ"
        
        sheet.autofit(axis="columns")
        
        db.disconnect()
        logger.info("HRMS finished successfully")
        
    except Exception as e:
        error_msg = traceback.format_exc()
        logger.error(f"Error: {e}")
        
        try:
            sheet.range('A10').value = f"ОШИБКА: {str(e)}"
        except:
            pass
        
        raise


if __name__ == "__main__":
    # Прямой запуск для теста (нужен открытый Excel)
    run()
