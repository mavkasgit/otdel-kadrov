"""
HRMS System - Main Entry Point
Запускается из Excel через VBA макрос (Shell)
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


def run(excel_file_path: str = None):
    """
    Основная функция.
    
    Args:
        excel_file_path: Путь к Excel файлу (передаётся из VBA)
    """
    ensure_directories()
    
    try:
        logger.info("HRMS started")
        
        if excel_file_path and os.path.exists(excel_file_path):
            logger.info(f"Using Excel file: {excel_file_path}")
            db = ExcelDatabase(workbook_path=excel_file_path)
        else:
            logger.info("No file path provided, using caller workbook")
            import xlwings as xw
            wb = xw.Book.caller()
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
        except Exception as e:
            next_order = f"001-П (ошибка: {e})"
        
        # Тест: поиск сотрудника
        if not employees.empty:
            first_emp = employees.iloc[0]
            found = db.find_employee(str(first_emp.get("Таб. №", "")))
            found_by_name = db.find_employee(str(first_emp.get("ФИО", ""))[:5])
        
        logger.info(f"Test complete: {len(employees)} employees, next order: {next_order}")
        
        db.disconnect()
        
    except Exception as e:
        error_msg = traceback.format_exc()
        logger.error(f"Error: {e}")
        logger.error(error_msg)
        raise


if __name__ == "__main__":
    # Получаем путь к файлу из аргументов
    excel_file = None
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    
    run(excel_file)
