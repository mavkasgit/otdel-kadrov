"""
Excel Sheet Utilities
Утилиты для создания и форматирования листов Excel
"""
import sys
import os

# Добавляем корень проекта в путь для импортов
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import xlwings as xw
from core.logger import logger


def create_order_log_sheet(wb: xw.Book, year: int, data: list = None) -> xw.Sheet:
    """
    Создать лист журнала приказов для указанного года
    
    Args:
        wb: Workbook object
        year: Год для журнала
        data: Опциональные данные для заполнения
        
    Returns:
        Объект созданного листа
    """
    from settings import get_order_sheet_name, ORDER_LOG_COLUMNS, ORDER_SHEET_COLUMNS
    
    sheet_name = get_order_sheet_name(year)
    
    # Удаляем старый лист если есть
    for s in wb.sheets:
        if s.name == sheet_name:
            s.delete()
            break
    
    # Создаем новый лист
    sheet = wb.sheets.add(sheet_name)
    
    # Заголовки в колонке B (A - пустая)
    sheet.range("B1").value = ORDER_LOG_COLUMNS
    
    # Заполняем данные если есть
    if data:
        for i, row in enumerate(data, start=2):
            sheet.range(f"B{i}").value = row
    
    # Применяем форматирование
    format_order_sheet(sheet)
    
    logger.info(f"Created order log sheet: {sheet_name}")
    return sheet


def format_order_sheet(sheet: xw.Sheet):
    """
    Применить форматирование к листу журнала приказов
    
    Args:
        sheet: Лист для форматирования
    """
    try:
        from settings import ORDER_SHEET_COLUMNS
        
        last_row = sheet.used_range.last_cell.row
        if last_row < 1:
            return
        
        # Ширина колонок (A пустая)
        widths = ORDER_SHEET_COLUMNS
        col_letters = ["A", "B", "C", "D", "E", "F", "G"]
        for letter, width in zip(col_letters, widths):
            sheet.range(f"{letter}:{letter}").column_width = width
        
        # Заголовок - перенос и высота
        sheet.range("B1:G1").api.WrapText = True
        sheet.range("1:1").row_height = 30
        
        # Excel Table
        table_name = "ТаблицаПриказы"
        try:
            sheet.api.ListObjects(table_name).Delete()
        except:
            pass
        
        if last_row >= 2:
            table = sheet.api.ListObjects.Add(
                Source=sheet.range(f"B1:G{last_row}").api,
                XlListObjectHasHeaders=True
            )
            table.Name = table_name
            table.TableStyle = "TableStyleMedium21"
        
        # Выравнивание по центру
        if last_row >= 2:
            sheet.range(f"B1:G{last_row}").api.HorizontalAlignment = -4108
            sheet.range(f"B1:G{last_row}").api.VerticalAlignment = -4108
        
        # Freeze Panes - только верхняя строка (выбираем A2, не B2)
        try:
            sheet.activate()
            sheet.range("A2").select()
            sheet.api.Parent.Parent.Windows(1).FreezePanes = True
        except Exception as e:
            logger.warning(f"Could not freeze panes: {e}")
        
        logger.debug(f"Formatted order sheet: {sheet.name}")
        
    except Exception as e:
        logger.warning(f"Could not format sheet: {e}")


if __name__ == "__main__":
    print("Starting...")
    
    wb = xw.Book("Отдел Кадров.xlsm")
    print("Workbook opened")
    
    # Создать тестовый лист
    create_order_log_sheet(wb, 2026, [
        ["1", "Прием", "15.01.2026", "Иванов", "101", ""],
        ["2", "Отпуск", "20.01.2026", "Петров", "102", ""]
    ])
    print("Sheet created and formatted")
    
    wb.save()
    print("Saved")
    print("DONE!")
