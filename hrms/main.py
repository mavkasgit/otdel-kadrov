"""
HRMS System - Main Entry Point
Запускается из Excel через VBA макрос (Shell)
Одна кнопка в Excel - всё остальное в GUI
"""
import os
import sys

project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from settings import ensure_directories
from ui.views.main_menu import MainMenu


def main():
    """Запуск главного меню"""
    ensure_directories()
    
    try:
        app = MainMenu()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("Нажми Enter для выхода...")


if __name__ == "__main__":
    main()
