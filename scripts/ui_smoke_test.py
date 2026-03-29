import os
import sys
import time
import logging
import traceback
from unittest.mock import MagicMock, patch
import pandas as pd

# Add project root to sys.path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(project_root, 'smoke_test.log'), encoding='utf-8', mode='w'),
        logging.StreamHandler()
    ]
)

# Mock xlwings
mock_xlwings = MagicMock()
sys.modules['xlwings'] = mock_xlwings

# Mock pywintypes if it's imported somewhere to avoid issues
sys.modules['pywintypes'] = MagicMock()

def mock_excel_database(*args, **kwargs):
    db_mock = MagicMock()
    # Provide an empty DataFrame or dummy DataFrame for methods returning DataFrames
    dummy_df = pd.DataFrame(columns=["Tab. #", "FIO", "Dolzhnost", "Podrazdelenie", "Data priema", "Zarplata"])
    dummy_vacations_df = pd.DataFrame(columns=["Tab. #", "Data nachala", "Data okonchaniya", "Kolichestvo dnej", "Tip otpuska"])
    dummy_orders_df = pd.DataFrame(columns=["Nomer", "Data", "Tip", "Opisanie", "Tab. #"])

    db_mock.get_employees.return_value = dummy_df
    db_mock.get_employee_by_tab.return_value = pd.Series({"Tab. #": 1, "FIO": "Testov Test"})
    db_mock.get_vacations.return_value = dummy_vacations_df
    db_mock.get_orders.return_value = dummy_orders_df
    db_mock.get_units.return_value = ["ОТДЕЛ 1", "ОТДЕЛ 2"]
    return db_mock

def test_view(module_name, class_name=None, func_name=None):
    target_name = class_name or func_name
    logging.info(f"--- Testing {module_name}.{target_name} ---")
    root = None
    try:
        # Import the module dynamically
        import importlib
        module = importlib.import_module(f"hrms.ui.views.{module_name}")
        
        # Patch ExcelDatabase globally for this module if it imports it
        with patch('hrms.ui.views.{}.ExcelDatabase'.format(module_name), mock_excel_database, create=True):
            import ttkbootstrap as ttk
            
            # The view might start Mainloop, so we need to inject a close mechanism
            def auto_close():
                logging.info(f"Auto-closing {target_name} after 10 seconds...")
                try:
                    if 'root' in locals() and root:
                        root.quit()
                except Exception as e:
                    logging.info(f"Error closing: {e}")

            # Mock Tkinter Tk/Toplevel so we can close them if func_name is used
            import tkinter as tk
            original_mainloop = tk.Tk.mainloop
            def mock_mainloop(self):
                self.after(10000, self.quit)
                original_mainloop(self)
            
            root = ttk.Window(themename="yeti")
            root.withdraw() # hide the root
            
            logging.info(f"Executing {target_name}...")
            
            root.after(10000, auto_close)
            
            try:
                with patch('tkinter.Tk.mainloop', mock_mainloop):
                    if class_name:
                        # For classes that might block in __init__ we should make sure they don't block
                        ViewClass = getattr(module, class_name)
                        try:
                            view = ViewClass(parent=root)
                        except TypeError:
                            # Might be MainMenu which takes no args
                            view = ViewClass()
                        
                        # Wait for either 10 seconds or until window closed
                        root.mainloop()
                        
                        # Cleanup
                        if hasattr(view, 'root'):
                            view.root.destroy()
                        elif hasattr(view, 'dialog'):
                            view.dialog.destroy()
                    else:
                        target_func = getattr(module, func_name)
                        target_func()
                
                logging.info(f"SUCCESS: {target_name} executed successfully.")
                return True
            except Exception as e:
                logging.error(f"RUNTIME ERROR in {target_name}: {e}")
                logging.error(traceback.format_exc())
                if root:
                    root.destroy()
                return False
                
    except Exception as e:
        logging.error(f"LOAD/INIT ERROR in {module_name}.{class_name}: {e}")
        logging.error(traceback.format_exc())
        if root:
            try:
                root.destroy()
            except:
                pass
        return False


def run_smoke_tests():
    logging.info("Starting UI Smoke Tests...")
    
    views_to_test = [
        ("dashboard", None, "show_dashboard"),
        ("employee_card", "EmployeeCardDialog", None),
        ("main_menu", "MainMenu", None),
        ("order_generator", "OrderGeneratorDialog", None),
        ("sort_employees", None, "show_employee_selector"),
        ("vacation_mgr", "VacationManagerDialog", None)
    ]
    
    results = {}
    
    for mod, cls, func in views_to_test:
        succ = test_view(mod, class_name=cls, func_name=func)
        name = f"{mod}.{cls or func}"
        results[name] = succ
        # Sleep a bit between tests
        time.sleep(1)
        
    logging.info("=== SUMMARY ===")
    success_count = 0
    for name, succ in results.items():
        status = "PASSED" if succ else "FAILED"
        if succ:
            success_count += 1
        logging.info(f"{name}: {status}")
    
    logging.info(f"Total: {len(results)}, Passed: {success_count}, Failed: {len(results) - success_count}")

if __name__ == "__main__":
    run_smoke_tests()
