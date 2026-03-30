"""
HRMS Database Engine
Модуль доступа к данным Excel
"""
import xlwings as xw
import pandas as pd
from datetime import datetime
from typing import Optional, Dict, List, Any
import settings
from core.logger import logger
from core.exceptions import (
    DatabaseConnectionError,
    SheetNotFoundError,
    DataIntegrityError
)


class ExcelDatabase:
    """
    Main database interface for Excel workbook
    
    ВАЖНО: Запуск ТОЛЬКО из Excel через xlwings.Book.caller()
    Сотрудники добавляются ВРУЧНУЮ в Excel, Python только читает/редактирует
    """
    
    def __init__(self, workbook_path: Optional[str] = None):
        """
        Initialize database connection
        
        Args:
            workbook_path: Path to Excel file (optional, uses caller workbook if None)
        """
        self.workbook_path = workbook_path
        self.workbook: Optional[xw.Book] = None
        self._sheets_cache: Dict[str, xw.Sheet] = {}
        logger.debug(f"ExcelDatabase initialized with path: {workbook_path}")
    
    def connect(self) -> bool:
        """
        Establish connection to Excel workbook
        
        Returns:
            True if connection successful
            
        Raises:
            DatabaseConnectionError: If cannot connect to Excel
            SheetNotFoundError: If required sheets are missing
        """
        try:
            # Use caller workbook if launched from Excel, otherwise open file
            if self.workbook_path is None:
                try:
                    logger.info("Connecting to caller workbook (launched from Excel)")
                    self.workbook = xw.Book.caller()
                except Exception:
                    # Fallback: открываем файл напрямую (для отладки из консоли)
                    import settings
                    logger.info(f"Fallback: opening workbook directly: {settings.EXCEL_FILE}")
                    self.workbook = xw.Book(settings.EXCEL_FILE)
            else:
                logger.info(f"Opening workbook: {self.workbook_path}")
                self.workbook = xw.Book(self.workbook_path)
            
            if self.workbook is None:
                raise DatabaseConnectionError("Failed to open workbook")
            
            logger.info(f"Connected to workbook: {self.workbook.name}")
            
            # Verify required sheets exist
            self._verify_sheets()
            
            # Cache sheet references
            self._cache_sheets()
            
            logger.info("Database connection established successfully")
            return True
            
        except FileNotFoundError as e:
            error_msg = f"Excel file not found: {self.workbook_path}"
            logger.error(error_msg)
            raise DatabaseConnectionError(error_msg) from e
            
        except Exception as e:
            error_msg = f"Failed to connect to Excel: {str(e)}"
            logger.error(error_msg)
            raise DatabaseConnectionError(error_msg) from e
    
    def disconnect(self) -> None:
        """Close connection to Excel workbook"""
        if self.workbook:
            logger.info(f"Disconnecting from workbook: {self.workbook.name}")
            # Don't close the workbook if it was the caller
            if self.workbook_path is not None:
                self.workbook.close()
            self.workbook = None
            self._sheets_cache.clear()
            logger.info("Database connection closed")
    
    def _verify_sheets(self) -> None:
        """
        Verify that all required sheets exist in the workbook.
        If a required sheet is missing, it will be created automatically.
        """
        required_sheets = [
            settings.SHEET_EMPLOYEES,
            settings.SHEET_VACATIONS,
            settings.SHEET_SETTINGS
        ]
        
        try:
            existing_sheets = [sheet.name for sheet in self.workbook.sheets]
            logger.debug(f"Verifying sheets. Required: {required_sheets}. Existing: {existing_sheets}")
        except Exception as e:
            logger.error(f"Failed to get sheets list: {e}")
            return
        
        for sheet_name in required_sheets:
            try:
                if sheet_name not in existing_sheets:
                    logger.warning(f"Required sheet '{sheet_name}' not found. Creating it now.")
                    self.workbook.sheets.add(sheet_name)
                
                sheet = self.workbook.sheets[sheet_name]
                if not sheet.range("A1").value:
                    if sheet_name == settings.SHEET_VACATIONS:
                        sheet.range("A1").value = settings.VACATION_COLUMNS
                        logger.info(f"Initialized headers for {sheet_name}")
            except Exception as e:
                logger.error(f"Error processing sheet '{sheet_name}': {e}")

        logger.info("All required sheets verified or created.")
    
    def _cache_sheets(self) -> None:
        """Cache references to frequently used sheets"""
        try:
            self._sheets_cache[settings.SHEET_EMPLOYEES] = self.workbook.sheets[settings.SHEET_EMPLOYEES]
            self._sheets_cache[settings.SHEET_VACATIONS] = self.workbook.sheets[settings.SHEET_VACATIONS]
            self._sheets_cache[settings.SHEET_SETTINGS] = self.workbook.sheets[settings.SHEET_SETTINGS]
            logger.debug("Sheet references cached")
        except Exception as e:
            logger.warning(f"Failed to cache some sheets: {e}")
    
    def _get_sheet(self, sheet_name: str) -> xw.Sheet:
        """
        Get sheet by name from cache or workbook
        
        Args:
            sheet_name: Name of the sheet
            
        Returns:
            Sheet object
            
        Raises:
            SheetNotFoundError: If sheet doesn't exist
        """
        if sheet_name in self._sheets_cache:
            return self._sheets_cache[sheet_name]
        
        try:
            sheet = self.workbook.sheets[sheet_name]
            self._sheets_cache[sheet_name] = sheet
            return sheet
        except Exception as e:
            error_msg = f"Sheet '{sheet_name}' not found"
            logger.error(error_msg)
            raise SheetNotFoundError(error_msg) from e
    
    def _convert_dates(self, df: pd.DataFrame, date_columns: List[str]) -> pd.DataFrame:
        """
        Convert Excel date columns to Python datetime
        """
        for col in date_columns:
            if col in df.columns:
                try:
                    df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
                except Exception as e:
                    logger.warning(f"Failed to convert column '{col}' to datetime: {e}")
        return df

    def _clean_dataframe(self, df: pd.DataFrame, 
                         date_cols: Optional[List[str]] = None, 
                         numeric_cols: Optional[List[str]] = None) -> pd.DataFrame:
        """
        Centralized cleaning for all DataFrames coming from Excel.
        1. Convert dates.
        2. Convert numeric IDs/counts to nullable integers (strips .0).
        3. Replace NaN/NaT with None for safe UI display.
        """
        # 1. Dates
        if date_cols:
            df = self._convert_dates(df, date_cols)
            
        # 2. Numeric columns (Convert whole floats to int-strings, NaN to None)
        if numeric_cols:
            for col in numeric_cols:
                if col in df.columns:
                    try:
                        # For consistency and to avoid .0, we convert whole numbers to clean strings or None
                        def clean_id(v):
                            if pd.isna(v): return None
                            try:
                                f = float(v)
                                return str(int(f)) if f.is_integer() else str(v)
                            except:
                                return str(v)
                        
                        df[col] = df[col].apply(clean_id)
                    except Exception as e:
                        logger.warning(f"Failed to convert numeric column '{col}': {e}")
        
        # 3. Handle remaining empty cells (this must be last)
        df = df.where(pd.notna(df), None)
        return df
    
    def refresh_data(self) -> None:
        """Refresh data from Excel (clear cache if implemented)"""
        logger.info("Refreshing data from Excel")
        # For now, just log. In future, can implement caching and refresh logic
        pass
    
    def __enter__(self):
        """Context manager entry"""
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self.disconnect()
        return False

    def get_employees(self, filter_department: Optional[str] = None) -> pd.DataFrame:
        """
        Get all employees with optional department filter
        
        ОПТИМИЗАЦИЯ: Читает ВЕСЬ диапазон за раз до последней заполненной строки
        Использует sheet.used_range для автоматического определения диапазона
        
        Args:
            filter_department: Optional filter ("Завод КТМ" or "Основное")
            
        Returns:
            DataFrame with employee data
            
        Raises:
            DatabaseConnectionError: If not connected to workbook
            DataIntegrityError: If sheet structure is invalid
        """
        if not self.workbook:
            raise DatabaseConnectionError("Not connected to workbook. Call connect() first.")
        
        try:
            sheet = self._get_sheet(settings.SHEET_EMPLOYEES)
            logger.debug(f"Reading employees from sheet: {settings.SHEET_EMPLOYEES}")
            
            # Read entire used range (2D list)
            data = sheet.used_range.options(ndim=2).value
            
            if not data or len(data) < 2:
                logger.warning("Employee sheet is empty or has no data rows")
                return pd.DataFrame(columns=settings.EMPLOYEE_COLUMNS)
            
            # First row is headers
            headers = data[0]
            rows = data[1:]
            
            # Create DataFrame
            df = pd.DataFrame(rows, columns=headers)
            
            # Verify expected columns exist
            missing_cols = set(settings.EMPLOYEE_COLUMNS) - set(df.columns)
            if missing_cols:
                error_msg = f"Missing required columns in {settings.SHEET_EMPLOYEES}: {missing_cols}"
                logger.error(error_msg)
                raise DataIntegrityError(error_msg)
            
            # Centralized cleaning
            date_columns = ["Дата рождения", "Дата принятия", "Начало контракта", "Конец контракта"]
            numeric_columns = ["Таб. №", "Пенсионер"] # Tab No is the most important one
            df = self._clean_dataframe(df, date_cols=date_columns, numeric_cols=numeric_columns)
            
            # Apply department filter if specified
            if filter_department and filter_department in settings.VALID_DEPARTMENTS:
                df = df[df["Подразделение"] == filter_department]
                logger.debug(f"Filtered by department: {filter_department}, rows: {len(df)}")
            
            logger.info(f"Retrieved {len(df)} employees")
            return df
            
        except (DatabaseConnectionError, DataIntegrityError):
            raise
        except Exception as e:
            error_msg = f"Error reading employees: {str(e)}"
            logger.error(error_msg)
            raise DatabaseConnectionError(error_msg) from e
    
    def get_employee_by_tab_number(self, tab_number: int) -> Optional[Dict]:
        """
        Get employee data by tab number
        
        Args:
            tab_number: Employee tab number (Таб. №)
            
        Returns:
            Dictionary with employee data or None if not found
        """
        try:
            df = self.get_employees()
            
            # Robust numeric comparison for IDs
            tab_val = pd.to_numeric(tab_number, errors='coerce')
            df_tabs = pd.to_numeric(df["Таб. №"], errors='coerce')
            
            # Debug log
            logger.debug(f"Searching for Tab №: {tab_number} (numeric: {tab_val}). Available count: {len(df)}")
            
            # Filter
            employee_df = df[df_tabs == tab_val]
            
            if employee_df.empty:
                logger.warning(f"Employee with tab number {tab_number} not found after numeric comparison.")
                return None
            
            # Convert to dict
            employee = employee_df.iloc[0].to_dict()
            logger.debug(f"Retrieved employee: {employee.get('ФИО', 'Unknown')}")
            return employee
            
        except Exception as e:
            logger.exception(f"Error getting employee by tab number {tab_number}")
            raise

    def get_references(self) -> Dict[str, List[str]]:
        """
        Get reference lists from Settings sheet
        
        Лист "Настройки" содержит колонки: Должность, События, Подразделение, Форма оплаты, Ставка
        
        Returns:
            Dictionary with reference lists for dropdowns
        """
        try:
            sheet = self._get_sheet(settings.SHEET_SETTINGS)
            logger.debug(f"Reading references from sheet: {settings.SHEET_SETTINGS}")
            
            # Читаем данные как 2D список
            data = sheet.used_range.options(ndim=2).value
            
            if not data or len(data) < 2:
                logger.warning("Settings sheet is empty")
                return {}
            
            # First row is headers
            headers = data[0]
            rows = data[1:]
            
            # Create DataFrame
            df = pd.DataFrame(rows, columns=headers)
            
            # Build reference dictionary
            references = {}
            for col in settings.REFERENCE_COLUMNS:
                if col in df.columns:
                    # Get non-null unique values
                    values = df[col].dropna().unique().tolist()
                    references[col] = [str(v) for v in values if v]
                    logger.debug(f"Reference '{col}': {len(references[col])} values")
            
            logger.info(f"Retrieved {len(references)} reference lists")
            return references
            
        except Exception as e:
            logger.error(f"Error reading references: {e}")
            raise DatabaseConnectionError(f"Error reading references: {str(e)}") from e
    
    def get_vacations(self, tab_number: Optional[int] = None) -> pd.DataFrame:
        """
        Get vacation records, optionally filtered by employee
        
        Args:
            tab_number: Optional employee tab number to filter by
            
        Returns:
            DataFrame with vacation data
        """
        try:
            sheet = self._get_sheet(settings.SHEET_VACATIONS)
            logger.debug(f"Reading vacations from sheet: {settings.SHEET_VACATIONS}")
            
            # Read used range
            # Read used range as 2D list
            data = sheet.used_range.options(ndim=2).value
            
            if not data or len(data) < 2:
                logger.warning("Vacations sheet is empty")
                return pd.DataFrame(columns=settings.VACATION_COLUMNS)
            
            # First row is headers
            headers = data[0]
            rows = data[1:]
            
            # Create DataFrame
            try:
                # Fallback: если в первой строке пусто, используем стандартные заголовки
                if not any(headers) or pd.isna(headers[0]):
                    logger.warning("No headers found in Vacations sheet, using default column names")
                    headers = settings.VACATION_COLUMNS
                    df = pd.DataFrame(rows, columns=headers)
                else:
                    df = pd.DataFrame(rows, columns=headers)
            except Exception as e:
                logger.warning(f"DataFrame creation error for vacations: {e}")
                return pd.DataFrame(columns=settings.VACATION_COLUMNS)
            
            # Centralized cleaning
            date_columns = ["Дата начала", "Дата окончания"]
            numeric_columns = ["Таб. №", "Количество дней"]
            df = self._clean_dataframe(df, date_cols=date_columns, numeric_cols=numeric_columns)
            
            # Filter by tab number if specified
            if tab_number is not None:
                # Robust numeric comparison for filtering
                tab_val = pd.to_numeric(tab_number, errors='coerce')
                df_tabs = pd.to_numeric(df["Таб. №"], errors='coerce')
                df = df[df_tabs == tab_val]
                logger.debug(f"Filtered vacations for tab number {tab_number} (numeric: {tab_val}): {len(df)} records")
            
            logger.info(f"Retrieved {len(df)} vacation records")
            return df
            
        except Exception as e:
            logger.error(f"Error reading vacations: {e}")
            raise DatabaseConnectionError(f"Error reading vacations: {str(e)}") from e
    
    def add_vacation(self, vacation_data: Dict) -> int:
        """
        Add new vacation record
        
        Args:
            vacation_data: Dictionary with vacation data
                Required keys: Таб. №, Дата начала, Дата окончания, Тип отпуска
                
        Returns:
            ID of created vacation record
        """
        try:
            sheet = self._get_sheet(settings.SHEET_VACATIONS)
            
            # Get current data to find next ID
            df = self.get_vacations()
            
            # Generate new ID
            if df.empty or "ID записи" not in df.columns:
                new_id = 1
            else:
                new_id = int(df["ID записи"].max()) + 1
            
            # Calculate duration
            start_date = vacation_data["Дата начала"]
            end_date = vacation_data["Дата окончания"]
            duration = (end_date - start_date).days + 1
            
            # Get employee name using unified search (FIO primary, tab number backup)
            search_key = vacation_data.get("search_value") or vacation_data.get("ФИО") or vacation_data.get("Таб. №")
            employee = self.find_employee(str(search_key)) if search_key else None
            employee_name = employee["ФИО"] if employee else "Unknown"
            
            # Prepare row data
            new_row = [
                new_id,
                vacation_data["Таб. №"],
                employee_name,
                start_date,
                end_date,
                vacation_data["Тип отпуска"],
                duration
            ]
            
            # Find next empty row
            last_row = sheet.used_range.last_cell.row
            next_row = last_row + 1
            
            # Write data
            sheet.range(f"A{next_row}").value = new_row
            
            logger.info(f"Added vacation record ID {new_id} for {employee_name}")
            return new_id
            
        except Exception as e:
            logger.error(f"Error adding vacation: {e}")
            raise DatabaseConnectionError(f"Error adding vacation: {str(e)}") from e
    
    def update_vacation(self, vacation_id: int, vacation_data: Dict) -> bool:
        """
        Update existing vacation record
        
        Args:
            vacation_id: ID of vacation to update
            vacation_data: Dictionary with updated vacation data
            
        Returns:
            True if successful
        """
        try:
            sheet = self._get_sheet(settings.SHEET_VACATIONS)
            df = self.get_vacations()
            
            # Find vacation row
            vacation_row = df[df["ID записи"] == vacation_id]
            if vacation_row.empty:
                logger.warning(f"Vacation ID {vacation_id} not found")
                return False
            
            # Get row index in sheet (add 2: 1 for header, 1 for 0-based to 1-based)
            row_idx = vacation_row.index[0] + 2
            
            # Calculate duration
            start_date = vacation_data.get("Дата начала", vacation_row.iloc[0]["Дата начала"])
            end_date = vacation_data.get("Дата окончания", vacation_row.iloc[0]["Дата окончания"])
            duration = (end_date - start_date).days + 1
            
            # Update fields
            if "Дата начала" in vacation_data:
                sheet.range(f"D{row_idx}").value = vacation_data["Дата начала"]
            if "Дата окончания" in vacation_data:
                sheet.range(f"E{row_idx}").value = vacation_data["Дата окончания"]
            if "Тип отпуска" in vacation_data:
                sheet.range(f"F{row_idx}").value = vacation_data["Тип отпуска"]
            
            # Update duration
            sheet.range(f"G{row_idx}").value = duration
            
            logger.info(f"Updated vacation ID {vacation_id}")
            return True
            
        except Exception as e:
            logger.error(f"Error updating vacation: {e}")
            raise DatabaseConnectionError(f"Error updating vacation: {str(e)}") from e
    
    def delete_vacation(self, vacation_id: int) -> bool:
        """
        Delete vacation record
        
        Args:
            vacation_id: ID of vacation to delete
            
        Returns:
            True if successful
        """
        try:
            sheet = self._get_sheet(settings.SHEET_VACATIONS)
            df = self.get_vacations()
            
            # Find vacation row
            vacation_row = df[df["ID записи"] == vacation_id]
            if vacation_row.empty:
                logger.warning(f"Vacation ID {vacation_id} not found")
                return False
            
            # Get row index in sheet
            row_idx = vacation_row.index[0] + 2
            
            # Delete row
            sheet.range(f"{row_idx}:{row_idx}").api.Delete()
            
            logger.info(f"Deleted vacation ID {vacation_id}")
            return True
            
        except Exception as e:
            logger.error(f"Error deleting vacation: {e}")
            raise DatabaseConnectionError(f"Error deleting vacation: {str(e)}") from e

    def _get_order_sheet_for_year(self, year: int = None) -> Any:
        """
        Получить лист журнала приказов для указанного года.
        Если лист не существует - создать.
        
        Args:
            year: Год (если None, берется текущий)
            
        Returns:
            Объект листа Excel
        """
        from datetime import datetime
        from core.sheet_utils import create_order_log_sheet
        
        if year is None:
            year = datetime.now().year
        
        sheet_name = settings.get_order_sheet_name(year)
        existing_sheets = [s.name for s in self.workbook.sheets]
        
        if sheet_name not in existing_sheets:
            try:
                logger.info(f"Creating new order log sheet for year {year}: {sheet_name}")
                create_order_log_sheet(self.workbook, year)
                sheet = self.workbook.sheets[sheet_name]
            except Exception as e:
                logger.warning(f"Failed to create new sheet: {e}")
                sheet = self.workbook.sheets[0]
        else:
            sheet = self.workbook.sheets[sheet_name]
        
        return sheet
    
    def _get_order_log(self, year: int = None, order_type: Optional[str] = None) -> pd.DataFrame:
        """
        Чтение журнала приказов для указанного года
        
        Args:
            year: Год (если None, берется текущий)
            order_type: Опциональный фильтр по типу события
            
        Returns:
            DataFrame с записями журнала
        """
        from datetime import datetime
        
        if year is None:
            year = datetime.now().year
        
        sheet_name = settings.get_order_sheet_name(year)
        existing_sheets = [s.name for s in self.workbook.sheets]
        
        if sheet_name not in existing_sheets:
            logger.info(f"Order log sheet for {year} does not exist yet, returning empty")
            return pd.DataFrame(columns=settings.ORDER_LOG_COLUMNS)
        
        try:
            sheet = self._get_sheet(sheet_name)
            logger.debug(f"Reading order log from sheet: {sheet_name}")
            
            used_range = sheet.used_range
            data = used_range.value
            
            if not data:
                logger.warning("Order log sheet is empty")
                return pd.DataFrame(columns=settings.ORDER_LOG_COLUMNS)
            
            if isinstance(data, list) and len(data) > 0:
                if isinstance(data[0], str):
                    data = [data]
                
                if len(data) < 2:
                    logger.warning("Order log sheet has no data rows")
                    return pd.DataFrame(columns=settings.ORDER_LOG_COLUMNS)
                else:
                    headers = data[0]
                    rows = data[1:]
                    if len(headers) != len(rows[0]):
                        logger.warning(f"Header/row mismatch: {len(headers)} headers, {len(rows[0])} cols")
                        headers = headers[:len(rows[0])]
            else:
                return pd.DataFrame(columns=settings.ORDER_LOG_COLUMNS)
            
            try:
                # Fallback: если в первой строке пусто, используем стандартные заголовки
                if not any(headers) or pd.isna(headers[0]):
                    logger.warning("No headers found in Excel, using default column names")
                    headers = settings.ORDER_LOG_COLUMNS
                    # Если первая строка пуста, данные идут со второй
                    df = pd.DataFrame(rows, columns=headers)
                else:
                    df = pd.DataFrame(rows, columns=headers)
            except Exception as e:
                logger.warning(f"DataFrame creation error: {e}")
                return pd.DataFrame(columns=settings.ORDER_LOG_COLUMNS)
            
            # Centralized cleaning
            date_columns = ["Дата создания"]
            numeric_columns = ["Номер приказа", "Таб. №"]
            df = self._clean_dataframe(df, date_cols=date_columns, numeric_cols=numeric_columns)
            
            if order_type and "Тип события" in df.columns and len(df) > 0:
                try:
                    mask = df["Тип события"].astype(str).str.strip() == str(order_type).strip()
                    df = df[mask]
                    logger.debug(f"Filtered by order type: {order_type}, rows: {len(df)}")
                except Exception as e:
                    logger.warning(f"Filter error: {e}, returning all")
            
            logger.info(f"Retrieved {len(df)} order log entries")
            return df
            
        except Exception as e:
            logger.exception(f"Error reading order log: {e}")
            raise DatabaseConnectionError(f"Error reading order log: {str(e)}") from e
    
    def get_order_log(self, order_type: Optional[str] = None, year: Optional[int] = None) -> pd.DataFrame:
        """
        Чтение журнала приказов (публичный метод)
        
        Args:
            order_type: Опциональный фильтр по типу события
            year: Год (если None, берется текущий)
            
        Returns:
            DataFrame с записями журнала
        """
        return self._get_order_log(year=year, order_type=order_type)

    def _get_type_code(self, order_type: str) -> str:
        """
        Получение кода типа из справочника
        
        Args:
            order_type: Название типа из справочника
            
        Returns:
            Код типа (2 символа max)
            
        Raises:
            ValueError: Если тип не найден в справочнике
        """
        code = settings.ORDER_TYPE_CODES.get(order_type)
        if code is None:
            raise ValueError(f"Unknown order type: {order_type}")
        return code

    def get_next_order_number(self, order_type: str, year: Optional[int] = None) -> str:
        """
        Получение следующего номера для приказа (сквозная нумерация за год)
        
        Args:
            order_type: Тип события
            year: Год (если None, берется текущий)
            
        Returns:
            Следующее число (строка)
        """
        import re
        from datetime import datetime
        
        if year is None:
            year = datetime.now().year
        
        df = self._get_order_log(year=year)
        
        if df.empty:
            next_num = 1
        else:
            numbers = []
            if "Номер приказа" in df.columns:
                for order_num in df["Номер приказа"]:
                    if pd.notna(order_num):
                        match = re.search(r'(\d+)', str(order_num))
                        if match:
                            numbers.append(int(match.group(1)))
            
            if numbers:
                next_num = max(numbers) + 1
            else:
                next_num = 1
        
        return str(next_num)

    def find_employee(self, search_value: str) -> Optional[Dict]:
        """
        Единый поиск сотрудника по ФИО (основной) или Таб.№ (запасной)
        
        Args:
            search_value: Строка поиска (ФИО или Таб.№)
            
        Returns:
            Dict с данными сотрудника или None если не найден
        """
        try:
            df = self.get_employees()
            
            if df.empty:
                return None
            
            search_lower = search_value.lower().strip()
            matches = df[df["ФИО"].astype(str).str.lower().str.contains(search_lower, na=False)]
            
            if not matches.empty:
                if len(matches) > 1:
                    logger.warning(f"Multiple matches for FIO '{search_value}': {len(matches)} found, using first")
                return matches.iloc[0].to_dict()
            
            try:
                tab_num = int(search_value)
                matches = df[df["Таб. №"] == tab_num]
                
                if not matches.empty:
                    return matches.iloc[0].to_dict()
            except ValueError:
                pass
            
            logger.warning(f"Employee not found: {search_value}")
            return None
            
        except Exception as e:
            logger.error(f"Error searching employee '{search_value}': {e}")
            return None

    def add_order_log(self, order_data: Dict) -> str:
        """
        Добавление записи в журнал приказов
        
        Args:
            order_data: {
                'order_type': str,      # Обязательно
                'search_value': str,    # Единое поле поиска (ФИО или Таб.№)
                'file_path': str        # Опционально
            }
            
        Returns:
            Номер созданного приказа
            
        Raises:
            ValueError: Если сотрудник не найден или тип события некорректен
        """
        from datetime import date
        
        # 1. Получаем тип и номер приказа
        order_type = order_data.get("order_type")
        if not order_type:
            raise ValueError("order_type is required")
        
        self._get_type_code(order_type)
        
        # 3. Определяем дату и путь к файлу
        from datetime import date
        order_date = order_data.get("Дата создания") or order_data.get("order_date") or date.today()
        
        # Определяем год для журнала
        order_year = order_date.year if hasattr(order_date, 'year') else date.today().year
        
        order_number = order_data.get("Номер приказа") or order_data.get("order_number")
        if not order_number:
            order_number = self.get_next_order_number(order_type, year=order_year)
            
        # 2. Находим сотрудника
        search_value = order_data.get("tab_number") or order_data.get("search_value")
        if not search_value:
            raise ValueError("search_value or tab_number is required")
        
        employee = self.find_employee(str(search_value))
        if employee is None:
            raise ValueError(f"Сотрудник не найден: {search_value}")
        
        file_path = order_data.get("Путь к файлу") or order_data.get("file_path")
        
        # Получаем лист для нужного года (создастся автоматически если нет)
        sheet = self._get_order_sheet_for_year(order_year)
        
        # 4. Запись в журнал
        # Если лист пуст (кроме заголовков), пишем во вторую строку
        try:
            val_a2 = sheet.range("A2").value
            if not val_a2:
                next_row = 2
            else:
                last_row = sheet.used_range.last_cell.row
                next_row = last_row + 1
        except:
            next_row = 2
        
        new_row = [
            order_number,
            order_type,
            order_date,
            employee.get("ФИО"),
            employee.get("Таб. №"),
            file_path
        ]
        
        # Записываем все поля кроме пути к файлу (колонки B-G)
        sheet.range(f"B{next_row}").value = new_row[:-1]
        
        # Последнее поле (Путь к файлу) пишем как гиперссылку (колонка G)
        file_cell = sheet.range(f"G{next_row}")
        if file_path:
            file_cell.add_hyperlink(file_path, "Открыть файл", "Нажмите для открытия приказа")
        else:
            file_cell.value = ""
        
        # Применяем стилизацию таблицы
        from core.sheet_utils import format_order_sheet
        format_order_sheet(sheet)
        
        logger.info(f"Added order log entry to row {next_row}")
        
        logger.info(f"Added order log entry: {order_number} for {employee.get('ФИО')} in sheet {order_year}")
        return order_number
