# Requirements Document

## Introduction

Данный документ описывает требования к полноценной HRMS (Human Resource Management System) на базе Python и Excel. Система предназначена для автоматизации работы отдела кадров, включая управление сотрудниками, контроль контрактов, управление отпусками и генерацию кадровых документов. Архитектура системы основана на принципе: Excel = база данных (чистые таблицы), Python = обработчик данных и бизнес-логика.

## Glossary

- **HRMS_System**: Система управления человеческими ресурсами, включающая Python-приложение и Excel-базу данных
- **Excel_Database**: Excel файл "otdel-kadrov.xlsm", содержащий структурированные данные в виде таблиц
- **Employee_Registry**: Лист "Сотрудники" в Excel_Database, содержащий главный реестр сотрудников
- **Vacation_Log**: Лист "Отпуска" в Excel_Database, содержащий журнал транзакций отпусков
- **Staff_Table**: Лист "Штатка" в Excel_Database, содержащий структуру должностей и окладов
- **Reference_Lists**: Лист "Справочники" в Excel_Database, содержащий списки для выпадающих списков
- **Contract_Alert**: Уведомление о скором окончании контракта сотрудника
- **Birthday_Widget**: Виджет отображения предстоящих дней рождения сотрудников
- **Vacation_Overlap**: Пересечение периодов отпусков двух или более сотрудников
- **Order_Document**: Приказ (прием, увольнение, отпуск), сгенерированный из шаблона Word
- **Data_Validator**: Компонент проверки корректности данных и бизнес-правил
- **GUI_Application**: Графическое приложение на базе customtkinter
- **Python_Backend**: Серверная часть системы, обрабатывающая данные и бизнес-логику
- **Document_Generator**: Компонент генерации приказов из шаблонов Word
- **Analytics_Engine**: Компонент расчета метрик (стаж, возраст, остатки отпусков)

## Requirements

### Requirement 1: Структура данных в Excel

**User Story:** Как администратор системы, я хочу иметь структурированную базу данных в Excel, чтобы данные были организованы и доступны для обработки

#### Acceptance Criteria

1. THE Excel_Database SHALL contain a sheet named "Сотрудники" with columns: ID, ФИО, Дата_рождения, Дата_приема, Конец_контракта, Подразделение, Должность, Оклад, Гражданство, Площадка
2. THE Excel_Database SHALL contain a sheet named "Отпуска" with columns: ID_записи, ID_сотрудника, ФИО, Дата_начала, Дата_окончания, Тип_отпуска, Количество_дней
3. THE Excel_Database SHALL contain a sheet named "Штатка" with columns: Подразделение, Должность, Оклад_мин, Оклад_макс, Количество_ставок
4. THE Excel_Database SHALL contain a sheet named "Справочники" with columns: Подразделения, Должности, Типы_отпусков, Гражданства
5. THE Excel_Database SHALL NOT contain merged cells in data tables
6. THE Employee_Registry SHALL use "Завод" or "Офис" as valid values for Площадка column

### Requirement 2: Интеграция Python с Excel

**User Story:** Как разработчик, я хочу иметь надежную интеграцию между Python и Excel, чтобы обрабатывать данные программно

#### Acceptance Criteria

1. THE HRMS_System SHALL use xlwings library for bidirectional communication with Excel_Database
2. WHEN Excel_Database is opened, THE HRMS_System SHALL be able to read data using xlwings.Book.caller()
3. WHEN Python_Backend modifies data, THE Excel_Database SHALL reflect changes in real-time
4. THE HRMS_System SHALL use pandas library for data transformation and analysis
5. THE HRMS_System SHALL use openpyxl library for reading .xlsx files without Excel application
6. WHEN Excel_Database is not available, THE HRMS_System SHALL display an error message with file path

### Requirement 3: Контроль окончания контрактов

**User Story:** Как HR-менеджер, я хочу получать уведомления о скором окончании контрактов, чтобы своевременно продлевать или завершать трудовые отношения

#### Acceptance Criteria

1. THE HRMS_System SHALL calculate days remaining until contract end for each employee
2. WHEN contract end date is within 30 days, THE HRMS_System SHALL generate a Contract_Alert
3. WHEN contract end date is within 7 days, THE Contract_Alert SHALL be displayed with high priority indicator
4. THE GUI_Application SHALL display a list of Contract_Alerts on the dashboard
5. THE Contract_Alert SHALL include employee name, position, department, and days remaining
6. WHERE user configures custom alert threshold, THE HRMS_System SHALL use the configured value instead of 30 days

### Requirement 4: Отслеживание дней рождения

**User Story:** Как HR-менеджер, я хочу видеть предстоящие дни рождения сотрудников, чтобы поздравлять их своевременно

#### Acceptance Criteria

1. THE Analytics_Engine SHALL calculate upcoming birthdays within next 14 days
2. THE Birthday_Widget SHALL display employee name, birth date, and age on birthday
3. THE Birthday_Widget SHALL sort employees by date (earliest first)
4. WHEN employee birthday is today, THE Birthday_Widget SHALL highlight the entry
5. THE Birthday_Widget SHALL be visible on the main dashboard of GUI_Application

### Requirement 5: Управление отпусками

**User Story:** Как HR-менеджер, я хочу управлять отпусками сотрудников с контролем пересечений, чтобы избежать конфликтов в графике

#### Acceptance Criteria

1. WHEN user creates a new vacation record, THE Data_Validator SHALL check for Vacation_Overlap with existing records
2. IF Vacation_Overlap is detected for the same employee, THEN THE HRMS_System SHALL display an error and prevent creation
3. THE HRMS_System SHALL allow vacation records with different vacation types for the same employee if dates do not overlap
4. THE Vacation_Log SHALL store vacation start date, end date, type, and automatically calculate duration in days
5. THE GUI_Application SHALL provide a vacation management interface with calendar view
6. WHEN user selects an employee, THE GUI_Application SHALL display all vacation records for that employee

### Requirement 6: Расчет метрик сотрудников

**User Story:** Как HR-менеджер, я хочу автоматически рассчитывать стаж, возраст и остатки отпусков, чтобы не делать это вручную

#### Acceptance Criteria

1. THE Analytics_Engine SHALL calculate employee age based on birth date and current date
2. THE Analytics_Engine SHALL calculate employment tenure in years and months based on hire date
3. THE Analytics_Engine SHALL calculate vacation days used in current year from Vacation_Log
4. THE Analytics_Engine SHALL calculate vacation days remaining based on annual allowance minus used days
5. WHERE employee has less than 1 year tenure, THE Analytics_Engine SHALL calculate prorated vacation allowance
6. THE GUI_Application SHALL display calculated metrics in employee card view

### Requirement 7: Валидация данных

**User Story:** Как администратор системы, я хочу автоматически проверять корректность данных, чтобы предотвратить логические ошибки

#### Acceptance Criteria

1. WHEN user enters termination date, THE Data_Validator SHALL verify it is not earlier than hire date
2. WHEN user enters vacation dates, THE Data_Validator SHALL verify start date is not later than end date
3. WHEN user enters employee data, THE Data_Validator SHALL verify all required fields are filled
4. WHEN user selects department or position, THE Data_Validator SHALL verify the value exists in Reference_Lists
5. IF validation fails, THEN THE HRMS_System SHALL display a descriptive error message and prevent data save
6. THE Data_Validator SHALL verify salary is within min-max range defined in Staff_Table for the position

### Requirement 8: Генерация приказов

**User Story:** Как HR-менеджер, я хочу автоматически генерировать приказы из шаблонов, чтобы экономить время на оформлении документов

#### Acceptance Criteria

1. THE Document_Generator SHALL load Word templates from templates/ directory
2. WHEN user requests order generation, THE Document_Generator SHALL populate template with employee data from Excel_Database
3. THE HRMS_System SHALL support order types: прием (hire), увольнение (termination), отпуск (vacation)
4. THE Document_Generator SHALL save generated documents to приказы/ directory with filename format: {тип_приказа} {ФИО} {дата}.docx (БЕЗ подчеркиваний)
5. THE Document_Generator SHALL use docxtpl library for template processing
6. WHEN template variable is missing in data, THE Document_Generator SHALL display an error with variable name

### Requirement 9: Фильтрация по площадкам

**User Story:** Как HR-менеджер, я хочу фильтровать сотрудников по площадкам (Завод/Офис), чтобы работать с нужной группой

#### Acceptance Criteria

1. THE GUI_Application SHALL provide a filter control with options: "Все", "Завод", "Офис"
2. WHEN user selects a filter, THE GUI_Application SHALL display only employees matching the selected Площадка value
3. THE GUI_Application SHALL apply the filter to all views: employee list, dashboard widgets, reports
4. WHEN filter is "Все", THE GUI_Application SHALL display all employees regardless of Площадка value
5. THE GUI_Application SHALL persist filter selection during session

### Requirement 10: Архитектура проекта

**User Story:** Как разработчик, я хочу иметь четкую модульную архитектуру, чтобы код был поддерживаемым и расширяемым

#### Acceptance Criteria

1. THE HRMS_System SHALL organize code into modules: main.py, settings.py, core/, ui/, templates/, приказы/, личные_дела/
2. THE core/ module SHALL contain: db_engine.py, analytics.py, validator.py, doc_generator.py
3. THE ui/ module SHALL contain: styles.py, widgets.py, views/dashboard.py, views/employee_card.py, views/vacation_mgr.py
4. THE settings.py SHALL store configuration: Excel file path, alert thresholds, template paths
5. THE main.py SHALL serve as entry point for GUI_Application
6. THE HRMS_System SHALL use loguru library for logging with log files stored in logs/ directory

### Requirement 11: Графический интерфейс

**User Story:** Как пользователь, я хочу иметь современный и удобный графический интерфейс, чтобы комфортно работать с системой

#### Acceptance Criteria

1. THE GUI_Application SHALL use customtkinter library for modern UI components
2. THE GUI_Application SHALL display a dashboard with widgets: Contract_Alerts, Birthday_Widget, quick statistics
3. THE GUI_Application SHALL provide employee card view with all employee details and calculated metrics
4. THE GUI_Application SHALL provide vacation management view with calendar and list of vacations
5. THE GUI_Application SHALL use consistent color scheme and typography defined in ui/styles.py
6. THE GUI_Application SHALL be responsive and handle window resizing gracefully

### Requirement 12: Обработка ошибок и логирование

**User Story:** Как администратор системы, я хочу иметь подробное логирование и обработку ошибок, чтобы диагностировать проблемы

#### Acceptance Criteria

1. WHEN an error occurs, THE HRMS_System SHALL log error details with timestamp, module name, and stack trace
2. WHEN an error occurs in GUI_Application, THE HRMS_System SHALL display user-friendly error message
3. THE HRMS_System SHALL log all data modifications with user action, timestamp, and affected records
4. THE HRMS_System SHALL rotate log files when size exceeds 10 MB
5. THE HRMS_System SHALL store logs in logs/ directory with filename format: hrms_{date}.log
6. WHERE debug mode is enabled, THE HRMS_System SHALL log detailed information about data operations

### Requirement 13: Парсинг и сериализация данных Excel

**User Story:** Как разработчик, я хочу надежно читать и записывать данные в Excel, чтобы избежать потери или повреждения данных

#### Acceptance Criteria

1. WHEN Excel_Database is opened, THE Python_Backend SHALL parse all sheets into pandas DataFrames
2. WHEN data is modified in Python_Backend, THE Python_Backend SHALL serialize DataFrames back to Excel_Database
3. THE Python_Backend SHALL preserve data types: dates as datetime, numbers as numeric, text as string
4. THE Python_Backend SHALL handle empty cells as None values
5. FOR ALL valid DataFrames, parsing then serializing then parsing SHALL produce equivalent DataFrames (round-trip property)
6. IF Excel_Database structure is invalid, THEN THE Python_Backend SHALL display descriptive error with expected structure

### Requirement 14: Технологический стек

**User Story:** Как разработчик, я хочу использовать современные и надежные библиотеки, чтобы система была стабильной и производительной

#### Acceptance Criteria

1. THE HRMS_System SHALL require Python version 3.10 or higher
2. THE HRMS_System SHALL use xlwings for Excel integration
3. THE HRMS_System SHALL use pandas for data processing
4. THE HRMS_System SHALL use openpyxl for .xlsx file operations
5. THE HRMS_System SHALL use customtkinter for GUI
6. THE HRMS_System SHALL use Pillow for image processing in GUI
7. THE HRMS_System SHALL use docxtpl for Word document generation
8. THE HRMS_System SHALL use loguru for logging
9. THE HRMS_System SHALL provide requirements.txt with all dependencies and version constraints

### Requirement 15: Интеграция с существующей системой

**User Story:** Как пользователь, я хочу сохранить возможность запуска Python из Excel через макросы, чтобы использовать привычный workflow

#### Acceptance Criteria

1. THE HRMS_System SHALL maintain compatibility with existing macro.vba for launching Python scripts
2. THE HRMS_System SHALL support launching GUI_Application from Excel button via VBA macro
3. THE HRMS_System SHALL support launching specific functions (e.g., sort employees) from Excel via VBA
4. WHEN launched from Excel, THE HRMS_System SHALL automatically detect and use the calling workbook
5. THE HRMS_System SHALL work both when launched from Excel and as standalone application
6. THE HRMS_System SHALL preserve existing sort_employees.py functionality in new architecture
