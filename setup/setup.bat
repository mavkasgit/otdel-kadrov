@echo off
chcp 65001 >nul
echo ========================================
echo УСТАНОВКА PYTHON + XLWINGS ДЛЯ EXCEL
echo ========================================
echo.

REM Проверка Python
echo [1/5] Проверка Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python НЕ установлен!
    echo.
    echo Скачайте Python с https://www.python.org/downloads/
    echo ВАЖНО: При установке обязательно отметьте галочку "Add Python to PATH"!
    echo.
    pause
    exit /b 1
) else (
    python --version
    echo ✓ Python установлен
)
echo.

REM Проверка pip
echo [2/5] Проверка pip...
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ pip не найден! Убедитесь, что Python установлен корректно.
    pause
    exit /b 1
) else (
    echo ✓ pip установлен
)
echo.

REM Установка xlwings
echo [3/5] Установка библиотеки xlwings...
pip install xlwings
if %errorlevel% neq 0 (
    echo ❌ Ошибка установки xlwings
    pause
    exit /b 1
) else (
    echo ✓ Библиотека xlwings установлена
)
echo.

REM Установка xlwings add-in
echo [4/5] Интеграция xlwings в Excel (add-in)...
xlwings addin install
if %errorlevel% neq 0 (
    echo ⚠ Возможна ошибка установки надстройки. Закройте Excel и попробуйте снова.
) else (
    echo ✓ Надстройка xlwings добавлена в Excel
)
echo.

REM Проверка установки
echo [5/5] Финальная проверка...
python -c "import xlwings; print('xlwings версия:', xlwings.__version__)"
if %errorlevel% neq 0 (
    echo ❌ Ошибка при проверке.
    pause
    exit /b 1
)
echo.

echo ========================================
echo ✓ УСТАНОВКА ЗАВЕРШЕНА УСПЕШНО!
echo ========================================
echo.
echo Теперь прочитайте файл "ИНСТРУКЦИЯ.txt" для настройки самого Excel.
echo.
pause