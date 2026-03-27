@echo off
echo ========================================
echo HRMS System - Установка зависимостей
echo ========================================
echo.

REM Проверка наличия Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: Python не найден!
    echo Установите Python 3.10+ с галочкой "Add Python to PATH"
    echo Скачать: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Python найден:
python --version
echo.

REM Создание виртуального окружения
echo Создание виртуального окружения (.venv)...
if exist .venv (
    echo Виртуальное окружение уже существует
) else (
    python -m venv .venv
    if errorlevel 1 (
        echo ОШИБКА: Не удалось создать виртуальное окружение
        pause
        exit /b 1
    )
    echo Виртуальное окружение создано
)
echo.

REM Активация виртуального окружения
echo Активация виртуального окружения...
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo ОШИБКА: Не удалось активировать виртуальное окружение
    pause
    exit /b 1
)
echo.

REM Обновление pip
echo Обновление pip...
python -m pip install --upgrade pip
echo.

REM Установка зависимостей
echo Установка зависимостей из requirements.txt...
pip install -r requirements.txt
if errorlevel 1 (
    echo ОШИБКА: Не удалось установить зависимости
    pause
    exit /b 1
)
echo.

echo ========================================
echo Установка завершена успешно!
echo ========================================
echo.
echo Для активации виртуального окружения используйте:
echo   .venv\Scripts\activate
echo.
echo Для запуска системы откройте Excel файл otdel-kadrov.xlsm
echo и нажмите кнопку запуска HRMS
echo.
pause
