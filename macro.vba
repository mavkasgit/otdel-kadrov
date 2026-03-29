' HRMS System - VBA Macros
' Простой запуск Python без xlwings add-in

Public Const PYTHON_EXE = ".venv\Scripts\python.exe"
Public Const PROJECT_PATH = "D:\KTM\Excel\otdel-kadrov-xlwing"

Sub LaunchHRMS()
    ' Получаем путь к текущей книге
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim filePath As String
    filePath = wb.FullName
    
    ' Запуск HRMS с передачей пути к файлу
    Dim cmd As String
    cmd = PYTHON_EXE & " """ & PROJECT_PATH & "\main.py"" """ & filePath & """"
    
    Shell cmd, 1
End Sub

Sub TestInfrastructure()
    ' Тест инфраструктуры
    Dim cmd As String
    cmd = PYTHON_EXE & " """ & PROJECT_PATH & "\test_infrastructure.py"""
    
    Shell cmd, 1
End Sub

Sub SortEmployees()
    ' Сортировка сотрудников
    Dim cmd As String
    cmd = PYTHON_EXE & " """ & PROJECT_PATH & "\sort_employees.py"""
    
    Shell cmd, 1
End Sub
