' HRMS System - VBA Macros
' Одна кнопка в Excel - всё остальное в GUI

Public Const PYTHON_EXE = "python"

Function GetProjectPath() As String
    GetProjectPath = ThisWorkbook.Path & "\hrms"
End Function

Sub OpenHRMS()
    ' РЕЖИМ РЕЛИЗ - скрытая консоль (для пользователей)
    Dim cmd As String
    cmd = PYTHON_EXE & " """ & GetProjectPath() & "\main.py"""
    
    CreateObject("WScript.Shell").Run cmd, 0, False
End Sub

Sub OpenHRMS_Debug()
    ' РЕЖИМ ОТЛАДКИ - видимая консоль (видно ошибки)
    Dim cmd As String
    cmd = PYTHON_EXE & " """ & GetProjectPath() & "\main.py"""
    
    Shell cmd, 1
End Sub
