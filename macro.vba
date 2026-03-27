Sub SortEmployees()
    RunPython "import sort_employees; sort_employees.show_employee_selector()"
End Sub

' Test infrastructure
Sub TestInfrastructure()
    RunPython "import test_infrastructure; test_infrastructure.main()"
End Sub

' Launch HRMS GUI (будет реализовано позже)
Sub LaunchHRMS()
    RunPython "import main; main.run()"
End Sub
