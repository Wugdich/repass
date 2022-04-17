Attribute VB_Name = "tests"
Option Explicit

Sub main_test()

    Dim t As String
    t = Workbooks("Выгрузка для ежедневного отчета 1 (18.03.2022).xlsx").Worksheets(1).Range("C4").Value
    MsgBox (TimeValue(t) > TimeValue("8:00:00"))
    
End Sub
