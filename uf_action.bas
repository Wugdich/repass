Attribute VB_Name = "uf_action"
Option Explicit

Sub update_avaliable_data()

    ra_uf.data_file_cb.Clear
    Dim xl_regexp As New RegExp
    xl_regexp.Pattern = ".xls"
    
    ' curren working directory
    Dim cwd As String
    cwd = ThisWorkbook.Path
    ' current data directory
    Dim cdd As String
    cdd = cwd & "\" & ra_uf.r_type_cb.Value & "\data\"
    Dim filename As String
    filename = Dir(cdd)
    
    Do While filename <> ""
        If xl_regexp.test(filename) = True Then
            ra_uf.data_file_cb.AddItem (filename)
        End If
        filename = Dir()
    Loop

End Sub

Sub open_report_files()
    ' Open files that required to prepare reports: data and template (both excel files).
    
    ' template file path define - tfp
    ' crd - current report directory
    Dim crd As String
    crd = ThisWorkbook.Path & "\" & ra_uf.r_type_cb.Value
    Dim tfp As String
    tfp = crd & "\report_template.xls"
    
    ' data file path (dfp) define
    Dim dfp As String
    dfp = crd & "\data\" & ra_uf.data_file_cb.Value
    
    ' open template and data files
    With Workbooks
        .Open (tfp)
        .Open (dfp)
    End With
    
End Sub

Sub type1_report_prepare()

    ' data updating stuff
    Call type1_do_sheet1
    Call type1_do_sheet3
    Call type1_update_captions
    Call type1_copy_image
    ' memorize object that represent workbook to save
    Dim wb_to_close As Object
    Set wb_to_close = Workbooks("report_template.xls")
    ' save report workbooks as
    Call type1_save_workbook
    ' close report workbook
    wb_to_close.Close SaveChanges:=False
    ' close data workbooks
    Workbooks(ra_uf.data_file_cb.Value).Close SaveChanges:=False
    
    
End Sub

Sub type1_do_sheet1()
    ' SHEET "НТВ весь день" PREPARE
    
    Dim ws1 As Object
    Set ws1 = Workbooks("report_template.xls").Worksheets("НТВ весь день")
    Dim data_wb As Object
    Set data_wb = Workbooks(ra_uf.data_file_cb.Value)
    
    ' CLEAR TABLE
    ws1.Range("A3:E100").Clear
    
    ' DATA TRANSFER
    
    ' find data respons to NTV
    Dim firstrow As Long
    Dim lastrow As Long
    With data_wb.Worksheets("Программы")
        firstrow = .Range("A:A").Find(what:="НТВ", After:=.Range("A1")).row
        lastrow = .Range("A:A").Find(what:="НТВ", After:=.Range("A1"), searchdirection:=xlPrevious).row
    End With
    
    ' transfer values
    Dim dRow As Long    ' data row
    Dim rRow As Long    ' report row
    rRow = 3
    For dRow = firstrow To lastrow
        data_wb.Worksheets("Программы").Range("C" & dRow & ":D" & dRow).Copy
        ws1.Range("A" & rRow).PasteSpecial Paste:=xlPasteValues
        data_wb.Worksheets("Программы").Range("F" & dRow & ":H" & dRow).Copy
        ws1.Range("C" & rRow).PasteSpecial Paste:=xlPasteValues
        rRow = rRow + 1
    Next dRow

    ' FORMATTING TABLE
    Dim i As Long
    rRow = 3
    For i = firstrow To lastrow
        If rRow Mod 2 = 0 Then
            Workbooks("report_template.xls").Worksheets("Tools").Range("A2:E2").Copy
        Else
            Workbooks("report_template.xls").Worksheets("Tools").Range("A3:E3").Copy
        End If
        ws1.Range("A" & rRow).PasteSpecial Paste:=xlPasteFormats
        rRow = rRow + 1
    Next i
    
    With ws1.Range("A3:E" & (lastrow - firstrow + 3))
        ' borders
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
End Sub

Sub type1_do_sheet3()
    ' SHEET 3 "данные для итогов" UPDATE
    
    ' average shares
    Dim ws4 As Object
    Set ws4 = Workbooks("report_template.xls").Worksheets("данные для итогов")
    Dim wsData As Object
    Set wsData = Workbooks(ra_uf.data_file_cb.Value).Worksheets("Периоды")
    
    ws4.Range("N29:P40") = WorksheetFunction.Transpose(wsData.Range("N3:Y5"))
    
    ' channels rating during day
    wsData.Range("B7:H46").Copy
    ws4.Range("O48").PasteSpecial Paste:=xlPasteValues
    
    ' news programs
    ws4.Range("V4:AC40").Clear
    
    Dim irrelevant_programs(1 To 23) As String
    Dim i As Integer
    
    
    For i = 1 To 23
        irrelevant_programs(i) = Workbooks("report_template.xls").Worksheets("Tools").Range("G" & (i + 1)).Value
    Next i
    
    Dim relevant_channels(1 To 3) As String
    relevant_channels(1) = "НТВ"
    relevant_channels(2) = "РОССИЯ 1"
    relevant_channels(3) = "ПЕРВЫЙ КАНАЛ"
    
    Dim wsProgram As Object
    Set wsProgram = Workbooks(ra_uf.data_file_cb.Value).Worksheets("Программы")
    
    Dim lastrow As Long
    lastrow = wsProgram.Range("A4").CurrentRegion.Rows.Count
    
    Dim rRow As Long ' report row
    rRow = 4
    For i = 4 To lastrow
        If IsInArray(wsProgram.Range("A" & i).Value, relevant_channels) _
            And Not IsInArray(wsProgram.Range("D" & i).Value, irrelevant_programs) _
            And CInt(Left(wsProgram.Range("C" & i).Value, 2)) < 24 _
            And wsProgram.Range("E" & i).Value = "Новости" Then
                ' a liitle tricky, because i can't calculate time value more then 24:00:00
                If TimeValue(wsProgram.Range("C" & i).Value & ":00") > TimeValue("17:50:00") Then
                    wsProgram.Range("A" & i & ":H" & i).Copy
                    ws4.Range("V" & rRow).PasteSpecial
                    rRow = rRow + 1
                End If
        End If
    Next i
    
    Workbooks("report_template.xls").Worksheets("Итоги дня").ChartObjects(2).Chart.SetSourceData _
    Source:=Workbooks("report_template.xls").Worksheets("данные для итогов").Range("AE3:AG" & (rRow - 1))

End Sub

Sub type1_update_captions()

    Dim region_cell As String
    region_cell = Workbooks(ra_uf.data_file_cb.Value).Worksheets("Программы").Range("A1").Value
    Dim region_regexp As New RegExp
    region_regexp.Pattern = "Russia"
    Dim region As String
    
    If region_regexp.test(region_cell) = True Then
        region = "Россия"
    Else
        region = "Москва"
    End If
    
    Dim data_type As String
    Dim cb_data_type As String
    cb_data_type = ra_uf.data_type_cb.Value
    If cb_data_type = "Фактические" Then
        data_type = "(фактические данные)"
    ElseIf cb_data_type = "Предварительные" Then
        data_type = "(предварительные данные)"
    Else
        data_type = "(ускоренные предварительные данные)"
    End If
    
    Dim caption As String
    caption = region & ", Все, 18+ " & data_type
    Workbooks("report_template.xls").Worksheets("НТВ весь день").PageSetup.RightHeader = "&""Arial""&K000000&B&14" & caption
    With Workbooks("report_template.xls").Worksheets("Итоги дня").Cells(2, 1)
        .Value = caption
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
End Sub

Sub type1_copy_image()
    
    ' sheet "Каналы с 6:00 утра" prepare
    
    Workbooks(ra_uf.data_file_cb.Value).Worksheets("График телесмотрения").Shapes(1).Copy
    Application.Wait (Now + TimeValue("0:00:03"))
    Workbooks("report_template.xls").Worksheets("Каналы с 6.00 утра").Paste Destination:=Workbooks("report_template.xls").Worksheets("Каналы с 6.00 утра").Range("A1")
    
End Sub

Sub type1_save_workbook()
    
    Dim workbook_name As String
    ' date
    Dim report_date As String
    report_date = Replace(Workbooks("report_template.xls").Worksheets("Setup").Range("G3").Value, ".", "")
    
    ' region
    Dim region_cell As String
    region_cell = Workbooks(ra_uf.data_file_cb.Value).Worksheets("Программы").Range("A1").Value
    Dim region_regexp As New RegExp
    region_regexp.Pattern = "Russia"
    Dim region As String
    If region_regexp.test(region_cell) = True Then
        region = "R"
    Else
        region = "M"
    End If
    
    ' data type
    Dim data_type As String
    Dim data_type_cb As String
    data_type_cb = ra_uf.data_type_cb.Value
    If ra_uf.data_type_cb.Value = "Фактические" Then
        data_type = "(фактические данные)"
    ElseIf data_type_cb = "Предварительные" Then
        data_type = "(предварительные данные)"
    Else
        data_type = "(ускоренные предварительные данные)"
    End If
    
    workbook_name = report_date & region & " " & data_type & ".xls"
    
    ' directory to save
    Dim s_dir As String
    s_dir = ThisWorkbook.Path & "\" & ra_uf.r_type_cb.Value & "\reports\"
    
    Workbooks("report_template.xls").SaveAs filename:=(s_dir & workbook_name)
    
End Sub

Sub type2_report_prepare()
    
    MsgBox "empty"
    
End Sub
