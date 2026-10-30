Attribute VB_Name = "ExcelReportMacros"
Option Explicit

Private Const DATA_SHEET As String = "Consolidated"
Private Const CTRL_SHEET As String = "Control Panel"

Public Sub ImportExcelFiles()
    Dim picker As FileDialog, f As Variant, src As Workbook, ws As Worksheet
    Dim dst As Worksheet, lastRow As Long, lastCol As Long, destRow As Long
    Dim n As Long, total As Long, expected As Variant, i As Long, v As Variant
    Dim started As Double, originalCalc As XlCalculation
    On Error GoTo Fail
    Set picker = Application.FileDialog(msoFileDialogFilePicker)
    With picker
        .Title = "Select monthly sales Excel files"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel Workbooks", "*.xlsx;*.xlsm"
        If .Show <> -1 Then Exit Sub
        If .SelectedItems.Count = 0 Then Exit Sub
    End With
    started = Timer
    originalCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Set dst = ThisWorkbook.Worksheets(DATA_SHEET)
    dst.Range("A2:J" & dst.Rows.Count).ClearContents
    expected = Array("Transaction_ID", "Date", "Department", "Region", "Product", "Quantity", "Unit_Price", "Amount", "Sales_Rep")
    destRow = 2
    For Each f In picker.SelectedItems
        If StrComp(CStr(f), ThisWorkbook.FullName, vbTextCompare) = 0 Then Err.Raise vbObjectError + 100, , "Do not select the macro workbook itself."
        Set src = Workbooks.Open(Filename:=CStr(f), ReadOnly:=True, UpdateLinks:=0)
        Set ws = src.Worksheets(1)
        For i = 0 To 8
            If StrComp(Trim$(CStr(ws.Cells(1, i + 1).Value2)), expected(i), vbTextCompare) <> 0 Then
                Err.Raise vbObjectError + 101, , "Unexpected column " & (i + 1) & " in " & src.Name & ". Expected: " & expected(i)
            End If
        Next i
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow > 1 Then
            n = lastRow - 1
            If destRow + n > dst.Rows.Count Then Err.Raise vbObjectError + 102, , "Excel row limit exceeded."
            v = ws.Range("A2:I" & lastRow).Value2
            dst.Cells(destRow, 1).Resize(n, 9).Value2 = v
            dst.Cells(destRow, 10).Resize(n, 1).Value2 = src.Name
            destRow = destRow + n
            total = total + n
        End If
        src.Close SaveChanges:=False
        Set src = Nothing
    Next f
    dst.Columns("A:J").AutoFit
    dst.Columns("B").NumberFormat = "yyyy-mm-dd"
    dst.Columns("G:H").NumberFormat = "#,##0.00"
    ThisWorkbook.Worksheets(CTRL_SHEET).Range("C12").Value = "Imported " & total & " rows from " & picker.SelectedItems.Count & " files"
    With ThisWorkbook.Worksheets("Run_Summary")
        .Range("A1:B10").ClearContents
        .Range("A1").Value = "Metric": .Range("B1").Value = "Value"
        .Range("A2").Value = "Imported rows": .Range("B2").Value = total
    End With
CleanExit:
    On Error Resume Next
    If Not src Is Nothing Then src.Close SaveChanges:=False
    Application.Calculation = originalCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    If total > 0 Then MsgBox total & " rows imported. Next: Generate Pivot Report.", vbInformation
    Exit Sub
Fail:
    Dim msg As String
    msg = Err.Description
    ThisWorkbook.Worksheets(CTRL_SHEET).Range("C12").Value = "Import failed: " & msg
    total = 0
    Resume CleanExit
End Sub

Public Sub GeneratePivotReport()
    Dim wsData As Worksheet, wsPivot As Worksheet, wsRun As Worksheet
    Dim outBook As Workbook, outData As Worksheet, outPivot As Worksheet, outRun As Worksheet
    Dim cache As PivotCache, pt As PivotTable, pf As PivotField
    Dim groupField As String, valueField As String, agg As String
    Dim fn As Variant, lastRow As Long, summary As Worksheet, func As XlConsolidationFunction
    Dim startTime As Double, n As Long
    On Error GoTo Fail
    startTime = Timer
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then MsgBox "Import files first.", vbExclamation: Exit Sub
    groupField = Trim$(CStr(ThisWorkbook.Worksheets(CTRL_SHEET).Range("C8").Value2))
    valueField = Trim$(CStr(ThisWorkbook.Worksheets(CTRL_SHEET).Range("C9").Value2))
    agg = LCase$(Trim$(CStr(ThisWorkbook.Worksheets(CTRL_SHEET).Range("C10").Value2)))
    If Not IsValidField(groupField) Or Not IsValidField(valueField) Then Err.Raise vbObjectError + 120, , "Invalid grouping/value field."
    If StrComp(groupField, valueField, vbTextCompare) = 0 Then Err.Raise vbObjectError + 121, , "Group and value fields must differ."
    Select Case agg
        Case "sum": func = xlSum
        Case "count": func = xlCount
        Case "average": func = xlAverage
        Case "min": func = xlMin
        Case "max": func = xlMax
        Case Else: Err.Raise vbObjectError + 122, , "Invalid aggregation."
    End Select
    fn = Application.GetSaveAsFilename(InitialFileName:="Monthly_Excel_Report.xlsx", FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    If VarType(fn) = vbBoolean Then Exit Sub
    Application.ScreenUpdating = False
    Set outBook = Workbooks.Add(xlWBATWorksheet)
    Set outData = outBook.Worksheets(1): outData.Name = "Consolidated"
    wsData.Range("A1:J" & lastRow).Copy
    outData.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
    outData.Columns("A:J").AutoFit
    outData.Columns("B").NumberFormat = "yyyy-mm-dd"
    outData.Columns("G:H").NumberFormat = "#,##0.00"
    Set outPivot = outBook.Worksheets.Add(After:=outData): outPivot.Name = "Pivot_Summary"
    Set cache = outBook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=outData.Range("A1:J" & lastRow), Version:=xlPivotTableVersion15)
    Set pt = cache.CreatePivotTable(TableDestination:=outPivot.Range("A3"), TableName:="SalesPivot")
    With pt
        .PivotFields(groupField).Orientation = xlRowField
        .PivotFields(groupField).Position = 1
        Set pf = .AddDataField(.PivotFields(valueField), agg & " of " & valueField, func)
        pf.NumberFormat = "#,##0.00"
        .RowAxisLayout xlTabularRow
    End With
    outPivot.Range("A1").Value = "MONTHLY PIVOT SUMMARY"
    outPivot.Range("A1").Font.Bold = True
    outPivot.Range("A1").Font.Size = 16
    outPivot.Columns("A:D").AutoFit
    Set outRun = outBook.Worksheets.Add(After:=outPivot): outRun.Name = "Run_Summary"
    outRun.Range("A1:B6").Value = ""
    outRun.Range("A1").Value = "Metric": outRun.Range("B1").Value = "Value"
    outRun.Range("A2").Value = "Rows": outRun.Range("B2").Value = lastRow - 1
    outRun.Range("A3").Value = "Group field": outRun.Range("B3").Value = groupField
    outRun.Range("A4").Value = "Value field": outRun.Range("B4").Value = valueField
    outRun.Range("A5").Value = "Aggregation": outRun.Range("B5").Value = agg
    outRun.Range("A6").Value = "Generated": outRun.Range("B6").Value = Now
    outRun.Columns("A:B").AutoFit
    Application.DisplayAlerts = False
    outBook.SaveAs Filename:=CStr(fn), FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    ThisWorkbook.Worksheets(CTRL_SHEET).Range("C12").Value = "Saved " & (lastRow - 1) & " rows: " & CStr(fn)
    Application.ScreenUpdating = True
    MsgBox "Report generated: " & CStr(fn), vbInformation
    Exit Sub
Fail:
    Dim errText As String
    errText = Err.Description
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Report generation failed: " & errText, vbCritical
End Sub

Private Function IsValidField(ByVal fieldName As String) As Boolean
    Dim i As Long
    For i = 1 To 10
        If StrComp(fieldName, CStr(ThisWorkbook.Worksheets(DATA_SHEET).Cells(1, i).Value2), vbTextCompare) = 0 Then
            IsValidField = True: Exit Function
        End If
    Next i
End Function

Public Sub ResetImportedData()
    If MsgBox("Clear all imported data? Original files are unaffected.", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    With ThisWorkbook.Worksheets(DATA_SHEET)
        .Range("A2:J" & .Rows.Count).ClearContents
    End With
    ThisWorkbook.Worksheets(CTRL_SHEET).Range("C12").Value = "Ready"
    MsgBox "Imported data cleared.", vbInformation
End Sub
