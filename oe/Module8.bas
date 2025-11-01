Attribute VB_Name = "Module8"
Sub PastDUE()
Attribute PastDUE.VB_ProcData.VB_Invoke_Func = " \n14"

Application.ScreenUpdating = False
' Clear existing data
    Sheets("PastDUE").Select
    Cells.Select
    Selection.ClearContents
'
'  Format PastDUE sheet and copy data into
    Sheets("DELIVERY SCHEDULE").Select
    Range("A3:P700").Select
    Selection.Copy
    Sheets("PastDUE").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    Range("A1").Select
    ActiveSheet.Paste
' Remove unwanted columns
    Range("A:A,F:I").Select
    Range("A:A,F:I,K:O").Select
    Selection.Delete Shift:=xlToLeft
    Range("A2").Select
'Delete empty rows
On Error Resume Next
Columns("a").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
Range("A1").Select
    ActiveSheet.Range("$A$1:$F$346").AutoFilter Field:=1, Criteria1:="Job #:"
    ActiveSheet.Range("$A$1:$F$346").AutoFilter Field:=1
    ActiveWorkbook.Worksheets("PastDUE").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("PastDUE").AutoFilter.Sort.SortFields.Add Key:= _
        Range("F1:F346"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("PastDUE").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Search result for records between 2 dates
    
    Dim early As Date, late As Date, N As Long
    Dim dt As Date

    early = CDate(Application.InputBox(Prompt:="Please enter start date:", Type:=2))
    late = CDate(Application.InputBox(Prompt:="Please enter end date:", Type:=2))
    MsgBox early & vbCrLf & late

    N = Cells(Rows.Count, "F").End(xlUp).row
    
    On Error Resume Next

    For i = N To 2 Step -1
        dt = Cells(i, 6).Value
        If dt > late Or dt < early Then
            Cells(i, 6).EntireRow.Delete
        End If
    Next i
    
    ' PastDUEsortbyCustomer Macro
    ActiveWorkbook.Worksheets("PastDUE").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("PastDUE").AutoFilter.Sort.SortFields.Add Key:= _
        Range("B1:B700"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("PastDUE").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

