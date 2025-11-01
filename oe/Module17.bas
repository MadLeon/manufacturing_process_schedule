Attribute VB_Name = "Module17"
Sub KinectrisData()
    Application.ScreenUpdating = False
    Dim i As Range
    Dim num As Integer
    num = 1
    
'Filter for "Kinectrics" in "Delivery Schedule", starts
Workbooks("Order Entry Log.xlsm").Activate
Sheets("Delivery Schedule").Select

    Columns("G:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("K:K").Select
    Selection.EntireColumn.Hidden = True
    Columns("M:O").Select
    Selection.EntireColumn.Hidden = True
    Columns("Q:R").Select
    Selection.EntireColumn.Hidden = True
    
    ActiveSheet.Range("$A$3:$R$5000").AutoFilter Field:=3, Criteria1:= _
    "Kinectrics"
    'Sheets("Delivery Schedule").Select
    Range("A4").Select
'Filter for "Kinectrics" in "Delivery Schedule", ends


' Copy all "Kinectrics" Visibles to "Temp", starts

Selection.SpecialCells(xlCellTypeVisible).End(xlDown).Select
ActiveCell.CurrentRegion.Select
ActiveCell.CurrentRegion.Copy

Sheets("Temp").Select
Range("A1").Select
ActiveCell.PasteSpecial
Selection.FormatConditions.Delete

'Range(A1).EntireRow.Delete
' Copy all "Kinectrics" Visibles to "Temp", ends

Workbooks("Order Entry Log.xlsm").Activate
Sheets("Kinectrics").Select
'Clear all Filters if exist
Workbooks("Order Entry Log.xlsm").Activate
Sheets("Kinectrics").Select
'ActiveSheet.AutoFilter.ShowAllData


'Finding the last number in Col B from Raw 1500, then copy the value to "Cal" => "a1"
Range("B1500").End(xlUp).Select
ActiveCell.Copy
Sheets("Cal").Select
Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
False, Transpose:=True
   
    tgtVal = (Workbooks("Order Entry Log.xlsm").Sheets("Cal").Range("a1"))
    
'   Comparing value with numbers in Col B from "Delivery Schedule"
    Sheets("Temp").Activate
    'Selection.SpecialCells(xlCellTypeVisible).Select
        For Each i In Range("b2:b1000")
        If i.Value > (tgtVal) Then

            i.Select
            ActiveCell.Rows("1:1").EntireRow.Select
            'Selection.SpecialCells(xlCellTypeVisible).Select
            Selection.Copy
            Sheets("Kinectrics").Range("a1500").End(xlUp).Offset(num, 0).PasteSpecial Paste:=xlPasteValues
            'ActiveCell.Rows.Delete
            num = num + 1
'            Workbooks("Order Entry Log.xlsm").Activate
            Sheets("Temp").Activate


        End If
    Next i
    
'        Workbooks("Manufacturing Process Schedule.xlsm").Activate
            Sheets("Kinectrics").Activate

On Error Resume Next
Columns("a").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

Sheets("Cal").Select
Sheets("Cal").Cells.ClearContents

Sheets("Temp").Select
Sheets("Temp").Cells.ClearContents

Sheets("Delivery Schedule").Select
Columns("A:T").Select
    Selection.EntireColumn.Hidden = False
    
    'This Un-filter everything
        ActiveSheet.ShowAllData
        Range("A1").End(xlDown).Select
        
Sheets("Kinectrics").Select
Range("A2").End(xlDown).Select

    Application.ScreenUpdating = True
    MsgBox "DATA UPDATE COMPLETED"

End Sub



