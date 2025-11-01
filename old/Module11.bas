Attribute VB_Name = "Module11"
Sub MoveData()
    Application.ScreenUpdating = False
    Dim i As Range
    Dim num As Integer
    num = 1
    'Find the very last used cell  in a Column:
'Application.ScreenUpdating = False
On Error Resume Next
'Filter to show ALL
ActiveSheet.ShowAllData

Range("JOB#").Sort Key1:=Range("h1"), Order1:=xlDescending

' autoFilter Macro
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=1
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=2
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=3
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=4
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=5
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=6
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=7
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=8
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=13
    ActiveWorkbook.Worksheets("DELIVERY SCHEDULE TRACKING").AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("DELIVERY SCHEDULE TRACKING").AutoFilter.Sort. _
        SortFields.Add Key:=Range("H2:H1500"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DELIVERY SCHEDULE TRACKING").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Workbooks("Manufacturing Process Schedule.xlsm").Activate
Sheets("DELIVERY SCHEDULE TRACKING").Select
Range("H1500").End(xlUp).Select
ActiveCell.Copy
Sheets("Cal").Select
Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
   
    tgtVal = (Workbooks("Manufacturing Process Schedule.xlsm").Sheets("Cal").Range("a1"))

    
    Workbooks.Open ("\\RTDNAS2\oe\order entry log.xlsm")
    Workbooks("Order Entry Log.xlsm").Activate
    Sheets("Delivery Schedule").Activate
    
    On Error Resume Next
    'Filter to show ALL
    ActiveSheet.ShowAllData
    
    For Each i In Range("B4:B1500")
        'If i.Value > 55449 Then
        If i.Value > (tgtVal) Then

            i.Select
            ActiveCell.Rows("1:1").EntireRow.Select
            Selection.Copy
            Workbooks("Manufacturing Process Schedule.xlsm").Activate
            Sheets("Temp").Range("a250").End(xlUp).Offset(num, 0).PasteSpecial
            'ActiveCell.Rows.Delete
            num = num + 1
            Workbooks("Order Entry Log.xlsm").Activate
            Sheets("Delivery Schedule").Activate


        End If
    Next i
    
        Workbooks("Manufacturing Process Schedule.xlsm").Activate
            Sheets("Temp").Activate

On Error Resume Next
Columns("a").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

Columns("m:o").EntireColumn.Delete
Columns("n:p").EntireColumn.Delete
Columns("g").EntireColumn.Delete
Columns("h").EntireColumn.Delete
Columns("i").EntireColumn.Delete

' Move PO to Col B
    Columns("I:I").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight

' Move DWG Rel to Col C
    Columns("H:H").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
' Move Pasrt# to Col D
    Columns("G:G").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    
' Move Description to Col E
    Columns("I:I").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight

' Move Customer to Col F
    Columns("G:G").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    
' Move QTY to Col G
    Columns("H:H").Select
    Selection.Cut
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight

' Move Job# to Col H (no need)
    
' Move  Due Date to Col I
    Columns("J:J").Select
    Selection.Cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
'Delete Col "J" to finalize
Columns("J").EntireColumn.Delete

    
Workbooks("Manufacturing Process Schedule.xlsm").Activate
Sheets("temp").Select
Range("A1").End(xlDown).Select
'Range(ActiveCell, ActiveCell.End(xlUp)).Select
ActiveCell.CurrentRegion.Select
ActiveCell.CurrentRegion.Copy

Workbooks("Manufacturing Process Schedule.xlsm").Activate
Sheets("delivery schedule tracking").Select
Range("A2").End(xlDown).Select
ActiveCell.Offset(1, 0).Select
ActiveCell.PasteSpecial xlPasteValues

Workbooks("Manufacturing Process Schedule.xlsm").Activate
Sheets("temp").Select
Range("A1").End(xlDown).Select
'Range(ActiveCell, ActiveCell.End(xlUp)).Select
ActiveCell.CurrentRegion.Select
ActiveCell.CurrentRegion.ClearContents

Workbooks("Manufacturing Process Schedule.xlsm").Activate
Sheets("delivery schedule tracking").Select
Range("A2").End(xlDown).Select

' Shipped Macro
' This Macro will remove jobs that have already shipped.
'

'
    Windows("Manufacturing process Schedule.xlsm").Activate
    Sheets("Shipped").Select
    Columns("A").EntireColumn.Delete
   
    Windows("Order Entry Log.xlsm").Activate
    Sheets("DELIVERY SCHEDULE").Select
    Range("B4:B1500").Select
    Selection.Copy
    Windows("Manufacturing process Schedule.xlsm").Activate
    Sheets("Shipped").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Rows("1:3").Select
   ' Application.CutCopyMode = False
    'Selection.Delete Shift:=xlUp
    Range("A1").Select
    
    Columns("a").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
      
    Sheets("delivery schedule tracking").Select
    Range("h3:h500").Select
    Selection.Copy
    Sheets("List").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
          
        
    Dim iListCount As Integer
    Dim iCtr As Integer

' Turn off screen updating to speed up macro.
' Application.ScreenUpdating = False

' Get count of records to search through (list that will be deleted).
iListCount = Sheets("List").Cells(Rows.Count, "A").End(xlUp).Row

' Loop through the "master" list.
For Each x In Sheets("Shipped").Range("A1:A" & Sheets("shipped").Cells(Rows.Count, "A").End(xlUp).Row)
   ' Loop through all records in the second list.
   For iCtr = iListCount To 1 Step -1
      ' Do comparison of next record.
      ' To specify a different column, change 1 to the column number.
      If x.Value = Sheets("List").Cells(iCtr, 1).Value Then
         ' If match is true then delete row.
         Sheets("List").Cells(iCtr, 1).EntireRow.Delete
       End If
   Next iCtr
Next
'Application.ScreenUpdating = True
'MsgBox "Done!"


Sheets("DELIVERY SCHEDULE TRACKING").Select
Range("H1500").End(xlUp).Select

'MsgBox "DATA UPDATE COMPLETED"

'Sub RemoveShippedItems()
'Dim iListCount As Integer
'Dim iCtr As Integer

' Turn off screen updating to speed up macro.
' Application.ScreenUpdating = False

' Get count of records to search through (list that will be deleted).
iListCount = Sheets("DELIVERY SCHEDULE TRACKING").Cells(Rows.Count, "H").End(xlUp).Row

' Loop through the "master" list.
For Each x In Sheets("list").Range("A1:A" & Sheets("list").Cells(Rows.Count, "A").End(xlUp).Row)
   ' Loop through all records in the second list.
   For iCtr = iListCount To 3 Step -1
      ' Do comparison of next record.
      ' To specify a different column, change 1 to the column number.
      If x.Value = Sheets("DELIVERY SCHEDULE TRACKING").Cells(iCtr, 8).Value Then
         ' If match is true then delete row.
         Sheets("DELIVERY SCHEDULE TRACKING").Cells(iCtr, 1).EntireRow.Delete
       End If
   Next iCtr
Next

'copy original Due Date to the Act. DUE
'Range("I3:I500").Copy
 '   Range("J3").PasteSpecial
  '  Application.CutCopyMode = False
   ' Range("I500").End(xlUp).Select
    
    
 'define i as the range will be used
 '   Dim t As Long
    'find last cell in the column
  '  Range("J500").End(xlUp).Select
    'find the cell number of row
   ' d = ActiveCell.Row
    'loop too check the value of each cell from row 1 to activecell
   ' For t = 3 To d
        'if the date of each cell is earlier than today's date,then update
    '    If Range("J" & t).Value < Range("K1").Value Then

     '   Range("J" & t).Value = Range("K1").Value
      '  End If
   ' Next

'End Sub


Cells.Select
    Range("B1").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With

Windows("Order Entry Log.xlsm").Activate
Workbooks("order entry log.xlsm").Close SaveChanges:=False


    Range("I500").End(xlUp).Select

    Application.ScreenUpdating = True
    MsgBox "DATA UPDATE COMPLETED"

End Sub

