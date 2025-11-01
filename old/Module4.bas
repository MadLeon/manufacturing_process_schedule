Attribute VB_Name = "Module4"
Sub RemoveShippedItems()
Attribute RemoveShippedItems.VB_Description = "This Macro will remove jobs that have already shipped."
Attribute RemoveShippedItems.VB_ProcData.VB_Invoke_Func = " \n14"
Application.ScreenUpdating = False
Dim iListCount As Integer
Dim iCtr As Integer
'
' Shipped Macro
' This Macro will remove jobs that have already shipped.
'

'
    Windows("Manufacturing process Schedule.xlsm").Activate
    Sheets("Shipped").Select
    Columns("A").EntireColumn.Delete
   
    Workbooks.Open ("\\bvserver\oe\order entry log.xlsm")
    Windows("Order Entry Log.xlsm").Activate
    Sheets("DELIVERY SCHEDULE").Select
    Range("B4:B1000").Select
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
    Range("h3:h1000").Select
    Selection.Copy
    Sheets("List").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
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


Windows("Order Entry Log.xlsm").Activate
Workbooks("order entry log.xlsm").Close SaveChanges:=False
 
 
 
Sheets("DELIVERY SCHEDULE TRACKING").Select
Range("H1500").End(xlUp).Select

Application.ScreenUpdating = True
MsgBox "SHIPPED ITEM REMOVAL COMPLETED"

'Application.ScreenUpdating = True
'MsgBox "Done!"

End Sub

