Attribute VB_Name = "PartsShipped"
Sub Delete_Crossed_Out_Shipping()

    If MsgBox("Run this macro?", vbYesNo) = vbNo Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("DELIVERY SCHEDULE")
    'Set ws2 = ThisWorkbook.Worksheets("sheet2")
    'ws2.Activate
    ws.Activate
'Removes all Filter before running script'
    On Error Resume Next
        ws.ShowAllData
    On Error GoTo 0
'1.Filtering '
    'ws.Range("A2:S2000").AutoFilter Field:=17, Criteria1:=">0"
    'ws.Range("Q4").AutoFilter Field:=17, Criteria1:="<>"
    ws.Range("R4").AutoFilter Field:=18, Criteria1:="<>"
'2. Copies the Filtered Data'
    Application.DisplayAlerts = False
    
    
    'ws.Range("A4:S2000").SpecialCells(xlCellTypeVisible).Copy
    'ws.Range("A4:R4").End(xlDown).Copy
    ws.Range("A4", Range("R4").End(xlDown)).Copy
    Application.DisplayAlerts = True
'Removes the filter applied from code above'
   'ws.AutoFilter.ShowAllData
    
 
'3.Paste into another workbook'
    
'Declaring The secondary Workbook'
    Dim wb3 As Workbook
    
    Application.ScreenUpdating = False

    
  
'Declaring location of the new workbook'
    
    Set wb3 = Workbooks.Open(Filename:="\\rtdnas2\OE\Parts Shipped.xlsm", Password:="8849", WriteResPassword:="8849")

    wb3.Worksheets("parts shipped").Activate
'Selects the last row of the new sheet and pastes it at the next available spot'
    lastrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
    ActiveSheet.Cells(lastrow + 1, 1).Select
    ActiveSheet.Paste
'removes error displays'
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    ActiveWorkbook.Close savechanges = True
    Application.DisplayAlerts = True
'Small delay to reduce errors in the future'
    Application.Wait (Now + TimeValue("0:00:01"))
    Application.CutCopyMode = False
    Set wb3 = Nothing
    
    
    
'Deletes The data from First sheet that was copied and pasted into secondary sheet'
  '  ws.Range("A2:S2000").SpecialCells(xlCellTypeVisible).Delete'
    Application.ScreenUpdating = True

    Application.DisplayAlerts = False
    'ws.Range("A4", Range("R4").End(xlDown)).Delete
    Application.DisplayAlerts = True
        
   ws.ShowAllData
   ws.Range("A3:R1000").AutoFilter
   ws.Range("A3:R1000").AutoFilter
   
   
    
End Sub






