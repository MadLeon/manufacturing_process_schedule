Attribute VB_Name = "Module5"
Sub NewOE()
'Find the very last used cell in a Column:
Application.ScreenUpdating = False
Sheets("DELIVERY SCHEDULE").Select
Range("a65536").End(xlUp).Select
ActiveCell.Copy
Sheets("Input Form").Select
Range("g5").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

End Sub
Sub NewJOB()
'Find the very last used cell in a Column:
Application.ScreenUpdating = False
Sheets("DELIVERY SCHEDULE").Select
Range("b65536").End(xlUp).Select
ActiveCell.Copy
Sheets("Input Form").Select
Range("g7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

End Sub

Sub EMT()
'    Call NewOE

Dim vNewRow As Long
    ' Find the first empty row in the data table
    vNewRow = Sheets("DELIVERY SCHEDULE").Cells(Rows.Count, 1).End(xlUp).Offset(2, 0).row
    'ActiveCell.Offset(2, 0).Select
    
    
    ' Check for data in Field 1
    If Trim(Range("OE").Value) = "" Then
        Range("OE").Activate
        MsgBox "Please enter data in OE!"
        Exit Sub
    End If
    ' Check for data in Field 2
   ' If Trim(Range("JobNum").Value) = "" Then
    '    Range("JobNum").Activate
     '   MsgBox "Please enter data in Field 2!"
      '  Exit Sub
   ' End If
    ' Check for data in Field 3
   ' If Trim(Range("Customer").Value) = "" Then
   '     Range("Customer").Activate
    '    MsgBox "Please enter data in Field 3!"
     '   Exit Sub
  '  End If
    ' Copy the data to the data table
    With SCHEDULE
        Sheets("DELIVERY SCHEDULE").Cells(vNewRow, 1).Value = Range("OE").Value
        Sheets("DELIVERY SCHEDULE").Cells(vNewRow, 2).Value = Range("JobNum").Value
        Sheets("DELIVERY SCHEDULE").Cells(vNewRow, 3).Value = Range("Customer").Value
        Sheets("DELIVERY SCHEDULE").Cells(vNewRow, 5).Value = Range("Parts").Value
        Sheets("DELIVERY SCHEDULE").Cells(vNewRow, 6).Value = Range("revision").Value
        Sheets("DELIVERY SCHEDULE").Cells(vNewRow, 10).Value = Range("desc").Value
        Sheets("DELIVERY SCHEDULE").Cells(vNewRow, 4).Value = Range("qty").Value
        Sheets("DELIVERY SCHEDULE").Cells(vNewRow, 16).Value = Range("date").Value
        Sheets("DELIVERY SCHEDULE").Cells(vNewRow, 7).Value = Range("contact").Value
        Sheets("DELIVERY SCHEDULE").Cells(vNewRow, 12).Value = Range("po").Value
        Sheets("DELIVERY SCHEDULE").Cells(vNewRow, 9).Value = Range("poline").Value
        Sheets("DELIVERY SCHEDULE").Cells(vNewRow, 11).Value = Range("price").Value
        Sheets("DELIVERY SCHEDULE").Cells(vNewRow, 8).Value = Range("od").Value
        
        'Application.OnTime Now + TimeValue("00:02:00"), "SaveWB"
        Application.OnTime Now + TimeValue("00:00:00"), "Stop_Timer"
        
        Application.OnTime Now + TimeValue("00:00:10"), "Start_Timer"

            
    End With
        ' Clear data fields and reset the form
    With InputForm
        '.Range("OE").Value = ""
        '.Range("JobNum").Value = ""
       ' .Range("Customer").Value = ""
        .Range("Parts").Value = ""
        .Range("Revision").Value = ""
        .Range("desc").Value = ""
        .Range("qty").Value = ""
        .Range("date").Value = ""
        '.Range("contact").Value = ""
        '.Range("po").Value = ""
        .Range("poline").Value = ""
        .Range("price").Value = ""
        .Range("OE").Select
        
    End With
   ' With Sheets("DELIVERY SCHEDULE")
    '    .Activate
     '   .Cells(vNewRow, 1).Select
  '  End With
 '   Call NewOE
    Call NewJOB
 
End Sub


