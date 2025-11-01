Attribute VB_Name = "Module7"
Sub MPDataXferLMHI()
Application.ScreenUpdating = False
'Open MP Template for new order
'Workbooks.Open ("m:\sue's folder\oe\Order Entry Log.xlsm")
'Go back to OE Log
Workbooks("Order Entry Log.xlsm").Activate
'Copy "Part #" to "Part #" in MP with Value Only
ActiveCell.Select
ActiveCell.Copy
Workbooks("Order Entry Log.xlsm").Sheets("BUSHINGS").Range("j7").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
'Copy "OE #" to "OE #" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=-4).Activate
ActiveCell.Copy
Workbooks("Order Entry Log.xlsm").Sheets("BUSHINGS").Range("n6").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

'Copy "JOB #" to "Job #" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=1).Activate
ActiveCell.Copy
Workbooks("Order Entry Log.xlsm").Sheets("BUSHINGS").Range("q6").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
'Workbooks("Order Entry Log.xlsm").Sheets("BUSHINGS").Range("q6").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
'Copy "Customer" to "Customer" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=1).Activate
ActiveCell.Copy
Workbooks("Order Entry Log.xlsm").Sheets("BUSHINGS").Range("b6").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
'Copy "QTY" to "QTY" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=1).Activate
ActiveCell.Copy
Workbooks("Order Entry Log.xlsm").Sheets("BUSHINGS").Range("q9").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

'Copy "Contact" to "Contact" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=3).Activate
ActiveCell.Copy
Workbooks("Order Entry Log.xlsm").Sheets("BUSHINGS").Range("n9").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

'Copy "Date" to "Date" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=1).Activate
ActiveCell.Copy
Workbooks("Order Entry Log.xlsm").Sheets("BUSHINGS").Range("b8").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
'Copy "Rev" to "Rev" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=-2).Activate
ActiveCell.Copy
Workbooks("Order Entry Log.xlsm").Sheets("BUSHINGS").Range("r7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

'Copy "Ln#" to "Ln #" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=3).Activate
ActiveCell.Copy
Workbooks("Order Entry Log.xlsm").Sheets("BUSHINGS").Range("f7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True


'Copy "Desc" to "Desc" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=1).Activate
ActiveCell.Copy
Workbooks("Order Entry Log.xlsm").Sheets("BUSHINGS").Range("i8").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
'Copy "PO" to "PO" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=2).Activate
ActiveCell.Copy
Workbooks("Order Entry Log.xlsm").Sheets("BUSHINGS").Range("b7").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
'Copy "Del Date" to "Del Date" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=4).Activate
ActiveCell.Copy
Workbooks("Order Entry Log.xlsm").Sheets("BUSHINGS").Range("e9").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

        
'Keep working in MP ...
'Workbooks("Order Entry Log.xlsm").Activate
'Sheets("BUSHINGS").Select
'    Sheets("BUSHINGS").Copy

   
' Back to OE Log
Workbooks("Order Entry Log.xlsm").Activate


End Sub









