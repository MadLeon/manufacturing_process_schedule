Attribute VB_Name = "Module11"
Sub MPDataXferL()
Application.ScreenUpdating = False
'Open MP Template for new order
Workbooks.Open ("\\rtdnas2\oe\Manufacturing Processl.xls")
'Go back to OE Log
Workbooks("Order Entry Log.xlsm").Activate
'Copy "Part #" to "Part #" in MP with Value Only
ActiveCell.Select
ActiveCell.Copy
Workbooks("Manufacturing Processl.xls").Sheets("sheet1").Range("j7").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
'Copy "OE #" to "OE #" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=-4).Activate
ActiveCell.Copy
Workbooks("Manufacturing Processl.xls").Sheets("sheet1").Range("n6").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

'Copy "JOB #" to "Job #" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=1).Activate
ActiveCell.Copy
Workbooks("Manufacturing Processl.xls").Sheets("sheet1").Range("q6").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
'Workbooks("Manufacturing Processl.xls").Sheets("sheet1").Range("q6").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
'Copy "Customer" to "Customer" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=1).Activate
ActiveCell.Copy
Workbooks("Manufacturing Processl.xls").Sheets("sheet1").Range("b6").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
'Copy "QTY" to "QTY" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=1).Activate
ActiveCell.Copy
Workbooks("Manufacturing Processl.xls").Sheets("sheet1").Range("q9").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

'Copy "Contact" to "Contact" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=3).Activate
ActiveCell.Copy
Workbooks("Manufacturing Processl.xls").Sheets("sheet1").Range("n9").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

'Copy "Date" to "Date" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=1).Activate
ActiveCell.Copy
Workbooks("Manufacturing Processl.xls").Sheets("sheet1").Range("b8").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
'Copy "Rev" to "Rev" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=-2).Activate
ActiveCell.Copy
Workbooks("Manufacturing Processl.xls").Sheets("sheet1").Range("r7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

'Copy "Ln#" to "Ln #" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=3).Activate
ActiveCell.Copy
Workbooks("Manufacturing Processl.xls").Sheets("sheet1").Range("f7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True


'Copy "Desc" to "Desc" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=1).Activate
ActiveCell.Copy
Workbooks("Manufacturing Processl.xls").Sheets("sheet1").Range("i8").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
'Copy "PO" to "PO" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=2).Activate
ActiveCell.Copy
Workbooks("Manufacturing Processl.xls").Sheets("sheet1").Range("b7").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
'Copy "Del Date" to "Del Date" in MP with Value Only
ActiveCell.Offset(ColumnOffset:=4).Activate
ActiveCell.Copy
Workbooks("Manufacturing Processl.xls").Sheets("sheet1").Range("e9").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

        
'Keep working in MP ...
'Workbooks("Manufacturing Processl.xls").Activate
'Sheets("Sheet1").Select
'    Sheets("Sheet1").Copy

   
' Back to OE Log
Workbooks("Order Entry Log.xlsm").Activate


End Sub







