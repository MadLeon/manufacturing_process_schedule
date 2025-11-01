Attribute VB_Name = "Module4"
Sub LookForFile()
Application.ScreenUpdating = False
ActiveCell.Select
ActiveCell.Copy
Workbooks("order entry log.xlsm").Sheets("customers").Range("j1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

ActiveCell.Offset(ColumnOffset:=-2).Activate
ActiveCell.Copy
Workbooks("order entry log.xlsm").Sheets("customers").Range("d1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
ActiveCell.Offset(ColumnOffset:=2).Activate
ActiveCell.Copy

MsgBox ("Paste Your Part # in The Search Box")

   Dim Shex As Object
   Set Shex = CreateObject("Shell.Application")
   'tgtfile = "m:\manufacturing process\"
   tgtfile = (Sheets("customers").Range("k2"))
   Shex.Open (tgtfile)
End Sub


