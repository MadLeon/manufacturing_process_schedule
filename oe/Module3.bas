Attribute VB_Name = "Module3"
Sub Macro3()
Attribute Macro3.VB_Description = "Macro recorded 5/29/2009 by RTD"
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
' Macro recorded 5/29/2009 by RTD
'

'
    Rows("66:66").Select
    Selection.Insert Shift:=xlDown
    Sheets("Format").Select
    Selection.Copy
    Sheets("DELIVERY SCHEDULE").Select
    Range("A66").Select
    Selection.PasteSpecial Paste:=xlAll, Operation:=xlNone, SkipBlanks:=False _
        , Transpose:=False
    Range("A66").Select
End Sub
