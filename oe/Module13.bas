Attribute VB_Name = "Module13"
Sub Hide_GKMNOQR()
Attribute Hide_GKMNOQR.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Hide_GKMNOQR Macro
'

'
    Columns("G:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("K:K").Select
    Selection.EntireColumn.Hidden = True
    Columns("M:O").Select
    Selection.EntireColumn.Hidden = True
    Columns("Q:R").Select
    Selection.EntireColumn.Hidden = True
    
        ' This  select "Customer" = "Kinectrics" to be filtered
    
    ActiveSheet.Range("$A$3:$R$51965").AutoFilter Field:=3, Criteria1:= _
        "Kinectrics"

End Sub
Sub Unhide_GKMNOQR()
Attribute Unhide_GKMNOQR.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Unhide_GKMNOQR Macro
'

'
    Columns("A:T").Select
    Selection.EntireColumn.Hidden = False
    
    'This Un-filter everything
        ActiveSheet.ShowAllData

End Sub
