Attribute VB_Name = "Module10"
Sub SaveWB()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'alertTime = Now + TimeValue("00:02:00")
'Application.OnTime alertTime, "EventMacro"


ThisWorkbook.Save
'Application.Wait Now + TimeValue("00:0:10") , "SaveWB" (This step is only for Regularly savings)
End Sub

