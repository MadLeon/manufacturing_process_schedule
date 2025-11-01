Attribute VB_Name = "Module12"
Dim TimerActive As Boolean
Sub StartTimer()
    Start_Timer
End Sub

Private Sub Start_Timer()
    TimerActive = True
    Application.OnTime Now() + TimeValue("00:01:00"), "Timer"
End Sub

Private Sub Stop_Timer()
    TimerActive = False
End Sub

Private Sub Timer()
    If TimerActive Then
       ' ActiveSheet.Cells(1, 1).Value = Time
        Application.OnTime Now() + TimeValue("00:01:00"), "SaveWB"
    End If
End Sub
