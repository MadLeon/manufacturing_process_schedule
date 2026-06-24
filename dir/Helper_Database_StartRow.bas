' =============================================
' Initializes the StartRow global variable
' Logic: find first merged cell in column A, data starts below it
' If nothing found, fallback to 15
' =============================================
Public Sub InitializeStartRow(ws As Worksheet)
    Dim r As Long
    
    For r = 15 To 56
        If ws.Cells(r, "A").MergeCells Then
            Data_SharedValues.startRow = r + 1
            ' Debug.Print "InitializeStartRow: first merged header at row " & r & _
                        ", Data_SharedValues.StartRow set to " & Data_SharedValues.StartRow
            Exit Sub
        End If
    Next r
    
    ' Fallback
    Data_SharedValues.startRow = 15
    ' Debug.Print "InitializeStartRow: no merged header found, fallback StartRow = 15"
End Sub

