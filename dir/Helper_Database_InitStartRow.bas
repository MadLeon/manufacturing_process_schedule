' =========================================================
' Module: Helper_Database_InitStartRow
' Purpose: Determine the starting row of the data table
'          based on whether row 15 is merged.
' =========================================================
Option Explicit

Public Sub InitializeStartRow(ws As Worksheet)
    ' -------------------------
    ' Determine start row (row 15 merged or not)
    ' -------------------------
    If ws.Range("A15:I15").MergeCells Then
        Data_SharedValues.startRow = 16
    Else
        Data_SharedValues.startRow = 15
    End If
    
    Debug.Print "InitializeStartRow finished! Data_SharedValues.startRow = " & Data_SharedValues.startRow
End Sub

