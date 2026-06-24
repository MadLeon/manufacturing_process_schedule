' =========================================================
' Module: Helper_UpdatePageNumbers
' =========================================================

' Update page numbers in cell I5 on each worksheet.
' Example: If previously "Page 1 of 1", after update it becomes "Page 1 of 2", "Page 2 of 2", etc.
Public Sub UpdatePageNumbers()
    Dim ws As Worksheet
    Dim totalSheets As Integer
    Dim count As Integer

    ' Count all worksheets except the one named "Data"
    totalSheets = 0
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Data" Then totalSheets = totalSheets + 1
    Next ws

    ' Update each sheet's page number in I5
    count = 1
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Data" Then
            With ws.Range("I5")
                ' Format: "Page x of y"
                .Value = "Page " & count & " of " & totalSheets
                ' Align the text to the right
                .HorizontalAlignment = xlRight
            End With
            count = count + 1
        End If
    Next ws
End Sub