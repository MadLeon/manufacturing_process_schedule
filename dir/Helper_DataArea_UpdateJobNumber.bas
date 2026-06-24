' =========================================================
' Module: Helper_DataArea_UpdateJobNumber
' =========================================================

' === Update all merged cell prefixes (A-I range) in workbook based on new prefix from H6:I6 ===
Public Sub UpdateJobNumber(newPrefix As String)
    Dim ws As Worksheet
    Dim ma As Range
    Dim cellVal As String
    Dim pos As Long
    
    ' Loop through all worksheets (excluding "Data")
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Data" Then
            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
            
            Dim r As Long
            For r = 15 To lastRow
                Set ma = ws.Cells(r, "A").mergeArea
                ' Only handle merged cells covering A-I columns
                If ma.Columns.count = 9 Then
                    cellVal = ma.Cells(1, 1).Value
                    pos = InStr(cellVal, "-")
                    
                    If pos > 0 Then
                        ' Replace old prefix with new one, keep suffix (number part)
                        ma.Cells(1, 1).Value = newPrefix & Mid(cellVal, pos)
                    Else
                        ' No dash found ? assign new prefix directly
                        ma.Cells(1, 1).Value = newPrefix
                    End If
                    
                    ' Skip to the last row of this merged area to avoid redundant processing
                    r = ma.Rows(ma.Rows.count).Row
                End If
            Next r
        End If
    Next ws
End Sub


