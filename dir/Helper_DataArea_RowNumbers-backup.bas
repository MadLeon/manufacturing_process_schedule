' =========================================================
' Module: Helper_DataArea_RowNumbers
' =========================================================

Option Explicit

' === Update all merged cell numbers in column A across the entire workbook ===
Public Sub UpdateMergedRowNumbersWorkbook()
    Dim ws As Worksheet
    Dim mergedAreas As Collection
    Set mergedAreas = New Collection   ' Used to store all found merged areas
    
    Dim cell As Range, ma As Range
    Dim prefix As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary") ' Dictionary to track numbering per prefix
    
    ' === Loop through all worksheets in the workbook ===
    For Each ws In ThisWorkbook.Worksheets
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row  ' Find last used row in column A
        
        Dim r As Long
        For r = 15 To lastRow
            Set cell = ws.Cells(r, "A")
            If cell.MergeCells Then
                On Error Resume Next
                ' Add merged range to collection, key = sheet name + range address
                mergedAreas.Add cell.mergeArea, ws.Name & "!" & cell.mergeArea.Address
                On Error GoTo 0
            End If
        Next r
    Next ws
    
    ' === Loop through all collected merged ranges and update numbering ===
    For Each ma In mergedAreas
        ' Get parent worksheet of the merged area
        Dim parentWs As Worksheet
        Set parentWs = ma.Worksheet
        
        ' Check if range H9:I9 is empty
        Dim skipNumbering As Boolean
        skipNumbering = Application.WorksheetFunction.CountA(parentWs.Range("H9:I9")) = 0
        
        ' Get prefix from the merged cell (text before "-")
        If InStr(ma.Cells(1, 1).Value, "-") > 0 Then
            prefix = Left(ma.Cells(1, 1).Value, InStrRev(ma.Cells(1, 1).Value, "-") - 1)
        Else
            prefix = ma.Cells(1, 1).Value
        End If
        
        If skipNumbering Then
            ' If H9:I9 is empty ? keep only prefix (no numbering)
            ma.Cells(1, 1).Value = prefix
        Else
            ' Normal numbering process
            If dict.Exists(prefix) Then
                dict(prefix) = dict(prefix) + 1
            Else
                dict(prefix) = 1
            End If
            ' Assign prefix-number format like "ABC-1", "ABC-2", etc.
            ma.Cells(1, 1).Value = prefix & "-" & dict(prefix)
        End If
    Next ma
End Sub