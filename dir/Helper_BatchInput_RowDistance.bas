' =========================================================
' Module: Helper_BatchInput_RowDistance
' Returns the row distance between the selected cell and its nearest merged header
' =========================================================
Public Function GetRowDistanceToNearestHeader(selectedCell As Range) As Long
    Dim nearestHeader As Class_MergedHeaderInfo
    Dim candidate As Class_MergedHeaderInfo
    Dim selSheetIndex As Long, candidateSheetIndex As Long
    Dim distance As Long, minDistance As Long
    Dim wsIndex As Long

    If Data_SharedValues.MergedHeaders Is Nothing Then
        MsgBox "MergedHeaders collection is not initialized!", vbExclamation
        GetRowDistanceToNearestHeader = -1
        Exit Function
    End If

    minDistance = 999
    selSheetIndex = selectedCell.Worksheet.Index

    For Each candidate In Data_SharedValues.MergedHeaders
        candidateSheetIndex = Worksheets(candidate.SheetName).Index

        ' Only consider headers above or in previous sheets
        If candidateSheetIndex < selSheetIndex Or _
           (candidateSheetIndex = selSheetIndex And candidate.rowNum <= selectedCell.Row) Then

            distance = 0

            ' Add full data rows for all sheets strictly between candidate sheet and selected sheet
            For wsIndex = candidateSheetIndex + 1 To selSheetIndex - 1
                distance = distance + (56 - Data_SharedValues.startRow + 1)
            Next wsIndex

            ' Calculate distance from candidate header
            If candidateSheetIndex < selSheetIndex Then
                ' Header is on previous sheet
                distance = distance + (56 - candidate.rowNum)           ' rows remaining in header sheet
                distance = distance + (selectedCell.Row - 15 + 1)       ' rows in current sheet relative to previous header
            Else
                ' Header is in the same sheet
                distance = selectedCell.Row - candidate.rowNum
            End If

            ' Keep nearest header
            If distance < minDistance Then
                minDistance = distance
                Set nearestHeader = candidate
            End If
        End If
    Next candidate

    GetRowDistanceToNearestHeader = minDistance
End Function



