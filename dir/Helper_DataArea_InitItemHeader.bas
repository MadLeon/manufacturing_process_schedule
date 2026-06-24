' =========================================================
' Module: Helper_InitializeMergedHeader
' Purpose: Find and store merged headers across worksheets
' =========================================================
Option Explicit

Public Sub InitializeMergedHeaders()
    Dim ws As Worksheet, ma As Range
    Dim headerInfo As Class_MergedHeaderInfo
    Dim r As Long
    Dim foundCount As Long
    
    ' Reset collection
    Set Data_SharedValues.MergedHeaders = New Collection
    foundCount = 0
    
    ' Scan all worksheets except "Data"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Data" Then
            For r = 15 To 56
                If ws.Cells(r, "A").MergeCells Then
                    Set ma = ws.Cells(r, "A").mergeArea
                    If ma.Columns.count >= 9 Then
                        foundCount = foundCount + 1
                        
                        ' Populate class and add to collection
                        Set headerInfo = New Class_MergedHeaderInfo
                        headerInfo.SheetName = ws.Name
                        headerInfo.rowNum = r
                        headerInfo.RangeAddress = ma.Address
                        Data_SharedValues.MergedHeaders.Add headerInfo
                    End If
                End If
            Next r
        End If
    Next ws
End Sub

