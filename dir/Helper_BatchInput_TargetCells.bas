Option Explicit

' =========================================================
' Module: Helper_BatchInput_TargetCells
' Purpose: Find all cells that have the same row distance
'          from nearest header and match column A-E values
' =========================================================

Public Sub FindTargetCells(selectedCell As Range, rowDistance As Long)
    Dim ws As Worksheet
    Dim r As Long
    Dim targetCell As Range
    Dim matchCount As Long
    Dim colCheck As Long
    Dim isMatch As Boolean
    Dim mismatchInfo As String
    
    ' Initialize collection
    Set Data_SharedValues.TargetCells = New Collection
    matchCount = 0
    
    ' Loop through all worksheets except "Data"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Data" Then
            ' Loop through data rows
            For r = 15 To 56
                Set targetCell = ws.Cells(r, "F")
                
                ' Calculate distance to nearest header
                If GetRowDistanceToNearestHeader(targetCell) = rowDistance Then
                    
                    ' Check columns A-E for same properties
                    isMatch = True
                    mismatchInfo = ""
                    For colCheck = 1 To 5
                        If ws.Cells(r, colCheck).Value <> selectedCell.Worksheet.Cells(selectedCell.Row, colCheck).Value Then
                            isMatch = False
                            mismatchInfo = mismatchInfo & _
                                "Col " & Chr(64 + colCheck) & ": Expected = " & _
                                selectedCell.Worksheet.Cells(selectedCell.Row, colCheck).Value & _
                                ", Found = " & ws.Cells(r, colCheck).Value & vbCrLf
                        End If
                    Next colCheck
                    
                    ' If all match, store the target cell
                    If isMatch Then
                        Data_SharedValues.TargetCells.Add targetCell
                        matchCount = matchCount + 1
                        Debug.Print "Target cell #" & matchCount & ": " & ws.Name & "!" & targetCell.Address
                    Else
                        ' Alert user for mismatches in this row
                        Debug.Print "Property mismatch found in sheet '" & ws.Name & "' at row " & r & ":" & vbCrLf & _
                                    mismatchInfo, vbExclamation, "Mismatch Detected"
                    End If
                End If
            Next r
        End If
    Next ws
    
    Debug.Print "Total target cells found: " & matchCount
End Sub


