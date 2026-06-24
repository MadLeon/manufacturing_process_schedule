' =============================================
' Validates that the selected cell is in column F and within data area
' =============================================
Public Function ValidateSelectedCell(selectedCell As Range) As Boolean
    Dim selRow As Long
    Dim selCol As Long
    Dim endRow As Long
    
    selRow = selectedCell.Row
    selCol = selectedCell.Column
    
    ' Must be column F
    If selCol <> 6 Then
        MsgBox "Please select a cell in column F.", vbExclamation
        ValidateSelectedCell = False
        Exit Function
    End If
    
    ' Ensure StartRow is set
    If Data_SharedValues.startRow = 0 Then
        Helper_Database_InitStartRow.InitializeStartRow ActiveSheet
    End If
    endRow = 56
    
    If selRow < Data_SharedValues.startRow Or selRow > endRow Then
        MsgBox "Selected cell must be within the data area (rows " & _
               Data_SharedValues.startRow & "–" & endRow & ").", vbExclamation
        ValidateSelectedCell = False
        Exit Function
    End If
    
    ValidateSelectedCell = True
End Function

' =========================================================
' ValidateMergedCells: Find merged header cells and store info
' =========================================================
Public Function ValidateMergedCells(expectedCount As Long) As Boolean
    Dim foundCount As Long
    
    ' Initialize merged headers
    Call InitializeMergedHeaders
    
    ' Count the found headers
    foundCount = Data_SharedValues.MergedHeaders.count
    
    ' Validation check
    If foundCount <> expectedCount Then
        MsgBox "Merged cell count mismatch!" & vbCrLf & _
               "Expected: " & expectedCount & vbCrLf & _
               "Found: " & foundCount, vbExclamation
        ValidateMergedCells = False
    Else
        ValidateMergedCells = True
    End If
End Function

