' =========================================================
' Module: Helper_DataArea
' =========================================================

' -------------------------
' Helper: Check if value should be treated as special (requires measurement tool)
' -------------------------
Public Function IsSpecialKeyword(bValue As String) As Boolean
    Dim val As String
    val = Trim(bValue)
    
    ' If the value is empty, it is not special
    If val = "" Then
        IsSpecialKeyword = False
        Exit Function
    End If
    
    ' If the value is numeric (integer or decimal), it is not special
    If IsNumeric(val) Then
        IsSpecialKeyword = False
    ' If the value is a fraction (e.g., "1/2" or "2 1/4"), it is not special
    ElseIf IsFractionText(val) Then
        IsSpecialKeyword = False
    Else
        ' Any other text (e.g., thread/complex format like "7/8-9 UNC-2B") is special
        ' This indicates a measurement tool should be used
        IsSpecialKeyword = True
    End If
End Function


' -------------------------
' Helper: Match decimal places of column F with column G
' -------------------------
Sub MatchDecimalPlaces()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim startRow As Long
    Dim i As Long
    Dim fVal As Variant
    Dim decimalPlaces As Long
    Dim numberFormatStr As String
    
    ' Determine starting row
    If ws.Range("A15:I15").MergeCells Then
        startRow = 16
    Else
        startRow = 15
    End If
    
    ' Loop through rows to format column G based on F
    For i = startRow To 56
        fVal = ws.Cells(i, 6).Value  ' Column F
        decimalPlaces = 0
        
        ' Count decimal places
        If InStr(1, CStr(fVal), ".") > 0 Then
            decimalPlaces = Len(Split(CStr(fVal), ".")(1))
        End If
        
        ' Build number format string
        numberFormatStr = "0"
        If decimalPlaces > 0 Then numberFormatStr = numberFormatStr & "." & String(decimalPlaces, "0")
        
        ws.Cells(i, 7).NumberFormat = numberFormatStr ' Column G
    Next i
End Sub

' -------------------------
' Helper: Check decimal consistency across all sheets
' Uses: ValidateMergedCells, GetRowDistanceToNearestHeader, FindTargetCells
' -------------------------
Sub CheckFDecimalConsistency(ws As Worksheet, Target As Range)
    Dim changedCell As Range
    Dim rowDistance As Long
    Dim i As Long
    Dim targetCol As Long
    Dim tc As Range
    Dim standardDecimals As Long
    Dim foundStandard As Boolean
    Dim thisDecimals As Long
    Dim msgDebug As String

    On Error GoTo CleanUp

    ' Ensure MergedHeaders is valid and matches expected structure
    Dim expectedCount As Long
    expectedCount = ws.Range("H9").Value

    ' Ensure MergedHeaders collection exists — initialize if missing
    If Data_SharedValues.MergedHeaders Is Nothing Then
        InitializeMergedHeaders
    End If

    ' Loop through each changed cell in the Target range
    For Each changedCell In Target.Cells
        ' Only process cells within the data area (startRow..56) and in the target column
        If changedCell.Row < Data_SharedValues.startRow Or changedCell.Row > 56 Then GoTo NextChangedCell

        ' Compute distance to nearest header for this changed cell
        rowDistance = GetRowDistanceToNearestHeader(changedCell)
        Debug.Print "CheckFDecimalConsistency: Changed cell: " & changedCell.Worksheet.Name & "!" & changedCell.Address & " | Distance: " & rowDistance

        ' Find all target cells that share same distance and A-E properties
        FindTargetCells changedCell, rowDistance

        ' If no target cells found, skip
        If Data_SharedValues.TargetCells Is Nothing Then
            Debug.Print "  No TargetCells collection created for this distance."
            GoTo NextChangedCell
        End If

        If Data_SharedValues.TargetCells.count = 0 Then
            Debug.Print "  No target cells matched for this distance."
            GoTo NextChangedCell
        End If

        ' Determine decimal standard from first non-empty target cell (use cell.Text to preserve formatting)
        foundStandard = False
        standardDecimals = 0

        For i = 1 To Data_SharedValues.TargetCells.count
            Set tc = Data_SharedValues.TargetCells(i)
            ' Skip truly empty cells
            If Trim(CStr(tc.Value)) = "" Then
                ' clear any previous highlight (we don't treat empty as mismatch)
                tc.Interior.ColorIndex = xlNone
            Else
                ' Use the displayed text to count decimal places (preserves trailing zeros)
                standardDecimals = CountDecimalPlaces(CStr(tc.Text))
                foundStandard = True
                Debug.Print "  Standard chosen from: " & tc.Worksheet.Name & "!" & tc.Address & " | Text='" & tc.Text & "' | Decimals=" & standardDecimals
                Exit For
            End If
        Next i

        ' If no non-empty targets found, nothing to enforce
        If Not foundStandard Then
            Debug.Print "  No non-empty target cells to determine standard decimals. Skipping."
            GoTo NextChangedCell
        End If

        ' Compare other target cells against the standard
        For i = 1 To Data_SharedValues.TargetCells.count
            Set tc = Data_SharedValues.TargetCells(i)

            If Trim(CStr(tc.Value)) = "" Then
                ' Skip empties; ensure no highlight
                tc.Interior.ColorIndex = xlNone
            Else
                ' Count decimals using displayed text
                thisDecimals = CountDecimalPlaces(CStr(tc.Text))

                If thisDecimals <> standardDecimals Then
                    ' Highlight mismatch with cyan
                    tc.Interior.Color = RGB(0, 255, 255)
                    Debug.Print "  Highlighted mismatch: " & tc.Worksheet.Name & "!" & tc.Address & " | Text='" & tc.Text & "' | Decimals=" & thisDecimals & " (expected " & standardDecimals & ")"
                Else
                    ' Clear highlight if previously set
                    tc.Interior.ColorIndex = xlNone
                    Debug.Print "  OK: " & tc.Worksheet.Name & "!" & tc.Address & " | Decimals=" & thisDecimals
                End If
            End If
        Next i

NextChangedCell:
    Next changedCell

CleanUp:
    Exit Sub
End Sub


' -------------------------
' Helper: Apply formula logic based on B, C, D, E, F values
' -------------------------
Public Sub HandleFormulaLogic(ws As Worksheet, Target As Range, startRow As Long)
    Dim cell As Range, r As Long
    Dim kw As Variant, isSpecial As Boolean
    
    For Each cell In Target
        r = cell.Row
        If r >= startRow And r <= 56 Then
            Dim cellB As Range, cellC As Range, cellG As Range, cellH As Range
            Set cellB = ws.Cells(r, "B")
            Set cellC = ws.Cells(r, "C")
            Set cellG = ws.Cells(r, "G")
            Set cellH = ws.Cells(r, "H")
            
            ' Check if B contains special keyword
            isSpecial = IsSpecialKeyword(cellB.Value)
            
            ' Apply formulas based on conditions
            If cellB.Value = "" And cellC.Value = "" Then
                If cellG.HasFormula Then cellG.ClearContents
                If cellH.HasFormula Then cellH.ClearContents
            Else
                If Not isSpecial Then
                    If cellG.Value = "" And Not cellG.HasFormula Then
                        cellG.Formula = "=IF((F" & r & "<>""""),(ToNumber(F" & r & ")-ToNumber(B" & r & ")),"""")"
                    End If
                Else
                    If cellG.HasFormula Or cellG.Value <> "" Then cellG.ClearContents
                End If
                If cellH.Value = "" And Not cellH.HasFormula Then
                    cellH.Formula = "=IF(AND((ToNumber(D" & r & ")+0.00001)>=G" & r & ",(ToNumber(E" & r & ")-0.00001)<=G" & r & "),""A"",""R"")"
                End If
            End If
        End If
    Next cell
End Sub

' -------------------------
' Helper: Handle formula logic for GDT entries
' -------------------------
Public Sub HandleGDTFormula(ws As Worksheet, targetRow As Long)
    Dim r As Long
    r = targetRow
    
    Dim cellD As Range, cellF As Range, cellG As Range, cellH As Range
    Set cellD = ws.Cells(r, "D")
    Set cellF = ws.Cells(r, "F")
    Set cellG = ws.Cells(r, "G")
    Set cellH = ws.Cells(r, "H")
    
    ' Clear previous formulas
    If cellG.HasFormula Then cellG.ClearContents
    If cellH.HasFormula Then cellH.ClearContents
    
    ' Formula for G: difference between F and B (template 1)
    If Not cellG.HasFormula Then
        cellG.Formula = "=IF((F" & r & "<>""""),(ToNumber(F" & r & ")-ToNumber(B" & r & ")),"""")"
    End If
    
    ' Formula for H: check if G is between D and E (template 2)
    If Not cellH.HasFormula Then
        cellH.Formula = "=IF(AND((ToNumber(D" & r & ")+0.00001)>=G" & r & ",(ToNumber(E" & r & ")-0.00001)<=G" & r & "),""A"",""R"")"
    End If

End Sub


' -------------------------
' Helper: Handle tool logic based on B value, update I column
' -------------------------
Public Sub HandleToolLogic(ws As Worksheet, targetRow As Long, bValue As String)
    Dim iCell As Range
    Set iCell = ws.Cells(targetRow, "I")
    
    Dim dataSheet As Worksheet
    Set dataSheet = ThisWorkbook.Sheets("Data")
    
    ' Clear validation if B is blank
    If Trim(bValue) = "" Then
        On Error Resume Next
        iCell.Validation.Delete
        iCell.ClearContents
        On Error GoTo 0
        Exit Sub
    End If
    
    On Error Resume Next
    iCell.Validation.Delete
    On Error GoTo 0
    
    ' Standard vs special tool dropdown
    If Not IsSpecialKeyword(bValue) Then
        ' Standard dropdown (N2:N9)
        With iCell.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Formula1:="='" & dataSheet.Name & "'!N2:N15"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    Else
        ' Special dropdown from Data!P
        Dim lastRow As Long, r As Long, matchList As String
        matchList = ""
        lastRow = dataSheet.Cells(dataSheet.Rows.count, "P").End(xlUp).Row
        For r = 2 To lastRow
            If InStr(1, dataSheet.Cells(r, "P").Value, bValue) > 0 Then
                If matchList = "" Then
                    matchList = dataSheet.Cells(r, "P").Value
                Else
                    matchList = matchList & "," & dataSheet.Cells(r, "P").Value
                End If
            End If
        Next r
        
        If matchList <> "" Then
            With iCell.Validation
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Formula1:=matchList
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        End If
    End If
End Sub

Public Sub HandleGDTTool(ws As Worksheet, targetRow As Long)
    Dim iCell As Range
    Set iCell = ws.Cells(targetRow, "I")
    
    Dim dataSheet As Worksheet
    Set dataSheet = ThisWorkbook.Sheets("Data")
    
    ' Clear any existing validation first
    On Error Resume Next
    iCell.Validation.Delete
    On Error GoTo 0
    
    ' Always use standard dropdown (Data!N2:N9)
    With iCell.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="='" & dataSheet.Name & "'!N2:N9"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub


' -------------------------
' Column A changes (A15:A56)
' Add "°" below if value contains "C"
' -------------------------
Public Sub MarkDegreeBelow(ws As Worksheet, changedCell As Range, startRow As Long, lastRow As Long)
    ' ws: worksheet object
    ' changedCell: the single cell that changed
    ' startRow: starting data row (usually 15 or 16)
    ' lastRow: last row in the column (56)
    
    ' Only operate if the changed cell is in column A
    If changedCell.Column <> 1 Then Exit Sub
    
    ' Check if the value contains "C"
    If InStr(1, changedCell.Value, "C", vbTextCompare) > 0 Or _
       InStr(1, changedCell.Value, ChrW(&H2335) & "Ø", vbTextCompare) > 0 Then
        ' Check if the cell below is within the allowed range
        If changedCell.Row + 1 <= lastRow Then
            ws.Cells(changedCell.Row + 1, "A").Value = "°"
        End If
    End If
End Sub


