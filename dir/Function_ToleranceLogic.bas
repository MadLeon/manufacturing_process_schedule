' =========================================================
' Module: Function_ToleranceLogic
' =========================================================
' -------------------------
' Helper: Get table start row based on table number
' -------------------------
Private Function GetTableStartRow(tableNum As Long) As Long
    Select Case tableNum
        Case 1: GetTableStartRow = 1
        Case 2: GetTableStartRow = 10
        Case 3: GetTableStartRow = 19
        Case 4: GetTableStartRow = 28
        Case Else: GetTableStartRow = 0  ' Return 0 if invalid table number
    End Select
End Function

' -------------------------
' Helper: Get tolerance cell from LOCAL table (K4:P11) based on value, decimals, tableNum (ignored), and angle
' Logic kept same as original, but operates on the active/local sheet table K4:P11.
' -------------------------
Public Function GetToleranceValue(ByVal dimVal As Variant, _
                                  ByVal decimals As Long, _
                                  ByVal tableNum As Long, _
                                  ByVal isAngle As Boolean, _
                                  Optional ByVal srcCell As Range = Nothing) As Range
    Dim wsTbl As Worksheet
    Dim tblTopRow As Long, tblBottomRow As Long
    Dim rowStart As Long, rowEnd As Long
    Dim i As Long
    Dim overVal As Double, toVal As Double
    Dim numVal As Double
    Dim isFrac As Boolean
    Dim tblRange As Range
    Dim maxRow As Long

    On Error GoTo ErrHandler

    ' Determine worksheet to read table from: prefer srcCell sheet, otherwise use ActiveSheet
    If Not srcCell Is Nothing Then
        Set wsTbl = srcCell.Worksheet
    Else
        Set wsTbl = ThisWorkbook.ActiveSheet
    End If

    ' Set the local table range (fixed)
    Set tblRange = wsTbl.Range("K4:P11")

    ' If table area is empty, return Nothing
    If Application.WorksheetFunction.CountA(tblRange) = 0 Then
        Debug.Print "GetToleranceValue: local table K4:P11 is empty."
        Set GetToleranceValue = Nothing
        Exit Function
    End If

    ' Check if input is fraction (text like "1/2" or "2 1/2")
    isFrac = IsFractionText(CStr(dimVal))

    ' Convert to numeric for comparisons where possible
    On Error Resume Next
    If isFrac Then
        ' Evaluate fraction text to numeric (e.g. "1/2" -> 0.5, "2 1/2" -> 2.5)
        numVal = Evaluate("=" & dimVal)
    Else
        numVal = CDbl(dimVal)
    End If
    On Error GoTo 0

    ' Determine table scanning region
    tblTopRow = tblRange.Row          ' 4
    tblBottomRow = tblRange.Row + tblRange.Rows.count - 1 ' 11
    ' Original algorithm used tableStart + 3 as first data row
    rowStart = tblTopRow + 3
    If rowStart > tblBottomRow Then
        Debug.Print "GetToleranceValue: table too small or unexpected layout."
        Set GetToleranceValue = Nothing
        Exit Function
    End If

    ' Find rowEnd: iterate until blank or "angles" or until bottom of defined table
    rowEnd = rowStart
    Do While rowEnd <= tblBottomRow _
            And Trim(CStr(wsTbl.Cells(rowEnd, "K").Value)) <> "" _
            And LCase(Trim(CStr(wsTbl.Cells(rowEnd, "K").Value))) <> "angles"
        rowEnd = rowEnd + 1
    Loop

    ' Angle: original returned the cell in column I at rowEnd (the "angles" row)
    ' Mapping: old I -> local M
    If isAngle Then
        ' Ensure rowEnd within table area
        If rowEnd <= tblBottomRow Then
            Set GetToleranceValue = wsTbl.Cells(rowEnd, "M")
            Debug.Print "GetToleranceValue: angle -> returning " & GetToleranceValue.Address(False, False, xlA1, True)
            Exit Function
        Else
            Set GetToleranceValue = Nothing
            Exit Function
        End If
    End If

    ' Fraction: search Over/To ranges and return FRAC/WHOLE column (local M)
    If isFrac Then
        For i = rowStart To rowEnd - 1
            ' parse OVER (K) and TO (L), treat "-" or blank as -inf/+inf
            If Trim(wsTbl.Cells(i, "K").Value) = "-" Or Trim(wsTbl.Cells(i, "K").Value) = "" Then
                overVal = -1E+99
            Else
                overVal = CDbl(wsTbl.Cells(i, "K").Value)
            End If

            If Trim(wsTbl.Cells(i, "L").Value) = "-" Or Trim(wsTbl.Cells(i, "L").Value) = "" Then
                toVal = 1E+99
            Else
                toVal = CDbl(wsTbl.Cells(i, "L").Value)
            End If

            If numVal >= overVal And numVal <= toVal Then
                Set GetToleranceValue = wsTbl.Cells(i, "M") ' FRAC OR WHOLE column
                Debug.Print "GetToleranceValue: fraction match at row " & i & " -> " & GetToleranceValue.Address(False, False, xlA1, True)
                Exit Function
            End If
        Next i

        ' Not found
        Set GetToleranceValue = Nothing
        Debug.Print "GetToleranceValue: fraction not found for " & dimVal
        Exit Function
    End If

    ' Numeric: find Over/To row and return column based on decimals
    For i = rowStart To rowEnd - 1
        If Trim(wsTbl.Cells(i, "K").Value) = "-" Or Trim(wsTbl.Cells(i, "K").Value) = "" Then
            overVal = -1E+99
        Else
            overVal = CDbl(wsTbl.Cells(i, "K").Value)
        End If

        If Trim(wsTbl.Cells(i, "L").Value) = "-" Or Trim(wsTbl.Cells(i, "L").Value) = "" Then
            toVal = 1E+99
        Else
            toVal = CDbl(wsTbl.Cells(i, "L").Value)
        End If

        If numVal >= overVal And numVal <= toVal Then
            Select Case decimals
                Case 0
                    Set GetToleranceValue = wsTbl.Cells(i, "M") ' FRAC/WHOLE for 0 dp (original logic)
                Case 1
                    Set GetToleranceValue = wsTbl.Cells(i, "N") ' 1 decimal
                Case 2
                    Set GetToleranceValue = wsTbl.Cells(i, "O") ' 2 decimals
                Case 3
                    Set GetToleranceValue = wsTbl.Cells(i, "P") ' 3 decimals
                Case Else
                    Set GetToleranceValue = wsTbl.Cells(i, "P") ' default to 3 decimals
            End Select
            Debug.Print "GetToleranceValue: numeric match at row " & i & " -> " & GetToleranceValue.Address(False, False, xlA1, True)
            Exit Function
        End If
    Next i

    ' Not found
    Set GetToleranceValue = Nothing
    Debug.Print "GetToleranceValue: numeric not found for " & dimVal

    Exit Function

ErrHandler:
    Debug.Print "GetToleranceValue error: " & Err.Number & " - " & Err.Description
    Set GetToleranceValue = Nothing
End Function


' =========================================================
' Helper: Handle tolerance for B column input, populate D/E columns
' =========================================================
Public Sub HandleTolerance(ws As Worksheet, Target As Range, startRow As Long)
    Dim cell As Range, tolCell As Range
    Dim r As Long, dimVal As Variant
    Dim isAngle As Boolean
    Dim isFracInput As Boolean
    Dim valText As String
    Dim tblRange As Range
    Dim selTable As String
    Dim foundRow As Range
    Dim tolVal As Variant
    Dim overVal As Double, toVal As Double
    Dim numVal As Double
    Dim decimals As Long
    Dim colOffset As Long
    Dim i As Long

    ' -----------------------------
    ' Only process cells in column B (between startRow and 56)
    ' -----------------------------
    Dim changedB As Range
    Set changedB = Intersect(Target, ws.Range("B" & startRow & ":B56"))
    If changedB Is Nothing Then Exit Sub  ' nothing in column B to handle

    ' ========================================
    ' Step 1: Get current table name from L13
    ' ========================================
    selTable = Trim(ws.Range("L13").Value)
    If selTable = "" Then
        If Not Data_SharedValues.AlertToleranceShown Then
            MsgBox "No tolerance table is currently selected." & vbCrLf & _
                   "Tolerance auto-generation will not occur." & vbCrLf & _
                   "If you want this feature, please specify a table in L13.", _
                   vbExclamation, "Tolerance Table Missing"
            Data_SharedValues.AlertToleranceShown = True
        End If
        Exit Sub
    Else
        If Not Data_SharedValues.AlertToleranceShown Then
            MsgBox "Current tolerance table is: " & selTable & vbCrLf & _
                   "Please make sure you are using the correct table.", _
                   vbInformation, "Tolerance Table Selected"
            Data_SharedValues.AlertToleranceShown = True
        End If
    End If

    ' ========================================
    ' Step 2: Set the tolerance table range (current sheet)
    ' ========================================
    Set tblRange = ws.Range("K4:P11")
    If Application.WorksheetFunction.CountA(tblRange) = 0 Then
        Debug.Print "No data found in tolerance table area."
        Exit Sub
    End If

    ' ========================================
    ' Step 3: Loop through B-column inputs (only the changed B cells)
    ' ========================================
    For Each cell In changedB
        r = cell.Row
        ' boundary safety (should be redundant after Intersect check)
        If r < startRow Or r > 56 Then GoTo NextCell

        If Trim(cell.Value) = "" Then
            ws.Cells(r, "D").ClearContents
            ws.Cells(r, "E").ClearContents
            GoTo NextCell
        End If

        valText = Trim(CStr(cell.Value))

        ' Handle special keywords (e.g. UNC/UNF)
        If IsSpecialKeyword(valText) Then
            Debug.Print "Special keyword detected at row " & r & ": " & valText
            ws.Cells(r, "F").Value = cell.Value
            ws.Cells(r, "D").ClearContents
            ws.Cells(r, "E").ClearContents
            ws.Cells(r, "G").ClearContents
            GoTo NextCell
        End If

        ' Check for fraction input
        isFracInput = IsFractionText(valText)

        ' Check for angle (column A contains "°")
        isAngle = (InStr(1, CStr(ws.Cells(r, "A").Value), "°") > 0)

        ' ========================================
        ' Step 4: Find corresponding tolerance from current table (using local GetToleranceValue)
        ' ========================================
        On Error Resume Next
        If isFracInput Then
            numVal = Evaluate("=" & valText)
        ElseIf IsNumeric(valText) Then
            numVal = CDbl(valText)
        Else
            numVal = 0
        End If
        On Error GoTo 0

        decimals = CountDecimalPlaces(valText)

        Set tolCell = GetToleranceValue(numVal, decimals, 0, isAngle, cell)

        If tolCell Is Nothing Then
            Debug.Print "No tolerance match found for row " & r & ", value=" & valText
            ws.Cells(r, "D").ClearContents
            ws.Cells(r, "E").ClearContents
        Else
            ws.Cells(r, "D").Value = tolCell.Value
            ws.Cells(r, "E").Value = "-" & tolCell.Value
            Debug.Print "Tolerance applied for row " & r & ": +" & tolCell.Value & " / -" & tolCell.Value
        End If

NextCell:
    Next cell
End Sub


