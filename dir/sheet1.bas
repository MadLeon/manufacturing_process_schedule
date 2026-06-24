Option Explicit

' =========================================================
' Worksheet_Change Event Handler
' =========================================================
' This macro automatically reacts to changes in specific
' ranges on the worksheet and performs the following:
' - File name parsing and dropdown removal
' - Unit selection handling
' - Signature selection handling
' - Decimal place matching and tolerance checking
' - Formula logic updates based on B/C column changes
' - Tool logic and dependent dropdown updates
' - Data lookup replacements (I column)
' - Header synchronization across sheets
' =========================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = Me
    Dim cell As Range, iCell As Range
    Dim r As Long, lastRow As Long
    Dim selectedVal As String
    Dim dataSheet As Worksheet
    Dim syncRanges As Variant
    Dim syncCell As Variant
    Dim sh As Worksheet

    On Error GoTo CleanExit
    Application.EnableEvents = False
    
' -------------------------
' L13 dropdown change - display corresponding table
' -------------------------
If Not Intersect(Target, ws.Range("L13")) Is Nothing Then
    Set dataSheet = ThisWorkbook.Sheets("Data")
    
    ' Clear previous display area
    ws.Range("K4:P11").Clear
    
    ' Copy the corresponding table
    Select Case ws.Range("L13").Value
        Case "ATS"
            dataSheet.Range("G1:L8").Copy
        Case "CANDU - DETAILED"
            dataSheet.Range("G10:L17").Copy
        Case "CANDU - RAW"
            dataSheet.Range("G19:L26").Copy
        Case "CANDU - GENERAL"
            dataSheet.Range("G28:L35").Copy
        Case "ASME Y-14.5 (2009)"
            dataSheet.Range("G37:L44").Copy
        Case "KINECTRICS - MM"
            dataSheet.Range("G46:L53").Copy
        Case "CANDU - MM"
            dataSheet.Range("G55:L62").Copy
        Case Else
            Debug.Print "Invalid selection in L13: " & ws.Range("L13").Value
            GoTo SkipPaste
    End Select
    
    ' Paste values and formatting (including merged cells)
    With ws.Range("K4")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
    End With
    Application.CutCopyMode = False
    
    Debug.Print "Table displayed for selection: " & ws.Range("L13").Value

SkipPaste:
End If

    ' -------------------------
    ' File name parsing (C6:F6)
    ' -------------------------
    If Not Intersect(Target, ws.Range("C6:F6")) Is Nothing Then
        Call HandleFileNameParsing(ws)
        On Error Resume Next
        ws.Range("C6").Validation.Delete
        On Error GoTo 0
    End If

    ' -------------------------
    ' Unit selection change (C11:F11)
    ' -------------------------
    If Not Intersect(Target, ws.Range("C11:F11")) Is Nothing Then
        On Error Resume Next
        ws.Range("C11").Validation.Delete
        On Error GoTo 0
    End If

    ' -------------------------
    ' Signature selection (H57)
    ' -------------------------
    If Not Intersect(Target, ws.Range("H57")) Is Nothing Then
        Call HandleSignatureSelection(ws, Target)
    End If

    ' -------------------------
    ' Column F changes (F15:F56)
    ' Match decimal places and consistency
    ' -------------------------
    If Not Intersect(Target, ws.Range("F15:F56")) Is Nothing Then
        Call MatchDecimalPlaces
        Call CheckFDecimalConsistency(ws, Target)
    End If

    ' -------------------------
    ' Columns B/C changes (B:startRow–C56)
    ' Trigger formula logic
    ' -------------------------
    If Not Intersect(Target, ws.Range("B" & Data_SharedValues.startRow & ":C56")) Is Nothing Then
        Call HandleFormulaLogic(ws, Target, Data_SharedValues.startRow)
    End If

    ' -------------------------
    ' Column B changes (B:startRow–B56)
    ' Tolerance + tool logic
    ' -------------------------
    If Not Intersect(Target, ws.Range("B" & Data_SharedValues.startRow & ":B56")) Is Nothing Then
        ' Temporarily isolate events for HandleTolerance
        On Error GoTo ToleranceError
        Application.EnableEvents = False
        Call HandleTolerance(ws, Target, Data_SharedValues.startRow)
        Application.EnableEvents = True

        ' Run tool logic (these calls may write to sheet, but are safer)
        For Each cell In Intersect(Target, ws.Range("B" & Data_SharedValues.startRow & ":B56"))
            Call HandleToolLogic(ws, cell.Row, cell.Value)
        Next cell
    End If

ContinueAfterTolerance:
    On Error GoTo CleanExit

    ' -------------------------
    ' Column I changes (I:startRow–I56)
    ' Lookup values from Data sheet (P:R)
    ' -------------------------
    If Not Intersect(Target, ws.Range("I" & Data_SharedValues.startRow & ":I56")) Is Nothing Then
        For Each iCell In Intersect(Target, ws.Range("I" & Data_SharedValues.startRow & ":I56"))
            selectedVal = iCell.Value

            If selectedVal <> "" Then
                Set dataSheet = ThisWorkbook.Sheets("Data")
                lastRow = dataSheet.Cells(dataSheet.Rows.count, "P").End(xlUp).Row

                For r = 2 To lastRow
                    If dataSheet.Cells(r, "P").Value = selectedVal Then
                        Application.EnableEvents = False
                        iCell.Value = dataSheet.Cells(r, "R").Value
                        Application.EnableEvents = True
                        Exit For
                    End If
                Next r
            End If

            ' Remove validation to avoid repeated trigger
            On Error Resume Next
            iCell.Validation.Delete
            On Error GoTo 0
        Next iCell
    End If

    ' -------------------------
    ' Column A changes (A15:A56)
    ' Remove validation after selection
    ' Call module to mark "°" below if value contains "C"
    ' -------------------------
    If Not Intersect(Target, ws.Range("A" & Data_SharedValues.startRow & ":A56")) Is Nothing Then
        For Each cell In Intersect(Target, ws.Range("A" & Data_SharedValues.startRow & ":A56"))
            If cell.Value <> "" Then
                ' Remove validation
                On Error Resume Next
                cell.Validation.Delete
                On Error GoTo 0
                
                ' Call module procedure
                MarkDegreeBelow ws, cell, Data_SharedValues.startRow, 56
            End If
        Next cell
    End If

    ' -------------------------
    ' Header changes H9:I9
    ' Handle item number update
    ' -------------------------
    If Not Intersect(Target, ws.Range("H9:I9")) Is Nothing Then
        Call HandleItemNumber(ws)
        Helper_Database_InitStartRow.InitializeStartRow ws
    End If

    ' -------------------------
    ' Synchronize header ranges across all sheets
    ' H6:I6 ~ H10:I10
    ' -------------------------
    syncRanges = Array("H6:I6", "H7:I7", "H8:I8", "H9:I9", "H10:I10", _
                   "C6:F6", "C7:F7", "C8:F8", "C9:F9", "C10:F10", "C11:F11", _
                   "F59")
    
    For Each syncCell In syncRanges
        If Not Intersect(Target, ws.Range(syncCell)) Is Nothing Then
            ' Copy values to all sheets except Data and current
            For Each sh In ThisWorkbook.Worksheets
                If sh.Name <> "Data" And sh.Name <> ws.Name Then
                    sh.Range(syncCell).Value = ws.Range(syncCell).Value
                    
                    ' Special case: if H9:I9 changed, also update H10:I10
                    If syncCell = "H9:I9" Then
                        sh.Range("H10:I10").Value = ws.Range("H10:I10").Value
                    End If
                End If
            Next sh
    
            ' If H6:I6 changed, update Job Number
            If syncCell = "H6:I6" Then
                Call UpdateJobNumber(ws.Range("H6").Value)
            End If
        End If
    Next syncCell

CleanExit:
    Application.EnableEvents = True
    Exit Sub

ToleranceError:
    Debug.Print "HandleTolerance error: " & Err.Description
    Resume ContinueAfterTolerance
End Sub

