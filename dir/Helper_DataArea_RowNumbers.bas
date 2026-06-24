' =========================================================
' Module: Helper_DataArea_RowNumbers
' =========================================================

Option Explicit

' === Update all merged cell numbers in column A across the entire workbook ===
' Preserves the first merged row's existing format and auto-numbers the rest
Public Sub UpdateMergedRowNumbersWorkbook()
    Dim ws As Worksheet
    Dim mergedAreas As Collection
    Set mergedAreas = New Collection   ' Used to store all found merged areas
    
    Dim cell As Range, ma As Range
    Dim prefix As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary") ' Dictionary to track numbering per prefix
    
    Dim isFirst As Boolean
    isFirst = True  ' Flag to identify the first merged area
    
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
        
        ' Skip numbering for the first merged area (preserve user's manual input)
        If isFirst Then
            isFirst = False
            ' Do nothing, but record the first row's number in the dictionary for subsequent rows
            ' Extract the number part from first row (if it has one)
            Dim firstRowNumber As Long
            If InStr(ma.Cells(1, 1).Value, "-") > 0 Then
                On Error Resume Next
                firstRowNumber = CLng(Mid(ma.Cells(1, 1).Value, InStrRev(ma.Cells(1, 1).Value, "-") + 1))
                On Error GoTo 0
                If firstRowNumber > 0 Then
                    dict(prefix) = firstRowNumber  ' Start from this number
                End If
            End If
        ElseIf skipNumbering Then
            ' If H9:I9 is empty → keep only prefix (no numbering)
            ma.Cells(1, 1).Value = prefix
        Else
            ' Normal numbering process for subsequent rows
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

' =========================================================
' Helper function to parse merged item number from "prefix-number" format
' Returns True if parsing is successful, False otherwise
' Example: "72023-5" -> prefix="72023", number=5
' =========================================================
Private Function ParseMergedItemNumber(itemText As String, ByRef outPrefix As String, ByRef outNumber As Long) As Boolean
    Dim lastDashPos As Long
    Dim prefixPart As String
    Dim numberPart As String
    
    ParseMergedItemNumber = False
    On Error GoTo ParseError
    
    ' Trim whitespace
    itemText = Trim(itemText)
    
    ' Find the last dash position (in case prefix contains dash)
    lastDashPos = InStrRev(itemText, "-")
    
    ' Must have at least one dash
    If lastDashPos <= 0 Then
        GoTo ParseError
    End If
    
    ' Split into prefix and number
    prefixPart = Left(itemText, lastDashPos - 1)
    numberPart = Mid(itemText, lastDashPos + 1)
    
    ' Prefix must not be empty
    If Len(prefixPart) = 0 Then
        GoTo ParseError
    End If
    
    ' Convert number part to Long
    On Error Resume Next
    outNumber = CLng(numberPart)
    On Error GoTo ParseError
    
    ' Number must be positive
    If outNumber <= 0 Then
        GoTo ParseError
    End If
    
    outPrefix = prefixPart
    ParseMergedItemNumber = True
    Exit Function

ParseError:
    ParseMergedItemNumber = False
End Function
