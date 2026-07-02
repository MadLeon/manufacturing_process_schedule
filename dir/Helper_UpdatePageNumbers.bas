' =========================================================
' Module: Helper_UpdatePageNumbers
' =========================================================

' Update page numbers in cell I5 on each worksheet.
' Auto-numbered based on current sheet count: Page 1 of X, Page 2 of X, etc.
' Used for automatic page extension during normal workflow.
Public Sub UpdatePageNumbers()
    Dim ws As Worksheet
    Dim totalSheets As Integer
    Dim count As Integer

    ' Skip entirely while any visible sheet is locked (e.g. mid bulk-unlock);
    ' writing to I5 on a protected sheet raises a runtime error once
    ' UserInterfaceOnly resets on file reopen.
    If Function_SignatureLock.IsAnySheetLocked() Then Exit Sub

    ' Count all worksheets except the one named "Data"
    totalSheets = 0
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Data" Then totalSheets = totalSheets + 1
    Next ws

    ' Update each sheet's page number in I5
    count = 1
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Data" Then
            With ws.Range("I5")
                ' Format: "Page x of y"
                .Value = "Page " & count & " of " & totalSheets
                ' Align the text to the right
                .HorizontalAlignment = xlRight
            End With
            count = count + 1
        End If
    Next ws
End Sub

' =========================================================
' Update page numbers based on the first sheet's page number.
' Manual button trigger version that respects the starting page and total page limit.
' Example: If first sheet is "Page 2 of 10", subsequent sheets will be 2, 3, 4, ...
' Validates that the numbering does not exceed the total page limit.
' =========================================================
Public Sub UpdatePageNumbersFromFirstSheet()
    Dim ws As Worksheet
    Dim totalSheets As Integer
    Dim count As Integer
    Dim firstSheet As Worksheet
    Dim startingPage As Integer
    Dim totalPageLimit As Integer
    Dim firstPageValue As String
    Dim parsedParts() As String

    On Error GoTo ErrorHandler

    ' Find the first non-Data worksheet
    Set firstSheet = Nothing
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Data" Then
            Set firstSheet = ws
            Exit For
        End If
    Next ws

    ' If no valid first sheet found, exit
    If firstSheet Is Nothing Then
        MsgBox "No valid worksheets found (excluding Data sheet).", vbExclamation, "Update Page Numbers From First Sheet"
        Exit Sub
    End If

    ' Parse the first sheet's page number from I5
    firstPageValue = CStr(firstSheet.Range("I5").Value)
    If Not ParsePageNumber(firstPageValue, startingPage, totalPageLimit) Then
        MsgBox "First sheet page number format is invalid. Expected format: 'Page X of Y'", vbExclamation, "Update Page Numbers From First Sheet"
        Exit Sub
    End If

    ' Count all worksheets except the one named "Data"
    totalSheets = 0
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Data" Then totalSheets = totalSheets + 1
    Next ws

    ' Validate that the numbering will not exceed the total page limit
    If startingPage + totalSheets - 1 > totalPageLimit Then
        MsgBox "Cannot update page numbers. Starting at Page " & startingPage & ", " & totalSheets & " sheets would exceed the limit of Page " & totalPageLimit & " of " & totalPageLimit & ".", vbExclamation, "Update Page Numbers From First Sheet"
        Exit Sub
    End If

    ' Update each sheet's page number in I5
    count = startingPage
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Data" Then
            With ws.Range("I5")
                ' Format: "Page x of y"
                .Value = "Page " & count & " of " & totalPageLimit
                ' Align the text to the right
                .HorizontalAlignment = xlRight
            End With
            count = count + 1
        End If
    Next ws

    Exit Sub
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Update Page Numbers From First Sheet"
End Sub

' Helper function to parse page number from "Page X of Y" format
' Returns True if parsing is successful, False otherwise
Private Function ParsePageNumber(pageText As String, ByRef outStartPage As Integer, ByRef outTotalPages As Integer) As Boolean
    Dim parts() As String
    Dim i As Integer

    ParsePageNumber = False
    On Error GoTo ParseError

    ' Trim whitespace
    pageText = Trim(pageText)

    ' Expected format: "Page X of Y"
    ' Split by space
    parts = Split(pageText, " ")

    ' Should have exactly 4 parts: "Page", "X", "of", "Y"
    If UBound(parts) <> 3 Then
        GoTo ParseError
    End If

    ' Validate parts
    If LCase(parts(0)) <> "page" Or LCase(parts(2)) <> "of" Then
        GoTo ParseError
    End If

    ' Try to convert X and Y to integers
    On Error Resume Next
    outStartPage = CLng(parts(1))
    outTotalPages = CLng(parts(3))
    On Error GoTo ParseError

    If outStartPage <= 0 Or outTotalPages <= 0 Then
        GoTo ParseError
    End If

    ParsePageNumber = True
    Exit Function

ParseError:
    ParsePageNumber = False
End Function

' =========================================================
' Test function - verify the UpdatePageNumbersFromFirstSheet logic
' =========================================================
Public Sub TestUpdatePageNumbersFromFirstSheet()
    Dim result As String
    Dim firstSheet As Worksheet
    Dim ws As Worksheet
    Dim totalSheets As Integer
    Dim startingPage As Integer
    Dim totalPageLimit As Integer
    Dim firstPageValue As String

    ' Find first non-Data sheet
    Set firstSheet = Nothing
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Data" Then
            Set firstSheet = ws
            Exit For
        End If
    Next ws

    If firstSheet Is Nothing Then
        result = "No valid worksheets found."
        MsgBox result, vbInformation, "Test Result"
        Exit Sub
    End If

    ' Get first sheet's page number
    firstPageValue = CStr(firstSheet.Range("I5").Value)
    If Not ParsePageNumber(firstPageValue, startingPage, totalPageLimit) Then
        result = "Invalid page number format in first sheet: " & firstPageValue
        MsgBox result, vbInformation, "Test Result"
        Exit Sub
    End If

    ' Count sheets
    totalSheets = 0
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Data" Then totalSheets = totalSheets + 1
    Next ws

    ' Build result message
    result = "First Sheet Page Number: " & firstPageValue & vbCrLf
    result = result & "Starting Page: " & startingPage & vbCrLf
    result = result & "Total Page Limit: " & totalPageLimit & vbCrLf
    result = result & "Current Sheet Count: " & totalSheets & vbCrLf
    result = result & "Final Page Would Be: Page " & (startingPage + totalSheets - 1) & " of " & totalPageLimit & vbCrLf & vbCrLf

    ' Check if valid
    If startingPage + totalSheets - 1 > totalPageLimit Then
        result = result & "Status: INVALID - Would exceed page limit!"
    Else
        result = result & "Status: VALID - Ready to update."
    End If

    MsgBox result, vbInformation, "Test Result - UpdatePageNumbersFromFirstSheet"
End Sub