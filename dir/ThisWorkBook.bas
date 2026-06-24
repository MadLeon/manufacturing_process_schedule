Option Explicit

' =========================================================
' Module: ThisWorkbook
' Purpose: Handle workbook-level events such as opening,
'          sheet activation, and sheet changes.
' =========================================================


' =========================================================
' Triggered when the workbook is opened.
' =========================================================
Private Sub Workbook_Open()
    ' Existing startup tasks
    Data_SharedValues.SheetCount = CountVisibleSheets()
    Call UpdatePageNumbers
    Call UpdateMergedRowNumbersWorkbook
    Helper_Database_StartRow.InitializeStartRow ActiveSheet
End Sub


' =========================================================
' Triggered when the user switches to another worksheet.
' =========================================================
Private Sub Workbook_SheetActivate(ByVal sh As Object)
    Dim currentCount As Integer
    ' Get the current count of visible worksheets.
    currentCount = CountVisibleSheets()
    
    ' If the number of worksheets has changed (added or removed),
    ' update the counter and refresh page numbers.
    If currentCount <> Data_SharedValues.SheetCount Then
        ' A new sheet has been added or copied
        Debug.Print "New sheet detected: " & sh.Name
        
        Data_SharedValues.SheetCount = currentCount
        Call UpdatePageNumbers
        Call UpdateMergedRowNumbersWorkbook
    End If
        
    ' Re-initialize start row whenever a sheet is activated
    Helper_Database_InitStartRow.InitializeStartRow sh
End Sub


' =========================================================
' Count the number of visible worksheets in the workbook,
' excluding the sheet named "Data".
' =========================================================
Private Function CountVisibleSheets() As Integer
    Dim ws As Worksheet
    Dim count As Integer
    count = 0
    
    ' Loop through all worksheets.
    For Each ws In ThisWorkbook.Sheets
        ' Exclude the sheet named "Data".
        If ws.Name <> "Data" Then
            count = count + 1
        End If
    Next ws
    
    ' Return the number of visible worksheets.
    CountVisibleSheets = count
End Function


' =========================================================
' Triggered when any cell value is changed in any worksheet.
' Used here to automatically update numbering in merged cells
' if a change occurs in column A (starting from row 15).
' =========================================================
Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    On Error GoTo CleanExit
    ' Disable events temporarily to prevent recursion.
    Application.EnableEvents = False
    
    ' Keep SheetCount always updated when any change occurs
    Data_SharedValues.SheetCount = CountVisibleSheets()
    
    ' Only handle changes in column A from row 15 downward.
    ' Column A may contain merged cells that require renumbering.
    If Not Intersect(Target, sh.Range("A15:A" & sh.Rows.count)) Is Nothing Then
        ' Update merged cell numbering across the entire workbook.
        Call UpdateMergedRowNumbersWorkbook
    End If

CleanExit:
    ' Re-enable events after changes.
    Application.EnableEvents = True
End Sub