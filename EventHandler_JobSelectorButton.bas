' Class Module: EventHandler_JobSelectorButton
Option Explicit

Public WithEvents Btn As MSForms.CommandButton

Private Sub Btn_Click()
    Dim selectedDrawingNumber As String
    Dim oe_number As String, po_number As String, part_number As String, description As String
    Dim selectedRow As Long
    Dim tagParts() As String
    Dim JobSelector As Object
    Dim currentSheet As Worksheet
    Dim lastRow As Long
    Dim formula As String

    ' Split tag to retrieve oe_number, po_number
    tagParts = Split(Btn.Tag, "|")

    oe_number = tagParts(0)
    po_number = tagParts(1)
    part_number = tagParts(2)
    selectedRow = mod_PublicData.GetLastEditedRow() ' Get row from public data module

    ' Get selected drawing number from the JobSelector form
    Set JobSelector = Btn.Parent
    selectedDrawingNumber = JobSelector.cboDrawingNumber.Value
    
    ' Get description of the selected drawing number
    Dim dbPath As String, sql As String, results As Variant
    dbPath = "\\rtdnas2\OE\jobs.db"

    If Not InitializeSQLite(dbPath) Then
        MsgBox "Failed to initialize database connection.", vbCritical
        Exit Sub
    End If

    sql = "SELECT description FROM assemblies WHERE drawing_number = '" & selectedDrawingNumber & "'"
    results = ExecuteSQL(sql)

    If Not IsNull(results) Then
        description = results(0)(0)
    Else
        description = "" ' Set to empty string if description not found
    End If

    CloseSQLite
    
    ' Get the current active worksheet
    Set currentSheet = ThisWorkbook.ActiveSheet
    
    ' Get the last row of the current sheet
    lastRow = currentSheet.Rows.Count

    ' Validate selectedRow
    If selectedRow > 0 And selectedRow <= lastRow Then
        ' Write selected drawing number and other information to Excel
        currentSheet.Cells(selectedRow, 1).Value = selectedDrawingNumber ' Column A
        currentSheet.Cells(selectedRow, 5).Value = oe_number ' Column E
        currentSheet.Cells(selectedRow, 6).Value = po_number ' Column F
        currentSheet.Cells(selectedRow, 8).Value = description ' Column G

        ' Construct the formula
        formula = "=IF(AND(A" & selectedRow & "=""" & """, B" & selectedRow & "=""" & """, C" & selectedRow & "=""" & """), """", A" & selectedRow & " & "" REV. "" & B" & selectedRow & " & "" @"" & C" & selectedRow & ")"

        ' Apply the formula to column D
        currentSheet.Cells(selectedRow, 4).formula = formula

    Else
        Debug.Print "Error: selectedRow (" & selectedRow & ") is out of valid range (1 to " & lastRow & ")."
        MsgBox "Invalid row number. Please make sure to edit a cell in column C before selecting.", vbCritical
    End If

    Unload JobSelector
End Sub

