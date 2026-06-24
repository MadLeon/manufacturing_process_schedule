Attribute VB_Name = "mod_UpdateDocumentPackageTracker"
' =========================================================
' Module: mod_UpdateDocumentPackageTracker
' Purpose: Mark the current row's drawing as received (checkmark)
'          in the Document Package Tracker workbook.
' Trigger: Button click on the tracking sheet.
' =========================================================
Option Explicit

Private Const TRACKER_PATH As String = "\\rtdnas2\QCReports\FINAL REPORTS\Document Package Tracker.xlsm"

Public Sub UpdateDocumentPackageTracker()
    ' --- Validate selection ---
    If TypeName(Selection) <> "Range" Then
        MsgBox "No cell selected.", vbExclamation, "Error"
        Exit Sub
    End If

    Dim currentRow As Long
    currentRow = Selection.Row

    ' --- Extract Drawing Number (col A) and PO Number (col F) ---
    Dim drawingNumber As String
    Dim poNumber As String
    drawingNumber = Trim(CStr(ThisWorkbook.ActiveSheet.Cells(currentRow, 1).Value))
    poNumber = Trim(CStr(ThisWorkbook.ActiveSheet.Cells(currentRow, 6).Value))

    If drawingNumber = "" Then
        MsgBox "Drawing Number (Column A) is empty in row " & currentRow & ".", vbExclamation, "Missing Data"
        Exit Sub
    End If
    If poNumber = "" Then
        MsgBox "PO Number (Column F) is empty in row " & currentRow & ".", vbExclamation, "Missing Data"
        Exit Sub
    End If

    ' --- Check tracker file exists ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(TRACKER_PATH) Then
        MsgBox "Document Package Tracker not found:" & vbCrLf & TRACKER_PATH, vbCritical, "File Not Found"
        Set fso = Nothing
        Exit Sub
    End If
    Set fso = Nothing

    ' --- Open tracker (reuse if already open) ---
    Dim wbTracker As Workbook
    Dim openedByUs As Boolean
    openedByUs = False

    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.FullName = TRACKER_PATH Then
            Set wbTracker = wb
            Exit For
        End If
    Next wb

    If wbTracker Is Nothing Then
        On Error Resume Next
        Set wbTracker = Workbooks.Open(TRACKER_PATH)
        If Err.Number <> 0 Then
            MsgBox "Failed to open Document Package Tracker:" & vbCrLf & Err.Description, vbCritical, "Open Error"
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        openedByUs = True
    End If

    ' --- Locate worksheet named after PO Number ---
    Dim wsTarget As Worksheet
    On Error Resume Next
    Set wsTarget = wbTracker.Worksheets(poNumber)
    On Error GoTo 0

    If wsTarget Is Nothing Then
        MsgBox "Worksheet '" & poNumber & "' not found in Document Package Tracker.", vbExclamation, "Sheet Not Found"
        If openedByUs Then wbTracker.Close SaveChanges:=False
        Exit Sub
    End If

    ' --- Search for Drawing Number in columns A, E, B, G ---
    Dim foundCell As Range
    Dim searchCols As Variant
    searchCols = Array(1, 5, 2, 7)   ' A, E, B, G

    Dim col As Variant
    For Each col In searchCols
        Set foundCell = wsTarget.Columns(col).Find(What:=drawingNumber, LookAt:=xlWhole, MatchCase:=False)
        If Not foundCell Is Nothing Then Exit For
    Next col

    If foundCell Is Nothing Then
        MsgBox "Drawing Number '" & drawingNumber & "' not found in worksheet '" & poNumber & "'.", _
               vbExclamation, "Not Found"
        If openedByUs Then wbTracker.Close SaveChanges:=False
        Exit Sub
    End If

    ' --- Write checkmark to the cell immediately to the right ---
    Dim targetCell As Range
    Set targetCell = foundCell.Offset(0, 1)

    If targetCell.Value = ChrW(10003) Then
        If openedByUs Then wbTracker.Close SaveChanges:=False
        MsgBox "'" & drawingNumber & "' is already marked in sheet '" & poNumber & "'.", vbInformation, "Already Marked"
        Exit Sub
    End If

    targetCell.Value = ChrW(10003)   ' ✓

    ' --- Save and optionally close ---
    wbTracker.Save
    If openedByUs Then wbTracker.Close SaveChanges:=False

    Debug.Print "UpdateDocumentPackageTracker: marked '" & drawingNumber & "' in sheet '" & poNumber & "' at " & targetCell.Address
    MsgBox "'" & drawingNumber & "' has been marked (✓) in sheet '" & poNumber & "'.", vbInformation, "Success"
End Sub
