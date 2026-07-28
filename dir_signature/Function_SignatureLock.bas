Option Explicit

' =========================================================
' Module: Function_SignatureLock
' Purpose: Stamp a signature picture onto every worksheet and
'          fully lock every worksheet with a single password.
' =========================================================

' Entry point to be called from a Ribbon button.
' Runs the full flow: pick signature image -> stamp it on all
' sheets -> ask for a password -> lock all sheets with it.
Public Sub LaunchSignatureLock()
    If IsAnySheetLocked() Then
        MsgBox "Sheets are already signed and locked. Unprotect them first before signing again.", vbExclamation
        Exit Sub
    End If

    Dim picPath As String
    picPath = PickSignatureFile()
    If picPath = "" Then Exit Sub

    InsertSignatureToAllSheets picPath

    Dim pwd As String
    pwd = InputBox("Enter password to lock all sheets:", "Lock Sheets")
    If pwd = "" Then Exit Sub

    LockAllSheets pwd

    MsgBox "Signature inserted and all sheets locked.", vbInformation
End Sub

' Returns True if any visible worksheet is already protected, so the
' launch button can be clicked again safely without re-running
' the flow on an already-locked workbook. The hidden "Data" sheet is
' skipped since it is locked by default and shouldn't affect this check.
Public Function IsAnySheetLocked() As Boolean
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And ws.ProtectContents Then
            IsAnySheetLocked = True
            Exit Function
        End If
    Next ws
End Function

' Shows a file picker rooted at the shared signature folder and
' returns the selected image path, or "" if the user cancels.
Private Function PickSignatureFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select Signature Image"
        .InitialFileName = "\\rtdnas2\QCReports\Signature\"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Image Files", "*.jpg; *.jpeg; *.png; *.bmp; *.gif"

        If .Show = -1 Then
            PickSignatureFile = .SelectedItems(1)
        Else
            PickSignatureFile = ""
        End If
    End With
End Function

' Inserts the given image at cell H57 on every visible worksheet (the
' hidden "Data" sheet is skipped, see IsAnySheetLocked), sized to a
' fixed 0.5-inch height with aspect ratio locked, white set as the
' transparent color, then nudged down 9pt and right 2pt from the
' cell's top-left corner.
Private Sub InsertSignatureToAllSheets(picPath As String)
    Dim ws As Worksheet
    Dim pic As Shape
    Dim anchor As Range

    For Each ws In ThisWorkbook.Worksheets
        ' Skip the hidden "Data" sheet: it is already protected with its own
        ' password, so inserting a shape into it raises a runtime error.
        If ws.Visible = xlSheetVisible Then
            On Error Resume Next
            ws.Shapes("SignaturePic").Delete
            On Error GoTo 0

            Set anchor = ws.Range("H57")
            Set pic = ws.Shapes.AddPicture(picPath, msoFalse, msoTrue, anchor.Left, anchor.Top, -1, -1)
            pic.Name = "SignaturePic"
            pic.LockAspectRatio = msoTrue
            pic.Height = Application.InchesToPoints(0.5)

            ' Set Transparent Color (white) so the signature background disappears
            pic.PictureFormat.TransparencyColor = RGB(255, 255, 255)
            pic.PictureFormat.TransparentBackground = msoTrue

            pic.Top = anchor.Top + 9
            pic.Left = anchor.Left + 2
        End If
    Next ws
End Sub

' Fully locks every visible worksheet with the given password: no cell
' selection at all, and every native Protect Sheet option left unchecked.
' UserInterfaceOnly:=True keeps manual user edits blocked while still
' letting other macros (e.g. UpdatePageNumbers) write to cells in this session.
' The hidden "Data" sheet is skipped since it is already protected with
' its own separate password.
Public Sub LockAllSheets(pwd As String)
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.EnableSelection = xlNoSelection
            ws.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                UserInterfaceOnly:=True, _
                AllowFormattingCells:=False, AllowFormattingColumns:=False, AllowFormattingRows:=False, _
                AllowInsertingColumns:=False, AllowInsertingRows:=False, AllowInsertingHyperlinks:=False, _
                AllowDeletingColumns:=False, AllowDeletingRows:=False, AllowSorting:=False, _
                AllowFiltering:=False, AllowUsingPivotTables:=False
        End If
    Next ws
End Sub

' Unprotects every visible worksheet that is currently protected, using
' the given password. The hidden "Data" sheet is always skipped since
' it is locked by default with a separate password.
Public Sub UnlockAllSheets(pwd As String)
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Data" And ws.Visible = xlSheetVisible And ws.ProtectContents Then
            ws.Unprotect Password:=pwd
        End If
    Next ws
End Sub
