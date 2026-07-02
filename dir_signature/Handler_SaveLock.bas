Option Explicit

' =========================================================
' Module: Handler_SaveLock
' Purpose: Bound to Ctrl+S (see ThisWorkbook.Workbook_Open).
'          Locks all sheets with a fixed password before saving,
'          independent of the LaunchSignatureLock flow (no signing).
' =========================================================

Private Const LOCK_PASSWORD As String = "rtd#158849"

' Handles Ctrl+S: locks this workbook's sheets (if not already
' locked) before saving. If a different workbook is active, the
' key press falls through to a normal save untouched.
Public Sub HandleCtrlSSave()
    If ActiveWorkbook Is ThisWorkbook Then
        On Error GoTo LockFailed
        If Not Function_SignatureLock.IsAnySheetLocked() Then
            Function_SignatureLock.LockAllSheets LOCK_PASSWORD
        End If
        On Error GoTo 0
    End If

    ActiveWorkbook.Save
    Exit Sub

LockFailed:
    MsgBox "Failed to lock sheets before saving: " & Err.Description & vbCrLf & _
        "Skipping lock and saving anyway.", vbExclamation, "Save Lock"
    Resume Next
End Sub
