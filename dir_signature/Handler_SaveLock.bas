Option Explicit

' =========================================================
' Module: Handler_SaveLock
' Purpose: Bound to Ctrl+S and Ctrl+D (see ThisWorkbook.Workbook_Open).
'          Ctrl+S locks all sheets with a fixed password before saving;
'          Ctrl+D unlocks them again with that same password. Both are
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

' Handles Ctrl+D: if this workbook is active and any visible sheet
' is currently locked, unlocks all locked visible sheets (Data sheet
' excepted) with the same fixed password. Does nothing if no visible
' sheet is locked, or if a different workbook is active.
Public Sub HandleCtrlDUnlock()
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If Not Function_SignatureLock.IsAnySheetLocked() Then Exit Sub

    On Error GoTo UnlockFailed
    Function_SignatureLock.UnlockAllSheets LOCK_PASSWORD
    On Error GoTo 0
    Exit Sub

UnlockFailed:
    MsgBox "Failed to unlock sheets: " & Err.Description, vbExclamation, "Unlock Sheets"
End Sub
