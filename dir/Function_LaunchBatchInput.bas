Option Explicit

' =============================================
' Module: Helper_LaunchBatchInput
' Purpose: Launch batch input form with validations
' =============================================

' Launches the batch input form for the currently selected cell
Public Sub LaunchBatchInput()
    Dim ws As Worksheet
    Dim selectedCell As Range
    Dim h9Val As Long
    Dim i As Long
    
    Set ws = ActiveSheet
    Set selectedCell = ActiveCell
    
    ' 1?? Initialize StartRow if not already set
    If Data_SharedValues.startRow = 0 Then Helper_Database_InitStartRow.InitializeStartRow ws
    
    ' 2?? Validate selected cell (column F and within data area)
    If Not ValidateSelectedCell(selectedCell) Then Exit Sub
    
    ' 3?? Validate H9 value
    h9Val = CLng(ws.Range("H9").Value)
    If h9Val <= 1 Then
        MsgBox "H9 must have a number greater than 1.", vbExclamation
        Exit Sub
    End If
    
    ' 4?? Validate merged cells in workbook (always scan from row 15)
    If Not ValidateMergedCells(h9Val) Then Exit Sub
    
    ' 5?? Store selected column and item count
    Data_SharedValues.ItemCount = h9Val
    ReDim Data_SharedValues.BatchValues(1 To h9Val)
    
    Dim rowDistance As Long
    rowDistance = GetRowDistanceToNearestHeader(ActiveCell)
    
    ' Find all matching target cells
    FindTargetCells ActiveCell, rowDistance
    
    ' 6?? Show batch input form
    Dim frm As Form_BatchInput
    Set frm = New Form_BatchInput
    Load frm
    frm.PrepareForm h9Val
    frm.Show vbModal
    
End Sub


