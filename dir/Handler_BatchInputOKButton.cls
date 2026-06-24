' =========================================================
' Class Module: Handler_BatchInputOKButton
' =========================================================

Option Explicit

Public WithEvents Btn As MSForms.CommandButton

Private Sub Btn_Click()
    Dim i As Long
    Dim tb As MSForms.TextBox
    Dim targetCell As Range
    
    ' ==========================
    ' 1. Collect values from textboxes
    ' ==========================
    For i = 1 To Data_SharedValues.ItemCount
        Set tb = Btn.Parent.Controls("txtItem" & i)
        Data_SharedValues.BatchValues(i) = tb.Text
    Next i
    
    ' ==========================
    ' 2. Validate inputs
    ' ==========================
    If Not Helper_BatchInput_NumValidation.ValidateBatchInputNumbers() Then Exit Sub
    
    ' ==========================
    ' 3. Update target cells with input values
    ' ==========================
    If Not Data_SharedValues.TargetCells Is Nothing Then
        If Data_SharedValues.TargetCells.count <> Data_SharedValues.ItemCount Then
            MsgBox "Target cells count (" & Data_SharedValues.TargetCells.count & _
                   ") does not match number of inputs (" & Data_SharedValues.ItemCount & ")." & vbCrLf & vbCrLf & _
                   "Please check that the current dimension row of each item is consistent.", vbExclamation
            Exit Sub
        End If
        
        For i = 1 To Data_SharedValues.ItemCount
            Set targetCell = Data_SharedValues.TargetCells(i)
            targetCell.Value = Data_SharedValues.BatchValues(i)
        Next i
    Else
        MsgBox "No target cells found to update.", vbExclamation
        Exit Sub
    End If
    
    ' ==========================
    ' 4. Close the form
    ' ==========================
    Btn.Parent.Hide
    
    ' ==========================
    ' 5. Force recalculation across all sheets
    ' ==========================
    Application.CalculateFullRebuild
End Sub


