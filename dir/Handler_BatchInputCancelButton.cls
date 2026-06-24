' =========================================================
' Class Module: Handler_BatchInputCancelButton
' =========================================================

Option Explicit

Public WithEvents Btn As MSForms.CommandButton

Private Sub Btn_Click()
    ' Clear all batch values
    Dim i As Long
    For i = 1 To Data_SharedValues.ItemCount
        Data_SharedValues.BatchValues(i) = ""
    Next i
    
    ' Close the form
    Btn.Parent.Hide
End Sub