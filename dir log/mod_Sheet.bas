' In Sheet module
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Check if the change occurred in column C
    If Target.Column = 3 Then
        ' Set the row number of the edited cell
        mod_PublicData.SetLastEditedRow Target.row
        
        ' If the cell is not empty, call FindDrawingNumbers
        If Target.value <> "" Then
            Call FindDrawingNumbers(Target.value)
        End If
    End If
End Sub