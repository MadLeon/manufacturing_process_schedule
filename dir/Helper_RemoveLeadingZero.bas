' =========================================================
' Module: Helper_RemoveLeadingZero
' =========================================================

Sub RemoveLeadingZero()
    Dim cell As Range
    For Each cell In Selection
        If Not IsEmpty(cell.Value) Then
            If IsNumeric(cell.Value) Then
                If Left(cell.Value, 2) = "0." Then
                    cell.Value = Mid(cell.Value, 2)
                End If
            End If
        End If
    Next cell
End Sub
