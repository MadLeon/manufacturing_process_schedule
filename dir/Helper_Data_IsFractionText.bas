' -------------------------
' Module: Helper_Data_IsFractionText
' Check if input is a fraction (pure or mixed, e.g., "1/2" or "2 1/2")
' -------------------------
Public Function IsFractionText(val As Variant) As Boolean
    Dim s As String
    Dim parts() As String
    Dim fracParts() As String
    
    s = Trim(CStr(val))
    IsFractionText = False
    
    If s = "" Then Exit Function
    
    ' Check mixed fraction: integer + space + fraction
    If InStr(s, " ") > 0 Then
        parts = Split(s, " ")
        If UBound(parts) = 1 Then
            fracParts = Split(parts(1), "/")
            If UBound(fracParts) = 1 Then
                If IsNumeric(parts(0)) And IsNumeric(fracParts(0)) And IsNumeric(fracParts(1)) Then
                    IsFractionText = True
                    Debug.Print "Mixed fraction detected: " & s
                    Exit Function
                End If
            End If
        End If
    ElseIf InStr(s, "/") > 0 Then
        ' Pure fraction
        fracParts = Split(s, "/")
        If UBound(fracParts) = 1 Then
            If IsNumeric(fracParts(0)) And IsNumeric(fracParts(1)) Then
                IsFractionText = True
                Debug.Print "Pure fraction detected: " & s
                Exit Function
            End If
        End If
    End If
End Function
