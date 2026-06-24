' -------------------------
' Module: Helper_Data_CountDecimal
' -------------------------
Public Function CountDecimalPlaces(val As Variant) As Long
    Dim s As String
    s = CStr(val)                       ' Convert to string
    If InStr(s, ".") > 0 Then           ' If decimal exists
        CountDecimalPlaces = Len(Split(s, ".")(1))  ' Count digits after decimal
    Else
        CountDecimalPlaces = 0          ' No decimal => 0
    End If
End Function
