' =========================================================
' Module: Helper_Data_ConvertTextToNumber
' =========================================================

Public Function ToNumber(v As Variant) As Variant
    Dim s As String, parts As Variant, frac As Variant
    Dim intPart As Double, numPart As Double, denPart As Double
    Dim isNegative As Boolean
    On Error GoTo ErrHandler

    ' numeric input
    If IsNumeric(v) Then ToNumber = CDbl(v): Exit Function

    s = Trim(CStr(v))
    If s = "" Then ToNumber = CVErr(xlErrValue): Exit Function

    ' detect negative sign
    isNegative = (Left(s, 1) = "-")
    If isNegative Then s = Mid(s, 2) ' remove negative sign for processing

    ' mixed fraction (e.g. "1 3/4" or "-1 3/4")
    If InStr(s, " ") > 0 And InStr(s, "/") > 0 Then
        parts = Split(s, " ")
        If UBound(parts) = 1 Then
            intPart = CDbl(parts(0))
            frac = Split(parts(1), "/")
            If UBound(frac) = 1 Then
                numPart = CDbl(frac(0))
                denPart = CDbl(frac(1))
                ToNumber = intPart + numPart / denPart
                If isNegative Then ToNumber = -ToNumber
                Exit Function
            End If
        End If
    End If

    ' pure fraction (e.g. "1/16" or "-1/16")
    If InStr(s, "/") > 0 Then
        frac = Split(s, "/")
        If UBound(frac) = 1 Then
            numPart = CDbl(frac(0))
            denPart = CDbl(frac(1))
            ToNumber = numPart / denPart
            If isNegative Then ToNumber = -ToNumber
            Exit Function
        End If
    End If

    ' fallback: numeric text
    If isNegative Then s = "-" & s
    If IsNumeric(s) Then ToNumber = CDbl(s): Exit Function

ErrHandler:
    ToNumber = CVErr(xlErrValue)
End Function

' =========================================================
' Function: IsFraction
' Purpose: Determine if a string represents a fraction
' Returns: True if the input is a pure or mixed fraction, False otherwise
' =========================================================
Function IsFraction(s As String) As Boolean
    Dim parts() As String
    Dim frac() As String
    Dim numPart As Double
    Dim denPart As Double
    Dim isNegative As Boolean
    
    IsFraction = False
    
    If s = "" Then Exit Function
    
    ' Check for negative sign
    isNegative = (Left(s, 1) = "-")
    If isNegative Then
        s = Mid(s, 2)
    End If
    
    ' mixed fraction (e.g. "1 3/4" or "-1 3/4")
    If InStr(s, " ") > 0 And InStr(s, "/") > 0 Then
        parts = Split(s, " ")
        If UBound(parts) = 1 Then
            On Error Resume Next
            numPart = CDbl(parts(0))
            If Err.Number = 0 Then
                frac = Split(parts(1), "/")
                If UBound(frac) = 1 Then
                    Err.Clear
                    denPart = CDbl(frac(1))
                    If Err.Number = 0 And denPart <> 0 Then
                        IsFraction = True
                        On Error GoTo 0
                        Exit Function
                    End If
                End If
            End If
            On Error GoTo 0
        End If
    End If
    
    ' pure fraction (e.g. "1/16" or "-1/16")
    If InStr(s, "/") > 0 Then
        frac = Split(s, "/")
        If UBound(frac) = 1 Then
            On Error Resume Next
            numPart = CDbl(frac(0))
            denPart = CDbl(frac(1))
            If Err.Number = 0 And denPart <> 0 Then
                IsFraction = True
            End If
            On Error GoTo 0
        End If
    End If
End Function
