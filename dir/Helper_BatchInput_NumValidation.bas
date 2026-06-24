' =========================================================
' Module: Helper_BatchInput_NumValidation
' =========================================================

Option Explicit

' =============================================
' Validate batch input numbers from the form
' =============================================
Public Function ValidateBatchInputNumbers() As Boolean
    Dim i As Long
    Dim val As String
    Dim numericCount As Long, fractionCount As Long, stringCount As Long
    Dim maxDecimalPlaces As Long
    Dim valDecimalPlaces As Long
    Dim dominantType As String
    Dim msg As String

    ValidateBatchInputNumbers = False ' default to fail

    ' === First pass: gather stats ===
    For i = 1 To Data_SharedValues.ItemCount
        val = Trim(CStr(Data_SharedValues.BatchValues(i)))

        If val = "" Then
            MsgBox "Input #" & i & " is empty. Please enter a value.", vbExclamation
            Exit Function
        End If

        If IsFractionText(val) Then
            fractionCount = fractionCount + 1
        ElseIf IsNumeric(val) Then
            numericCount = numericCount + 1
            valDecimalPlaces = CountDecimalPlaces(val)
            If valDecimalPlaces > maxDecimalPlaces Then
                maxDecimalPlaces = valDecimalPlaces
            End If
        Else
            stringCount = stringCount + 1
        End If
    Next i

    ' === Determine dominant input type ===
    If numericCount >= fractionCount And numericCount >= stringCount Then
        dominantType = "Numeric"
    ElseIf fractionCount >= numericCount And fractionCount >= stringCount Then
        dominantType = "Fraction"
    Else
        dominantType = "String"
    End If

    ' === Second pass: validate against dominant type ===
    For i = 1 To Data_SharedValues.ItemCount
        val = Trim(CStr(Data_SharedValues.BatchValues(i)))

        Select Case dominantType
            Case "Numeric"
                If Not IsNumeric(val) Then
                    msg = "Input #" & i & " (" & val & ") must be numeric like the others."
                    GoTo ValidationError
                End If

                valDecimalPlaces = CountDecimalPlaces(val)
                ' only warn if inconsistent precision, not less decimal precision unless mixture matters
                If valDecimalPlaces < maxDecimalPlaces Then
                    msg = "Input #" & i & " (" & val & ") has fewer decimal places than the most precise entry (" & _
                          maxDecimalPlaces & ")."
                    GoTo ValidationError
                End If

            Case "Fraction"
                If Not IsFractionText(val) Then
                    msg = "Input #" & i & " (" & val & ") must be a fraction like the others."
                    GoTo ValidationError
                End If

            Case "String"
                If IsNumeric(val) Or IsFractionText(val) Then
                    msg = "Input #" & i & " (" & val & ") must be a text value like the others."
                    GoTo ValidationError
                End If
        End Select
    Next i

    ' === All checks passed ===
    ValidateBatchInputNumbers = True
    Exit Function

ValidationError:
    MsgBox msg, vbExclamation, "Batch Input Validation"
End Function
