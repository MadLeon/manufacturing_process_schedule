Private Sub GenRandomNumber_Click()

    Dim upperTol As Double, lowerTol As Double, resultVal As Double
    Dim upperDp As Long, lowerDp As Long, decPlaces As Long
    Dim resultStr As String
    Dim selectedCell As Range, rowNum As Long
    Dim colDValue As Variant, colEValue As Variant

    ' Get the selected row for reference to columns B/D/E
    Set selectedCell = ActiveCell
    rowNum = selectedCell.Row
    colDValue = ActiveSheet.Cells(rowNum, 4).Value  ' Column D
    colEValue = ActiveSheet.Cells(rowNum, 5).Value  ' Column E

    ' -------------------- Decision logic --------------------
    If Trim(Me.UpperTolerance.Text) = "" And Trim(Me.LowerTolerance.Text) = "" Then
        ' Case 1: Both boundaries are empty → use B/D/E columns
        Dim baseValue As Double
        baseValue = ToNumber(ActiveSheet.Cells(rowNum, 2).Value) ' Column B
        upperTol = baseValue + IIf(colDValue = "", 0, ToNumber(colDValue))
        lowerTol = baseValue + IIf(colEValue = "", 0, ToNumber(colEValue))

        ' If either is a fraction, set decimal places to 2
        If IsFraction(CStr(colDValue)) Or IsFraction(CStr(colEValue)) Then
            decPlaces = 2
        Else
            upperDp = CountDecimalPlaces(ToNumber(colDValue)) + 1
            lowerDp = CountDecimalPlaces(ToNumber(colEValue)) + 1
            decPlaces = IIf(upperDp > lowerDp, upperDp, lowerDp)
            
            ' Limit decimal places to 4
            If decPlaces > 4 Then
                decPlaces = 4
            End If
        End If

    ElseIf Trim(Me.UpperTolerance.Text) <> "" And Trim(Me.LowerTolerance.Text) <> "" Then
        ' Case 4: Both boundaries have values → use these as the range
        upperTol = ToNumber(Me.UpperTolerance.Text)
        lowerTol = ToNumber(Me.LowerTolerance.Text)

        ' If either is a fraction, set decimal places to 2
        If IsFraction(Me.UpperTolerance.Text) Or IsFraction(Me.LowerTolerance.Text) Then
            decPlaces = 2
        Else
            upperDp = CountDecimalPlaces(Me.UpperTolerance.Text)
            lowerDp = CountDecimalPlaces(Me.LowerTolerance.Text)
            decPlaces = IIf(upperDp > lowerDp, upperDp, lowerDp)
        End If

    Else
        ' Only one boundary has value → fill all boxes with fixed value
        If Trim(Me.UpperTolerance.Text) <> "" Then
            upperTol = ToNumber(Me.UpperTolerance.Text)
            lowerTol = upperTol
            ' If it's a fraction, set decimal places to 2
            If IsFraction(Me.UpperTolerance.Text) Then
                decPlaces = 2
            Else
                decPlaces = CountDecimalPlaces(Me.UpperTolerance.Text)
            End If
        Else
            lowerTol = ToNumber(Me.LowerTolerance.Text)
            upperTol = lowerTol
            ' If it's a fraction, set decimal places to 2
            If IsFraction(Me.LowerTolerance.Text) Then
                decPlaces = 2
            Else
                decPlaces = CountDecimalPlaces(Me.LowerTolerance.Text)
            End If
        End If
    End If

    ' -------------------- Validate tolerances --------------------
    If upperTol < lowerTol Then
        MsgBox "Upper Tolerance must be greater than or equal to Lower Tolerance.", vbExclamation, "Input Error"
        Exit Sub
    End If

    Dim bothWholeNumbers As Boolean
    bothWholeNumbers = (upperTol = Int(upperTol)) And (lowerTol = Int(lowerTol))

    ' -------------------- Generate random numbers --------------------
    Dim idx As Long
    For idx = 1 To Data_SharedValues.ItemCount
        Randomize
        If upperTol = lowerTol Then
            resultVal = lowerTol
        ElseIf bothWholeNumbers Then
            resultVal = Int((upperTol - lowerTol + 1) * Rnd + lowerTol)
        Else
            resultVal = lowerTol + (upperTol - lowerTol) * Rnd
        End If

        ' Format result
        If bothWholeNumbers Then
            resultStr = CStr(Int(resultVal))
        ElseIf decPlaces > 0 Then
            resultStr = Format(resultVal, "0." & String(decPlaces, "0"))
            ' Remove leading zero
            If Left(resultStr, 2) = "0." Then
                resultStr = Mid(resultStr, 2)
            ElseIf Left(resultStr, 3) = "-0." Then
                resultStr = "-" & Mid(resultStr, 3)
            End If
        Else
            resultStr = CStr(resultVal)
        End If

        Me.Controls("txtItem" & idx).Text = resultStr
    Next idx

    ' Set focus to OK button
    On Error Resume Next
    Me.Controls("btnOK").SetFocus
    On Error GoTo 0

End Sub

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