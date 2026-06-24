' ------- Random number button click event ---------
Private Sub GenRandomNumber_Click()

    Dim upperTol As Double, lowerTol As Double, resultVal As Double
    Dim upperDp As Long, lowerDp As Long, decPlaces As Long
    Dim resultStr As String
    Dim baseValue As Double
    Dim isDynamicMode As Boolean
    Dim selectedCell As Range, rowNum As Long
    Dim colDValue As Variant, colEValue As Variant

    On Error Resume Next
    isDynamicMode = Me.CheckBox1.Value
    On Error GoTo 0

    ' Get the selected row for reference to columns B/D/E
    Set selectedCell = ActiveCell
    rowNum = selectedCell.Row
    colDValue = ActiveSheet.Cells(rowNum, 4).Value  ' Column D
    colEValue = ActiveSheet.Cells(rowNum, 5).Value  ' Column E

    ' -------------------- Decision logic --------------------
    If Trim(Me.UpperTolerance.Text) = "" And Trim(Me.LowerTolerance.Text) = "" Then
        ' Case 1: Both boundaries are empty → use B/D/E columns
        baseValue = ToNumber(ActiveSheet.Cells(rowNum, 2).Value) ' Column B
        upperTol = baseValue + IIf(colDValue = "", 0, ToNumber(colDValue))
        lowerTol = baseValue + IIf(colEValue = "", 0, ToNumber(colEValue))

        upperDp = CountDecimalPlaces(ToNumber(colDValue)) + 1
        lowerDp = CountDecimalPlaces(ToNumber(colEValue)) + 1
        decPlaces = IIf(upperDp > lowerDp, upperDp, lowerDp)
        
        ' Limit decimal places to 4
        If decPlaces > 4 Then
            decPlaces = 4
        End If

    ElseIf Trim(Me.UpperTolerance.Text) <> "" And Trim(Me.LowerTolerance.Text) <> "" Then
        ' Case 4: Both boundaries have values → use these as the range
        upperTol = ToNumber(Me.UpperTolerance.Text)
        lowerTol = ToNumber(Me.LowerTolerance.Text)

        upperDp = CountDecimalPlaces(Me.UpperTolerance.Text)
        lowerDp = CountDecimalPlaces(Me.LowerTolerance.Text)
        decPlaces = IIf(upperDp > lowerDp, upperDp, lowerDp)

    ElseIf isDynamicMode Then
        ' Case 3: Dynamic Mode checked and only one boundary has value → use it as base, D/E as offset
        If Trim(Me.LowerTolerance.Text) <> "" Then
            baseValue = ToNumber(Me.LowerTolerance.Text)
        Else
            baseValue = ToNumber(Me.UpperTolerance.Text)
        End If

        upperTol = baseValue + IIf(colDValue = "", 0, ToNumber(colDValue))
        lowerTol = baseValue + IIf(colEValue = "", 0, ToNumber(colEValue))

        upperDp = CountDecimalPlaces(ToNumber(colDValue)) + 1
        lowerDp = CountDecimalPlaces(ToNumber(colEValue)) + 1
        decPlaces = IIf(upperDp > lowerDp, upperDp, lowerDp)

    Else
        ' Case 2: Dynamic Mode not checked and only one boundary has value → fill all boxes with fixed value
        If Trim(Me.UpperTolerance.Text) <> "" Then
            upperTol = ToNumber(Me.UpperTolerance.Text)
            lowerTol = upperTol
            decPlaces = CountDecimalPlaces(Me.UpperTolerance.Text)
        Else
            lowerTol = ToNumber(Me.LowerTolerance.Text)
            upperTol = lowerTol
            decPlaces = CountDecimalPlaces(Me.LowerTolerance.Text)
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