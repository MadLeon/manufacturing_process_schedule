
' ------ UserForm: Form_BatchInput ------
Option Explicit

Private ButtonHandlers() As Object

' ------- Core: Generate dynamic inputs -----
Public Sub PrepareForm(ByVal numItems As Long)
    ' Layout and sizing config
    Dim colWidth As Long: colWidth = 120
    Dim rowHeight As Long: rowHeight = 25
    Dim formPadding As Long: formPadding = 50
    Dim btnWidth As Long: btnWidth = 60
    Dim btnHeight As Long: btnHeight = 20
    Dim btnSpacing As Long: btnSpacing = 10
    Dim componentTop As Long: componentTop = 24  ' Top Y of static controls
    Dim textboxHeight As Long: textboxHeight = 24

    ' Determine columns based on item count
    Dim numCols As Long
    Select Case numItems
        Case 1 To 20: numCols = 2
        Case 21 To 30: numCols = 3
        Case 31 To 40: numCols = 4
        Case Else: numCols = Int((numItems - 1) / 10) + 1
    End Select

    ' Remove previous dynamic controls
    Dim ctl As control
    For Each ctl In Me.Controls
        If Left(ctl.Name, 7) = "txtItem" Or Left(ctl.Name, 6) = "lblItem" Or _
           ctl.Name = "btnOK" Or ctl.Name = "btnCancel" Then
            Me.Controls.Remove ctl.Name
        End If
    Next ctl

    ' Store item count
    Data_SharedValues.ItemCount = numItems
    ReDim Data_SharedValues.BatchValues(1 To numItems)

    ' Calculate rows per column (ceil division)
    Dim numRowsPerCol As Long
    numRowsPerCol = (numItems + numCols - 1) \ numCols

    ' Add dynamic txt/lbl controls below static area
    Dim i As Long, currentCol As Long, currentRow As Long
    Dim lbl As MSForms.Label, tb As MSForms.TextBox
    For i = 1 To numItems
        currentCol = (i - 1) \ numRowsPerCol
        currentRow = (i - 1) Mod numRowsPerCol

        Set lbl = Me.Controls.Add("Forms.Label.1", "lblItem" & i, True)
        lbl.Caption = i & "."
        lbl.Top = componentTop + textboxHeight + 8 + currentRow * rowHeight
        lbl.Left = 10 + currentCol * colWidth
        lbl.Width = 15

        Set tb = Me.Controls.Add("Forms.TextBox.1", "txtItem" & i, True)
        tb.Top = lbl.Top - 3
        tb.Left = lbl.Left + lbl.Width + 3
        tb.Width = 80
        tb.Text = ""
    Next i

    ' Adjust form size (width/height)
    Me.Width = 30 + numCols * colWidth
    Dim lastRowTop As Long
    lastRowTop = componentTop + textboxHeight + 8 + numRowsPerCol * rowHeight
    Me.Height = lastRowTop + btnHeight + formPadding

    ' Add OK/Cancel buttons at form bottom right
    Dim btnOK As MSForms.CommandButton, btnCancel As MSForms.CommandButton
    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    btnOK.Caption = "OK"
    btnOK.Width = btnWidth
    btnOK.Height = btnHeight
    btnOK.Left = Me.Width - btnWidth * 2 - btnSpacing * 2 - 15
    btnOK.Top = Me.Height - btnHeight - btnSpacing - 30

    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    btnCancel.Caption = "Cancel"
    btnCancel.Width = btnWidth
    btnCancel.Height = btnHeight
    btnCancel.Left = Me.Width - btnWidth - btnSpacing - 15
    btnCancel.Top = btnOK.Top

    ReDim ButtonHandlers(1 To 2)
    Set ButtonHandlers(1) = New Handler_BatchInputOKButton
    Set ButtonHandlers(1).Btn = btnOK
    Set ButtonHandlers(2) = New Handler_BatchInputCancelButton
    Set ButtonHandlers(2).Btn = btnCancel

    Me.Controls("txtItem1").SetFocus
End Sub

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
        ' Case 1: Both boundaries are empty ¡ú use B/D/E columns
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
        ' Case 4: Both boundaries have values ¡ú use these as the range
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
        ' Only one boundary has value ¡ú fill all boxes with fixed value
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
