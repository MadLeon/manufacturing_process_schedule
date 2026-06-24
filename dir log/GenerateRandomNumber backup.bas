' =========================================================
' Module: Function_GenerateRandomNumber
' =========================================================

Sub GenerateRandomNumber()
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim rowNum As Long
    Dim idealVal As Double
    Dim upperTol As Double
    Dim lowerTol As Double
    Dim minVal As Double, maxVal As Double
    Dim stepSize As Double
    Dim possibleVals() As Double
    Dim numSteps As Long, randIndex As Long
    Dim resultVal As Double
    Dim resultStr As String
    Dim unitStr As String
    Dim decimalPlaces As Integer
    Dim tempStr As String
    Dim i As Long
    
    Set ws = ActiveSheet
    Set rng = Selection
    
    ' -------------------------------------------------
    ' Ensure all selected cells are in column F
    ' -------------------------------------------------
    For Each cell In rng
        If cell.Column <> 6 Then
            MsgBox "Please select only cells in column F.", vbExclamation
            Exit Sub
        End If
    Next cell
    
    ' -------------------------------------------------
    ' Loop through each selected cell
    ' -------------------------------------------------
    For Each cell In rng
        rowNum = cell.Row
        
        ' Skip if Ideal (B) or Upper Tolerance (D) is empty
        If IsEmpty(ws.Cells(rowNum, 2)) Or IsEmpty(ws.Cells(rowNum, 4)) Then GoTo NextCell
        
        ' -------------------------------------------------
        ' Read input values
        ' -------------------------------------------------
        idealVal = ToNumber(ws.Cells(rowNum, 2).Value)
        upperTol = ToNumber(ws.Cells(rowNum, 4).Value)

        If IsEmpty(ws.Cells(rowNum, 5)) Then
            lowerTol = 0
        Else
            lowerTol = ToNumber(ws.Cells(rowNum, 5).Value)
        End If
        
        ' -------------------------------------------------
        ' Determine decimal precision from Upper Tolerance
        ' -------------------------------------------------
        tempStr = CStr(ws.Cells(rowNum, 4).Text)
        If InStr(tempStr, ".") > 0 Then
            decimalPlaces = Len(tempStr) - InStr(tempStr, ".") + 1
        Else
            decimalPlaces = 0
        End If
        
        stepSize = 1 / (10 ^ decimalPlaces)
        
        ' -------------------------------------------------
        ' Calculate min and max bounds
        ' -------------------------------------------------
        minVal = idealVal + lowerTol
        maxVal = idealVal + upperTol
        
        ' Count the number of steps within range
        numSteps = Round((maxVal - minVal) / stepSize)
        ReDim possibleVals(0 To numSteps)
        
        ' Generate all possible values
        For i = 0 To numSteps
            possibleVals(i) = WorksheetFunction.Round(minVal + i * stepSize, decimalPlaces)
        Next i
        
        ' -------------------------------------------------
        ' Randomly pick one of the possible values
        ' -------------------------------------------------
        Randomize
        randIndex = Int((UBound(possibleVals) - LBound(possibleVals) + 1) * Rnd + LBound(possibleVals))
        resultVal = possibleVals(randIndex)
        
        ' -------------------------------------------------
        ' Format result string with correct decimal places
        ' -------------------------------------------------
        If decimalPlaces > 0 Then
            resultStr = WorksheetFunction.Text(resultVal, "0." & String(decimalPlaces, "0"))
        Else
            resultStr = CStr(resultVal)
        End If
        
        ' -------------------------------------------------
        ' If C11 = "INCH" and result starts with "0.", remove leading zero
        ' -------------------------------------------------
        unitStr = LCase(Trim(ws.Range("C11").Value))
        If InStr(unitStr, "inch") > 0 Then
            If Left(resultStr, 2) = "0." Then
                resultStr = Mid(resultStr, 2)
            End If
        End If
        
        ' -------------------------------------------------
        ' Write final result to column F
        ' -------------------------------------------------
        ws.Cells(rowNum, 6).Value = resultStr
        
NextCell:
    Next cell
End Sub

