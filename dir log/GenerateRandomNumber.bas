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
    Dim resultVal As Double
    Dim resultStr As String
    Dim unitStr As String
    Dim decimalPlaces As Integer
    Dim decimalPlacesD As Integer, decimalPlacesE As Integer
    
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
        ' 当上下tolerance同时为1, -1时，硬编码其为.1, -.1
        ' -------------------------------------------------
        If upperTol = 1 And lowerTol = -1 Then
            upperTol = 0.1
            lowerTol = -0.1
        End If
        
        ' -------------------------------------------------
        ' Determine decimal precision from Upper and Lower Tolerance
        ' -------------------------------------------------
        ' Get decimal places from the original text (before numeric conversion)
        ' This preserves trailing zeros like .000 or .010
        Dim cellDText As String, cellEText As String
        cellDText = CStr(ws.Cells(rowNum, 4).Text)
        cellEText = CStr(ws.Cells(rowNum, 5).Text)
        
        ' Get decimal places from Column D (Upper Tolerance) - from original text
        decimalPlacesD = CountDecimalPlaces(upperTol)
        If InStr(cellDText, ".") > 0 Then
            decimalPlacesD = Len(cellDText) - InStr(cellDText, ".")
        End If
        
        ' Get decimal places from Column E (Lower Tolerance) - from original text
        decimalPlacesE = CountDecimalPlaces(lowerTol)
        If InStr(cellEText, ".") > 0 Then
            decimalPlacesE = Len(cellEText) - InStr(cellEText, ".")
        End If
        
        ' If either is a fraction, set decimal places to 2
        If IsFraction(cellDText) Or IsFraction(cellEText) Then
            decimalPlaces = 2
        Else
            ' Use the greater decimal places
            decimalPlaces = IIf(decimalPlacesD > decimalPlacesE, decimalPlacesD, decimalPlacesE)
            
            ' Limit decimal places to 4
            If decimalPlaces > 4 Then
                decimalPlaces = 4
            End If
        End If
        
        ' -------------------------------------------------
        ' Calculate min and max bounds
        ' -------------------------------------------------
        minVal = idealVal + lowerTol
        maxVal = idealVal + upperTol
        
        ' -------------------------------------------------
        ' Generate random value within range
        ' -------------------------------------------------
        Randomize
        If maxVal = minVal Then
            resultVal = minVal
        Else
            resultVal = minVal + (maxVal - minVal) * Rnd
        End If
        
        ' Round to correct decimal places
        resultVal = WorksheetFunction.Round(resultVal, decimalPlaces)
        
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