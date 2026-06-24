Sub GenerateDiscusString()
    ' This sub generates an output string based on values from columns A, G, and C,
    ' and writes it to column I and the clipboard.

    Dim selectedCell As Range
    Dim currentRow As Long
    Dim colAValue As String, colBValue As String, colGValue As String, colCValue As String
    Dim outputString As String
    Dim wordArray() As String, word As Variant
    Dim formattedWord As String
    Dim i As Long

    ' 1. Record the current selected cell row as the current row
    Set selectedCell = Application.Selection.Cells(1, 1) ' Get the first selected cell
    currentRow = selectedCell.row

    ' 2. Get values from columns A, G, and C
    colAValue = Trim(Cells(currentRow, 1).Value) ' Column A
    colBValue = Trim(Cells(currentRow, 2).Value) ' Column B
    colGValue = Trim(Cells(currentRow, 7).Value) ' Column G
    colCValue = Trim(Cells(currentRow, 3).Value) ' Column C

    ' 3. Build the output string
    outputString = colAValue
    
    ' 3.1 Append " Rev{col B}"
    outputString = outputString & " Rev" & colBValue

    ' 4. Append " {col G}" with formatted colGValue
    If colGValue <> "" Then
        ' Uppercase the entire string
        colGValue = UCase(colGValue)
        
        ' Split the string into words
        wordArray = Split(colGValue, " ")
        
        ' Format each word
        For Each word In wordArray
            ' Ensure the word is not empty
            If Len(word) > 0 Then
                ' Extract the first character and convert it to uppercase
                Dim firstChar As String: firstChar = UCase(Left(word, 1))
                
                ' Extract the remaining characters and convert them to lowercase
                Dim restOfWord As String: restOfWord = LCase(Mid(word, 2))
                
                ' Concatenate the first character with the remaining characters
                formattedWord = firstChar & restOfWord
                
                ' Append the formatted word to the output string
                outputString = outputString & " " & formattedWord
            End If
        Next word
    End If

    ' 5. Append " @{col C}"
    outputString = outputString & " @" & colCValue

    ' Copy to clipboard
    With New DataObject
        .SetText outputString
        .PutInClipboard
    End With
    
    ' 6. Wite current system time to column J (format: h.mm)
    ' Cells(currentRow, 10).Value = Format(Time, "h.mm")
End Sub

Function GeneratePDFStringValue(currentRow As Long, Optional ws As Worksheet) As String
    ' This function generates a PDF filename string based on values from columns A, B, G, and C
    ' Parameters:
    '   currentRow: The row number to extract values from
    '   ws: Optional worksheet reference (defaults to active sheet)
    ' Returns: The formatted PDF string

    Dim colAValue As String, colBValue As String, colGValue As String, colCValue As String
    Dim outputString As String
    Dim wordArray() As String, word As Variant
    Dim formattedWord As String

    ' Use provided worksheet or default to active sheet
    If ws Is Nothing Then
        Set ws = ThisWorkbook.ActiveSheet
    End If

    ' Get values from columns A, B, G, and C
    colAValue = Trim(ws.Cells(currentRow, 1).Value) ' Column A
    colBValue = Trim(ws.Cells(currentRow, 2).Value) ' Column B
    colGValue = Trim(ws.Cells(currentRow, 7).Value) ' Column G
    colCValue = Trim(ws.Cells(currentRow, 3).Value) ' Column C

    ' Build the output string: "{col A} Rev{col B} {col G}-ballooned"
    outputString = colAValue
    
    ' Append Rev{col B}
    outputString = outputString & " Rev" & colBValue
    
    ' Append " {col G}" with formatted colGValue
    If colGValue <> "" Then
        ' Uppercase the entire string
        colGValue = UCase(colGValue)
        
        ' Split the string into words
        wordArray = Split(colGValue, " ")
        
        ' Format each word
        For Each word In wordArray
            ' Ensure the word is not empty
            If Len(word) > 0 Then
                ' Extract the first character and convert it to uppercase
                Dim firstChar As String: firstChar = UCase(Left(word, 1))
                
                ' Extract the remaining characters and convert them to lowercase
                Dim restOfWord As String: restOfWord = LCase(Mid(word, 2))
                
                ' Concatenate the first character with the remaining characters
                formattedWord = firstChar & restOfWord
                
                ' Append the formatted word to the output string
                outputString = outputString & " " & formattedWord
            End If
        Next word
    End If
    
    outputString = outputString & "-ballooned"
    GeneratePDFStringValue = outputString
End Function

Sub GeneratePDFString()
    ' This sub generates an output string based on values from columns A, G, and C
    ' and writes it to column J and the clipboard.

    Dim selectedCell As Range
    Dim currentRow As Long
    Dim outputString As String

    ' 1. Record the current selected cell row as the current row
    Set selectedCell = Application.Selection.Cells(1, 1) ' Get the first selected cell
    currentRow = selectedCell.row

    ' 2. Generate the PDF string using the function
    outputString = GeneratePDFStringValue(currentRow)

    ' 3. Copy to clipboard
    With New DataObject
        .SetText outputString
        .PutInClipboard
    End With
End Sub
