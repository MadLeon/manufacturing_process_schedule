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

    ' 6. Put the output string to both col I and clipboard
    Cells(currentRow, 9).Value = outputString ' Column I

    ' Copy to clipboard
    With New DataObject
        .SetText outputString
        .PutInClipboard
    End With

    MsgBox "Output string generated and copied to clipboard!", vbInformation
End Sub
