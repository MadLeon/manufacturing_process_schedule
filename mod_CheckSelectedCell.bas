Sub CheckSelectedCell()
    Dim rng As Range
    Dim sourceFilePath As String, destFilePath As String
    Dim wbSource As Workbook, wbDest As Workbook
    Dim colEValue As Variant, colFValue As Variant

    ' --- Configuration ---
    Const sourceFileName As String = "ver. 1.03.xlsm"  ' File name to copy
    sourceFilePath = "C:\Users\ee\Desktop\Dir History\" & sourceFileName
    Dim dirHistoryPath As String
    dirHistoryPath = "C:\Users\ee\Desktop\Dir History\"

    ' --- Check if a cell is selected ---
    If TypeName(Selection) <> "Range" Then
        Debug.Print "No cell selected"
        Exit Sub  ' Nothing selected
    End If

    ' Get the selected range
    Set rng = Selection

    ' --- Check if the selected cell is in column D and has a value ---
    If rng.Column = 4 And Not IsEmpty(rng.value) Then  ' Column D is column 4

        ' --- Build Destination File Path ---
        destFilePath = dirHistoryPath & "\" & rng.value & ".xlsm"  ' Copy name

        ' --- Copy the source file ---
        On Error Resume Next
        Kill destFilePath  ' Delete destination if it exists
        On Error GoTo 0
        FileCopy sourceFilePath, destFilePath

        ' --- Open the copied file ---
        Set wbDest = Workbooks.Open(destFilePath)
        If wbDest Is Nothing Then
            Debug.Print "Could not open the destination file: " & destFilePath
            Exit Sub
        End If

        ' --- Get values from *current active sheet* ---
        colEValue = ThisWorkbook.ActiveSheet.Cells(rng.row, 5).value  ' Column E
        colFValue = ThisWorkbook.ActiveSheet.Cells(rng.row, 6).value  ' Column F

        Debug.Print "colEValue = " & colEValue
        Debug.Print "colFValue = " & colFValue

        ' --- Write to the destination workbook ---
        With wbDest.Sheets(1)
            .Range("H7").value = colEValue
            .Range("H8").value = colFValue
            Debug.Print "H7 value set to: " & colEValue
            Debug.Print "H8 value set to: " & colFValue
        End With

        ' --- Leave the destination workbook open ---
        ' wbDest.Close SaveChanges:=True  ' Commented out to leave open
        Set wbDest = Nothing  ' Release the object, but keep the workbook open
        Debug.Print "File copied and values updated."
    Else
        Debug.Print "Selected cell not in column D or is empty."
    End If
End Sub
