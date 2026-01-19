Sub CheckSelectedCell()
    Dim rng As Range
    Dim sourceFilePath As String, destFilePath As String
    Dim wbSource As Workbook, wbDest As Workbook
    Dim colEValue As Variant, colFValue As Variant
    Dim fso As Object

    ' --- Configuration ---
    Const sourceFileName As String = "ver. 1.03.xlsm"  ' File name to copy
    sourceFilePath = "C:\Users\ee\Desktop\Dir History\" & sourceFileName
    Dim dirHistoryPath As String
    dirHistoryPath = "C:\Users\ee\Desktop\Dir History\"
    Dim destBasePath As String
    Dim colHValue As Variant
    
    ' Initialize FileSystemObject for directory operations
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' --- Check if a cell is selected ---
    If TypeName(Selection) <> "Range" Then
        Debug.Print "No cell selected"
        Exit Sub  ' Nothing selected
    End If

    ' Get the selected range
    Set rng = Selection

    ' --- Get the value from column D and H in the selected row ---
    Dim colDValue As Variant
    colDValue = ThisWorkbook.ActiveSheet.Cells(rng.row, 4).Value  ' Column D (column 4)
    colHValue = ThisWorkbook.ActiveSheet.Cells(rng.row, 8).Value  ' Column H (column 8)
    
    If Not IsEmpty(colDValue) Then  ' Check if column D has a value

        ' --- Determine destination base path based on column H value ---
        If IsEmpty(colHValue) Then
            MsgBox "Customer (Column H) is empty", vbExclamation, "Warning"
            Exit Sub
        End If
        
        Select Case CStr(colHValue)
            Case "Candu"
                destBasePath = "\\rtdnas2\QCReports\FINAL REPORTS\CANDU  ENERGY"
            Case "ATS"
                destBasePath = "\\rtdnas2\QCReports\FINAL REPORTS\ATS  Energy"
            Case "Kinectrics"
                destBasePath = "\\rtdnas2\QCReports\FINAL REPORTS\KINECTRICS INC"
            Case Else
                destBasePath = "\\rtdnas2\QCReports\FINAL REPORTS\CANDU  ENERGY"
                MsgBox "Unknown customer value: " & colHValue & vbNewLine & "Defaulting to CANDU ENERGY", vbExclamation, "Warning"
        End Select

        ' --- Get values from *current active sheet* for folder creation logic ---
        colFValue = ThisWorkbook.ActiveSheet.Cells(rng.row, 6).Value  ' Column F
        
        ' --- Build Destination File Path ---
        ' If column H is "Candu" and column F has a value, create subfolder path
        If CStr(colHValue) = "Candu" And Not IsEmpty(colFValue) Then
            Dim subfolderPath As String
            subfolderPath = destBasePath & "\" & CStr(colFValue)
            
            ' Create the subfolder if it doesn't exist
            If Not fso.FolderExists(subfolderPath) Then
                On Error Resume Next
                fso.CreateFolder subfolderPath
                On Error GoTo 0
                Debug.Print "Created folder: " & subfolderPath
            End If
            
            destFilePath = subfolderPath & "\" & colDValue & ".xlsm"
        Else
            destFilePath = destBasePath & "\" & colDValue & ".xlsm"  ' Copy name
        End If

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
        colEValue = ThisWorkbook.ActiveSheet.Cells(rng.row, 5).Value  ' Column E
        colGValue = ThisWorkbook.ActiveSheet.Cells(rng.row, 7).Value  ' Column G
        Dim colAValue As Variant, colBValue As Variant, colCValue As Variant
        colAValue = ThisWorkbook.ActiveSheet.Cells(rng.row, 1).Value  ' Column A
        colBValue = ThisWorkbook.ActiveSheet.Cells(rng.row, 2).Value  ' Column B
        colCValue = ThisWorkbook.ActiveSheet.Cells(rng.row, 3).Value  ' Column C

        ' --- Write to the destination workbook ---
        With wbDest.Sheets(1)
            .Range("H7").Value = colEValue
            .Range("H8").Value = colFValue
            .Range("C8").Value = colGValue
            .Range("C6").Value = UCase(CStr(colHValue))
            .Range("H6").Value = colCValue
            .Range("C7").Value = colAValue & " REV." & colBValue
        End With

        ' --- Leave the destination workbook open ---
        ' wbDest.Close SaveChanges:=True  ' Commented out to leave open
        Set wbDest = Nothing  ' Release the object, but keep the workbook open
        Set fso = Nothing  ' Release FileSystemObject
        Debug.Print "File copied and values updated."
    Else
        Debug.Print "Selected cell not in column D or is empty."
    End If
End Sub