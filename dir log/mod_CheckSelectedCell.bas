' --- Helper function to clean illegal filename characters ---
Function CleanFileName(fileName As String) As String
    Dim result As String
    result = fileName
    ' Replace illegal filename characters (but NOT backslash to preserve path structure)
    result = Replace(result, "<", "_")
    result = Replace(result, ">", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, """", "_")
    result = Replace(result, "/", "_")
    ' Do NOT replace backslash - it's used for paths
    result = Replace(result, "|", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, "*", "_")
    CleanFileName = result
End Function

Sub CheckSelectedCell()
    Dim rng As Range
    Dim sourceFilePath As String, destFilePath As String
    Dim wbSource As Workbook, wbDest As Workbook
    Dim colEValue As Variant, colFValue As Variant
    Dim fso As Object

    ' --- Configuration ---
    Const sourceFileName As String = "DIR Template.xlsm"  ' File name to copy
    sourceFilePath = "\\rtdnas2\QCReports\FINAL REPORTS\" & sourceFileName
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
        ' Clean the filename to remove illegal characters
        Dim cleanedColDValue As String
        cleanedColDValue = CleanFileName(CStr(colDValue))
        
        If CStr(colHValue) = "Candu" And Not IsEmpty(colFValue) Then
            Dim subfolderPath As String
            subfolderPath = destBasePath & "\" & CStr(colFValue)
            
            Debug.Print "Subfolder logic triggered for Candu"
            Debug.Print "Subfolder path: " & subfolderPath
            Debug.Print "Subfolder exists before creation: " & fso.FolderExists(subfolderPath)
            
            ' Create the subfolder if it doesn't exist
            If Not fso.FolderExists(subfolderPath) Then
                On Error Resume Next
                fso.CreateFolder subfolderPath
                Dim createErr As Integer
                createErr = Err.Number
                On Error GoTo 0
                Debug.Print "Attempted to create subfolder. Error code: " & createErr
                Debug.Print "Subfolder exists after creation attempt: " & fso.FolderExists(subfolderPath)
            Else
                Debug.Print "Subfolder already exists"
            End If
            
            destFilePath = subfolderPath & "\" & cleanedColDValue & ".xlsm"
        Else
            destFilePath = destBasePath & "\" & cleanedColDValue & ".xlsm"  ' Copy name
        End If

        ' --- Copy the source file ---
        Debug.Print "========== COPY DEBUG INFO =========="
        Debug.Print "Time: " & Format(Now, "HH:MM:SS")
        Debug.Print "Source file path: " & sourceFilePath
        Debug.Print "Source file path length: " & Len(sourceFilePath)
        Debug.Print "Source file exists (FSO): " & fso.FileExists(sourceFilePath)
        Debug.Print "Destination path: " & destFilePath
        Debug.Print "Destination path length: " & Len(destFilePath)
        Dim destFolderPath As String
        destFolderPath = Left(destFilePath, InStrRev(destFilePath, "\") - 1)
        Debug.Print "Destination folder path: " & destFolderPath
        Debug.Print "Destination folder path length: " & Len(destFolderPath)
        Debug.Print "Destination folder exists (FSO): " & fso.FolderExists(destFolderPath)
        Debug.Print "Destination base folder exists (FSO): " & fso.FolderExists(destBasePath)
        Debug.Print "======================================"
        
        ' --- Check if destination file already exists ---
        If fso.FileExists(destFilePath) Then
            MsgBox "Destination file already exists: " & destFilePath & vbNewLine & "Operation cancelled.", vbExclamation, "File Already Exists"
            Set fso = Nothing
            Exit Sub
        End If
        
        On Error Resume Next
        Kill destFilePath  ' Delete destination if it exists
        Err.Clear
        On Error GoTo 0
        
        ' Use FileCopy for better network path compatibility
        On Error Resume Next
        FileCopy sourceFilePath, destFilePath
        If Err.Number <> 0 Then
            Debug.Print "FileCopy ERROR: " & Err.Number & " - " & Err.description
            Debug.Print "Error Time: " & Format(Now, "HH:MM:SS")
            Debug.Print "Source file exists at error time: " & (Dir(sourceFilePath) <> "")
            MsgBox "Copy failed: " & Err.description, vbCritical, "Error"
            Err.Clear
            On Error GoTo 0
            Set fso = Nothing
            Exit Sub
        End If
        On Error GoTo 0
        Debug.Print "File copied successfully to: " & destFilePath
        
        ' Add delay to ensure file is released from network
        Application.Wait (Now + TimeValue("0:00:01"))

        ' --- Open the copied file ---
        On Error Resume Next
        Set wbDest = Workbooks.Open(destFilePath)
        If Err.Number <> 0 Then
            Debug.Print "Open file ERROR: " & Err.Number & " - " & Err.description
            Err.Clear
            On Error GoTo 0
        End If
        On Error GoTo 0
        
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
            .Range("C7").Value = colAValue & " REV. " & colBValue
            .Range("C11").Value = "INCH"
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

