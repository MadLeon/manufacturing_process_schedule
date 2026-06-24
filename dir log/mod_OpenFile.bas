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

Sub OpenFile()
    Dim rng As Range
    Dim filePath As String
    Dim wb As Workbook
    Dim colDValue As Variant, colFValue As Variant, colHValue As Variant
    Dim fso As Object
    Dim destBasePath As String

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
        
        ' --- Build File Path ---
        ' Clean the filename to remove illegal characters
        Dim cleanedColDValue As String
        cleanedColDValue = CleanFileName(CStr(colDValue))
        
        If CStr(colHValue) = "Candu" And Not IsEmpty(colFValue) Then
            Dim subfolderPath As String
            subfolderPath = destBasePath & "\" & CStr(colFValue)
            
            Debug.Print "Subfolder logic triggered for Candu"
            Debug.Print "Subfolder path: " & subfolderPath
            
            filePath = subfolderPath & "\" & cleanedColDValue & ".xlsm"
        Else
            filePath = destBasePath & "\" & cleanedColDValue & ".xlsm"
        End If

        ' --- Debug Info ---
        Debug.Print "========== OPEN FILE DEBUG INFO =========="
        Debug.Print "Time: " & Format(Now, "HH:MM:SS")
        Debug.Print "File path: " & filePath
        Debug.Print "File path length: " & Len(filePath)
        Debug.Print "File exists (FSO): " & fso.FileExists(filePath)
        Debug.Print "=========================================="
        
        ' --- Check if file exists ---
        If Not fso.FileExists(filePath) Then
            MsgBox "File not found: " & filePath, vbCritical, "Error"
            Set fso = Nothing
            Exit Sub
        End If

        ' --- Open the file ---
        On Error Resume Next
        Set wb = Workbooks.Open(filePath)
        If Err.Number <> 0 Then
            Debug.Print "Open file ERROR: " & Err.Number & " - " & Err.description
            MsgBox "Failed to open file: " & Err.description, vbCritical, "Error"
            Err.Clear
            On Error GoTo 0
            Set fso = Nothing
            Exit Sub
        End If
        On Error GoTo 0
        
        If wb Is Nothing Then
            Debug.Print "Could not open the file: " & filePath
            Set fso = Nothing
            Exit Sub
        End If

        Set wb = Nothing  ' Release the object, but keep the workbook open
        Set fso = Nothing  ' Release FileSystemObject
        Debug.Print "File opened successfully: " & filePath
    Else
        Debug.Print "Selected cell not in column D or is empty."
    End If
End Sub
