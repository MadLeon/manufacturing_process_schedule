' --- Helper function to recursively search for PDF files starting with a given prefix ---
Sub SearchPDFFilesRecursive(folder As Object, colAValue As String, ByRef latestFile As Object, ByRef latestDate As Date, fso As Object)
    Dim file As Object
    Dim subFolder As Object
    
    ' Search in current folder
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".pdf" Then
            If Left(file.Name, Len(colAValue)) = colAValue Then
                Debug.Print "Found matching file: " & file.Path & " (Modified: " & file.DateLastModified & ")"
                ' Keep track of the most recently modified file
                If file.DateLastModified > latestDate Then
                    latestDate = file.DateLastModified
                    Set latestFile = file
                End If
            End If
        End If
    Next file
    
    ' Recursively search in subfolders
    On Error Resume Next
    For Each subFolder In folder.SubFolders
        SearchPDFFilesRecursive subFolder, colAValue, latestFile, latestDate, fso
    Next subFolder
    On Error GoTo 0
End Sub

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

Sub OpenPDFFileInCompanyFolder()
    Dim rng As Range
    Dim filePath As String
    Dim objShell As Object
    Dim colHValue As Variant
    Dim fso As Object
    Dim destBasePath As String
    Dim pdfFileName As String
    Dim currentRow As Long
    Dim colAValue As String, colGValue As String, colCValue As String

    ' Initialize FileSystemObject for directory operations
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' --- Check if a cell is selected ---
    If TypeName(Selection) <> "Range" Then
        Debug.Print "No cell selected"
        Exit Sub
    End If

    ' Get the selected range
    Set rng = Selection
    currentRow = rng.row

    ' --- Get the value from column H in the selected row ---
    colHValue = ThisWorkbook.ActiveSheet.Cells(currentRow, 8).Value  ' Column H (column 8)
    
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

    ' --- Generate PDF filename using GeneratePDFStringValue function ---
    Dim colFValue As Variant
    colFValue = ThisWorkbook.ActiveSheet.Cells(currentRow, 6).Value  ' Column F

    ' Generate the PDF filename using the function from mod_GenerateString (new version)
    pdfFileName = GeneratePDFStringValue(currentRow, ThisWorkbook.ActiveSheet)

    ' Clean the filename to remove illegal characters
    Dim cleanedPDFFileName As String
    cleanedPDFFileName = CleanFileName(pdfFileName)
    
    ' Build the complete file path with .pdf extension
    ' Include column F as subfolder for Candu customer
    Dim basePath As String
    If CStr(colHValue) = "Candu" And Not IsEmpty(colFValue) Then
        basePath = destBasePath & "\" & CStr(colFValue)
    Else
        basePath = destBasePath
    End If
    
    filePath = basePath & "\" & cleanedPDFFileName & ".pdf"

    ' --- Debug Info ---
    Debug.Print "========== OPEN PDF DEBUG INFO =========="
    Debug.Print "Time: " & Format(Now, "HH:MM:SS")
    Debug.Print "Generated PDF filename (new version): " & pdfFileName
    Debug.Print "Cleaned PDF filename: " & cleanedPDFFileName
    Debug.Print "File path: " & filePath
    Debug.Print "File path length: " & Len(filePath)
    Debug.Print "File exists (FSO): " & fso.FileExists(filePath)
    Debug.Print "=========================================="
    
    ' --- Try to find and open the file ---
    Dim fileFound As Boolean
    fileFound = False
    
    ' Get colAValue early for all search steps
    colAValue = Trim(ThisWorkbook.ActiveSheet.Cells(currentRow, 1).Value)  ' Column A
    
    ' Step 1: Try searching for new version PDF in Bubble Drawing folder
    Dim bubbleDrawingPath As String
    bubbleDrawingPath = destBasePath & "\Bubble Drawing"
    
    Debug.Print "Step 1: Searching for new version PDF in Bubble Drawing folder: " & bubbleDrawingPath
    
    Dim bubbleNewVersionPath As String
    bubbleNewVersionPath = bubbleDrawingPath & "\" & cleanedPDFFileName & ".pdf"
    
    Debug.Print "Step 1: Looking for: " & bubbleNewVersionPath
    
    If fso.FileExists(bubbleNewVersionPath) Then
        filePath = bubbleNewVersionPath
        fileFound = True
        Debug.Print "Found new version PDF in Bubble Drawing folder: " & cleanedPDFFileName & ".pdf"
    Else
        Debug.Print "New version PDF not found in Bubble Drawing folder"
    End If
    
    ' Step 2: Try new version in basePath
    If fso.FileExists(filePath) Then
        fileFound = True
        Debug.Print "Found new version PDF file in basePath"
    Else
        ' Step 3: Try old version in basePath
        colAValue = Trim(ThisWorkbook.ActiveSheet.Cells(currentRow, 1).Value)  ' Column A
        colGValue = Trim(ThisWorkbook.ActiveSheet.Cells(currentRow, 7).Value)  ' Column G
        colCValue = Trim(ThisWorkbook.ActiveSheet.Cells(currentRow, 3).Value)  ' Column C
        
        Dim oldVersionFileName As String
        oldVersionFileName = colAValue
        
        ' Format and append colGValue
        If colGValue <> "" Then
            colGValue = UCase(colGValue)
            Dim wordArray() As String, word As Variant, formattedWord As String
            wordArray = Split(colGValue, " ")
            
            For Each word In wordArray
                If Len(word) > 0 Then
                    Dim firstChar As String: firstChar = UCase(Left(word, 1))
                    Dim restOfWord As String: restOfWord = LCase(Mid(word, 2))
                    formattedWord = firstChar & restOfWord
                    oldVersionFileName = oldVersionFileName & " " & formattedWord
                End If
            Next word
        End If
        
        oldVersionFileName = oldVersionFileName & "-ballooned @" & colCValue
        Dim cleanedOldFileName As String
        cleanedOldFileName = CleanFileName(oldVersionFileName)
        
        Dim oldFilePath As String
        oldFilePath = basePath & "\" & cleanedOldFileName & ".pdf"
        
        Debug.Print "Old version filename: " & oldVersionFileName
        Debug.Print "Old version cleaned: " & cleanedOldFileName
        Debug.Print "Old version file path: " & oldFilePath
        Debug.Print "Old version file exists: " & fso.FileExists(oldFilePath)
        
        If fso.FileExists(oldFilePath) Then
            fileFound = True
            filePath = oldFilePath
            Debug.Print "Found old version PDF file"
        End If
    End If
    
    If Not fileFound Then
        MsgBox "PDF file not found (tried Bubble Drawing, basePath new/old version)!", vbExclamation, "File Not Found"
        Set fso = Nothing
        Exit Sub
    End If

    ' --- Open the PDF file using default application ---
    On Error Resume Next
    Set objShell = CreateObject("Shell.Application")
    objShell.ShellExecute filePath
    If Err.Number <> 0 Then
        Debug.Print "Open PDF ERROR: " & Err.Number & " - " & Err.description
        MsgBox "Failed to open PDF file: " & Err.description, vbCritical, "Error"
        Err.Clear
        On Error GoTo 0
        Set fso = Nothing
        Set objShell = Nothing
        Exit Sub
    End If
    On Error GoTo 0
    
    Set objShell = Nothing
    Set fso = Nothing
    Debug.Print "PDF file opened successfully: " & filePath
End Sub

