' modScanGdriveAndStoreDrawings.bas
Option Explicit

Sub ScanGdriveAndStoreDrawings()
    ' This subroutine scans the G drive, finds all PDF files, and stores their information in the database.
    ' Table structure:
    '   - drawing_name (TEXT): Drawing file name
    '   - drawing_number (TEXT): Drawing number (initially empty)
    '   - file_location (TEXT): File location (for hyperlinks)

    Dim dbPath As String, dbHandle As LongPtr, result As Long, stmtHandle As LongPtr
    Dim sqlInsert As String
    Dim fso As Object, folder As Object, file As Object
    Dim gDrivePath As String ' Replace with your G drive path
    Dim drawingName As String, fileLocation As String

    ' 1. Set the G drive path  (Please modify according to your actual situation)
    gDrivePath = "G:\"  '  <--  **Please modify to your actual G drive path**
    If Right(gDrivePath, 1) <> "\" Then gDrivePath = gDrivePath & "\"  ' Ensure the path ends with "\"

    ' 2. Check if the G drive path is valid
    If Dir(gDrivePath, vbDirectory) = "" Then
        MsgBox "G drive path is invalid or inaccessible: " & gDrivePath, vbCritical
        Exit Sub
    End If

    ' 3. Initialize SQLite3
    result = SQLite3Initialize(ThisWorkbook.Path)
    If result <> SQLITE_INIT_OK Then
        MsgBox "SQLite3 initialization failed. Please check if SQLite3.dll and SQLite3_StdCall.dll are in the same directory.": Exit Sub
    End If

    ' 4. Connect to the database
    dbPath = ThisWorkbook.Path & "\jobs.db"
    result = SQLite3Open(dbPath, dbHandle)
    If result <> SQLITE_OK Then
        MsgBox "Failed to open the database jobs.db. Please check if the file exists and the permissions.": SQLite3Free: Exit Sub
    End If

    ' 5. Prepare the SQL insert statement
    sqlInsert = "INSERT INTO drawings (drawing_name, file_location) VALUES (?, ?)"
    result = SQLite3PrepareV2(dbHandle, sqlInsert, stmtHandle)
    If result <> SQLITE_OK Then
        MsgBox "Failed to prepare the insert statement: " & SQLite3ErrMsg(dbHandle): SQLite3Close dbHandle: SQLite3Free: Exit Sub
    End If

    ' 6. Use FileSystemObject to traverse the G drive
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next  ' Ignore errors for inaccessible folders
    Set folder = fso.GetFolder(gDrivePath)

    ' Traverse folders and subfolders
    ProcessFolder folder, fso, stmtHandle

    On Error GoTo 0  ' Restore error handling
    ' 7. Clean up after completion
    SQLite3Finalize stmtHandle
    result = SQLite3Close(dbHandle)
    If result <> SQLITE_OK Then
        Debug.Print "Error closing the database connection: " & SQLite3ErrMsg(dbHandle)
    End If

    SQLite3Free
    Set fso = Nothing
    Set folder = Nothing
    Set file = Nothing
    MsgBox "G drive scanning complete, drawing information has been stored in the database.", vbInformation

End Sub

' Recursive function to process folders
Sub ProcessFolder(ByRef folder As Object, ByRef fso As Object, ByRef stmtHandle As LongPtr)
    Dim subfolder As Object, file As Object
    Dim drawingName As String, fileLocation As String
    Dim result As Long
    
    ' Process files in the current folder
    For Each file In folder.Files
        If LCase(Right(file.Name, 3)) = "pdf" Then ' Check if the file is a PDF
            drawingName = file.Name
            fileLocation = file.Path

            ' Insert into the database
            SQLite3BindText stmtHandle, 1, drawingName
            SQLite3BindText stmtHandle, 2, fileLocation
            result = SQLite3Step(stmtHandle)
            If result <> SQLITE_DONE Then
                Debug.Print "Error inserting data: " & SQLite3ErrMsg(0) & ", File: " & drawingName & ", Path: " & fileLocation
            End If
            SQLite3Reset stmtHandle ' Reset the statement to prepare for the next insertion
        End If
    Next file

    ' Recursively process subfolders
    For Each subfolder In folder.subfolders
        If subfolder.Attributes <> 2 Then  ' Exclude hidden folders (if needed)
            ProcessFolder subfolder, fso, stmtHandle
        End If
    Next subfolder
End Sub
