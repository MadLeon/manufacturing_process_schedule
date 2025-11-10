' modScanGdriveAndStoreDrawings.bas
Option Explicit

Sub ScanGdriveAndStoreDrawings()
    ' This subroutine scans the G drive, finds all PDF files, and stores their information in the database.
    ' Table structure:
    '   - drawing_name (TEXT): Drawing file name
    '   - drawing_number (TEXT): Drawing number (initially empty)
    '   - file_location (TEXT): File location (for hyperlinks)

    Dim fso As Object, folder As Object, file As Object
    Dim gDrivePath As String ' Replace with your G drive path
    Dim dbPath As String
    Dim sqlInsert As String

    ' 1. Set the G drive path  (Please modify according to your actual situation)
    gDrivePath = "G:\"  '  <--  **Please modify to your actual G drive path**
    If Right(gDrivePath, 1) <> "\" Then gDrivePath = gDrivePath & "\"  ' Ensure the path ends with "\"

    ' 2. Check if the G drive path is valid
    If Dir(gDrivePath, vbDirectory) = "" Then
        MsgBox "G drive path is invalid or inaccessible: " & gDrivePath, vbCritical
        Exit Sub
    End If

    ' 3. Initialize SQLite database
    dbPath = ThisWorkbook.Path & "\jobs.db"
    If Not InitializeSQLite(dbPath) Then
        MsgBox "Failed to initialize SQLite database": Exit Sub
    End If

    ' 4. Prepare the SQL insert statement
    sqlInsert = "INSERT INTO drawings (drawing_name, file_location) VALUES (?, ?)"
    
    ' 5. Use FileSystemObject to traverse the G drive
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next  ' Ignore errors for inaccessible folders
    Set folder = fso.GetFolder(gDrivePath)

    ' Traverse folders and subfolders
    ProcessFolder folder, fso, sqlInsert

    On Error GoTo 0  ' Restore error handling
    
    ' 6. Clean up after completion
    CloseSQLite
    Set fso = Nothing
    Set folder = Nothing
    Set file = Nothing
    MsgBox "G drive scanning complete, drawing information has been stored in the database.", vbInformation
End Sub

' Recursive function to process folders
Sub ProcessFolder(ByRef folder As Object, ByRef fso As Object, ByVal sqlInsert As String)
    Dim subfolder As Object, file As Object
    Dim drawingName As String, fileLocation As String
    Dim stmtHandle As LongPtr, result As Long
    
    ' Process files in the current folder
    For Each file In folder.Files
        If LCase(Right(file.Name, 3)) = "pdf" Then ' Check if the file is a PDF
            drawingName = file.Name
            fileLocation = file.Path

            ' Build the SQL string with values
            Dim safeDrawingName As String, safeFileLocation As String
            safeDrawingName = Replace(drawingName, "'", "''")
            safeFileLocation = Replace(fileLocation, "'", "''")
            Dim sql As String
            sql = "INSERT INTO drawings (drawing_name, file_location) VALUES ('" & safeDrawingName & "', '" & safeFileLocation & "')"

            If Not ExecuteNonQuery(sql) Then
                Debug.Print "Error inserting data: File: " & drawingName & ", Path: " & fileLocation
            End If
        End If
    Next file

    ' Recursively process subfolders
    For Each subfolder In folder.subfolders
        If subfolder.Attributes <> 2 Then  ' Exclude hidden folders (if needed)
            ProcessFolder subfolder, fso, sqlInsert
        End If
    Next subfolder
End Sub