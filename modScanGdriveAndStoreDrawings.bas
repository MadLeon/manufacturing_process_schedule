' modScanGdriveAndStoreDrawings.bas
Option Explicit

Sub ScanGdriveAndStoreDrawings()
    ' Scans G drive for PDF files and stores their info in the database
    ' Implements caching and batch insertion for better performance
    
    Dim dbPath As String, dbHandle As LongPtr, result As Long
    Dim sqlInsert As String, sqlSelect As String
    Dim fso As Object, folder As Object
    Dim gDrivePath As String
    Dim cache As Object ' For storing existing filenames
    Dim batch As Collection ' For batch insertion
    Dim stmtHandle As LongPtr
    
    ' 1. Set G drive path (modify as needed)
    gDrivePath = "G:\"
    If Right(gDrivePath, 1) <> "\" Then gDrivePath = gDrivePath & "\"
    
    ' 2. Validate G drive path
    If Dir(gDrivePath, vbDirectory) = "" Then
        MsgBox "G drive path is invalid or inaccessible: " & gDrivePath, vbCritical
        Exit Sub
    End If
    
    ' 3. Initialize SQLite
    result = SQLite3Initialize(ThisWorkbook.Path)
    If result <> SQLITE_INIT_OK Then
        MsgBox "SQLite3 initialization failed. Check if DLLs are present.", vbExclamation
        Exit Sub
    End If
    
    ' 4. Connect to database
    dbPath = ThisWorkbook.Path & "\jobs.db"
    result = SQLite3Open(dbPath, dbHandle)
    If result <> SQLITE_OK Then
        MsgBox "Failed to open database: " & SQLite3ErrMsg(dbHandle), vbExclamation
        SQLite3Free
        Exit Sub
    End If
    
    ' 5. Prepare SQL statements
    sqlInsert = "INSERT OR IGNORE INTO drawings (drawing_name, file_location) VALUES (?, ?)"
    result = SQLite3PrepareV2(dbHandle, sqlInsert, stmtHandle)
    If result <> SQLITE_OK Then
        MsgBox "Failed to prepare insert statement: " & SQLite3ErrMsg(dbHandle), vbExclamation
        SQLite3Close dbHandle
        SQLite3Free
        Exit Sub
    End If
    
    ' 6. Initialize cache and batch
    Set cache = CreateObject("Scripting.Dictionary")
    Set batch = New Collection
    LoadExistingFiles dbHandle, cache ' Preload existing filenames
    
    ' 7. Scan files
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next ' Skip inaccessible folders
    Set folder = fso.GetFolder(gDrivePath)
    ProcessFolder folder, fso, cache, batch, stmtHandle
    On Error GoTo 0
    
    ' 8. Insert remaining batch items
    If batch.Count > 0 Then
        InsertBatch batch, stmtHandle
    End If
    
    ' 9. Clean up
    SQLite3Finalize stmtHandle
    SQLite3Close dbHandle
    SQLite3Free
    Set fso = Nothing
    
    MsgBox "Scan complete. Processed " & batch.Count & " new files.", vbInformation
End Sub

Private Sub LoadExistingFiles(ByVal dbHandle As LongPtr, ByRef cache As Object)
    ' Loads existing filenames into cache
    Dim sql As String, stmtHandle As LongPtr, result As Long
    
    sql = "SELECT drawing_name FROM drawings"
    result = SQLite3PrepareV2(dbHandle, sql, stmtHandle)
    If result = SQLITE_OK Then
        Do While SQLite3Step(stmtHandle) = 100
            cache(Trim(SQLite3ColumnText(stmtHandle, 0))) = True
        Loop
        SQLite3Finalize stmtHandle
    End If
    Debug.Print "Preloaded " & cache.Count & " existing filenames"
End Sub

Private Sub ProcessFolder( _
    ByRef folder As Object, _
    ByRef fso As Object, _
    ByRef cache As Object, _
    ByRef batch As Collection, _
    ByRef stmtHandle As LongPtr)
    
    Dim file As Object, subfolder As Object
    Static fileCount As Long
    
    ' Process files in current folder
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".pdf" Then
            fileCount = fileCount + 1
            
            ' Skip if already exists
            If Not cache.exists(file.Name) Then
                cache.Add file.Name, True
                batch.Add Array(file.Name, file.Path)
                
                ' Insert in batches of 100
                If batch.Count >= 100 Then
                    InsertBatch batch, stmtHandle
                End If
            End If
        End If
    Next
    
    ' Process subfolders
    For Each subfolder In folder.SubFolders
        If subfolder.Attributes <> 2 Then ' Skip hidden folders
            ProcessFolder subfolder, fso, cache, batch, stmtHandle
        End If
    Next
End Sub

Private Sub InsertBatch( _
    ByRef batch As Collection, _
    ByRef stmtHandle As LongPtr)
    
    Dim entry As Variant, result As Long
    
    On Error Resume Next ' Skip any errors in batch
    For Each entry In batch
        SQLite3BindText stmtHandle, 1, entry(0) ' drawing_name
        SQLite3BindText stmtHandle, 2, entry(1) ' file_location
        result = SQLite3Step(stmtHandle)
        If result <> SQLITE_DONE Then
            Debug.Print "Insert failed for: " & entry(0) & " - " & SQLite3ErrMsg(0)
        End If
        SQLite3Reset stmtHandle
    Next
    On Error GoTo 0
    
    Debug.Print "Inserted batch of " & batch.Count & " files"
    Set batch = New Collection ' Reset batch
End Sub