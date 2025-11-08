' modCreateDrawingTable.bas
Option Explicit

Sub CreateDrawingTable()
    ' This module is used to create or initialize the drawing information table.
    ' Table structure:
    '   - drawing_name (TEXT): Drawing file name
    '   - drawing_number (TEXT): Drawing number (initially empty)
    '   - file_location (TEXT): File location (for hyperlinks)
    '
    Dim dbPath As String, dbHandle As LongPtr, result As Long, sqlCreate As String, stmtHandle As LongPtr

    dbPath = ThisWorkbook.Path & "\jobs.db"  ' Database file path
    
    ' 1. Initialize SQLite3
    result = SQLite3Initialize(ThisWorkbook.Path)
    If result <> SQLITE_INIT_OK Then
        MsgBox "SQLite3 initialization failed. Please check if SQLite3.dll and SQLite3_StdCall.dll are in the same directory.": Exit Sub
    End If

    ' 2. Connect to the database
    result = SQLite3Open(dbPath, dbHandle)
    If result <> SQLITE_OK Then
        MsgBox "Failed to open the database jobs.db. Please check if the file exists and the permissions.": SQLite3Free: Exit Sub
    End If

    ' 3. Create the table (if it doesn't exist)
    sqlCreate = "CREATE TABLE IF NOT EXISTS drawings (" & _
                "drawing_name TEXT, " & _
                "drawing_number TEXT, " & _
                "file_location TEXT)"

    result = SQLite3PrepareV2(dbHandle, sqlCreate, stmtHandle)
    If result = SQLITE_OK Then
        result = SQLite3Step(stmtHandle)
        SQLite3Finalize stmtHandle  ' Release the statement
        If result = SQLITE_DONE Then
            Debug.Print "Drawing information table 'drawings' created or already exists."
        Else
            Debug.Print "Error creating the table: " & result
        End If
    Else
        Debug.Print "Error preparing the SQL statement: " & SQLite3ErrMsg(dbHandle)
    End If

    ' 4. Close the database connection
    result = SQLite3Close(dbHandle)
    If result <> SQLITE_OK Then
        Debug.Print "Error closing the database connection: " & SQLite3ErrMsg(dbHandle)
    End If

    ' 5. Release SQLite3 resources
    SQLite3Free

    MsgBox "Drawing information table created or already exists, initialization successful!", vbInformation

End Sub