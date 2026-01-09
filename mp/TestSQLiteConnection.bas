' TestSQLiteConnection.bas
' -----------------------------------------------------------------------------
' Module Functionality:
'   - Test SQLite database connection using mod_SQLite.bas API
' -----------------------------------------------------------------------------

Option Explicit

Sub TestSQLiteConnection()
    Dim dbPath As String
    Dim isInitialized As Boolean

    ' Path to the SQLite database file
    dbPath = ThisWorkbook.Path & "\record.db"

    ' Initialize SQLite connection
    isInitialized = InitializeSQLite(dbPath)

    If isInitialized Then
        MsgBox "Successfully connected to the database: " & dbPath, vbInformation, "SQLite Connection Test"
    Else
        MsgBox "Failed to connect to the database: " & dbPath, vbCritical, "SQLite Connection Test"
    End If

    ' Close SQLite connection
    CloseSQLite
End Sub