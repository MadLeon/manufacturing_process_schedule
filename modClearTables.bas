Option Explicit

' --- Configuration Variables ---
Const DB_PATH As String = "\\rtdnas2\QCReports\jobs.db"

Sub ClearCustomerFolderMapTable()
    ' Clear all content from the customer_folder_map table in jobs.db
    Dim dbPath As String
    dbPath = DB_PATH

    If Not InitializeSQLite(dbPath) Then Exit Sub

    Dim sqlClear As String
    sqlClear = "DELETE FROM customer_folder_map"

    Dim result As Boolean
    result = ExecuteNonQuery(sqlClear)

    CloseSQLite

    If result Then
        MsgBox "All records in customer_folder_map table have been cleared!", vbInformation
    Else
        MsgBox "Failed to clear customer_folder_map table.", vbExclamation
    End If
End Sub

Sub ClearAssemblyTable()
    ' Clear contents of assemblies table in jobs.db database

    Dim dbPath As String
    dbPath = DB_PATH  ' Database file path

    ' Initialize SQLite database
    If Not InitializeSQLite(dbPath) Then Exit Sub

    ' Execute DELETE statement
    If ExecuteNonQuery("DELETE FROM assemblies") Then
        Debug.Print "Assembly table 'assemblies' cleared successfully."
        MsgBox "Assembly table cleared successfully!", vbInformation
    Else
        Debug.Print "Error clearing assemblies table."
        MsgBox "Failed to clear assemblies table.", vbExclamation
    End If

    ' Close SQLite database
    CloseSQLite
End Sub

Sub ClearDrawingNumbers()
    ' Clear drawing_number column values in drawings table

    Dim dbPath As String
    dbPath = DB_PATH  ' Database file path

    ' Initialize SQLite database
    If Not InitializeSQLite(dbPath) Then Exit Sub

    ' Execute UPDATE statement
    If ExecuteNonQuery("UPDATE drawings SET drawing_number = ''") Then
        Debug.Print "Successfully cleared drawing_number in drawings table."
        MsgBox "Drawing numbers cleared successfully!", vbInformation
    Else
        Debug.Print "Error clearing drawing numbers."
        MsgBox "Failed to clear drawing numbers.", vbExclamation
    End If

    ' Close SQLite database
    CloseSQLite
End Sub
