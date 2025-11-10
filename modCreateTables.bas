' modCreateTable.bas
Option Explicit

Sub CreateCustomerFolderMapTable()
    ' Create a table for mapping customer names to G: drive folder names
    ' Table structure:
    '   - folder_name (TEXT): G: drive folder name
    '   - customer_name (TEXT): Customer name as in workbook

    Dim dbPath As String
    dbPath = ThisWorkbook.Path & "\jobs.db"

    If Not InitializeSQLite(dbPath) Then Exit Sub

    Dim sqlCreate As String
    sqlCreate = "CREATE TABLE IF NOT EXISTS customer_folder_map (" & _
                "folder_name TEXT, " & _
                "customer_name TEXT)"

    Dim result As Boolean
    result = ExecuteNonQuery(sqlCreate)

    CloseSQLite

    If result Then
        MsgBox "Customer-folder map table created or already exists!", vbInformation
    Else
        MsgBox "Failed to create customer-folder map table.", vbExclamation
    End If
End Sub

Sub CreateDrawingTable()
    ' Create or initialize the drawing information table
    ' Table structure:
    '   - drawing_name (TEXT): Drawing file name
    '   - drawing_number (TEXT): Drawing number
    '   - file_location (TEXT): File location (for hyperlinks)
    
    Dim dbPath As String
    
    dbPath = ThisWorkbook.Path & "\jobs.db"
    
    If Not InitializeSQLite(dbPath) Then Exit Sub
    
    Dim sqlCreate As String
    sqlCreate = "CREATE TABLE IF NOT EXISTS drawings (" & _
                "drawing_name TEXT, " & _
                "drawing_number TEXT, " & _
                "file_location TEXT)"
    
    Dim result As Boolean
    result = ExecuteNonQuery(sqlCreate)
    
    CloseSQLite
    
    If result Then
        MsgBox "Drawing information table created or already exists, initialization successful!", vbInformation
    Else
        MsgBox "Failed to create drawing table.", vbExclamation
    End If
End Sub

Sub CreateAssemblyTable()
    ' Create the assembly information table
    ' Table structure:
    '   - part_number (TEXT): Assembly part number
    '   - drawing_number (TEXT): Corresponding drawing number
    
    Dim dbPath As String
    
    dbPath = ThisWorkbook.Path & "\jobs.db"
    
    If Not InitializeSQLite(dbPath) Then Exit Sub
    
    Dim sqlCreate As String
    sqlCreate = "CREATE TABLE IF NOT EXISTS assemblies (" & _
                "part_number TEXT, " & _
                "drawing_number TEXT)"
    
    Dim result As Boolean
    result = ExecuteNonQuery(sqlCreate)
    
    CloseSQLite
    
    If result Then
        MsgBox "Assembly information table created or already exists, initialization successful!", vbInformation
    Else
        MsgBox "Failed to create assembly table.", vbExclamation
    End If
End Sub