Option Explicit

' -------------------------------------------------------------------------------------------------
' Module Functionality:
'   - Reads Job and Drawing Number relationships from "PrioritySheet" worksheet
'   - Stores the data into "assemblies" table (part_number, drawing_number)
'   - Procedures in this module can be called from other procedures
' -------------------------------------------------------------------------------------------------

Public Sub UpdateAssemblies()
    ' Main procedure: Processes each row of PrioritySheet, extracts Job and Drawing Number information,
    ' and stores it in the database.

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Priority Sheet")
    Dim lastRow As Long: lastRow = GetLastDataRow(ws)
    Dim i As Long, partCount As Long
    Dim part_number As String, drawing_number As String
    
    If Not InitializeSQLite(ThisWorkbook.Path & "\jobs.db") Then Exit Sub
    
    For i = 2 To lastRow
        If Not IsEmpty(ws.Cells(i, "A").Value) Then
            part_number = ws.Cells(i, "E").Value
            partCount = CountParts(ws, i)
            
            Dim j As Long
            For j = i + 1 To i + partCount
                If Not IsEmpty(ws.Cells(j, "E").Value) Then
                    drawing_number = ws.Cells(j, "E").Value
                    Call InsertAssemblyData(part_number, drawing_number)
                End If
            Next j
            
            i = i + partCount
        End If
    Next i
    
    CloseSQLite
    MsgBox "PrioritySheet processing complete. Data has been stored in the database.", vbInformation
End Sub

Private Sub InsertAssemblyData(part_number As String, drawing_number As String)
    ' Subprocedure: Inserts part_number and drawing_number into "assemblies" table.
    ' Checks for existing part_number and drawing_number combination before inserting.

    Dim sqlCheck As String, sqlInsert As String
    Dim recordExists As Boolean
    
    ' Check if record already exists
    sqlCheck = "SELECT COUNT(*) FROM assemblies WHERE part_number = ? AND drawing_number = ?"
    Dim result As Variant
    result = ExecuteSQL(sqlCheck, Array(part_number, drawing_number))
    
    If IsArray(result) Then
        recordExists = (result(0)(0) > 0)
    Else
        MsgBox "Error checking for existing record."
        Exit Sub
    End If

    ' Insert if record doesn't exist
    If Not recordExists Then
        sqlInsert = "INSERT INTO assemblies (part_number, drawing_number) VALUES (?, ?)"
        If ExecuteNonQuery(sqlInsert, Array(part_number, drawing_number)) Then
            Debug.Print "Data inserted successfully: part_number = " & part_number & ", drawing_number = " & drawing_number
        Else
            Debug.Print "Error inserting data"
        End If
    Else
        Debug.Print "Record already exists: part_number = " & part_number & ", drawing_number = " & drawing_number
    End If
End Sub

Private Function FileExists(ByVal filePath As String) As Boolean
    ' Check if file exists
    FileExists = (Dir(filePath, vbNormal) <> "")
End Function
