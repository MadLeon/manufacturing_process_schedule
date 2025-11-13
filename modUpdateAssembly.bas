Option Explicit

' -------------------------------------------------------------------------------------------------
' Module Functionality:
'   - Reads Job and Drawing Number relationships from "PrioritySheet" worksheet
'   - Stores the data into "assemblies" table (part_number, drawing_number)
'   - Procedures in this module can be called from other procedures
' -------------------------------------------------------------------------------------------------

' -------------------------------------------------------------------------------------------------
' Updates:
'   - Drawing description in assemblies table will be also updated.   2025.11.13
' -------------------------------------------------------------------------------------------------

' --- Configuration Variables ---
Const DB_PATH As String = "\\rtdnas2\OE\jobs.db"

Public Sub UpdateAssemblies()
    ' Main procedure: Processes each row of PrioritySheet, extracts Job and Drawing Number information,
    ' and stores it in the database.

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Priority Sheet")
    Dim lastRow As Long: lastRow = GetLastDataRow(ws)
    Dim i As Long, partCount As Long
    Dim part_number As String, drawing_number As String, description As String
    
    If Not InitializeSQLite(DB_PATH) Then Exit Sub
    
    For i = 2 To lastRow
        If Not IsEmpty(ws.Cells(i, "A").Value) Then
            part_number = ws.Cells(i, "E").Value
            partCount = CountParts(ws, i)
            
            Dim j As Long
            For j = i + 1 To i + partCount
                If Not IsEmpty(ws.Cells(j, "E").Value) Then
                    drawing_number = ws.Cells(j, "E").Value
                    description = ws.Cells(j, "D").Value
                    Call InsertAssemblyData(part_number, drawing_number, description)
                End If
            Next j
            
            i = i + partCount
        End If
    Next i
    
    CloseSQLite
    MsgBox "PrioritySheet processing complete. Data has been stored in the database.", vbInformation
End Sub

Private Sub InsertAssemblyData(part_number As String, drawing_number As String, description As String)
    ' Subprocedure: Inserts part_number and drawing_number into "assemblies" table.
    ' Checks for existing part_number and drawing_number combination before inserting.

    Dim sqlCheck As String, sqlInsert As String, sqlUpdate As String
    Dim recordExists As Boolean
    
    ' Check if record already exists
    sqlCheck = "SELECT COUNT(*) FROM assemblies WHERE part_number = '" & part_number & "' AND drawing_number = '" & drawing_number & "'"
    Dim result As Variant
    result = ExecuteSQL(sqlCheck)
    
    If IsArray(result) Then
        If UBound(result) >= 0 And UBound(result(0)) >= 0 Then
            recordExists = (result(0)(0) > 0)
        Else
            recordExists = False ' Handle the case where no rows are returned
        End If
    Else
        MsgBox "Error checking for existing record."
        Exit Sub
    End If

    ' Insert if record doesn't exist
    If Not recordExists Then
        sqlInsert = "INSERT INTO assemblies (part_number, drawing_number, description) VALUES ('" & part_number & "', '" & drawing_number & "', '" & description & "')"
        If ExecuteNonQuery(sqlInsert) Then
            Debug.Print "Data inserted successfully: part_number = " & part_number & ", drawing_number = " & drawing_number & ", description = " & description
        Else
            Debug.Print "Error inserting data"
        End If
    Else
        ' If record exists, check if description needs to be updated
        sqlCheck = "SELECT description FROM assemblies WHERE part_number = '" & part_number & "' AND drawing_number = '" & drawing_number & "'"
        Dim currentDescription As Variant
        currentDescription = ExecuteSQL(sqlCheck)
        
        If IsArray(currentDescription) Then
            If UBound(currentDescription) >= 0 And UBound(currentDescription(0)) >= 0 Then
                If IsNull(currentDescription(0)(0)) Or currentDescription(0)(0) = "" Then
                    ' if description is empty, update it
                    sqlUpdate = "UPDATE assemblies SET description = '" & description & "' WHERE part_number = '" & part_number & "' AND drawing_number = '" & drawing_number & "'"
                    If ExecuteNonQuery(sqlUpdate) Then
                        Debug.Print "Description updated successfully: part_number = " & part_number & ", drawing_number = " & drawing_number & ", description = " & description
                    Else
                        Debug.Print "Error updating description"
                    End If
                Else
                    Debug.Print "Record already exists with description: part_number = " & part_number & ", drawing_number = " & drawing_number
                End If
            Else
                Debug.Print "No description found for existing record."
            End If
        Else
            MsgBox "Error checking for existing description."
            Exit Sub
        End If
    End If
End Sub