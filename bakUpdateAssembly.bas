Option Explicit

' -------------------------------------------------------------------------------------------------
' 模块功能：
'   - 从 "PrioritySheet" 工作表中读取 Job 和 Drawing Number 的对应关系。
'   - 将读取到的数据存入 "assemblies" 表中 (part_number, drawing_number)。
'   - 该模块中的过程可以被其他过程调用。
' -------------------------------------------------------------------------------------------------

Public Sub UpdateAssemblies()
    ' 主过程：处理 PrioritySheet 的每一行，提取 Job 和 Drawing Number 信息，并存储到数据库。

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Priority Sheet") ' 获取 PrioritySheet 工作表
    Dim lastRow As Long: lastRow = GetLastDataRow(ws) ' 获取 A 列的最后一行
    Dim i As Long, partCount As Long
    Dim part_number As String, drawing_number As String
    Dim dbPath As String
    dbPath = ThisWorkbook.Path & "\jobs.db"
    
    If Not FileExists(dbPath) Then
        MsgBox "Database file not found. Please ensure that the 'jobs.db' exists in the workbook directory."
        Exit Sub
    End If
    
    For i = 2 To lastRow ' 从第二行开始遍历 (假设第一行是标题)
        If Not IsEmpty(ws.Cells(i, "A").Value) Then ' A 列有值，表示 Job 行
            part_number = ws.Cells(i, "E").Value ' 获取 Job 的 part_number
            
            partCount = CountParts(ws, i)
            
            ' 遍历零件行，提取 Drawing Number
            Dim j As Long
            For j = i + 1 To i + partCount
                If Not IsEmpty(ws.Cells(j, "E").Value) Then ' 零件行的 E 列有值，表示 Drawing Number
                    drawing_number = ws.Cells(j, "E").Value ' 获取 Drawing Number
                    
                    ' 将 part_number 和 drawing_number 存入数据库
                    Call InsertAssemblyData(dbPath, part_number, drawing_number)
                End If
            Next j
            
            i = i + partCount ' 跳过已处理的零件行
        End If
    Next i
    
    MsgBox "PrioritySheet 处理完成，数据已存入数据库。", vbInformation
    
End Sub

Private Sub InsertAssemblyData(dbPath As String, part_number As String, drawing_number As String)
    ' 子过程：将 part_number 和 drawing_number 插入 "assemblies" 表。
    ' 在插入之前检查是否已存在相同的 part_number 和 drawing_number 组合。

    Dim dbHandle As LongPtr, stmtHandle As LongPtr, result As Long, sqlInsert As String, sqlCheck As String
    Dim recordExists As Boolean

    ' 1. Initialize SQLite3
    result = SQLite3Initialize(ThisWorkbook.Path)
    If result <> SQLITE_INIT_OK Then
        MsgBox "SQLite3 initialization failed. Please check if SQLite3.dll and SQLite3_StdCall.dll are in the same directory.": Exit Sub
    End If

    ' 2. Connect to the database
    result = SQLite3Open(dbPath, dbHandle)
    If result <> SQLITE_OK Then
        Debug.Print "Failed to open the database jobs.db." & result: Exit Sub
    End If
    
    ' 3. 检查记录是否已存在
    sqlCheck = "SELECT COUNT(*) FROM assemblies WHERE part_number = '" & part_number & "' AND drawing_number = '" & drawing_number & "'"
    result = SQLite3PrepareV2(dbHandle, sqlCheck, stmtHandle)
    If result = SQLITE_OK Then
        If SQLite3Step(stmtHandle) = SQLITE_ROW Then
            If SQLite3ColumnInt32(stmtHandle, 0) > 0 Then
                recordExists = True
            Else
                recordExists = False
            End If
        End If
        SQLite3Finalize stmtHandle
    Else
        Debug.Print "Error preparing the SQL check statement: " & SQLite3ErrMsg(dbHandle)
        SQLite3Close dbHandle
        SQLite3Free
        Exit Sub
    End If

    ' 4. 如果记录不存在，则插入
    If Not recordExists Then
        sqlInsert = "INSERT INTO assemblies (part_number, drawing_number) VALUES ('" & part_number & "', '" & drawing_number & "')"
        result = SQLite3PrepareV2(dbHandle, sqlInsert, stmtHandle)
        If result = SQLITE_OK Then
            result = SQLite3Step(stmtHandle)
            SQLite3Finalize stmtHandle  ' Release the statement
            If result = SQLITE_DONE Then
                Debug.Print "Data inserted successfully: part_number = " & part_number & ", drawing_number = " & drawing_number
            Else
                Debug.Print "Error inserting data: " & result
            End If
        Else
            Debug.Print "Error preparing the SQL insert statement: " & SQLite3ErrMsg(dbHandle)
        End If
    Else
        Debug.Print "Record already exists: part_number = " & part_number & ", drawing_number = " & drawing_number
    End If

    ' 5. 关闭数据库连接
    result = SQLite3Close(dbHandle)
    If result <> SQLITE_OK Then
        Debug.Print "Error closing the database connection: " & SQLite3ErrMsg(dbHandle)
    End If

    ' 6. 释放 SQLite3 资源
    SQLite3Free
End Sub

Private Function FileExists(ByVal filePath As String) As Boolean
    ' 检查文件是否存在
    FileExists = (Dir(filePath, vbNormal) <> "")
End Function
