Option Explicit

Sub ClearAssemblyTable()
    ' 清除 jobs.db 数据库中 assemblies 表的内容

    Dim dbPath As String, dbHandle As LongPtr, result As Long, sqlDelete As String, stmtHandle As LongPtr

    dbPath = ThisWorkbook.Path & "\jobs.db"  ' 数据库文件路径

    ' 1. 初始化 SQLite3
    result = SQLite3Initialize(ThisWorkbook.Path)
    If result <> SQLITE_INIT_OK Then
        MsgBox "SQLite3 初始化失败。请检查 SQLite3.dll 和 SQLite3_StdCall.dll 是否在同一目录下。": Exit Sub
    End If

    ' 2. 连接到数据库
    result = SQLite3Open(dbPath, dbHandle)
    If result <> SQLITE_OK Then
        MsgBox "无法打开数据库 jobs.db。请检查文件是否存在以及权限。": SQLite3Free: Exit Sub
    End If

    ' 3. 构建 SQL DELETE 语句
    sqlDelete = "DELETE FROM assemblies"

    ' 4. 准备 SQL 语句
    result = SQLite3PrepareV2(dbHandle, sqlDelete, stmtHandle)
    If result = SQLITE_OK Then
        ' 5. 执行 SQL 语句
        result = SQLite3Step(stmtHandle)
        SQLite3Finalize stmtHandle  ' 释放语句

        If result = SQLITE_DONE Then
            Debug.Print "Assembly information table 'assemblies' 已清空。"
        Else
            Debug.Print "清除表格时出错: " & result
        End If
    Else
        Debug.Print "准备 SQL 语句时出错: " & SQLite3ErrMsg(dbHandle)
    End If

    ' 6. 关闭数据库连接
    result = SQLite3Close(dbHandle)
    If result <> SQLITE_OK Then
        Debug.Print "关闭数据库连接时出错: " & SQLite3ErrMsg(dbHandle)
    End If

    ' 7. 释放 SQLite3 资源
    SQLite3Free

    MsgBox "Assembly information table 已成功清空！", vbInformation

End Sub

Sub ClearDrawingNumbers()
    ' 清除 drawings 表中所有记录的 drawing_number 列的值

    Dim dbPath As String, dbHandle As LongPtr, result As Long, sqlUpdate As String, stmtHandle As LongPtr

    dbPath = ThisWorkbook.Path & "\jobs.db"  ' 数据库文件路径

    ' 1. 初始化 SQLite3
    result = SQLite3Initialize(ThisWorkbook.Path)
    If result <> SQLITE_INIT_OK Then
        MsgBox "SQLite3 初始化失败。请检查 SQLite3.dll 和 SQLite3_StdCall.dll 是否在同一目录下。": Exit Sub
    End If

    ' 2. 连接到数据库
    result = SQLite3Open(dbPath, dbHandle)
    If result <> SQLITE_OK Then
        MsgBox "无法打开数据库 jobs.db。请检查文件是否存在以及权限。": SQLite3Free: Exit Sub
    End If

    ' 3. 构建 SQL UPDATE 语句
    sqlUpdate = "UPDATE drawings SET drawing_number = ''"

    ' 4. 准备 SQL 语句
    result = SQLite3PrepareV2(dbHandle, sqlUpdate, stmtHandle)
    If result = SQLITE_OK Then
        ' 5. 执行 SQL 语句
        result = SQLite3Step(stmtHandle)
        SQLite3Finalize stmtHandle  ' 释放语句

        If result = SQLITE_DONE Then
            Debug.Print "成功清除 drawings 表中所有记录的 drawing_number。"
        Else
            Debug.Print "清除 drawing_number 时出错: " & result
        End If
    Else
        Debug.Print "准备 SQL 语句时出错: " & SQLite3ErrMsg(dbHandle)
    End If

    ' 6. 关闭数据库连接
    result = SQLite3Close(dbHandle)
    If result <> SQLITE_OK Then
        Debug.Print "关闭数据库连接时出错: " & SQLite3ErrMsg(dbHandle)
    End If

    ' 7. 释放 SQLite3 资源
    SQLite3Free

    MsgBox "成功清除 drawings 表中所有记录的 drawing_number!", vbInformation

End Sub