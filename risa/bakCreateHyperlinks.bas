Option Explicit

' ------------------ Helper Subroutine ------------------
Sub AddHyperlinkToCell(cell As Range, drawingsDict As Object, dbHandle As LongPtr)
    ' This subroutine encapsulates the logic for adding a hyperlink to a given cell, using cached data.
    Dim drawingNumber As String, fileLocation As String
    Dim sql As String, result As Long, stmtHandle As LongPtr
    Dim key As Variant

    drawingNumber = Trim(cell.Value)

    If drawingNumber <> "" Then
        ' 检查单元格是否已经是超链接
        On Error Resume Next
        Dim hyp As Hyperlink
        Set hyp = cell.Worksheet.Hyperlinks(cell.Address)
        On Error GoTo 0

        If Not hyp Is Nothing Then
            ' 如果已经是超链接，则跳过该行
            Set hyp = Nothing
        Else
            fileLocation = "" ' Reset fileLocation for each drawingNumber

            ' -------- 1. 首先在缓存中查找 drawing_number 列 --------
            For Each key In drawingsDict.Keys
                If drawingsDict(key).Item("drawing_number") = drawingNumber Then
                    fileLocation = drawingsDict(key).Item("file_location")
                    Exit For
                End If
            Next key

            ' -------- 2. 如果 drawing_number 不存在，则查找 drawing_name 列 --------
            If fileLocation = "" Then
                Dim customerName As String
                customerName = Trim(cell.Worksheet.Cells(cell.Row, 3).Value) ' 获取 Priority Sheet 第三列的 Customer

                For Each key In drawingsDict.Keys
                    If InStr(1, key, drawingNumber, vbTextCompare) > 0 And InStr(1, drawingsDict(key).Item("file_location"), customerName, vbTextCompare) > 0 Then
                        fileLocation = drawingsDict(key).Item("file_location")
                        Exit For
                    End If
                Next key
            End If

            ' -------- 3. 如果找到 file_location，则添加超链接并更新 drawing_number --------
            If fileLocation <> "" Then
                On Error Resume Next
                cell.Worksheet.Hyperlinks.Add Anchor:=cell, Address:=fileLocation, TextToDisplay:=drawingNumber
                On Error GoTo 0

                ' 设置字体和字号
                With cell.Font
                    .Name = "Cambria"
                    .Size = 16
                End With

                ' 更新 drawing_number
                sql = "UPDATE drawings SET drawing_number = '" & drawingNumber & "' WHERE drawing_name = '" & key & "'"
                result = SQLite3PrepareV2(dbHandle, sql, stmtHandle)
                If result = SQLITE_OK Then
                    SQLite3Step stmtHandle
                    SQLite3Finalize stmtHandle
                Else
                    Debug.Print "Error preparing update SQL: " & SQLite3ErrMsg(dbHandle)
                End If
            End If
        End If
    End If
End Sub

Sub CreateSingleHyperlink(cell As Range, dbHandle As LongPtr)
    ' This subroutine directly accesses the database to add a hyperlink to a given cell, without caching.
    Dim drawingNumber As String, fileLocation As String, dName As String
    Dim sql As String, result As Long, stmtHandle As LongPtr

    drawingNumber = Trim(cell.Value)

    If drawingNumber <> "" Then
        ' 检查单元格是否已经是超链接
        On Error Resume Next
        Dim hyp As Hyperlink
        Set hyp = cell.Worksheet.Hyperlinks(cell.Address)
        On Error GoTo 0

        If Not hyp Is Nothing Then
            ' 如果已经是超链接，则跳过该行
            Set hyp = Nothing
        Else
            fileLocation = "" ' Reset fileLocation for each drawingNumber

            ' -------- 1. 首先在数据库中查找 drawing_number 列 --------
            sql = "SELECT file_location FROM drawings WHERE drawing_number = '" & drawingNumber & "'"
            result = SQLite3PrepareV2(dbHandle, sql, stmtHandle)
            If result = SQLITE_OK Then
                If SQLite3Step(stmtHandle) = SQLITE_ROW Then
                    fileLocation = Trim(SQLite3ColumnText(stmtHandle, 0))
                End If
                SQLite3Finalize stmtHandle
            Else
                Debug.Print "Error preparing SQL: " & SQLite3ErrMsg(dbHandle)
            End If

            ' -------- 2. 如果 drawing_number 不存在，则查找 drawing_name 列 --------
            If fileLocation = "" Then
                Dim customerName As String
                customerName = Trim(cell.Worksheet.Cells(cell.Row, 3).Value) ' 获取 Priority Sheet 第三列的 Customer

                sql = "SELECT file_location, drawing_name FROM drawings WHERE drawing_name LIKE '%" & drawingNumber & "%' AND file_location LIKE '%" & customerName & "%'"
                result = SQLite3PrepareV2(dbHandle, sql, stmtHandle)
                If result = SQLITE_OK Then
                    If SQLite3Step(stmtHandle) = SQLITE_ROW Then
                        fileLocation = Trim(SQLite3ColumnText(stmtHandle, 0))
                        dName = Trim(SQLite3ColumnText(stmtHandle, 1))
                    End If
                    SQLite3Finalize stmtHandle
                Else
                    Debug.Print "Error preparing SQL: " & SQLite3ErrMsg(dbHandle)
                End If
            End If

            ' -------- 3. 如果找到 file_location，则添加超链接并更新 drawing_number --------
            If fileLocation <> "" Then
                On Error Resume Next
                cell.Worksheet.Hyperlinks.Add Anchor:=cell, Address:=fileLocation, TextToDisplay:=drawingNumber
                On Error GoTo 0

                ' 设置字体和字号
                With cell.Font
                    .Name = "Cambria"
                    .Size = 16
                End With

                ' 更新 drawing_number
                If dName <> "" Then
                    sql = "UPDATE drawings SET drawing_number = '" & drawingNumber & "' WHERE drawing_name = '" & dName & "'"
                Else
                    sql = "UPDATE drawings SET drawing_number = '" & drawingNumber & "' WHERE drawing_number = '" & drawingNumber & "'"
                End If

                result = SQLite3PrepareV2(dbHandle, sql, stmtHandle)
                If result = SQLITE_OK Then
                    SQLite3Step stmtHandle
                    SQLite3Finalize stmtHandle
                Else
                    Debug.Print "Error preparing update SQL: " & SQLite3ErrMsg(dbHandle)
                End If
            End If
        End If
    End If
End Sub

Sub CreateAllHyperlinks()
    ' Creates hyperlinks for all applicable cells in the Priority Sheet, using caching.
    Dim dbPath As String, dbHandle As LongPtr, result As Long, stmtHandle As LongPtr
    Dim curBook As Workbook, curWS As Worksheet
    Dim lastRow As Long, r As Long
    Dim sql As String
    Dim drawingsDict As Object

    ' -------- 1. 初始化工作簿和工作表 --------
    Set curBook = ThisWorkbook
    On Error Resume Next
    Set curWS = curBook.Sheets("Priority Sheet")
    If curWS Is Nothing Then
        MsgBox "找不到名为 'Priority Sheet' 的工作表!", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' -------- 2. 初始化数据库连接 --------
    dbPath = ThisWorkbook.Path & "\jobs.db"
    result = SQLite3Initialize(ThisWorkbook.Path)
    If result <> SQLITE_INIT_OK Then
        MsgBox "SQLite3 初始化失败!", vbCritical
        Exit Sub
    End If
    result = SQLite3Open(dbPath, dbHandle)
    If result <> SQLITE_OK Then
        MsgBox "无法打开数据库!", vbCritical
        Exit Sub
    End If

    ' -------- 3. 缓存数据库中的 drawings 表数据 --------
    Set drawingsDict = CreateObject("Scripting.Dictionary")
    sql = "SELECT drawing_name, drawing_number, file_location FROM drawings"
    result = SQLite3PrepareV2(dbHandle, sql, stmtHandle)
    If result = SQLITE_OK Then
        Do While SQLite3Step(stmtHandle) = SQLITE_ROW
            Dim dName As String, dNumber As String, fLocation As String
            dName = Trim(SQLite3ColumnText(stmtHandle, 0))
            dNumber = Trim(SQLite3ColumnText(stmtHandle, 1))
            fLocation = Trim(SQLite3ColumnText(stmtHandle, 2))

            ' 使用 drawing_name 作为主键，如果存在相同的 drawing_name，则用 Dictionary 存储多个值
            If Not drawingsDict.Exists(dName) Then
                Set drawingsDict(dName) = CreateObject("Scripting.Dictionary")
            End If

            With drawingsDict(dName)
                .Item("drawing_number") = dNumber
                .Item("file_location") = fLocation
            End With
        Loop
        SQLite3Finalize stmtHandle
    Else
        Debug.Print "Error preparing SQL: " & SQLite3ErrMsg(dbHandle)
    End If

    ' -------- 4. 获取 Priority Sheet 的最后一行 --------
    lastRow = GetLastDataRow(curWS)

    ' -------- 5. 遍历 Priority Sheet 的 E 列（从第二行开始） --------
    For r = 2 To lastRow
        ' E 列是第 5 列 (Part #)
        Call AddHyperlinkToCell(curWS.Cells(r, 5), drawingsDict, dbHandle)
    Next r

    ' -------- 6. 清理和关闭数据库连接 --------
    SQLite3Close dbHandle
    SQLite3Free
    Set drawingsDict = Nothing
    MsgBox "超链接创建完成!", vbInformation
End Sub

Sub CreateHyperlinks()
    ' This sub creates hyperlinks only for cells in E column that are within the current selection, without caching.
    Dim dbPath As String, dbHandle As LongPtr, result As Long
    Dim curBook As Workbook, curWS As Worksheet
    Dim selectedRange As Range, cell As Range
    Dim eCells As Collection
    Dim c As Variant

    ' -------- 1. 初始化工作簿和工作表 --------
    Set curBook = ThisWorkbook
    On Error Resume Next
    Set curWS = curBook.Sheets("Priority Sheet")
    If curWS Is Nothing Then
        MsgBox "找不到名为 'Priority Sheet' 的工作表!", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' -------- 2. 初始化数据库连接 --------
    dbPath = ThisWorkbook.Path & "\jobs.db"
    result = SQLite3Initialize(ThisWorkbook.Path)
    If result <> SQLITE_INIT_OK Then
        MsgBox "SQLite3 初始化失败!", vbCritical
        Exit Sub
    End If
    result = SQLite3Open(dbPath, dbHandle)
    If result <> SQLITE_OK Then
        MsgBox "无法打开数据库!", vbCritical
        Exit Sub
    End If

    ' -------- 3. 获取选定的区域 --------
    Set selectedRange = Selection
    Set eCells = New Collection

    ' -------- 4. 遍历选定区域内的单元格，找到 E 列的单元格 (从第二行开始) --------
    For Each cell In selectedRange
        If cell.Column = 5 And cell.Row > 1 Then ' 确保是 E 列且不是标题行
            eCells.Add cell
        End If
    Next cell

    ' -------- 5. 遍历 E 列的单元格，添加超链接 --------
    For Each c In eCells
        Set cell = c
        Call CreateSingleHyperlink(cell, dbHandle)
    Next c

    ' -------- 6. 清理和关闭数据库连接 --------
    SQLite3Close dbHandle
    SQLite3Free
    Set eCells = Nothing
    MsgBox "超链接创建完成!", vbInformation
End Sub

