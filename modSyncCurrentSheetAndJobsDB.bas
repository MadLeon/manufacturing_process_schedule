Sub SyncCurrentSheetAndJobsDB()
    ' 本地 test.xlsm 根据 jobs.db 内容同步：
    ' 1. 本地 Priority Sheet 有但数据库无的，移到shipped sheet
    ' 2. 数据库有但本地 Priority Sheet 无的，摘取部分字段添加到 Priority Sheet 末尾，并设置样式

    Dim dbPath As String, dbHandle As LongPtr, result As Long, stmtHandle As LongPtr
    Dim curBook As Workbook, curWS As Worksheet, shippedWS As Worksheet
    Dim lastRowCur As Long
    Dim dbDict As Object, curDict As Object
    Dim selectSQL As String, r As Long, k As Variant
    Dim jobNum As String, firstPartRow As Long

    dbPath = ThisWorkbook.Path & "\jobs.db"
    Set curBook = ThisWorkbook
    
    ' 获取Priority Sheet
    On Error Resume Next
    Set curWS = curBook.Sheets("Priority Sheet")
    If curWS Is Nothing Then
        Set curWS = curBook.Sheets.Add(After:=curBook.Sheets(curBook.Sheets.Count))
        curWS.Name = "Priority Sheet"
        curWS.Range("A1:F1").Value = Array("JOB #", "PO #", "Customer", "Description", "Part #", "Qty.")
    End If
    On Error GoTo 0

    lastRowCur = curWS.Cells(curWS.Rows.Count, 1).End(-4162).Row

    ' 设置列头
    curWS.Cells(1, 7).Value = "Ship Date"
    curWS.Cells(1, 8).Value = "Memo"
    curWS.Cells(1, 9).Value = "Status"

    ' 1. 初始化DLL
    result = SQLite3Initialize(ThisWorkbook.Path)
    If result <> 0 Then
        MsgBox "SQLite3初始化失败": Exit Sub
    End If
    result = SQLite3Open(dbPath, dbHandle)
    If result <> 0 Then
        MsgBox "无法打开数据库": Exit Sub
    End If

    ' 2. 获取数据库Job Number清单
    Set dbDict = CreateObject("Scripting.Dictionary")
    selectSQL = "SELECT Job_Number FROM jobs"
    result = SQLite3PrepareV2(dbHandle, selectSQL, stmtHandle)
    If result = 0 Then
        Do While SQLite3Step(stmtHandle) = 100
            jobNum = Trim(SQLite3ColumnText(stmtHandle, 0))
            If jobNum <> "" Then
                dbDict(jobNum) = 1 ' 仅需key
            End If
        Loop
        SQLite3Finalize stmtHandle
    End If

    ' 3. 查找PrioritySheet
    Set curDict = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRowCur
        jobNum = Trim(curWS.Cells(r, 1).Value) 'JOB # 在第1列
        If jobNum <> "" Then
            curDict(jobNum) = r
        End If
    Next

    ' 4.  获取/创建 Shipped sheet
    On Error Resume Next
    Set shippedWS = curBook.Sheets("Shipped")
    If shippedWS Is Nothing Then
        Set shippedWS = curBook.Sheets.Add(After:=curBook.Sheets(curBook.Sheets.Count))
        shippedWS.Name = "Shipped"
        ' 设置 shipped sheet 表头
        shippedWS.Range("A1:I1").Value = Array("JOB #", "PO #", "Customer", "Description", "Part #", "Qty.", "Ship Date", "Memo", "Status")
        With shippedWS.Range(shippedWS.Cells(1, 1), shippedWS.Cells(1, 10))
            .Interior.Color = RGB(255, 199, 206) ' 淡粉色
            .Font.Bold = True
            .Font.Size = 16
            .Font.Name = "Cambria"
            .HorizontalAlignment = xlCenter
            With .Borders
                .LineStyle = xlContinuous
                .Color = vbBlack
                .Weight = xlThin
            End With
        End With
        shippedWS.Range(shippedWS.Cells(1, 1), shippedWS.Cells(1, 10)).EntireColumn.AutoFit
    End If
    On Error GoTo 0

    ' -------- 5. 处理从 Priority Sheet 移动到 Shipped 的行 --------
    Dim movedCount As Long: movedCount = 0
    r = 2 ' 从数据区域第一行开始
    Do While r <= lastRowCur
        jobNum = Trim(curWS.Cells(r, 1).Value)
        If jobNum <> "" Then  ' 如果是 Job 行
            If Not dbDict.Exists(jobNum) Then
                ' 获取该job下面所有的part数
                Dim partsCount As Long
                partsCount = CountParts(curWS, r)

                ' 计算要移动的总行数
                Dim totalRowsToMove As Long
                totalRowsToMove = 1 + partsCount

                ' 移动
                shipLastRow = shippedWS.Cells(shippedWS.Rows.Count, 1).End(-4162).Row + 1
                curWS.Rows(r).Resize(totalRowsToMove).Copy shippedWS.Rows(shipLastRow)

                ' 删除
                curWS.Rows(r).Resize(totalRowsToMove).Delete

                ' 更新行数
                movedCount = movedCount + totalRowsToMove
                lastRowCur = curWS.Cells(curWS.Rows.Count, 1).End(-4162).Row

                ' !!!!! 重要: 同步 r 的值, 避免跳过行 !!!!!
                r = r - 1 ' 因为删除了, 本行要重新检查
                If r < 2 Then r = 2 ' 防止r小于表头

             Else
                 r = r+1 '跳到下个job开始位置
            End If
             
        End If
       
         r = r + 1
        If r > lastRowCur Then Exit Do
    Loop
     Debug.Print "总共移动 " & movedCount & " 行"

    ' -------- 7. 只有发生了数据移动才触发列宽自动调整 --------
    If movedCount > 0 Then
        Dim shippedLastRow As Long
        Dim shippedUsedRng As Range
        shippedLastRow = shippedWS.Cells(shippedWS.Rows.Count, 1).End(-4162).Row
        Set shippedUsedRng = shippedWS.Range(shippedWS.Cells(1, 1), shippedWS.Cells(shippedLastRow, 10)) ' 包含J列
        shippedUsedRng.Columns.AutoFit
    End If

    ' -------- 9. 完成清理工作 --------
    SQLite3Close dbHandle
    SQLite3Free
    Set curDict = Nothing: Set dbDict = Nothing
    Set srcEntryBook = Nothing: Set srcEntryWS = Nothing: Set srcEntryApp = Nothing

    Debug.Print "本地文件与数据库同步已完成"
End Sub

'统计零部件方法：
Function CountParts(ws As Worksheet, startRow As Long) As Long
    Dim r As Long
    Dim lastRow As Long
    CountParts = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If startRow >= lastRow Then Exit Function

    r = startRow + 1
    Do While r <= lastRow
        If Trim(ws.Cells(r, 1).Value) <> "" Then Exit Do ' 遇到下一个 Job,退出
        CountParts = CountParts + 1
        r = r + 1
    Loop

    Debug.Print "CountParts: Found " & CountParts & " parts for Job at row " & startRow
End Function