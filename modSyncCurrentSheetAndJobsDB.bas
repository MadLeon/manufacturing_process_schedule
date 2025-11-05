Sub SyncCurrentSheetAndJobsDB()
    ' 本地 test.xlsm 根据 jobs.db 内容同步：
    ' 1. 本地 Priority Sheet 有但数据库无的，移到shipped sheet，并增加下拉菜单操作
    ' 2. 数据库有但本地 Priority Sheet 无的，摘取部分字段添加到 Priority Sheet 末尾，并设置样式

    Dim dbPath As String, dbHandle As LongPtr, result As Long, stmtHandle As LongPtr
    Dim curBook As Workbook, curWS As Worksheet, shippedWS As Worksheet
    Dim lastRowCur As Long, shipLastRow As Long
    Dim dbDict As Object, curDict As Object
    Dim selectSQL As String, r As Long, k As Variant
    Dim jobNum As String

    dbPath = ThisWorkbook.Path & "\jobs.db"
    Set curBook = ThisWorkbook

    ' -------- 获取/新建 Priority Sheet --------
    On Error Resume Next
    Set curWS = curBook.Sheets("Priority Sheet")
    If curWS Is Nothing Then
        Set curWS = curBook.Sheets.Add(After:=curBook.Sheets(curBook.Sheets.Count))
        curWS.Name = "Priority Sheet"
        curWS.Range("A1:F1").Value = Array("JOB #", "PO #", "Customer", "Description", "Part #", "Qty.")
    End If
    On Error GoTo 0

    lastRowCur = curWS.Cells(curWS.Rows.Count, 1).End(-4162).Row

    ' -------- 设置 G,H,I 列表头 --------
    curWS.Cells(1, 7).Value = "Ship Date"
    curWS.Cells(1, 8).Value = "Memo"
    curWS.Cells(1, 9).Value = "Status"

    ' -------- 1. 初始化DLL连接 --------
    result = SQLite3Initialize(ThisWorkbook.Path)
    If result <> 0 Then
        MsgBox "SQLite3初始化失败": Exit Sub
    End If
    result = SQLite3Open(dbPath, dbHandle)
    If result <> 0 Then
        MsgBox "无法打开数据库": Exit Sub
    End If

    ' -------- 2. 获取 jobs.db 所有 Job_Number 记录 --------
    Set dbDict = CreateObject("Scripting.Dictionary")
    selectSQL = "SELECT Job_Number, PO_Number, Customer_Name, Part_Description, Part_Number, Job_Quantity, Delivery_Shipped_Date FROM jobs"
    result = SQLite3PrepareV2(dbHandle, selectSQL, stmtHandle)
    If result = 0 Then
        Do While SQLite3Step(stmtHandle) = 100
            jobNum = Trim(SQLite3ColumnText(stmtHandle, 0))
            If jobNum <> "" Then
                dbDict(jobNum) = Array( _
                    jobNum, _
                    SQLite3ColumnText(stmtHandle, 1), _
                    SQLite3ColumnText(stmtHandle, 2), _
                    SQLite3ColumnText(stmtHandle, 3), _
                    SQLite3ColumnText(stmtHandle, 4), _
                    SQLite3ColumnText(stmtHandle, 5), _
                    SQLite3ColumnText(stmtHandle, 6) _
                )
            End If
        Loop
        SQLite3Finalize stmtHandle
    End If

    ' -------- 3. 获取 Priority Sheet 的所有 Job_Number --------
    Set curDict = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRowCur
        jobNum = Trim(curWS.Cells(r, 1).Value) 'JOB # 在第1列
        If jobNum <> "" Then
            curDict(jobNum) = r
        End If
    Next

    ' -------- 4. 获取/新建 shipped sheet，并设置表头与格式 --------
    On Error Resume Next
    Set shippedWS = curBook.Sheets("Shipped")
    If shippedWS Is Nothing Then
        Set shippedWS = curBook.Sheets.Add(After:=curBook.Sheets(curBook.Sheets.Count))
        shippedWS.Name = "Shipped"
        ' 一次写入 A~I 表头
        shippedWS.Range("A1:I1").Value = Array("JOB #", "PO #", "Customer", "Description", "Part #", "Qty.", "Ship Date", "Memo", "Status")
        ' 设置表头样式
        With shippedWS.Range(shippedWS.Cells(1, 1), shippedWS.Cells(1, 10))  ' 包含J列
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

    ' -------- 5. 移动 Priority Sheet 有但数据库无的条目至 Shipped，并添加下拉菜单 --------
    Dim movedCount As Long: movedCount = 0
    For r = lastRowCur To 2 Step -1
        jobNum = Trim(curWS.Cells(r, 1).Value)
        If jobNum <> "" Then
            If Not dbDict.Exists(jobNum) Then
                shipLastRow = shippedWS.Cells(shippedWS.Rows.Count, 1).End(-4162).Row + 1
                curWS.Rows(r).Copy shippedWS.Rows(shipLastRow)
                curWS.Rows(r).Delete
                movedCount = movedCount + 1
                ' ----------- 在 J 列插入下拉菜单 (数据验证) -----------
                With shippedWS.Cells(shipLastRow, 10).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                         Formula1:="Return,Delete"  ' 下拉选项
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = True
                    .ShowError = True
                End With
                shippedWS.Cells(shipLastRow, 10).Value = ""  ' 默认空
            End If
        End If
    Next

    ' -------- 6. 只有发生了数据移动才触发列宽自动调整 --------
    If movedCount > 0 Then
        Dim shippedLastRow As Long
        Dim shippedUsedRng As Range
        shippedLastRow = shippedWS.Cells(shippedWS.Rows.Count, 1).End(-4162).Row
        Set shippedUsedRng = shippedWS.Range(shippedWS.Cells(1, 1), shippedWS.Cells(shippedLastRow, 10)) ' 包含J列
        shippedUsedRng.Columns.AutoFit
    End If

    ' -------- 7. 查找Priority Sheet 当前末行 --------
    lastRowCur = curWS.Cells(curWS.Rows.Count, 1).End(-4162).Row
    If lastRowCur < 2 Then lastRowCur = 1

    ' -------- 8. 数据库有但本地 Priority Sheet 无的，添加到 Priority Sheet 末尾，并设置橙色与边框 --------
    For Each k In dbDict.Keys
        If Not curDict.Exists(k) Then
            lastRowCur = lastRowCur + 1
            curWS.Cells(lastRowCur, 1).Value = dbDict(k)(0) 'JOB #
            curWS.Cells(lastRowCur, 2).Value = dbDict(k)(1) 'PO #
            curWS.Cells(lastRowCur, 3).Value = dbDict(k)(2) 'Customer
            curWS.Cells(lastRowCur, 4).Value = dbDict(k)(3) 'Description
            curWS.Cells(lastRowCur, 5).Value = dbDict(k)(4) 'Part #
            curWS.Cells(lastRowCur, 6).Value = dbDict(k)(5) 'Qty.
            curWS.Cells(lastRowCur, 7).Value = dbDict(k)(6) 'Ship Date
            ' H, I 列空

            ' 设置A-G列橙色（255,199,44），加粗外内全部边框
            Dim rg As Range
            Set rg = curWS.Range(curWS.Cells(lastRowCur, 1), curWS.Cells(lastRowCur, 7))
            rg.Interior.Color = RGB(255, 199, 44) ' 橙色
            With rg.Borders
                .LineStyle = xlContinuous
                .Color = vbBlack
                .Weight = xlThin
            End With
            With rg.Borders(xlEdgeLeft)
                .Weight = xlThin
            End With
            With rg.Borders(xlEdgeRight)
                .Weight = xlThin
            End With
            With rg.Borders(xlEdgeTop)
                .Weight = xlThin
            End With
            With rg.Borders(xlEdgeBottom)
                .Weight = xlThin
            End With
        End If
    Next

    ' -------- 9. 关闭数据库 --------
    SQLite3Close dbHandle
    SQLite3Free

    Debug.Print "本地文件与数据库同步已完成"
End Sub