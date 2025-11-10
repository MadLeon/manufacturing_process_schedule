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
    
    ' 获取/新建 Priority Sheet 并设置表头
    On Error Resume Next
    Set curWS = curBook.Sheets("Priority Sheet")
    If curWS Is Nothing Then
        Set curWS = curBook.Sheets.Add(After:=curBook.Sheets(curBook.Sheets.Count))
        curWS.Name = "Priority Sheet"
    End If
    On Error GoTo 0

    ' 设置表头 A1:I1
    With curWS
        .Range("A1:I1").Value = Array("JOB #", "PO #", "Customer", "Description", "Part #", "Qty.", "Ship Date", "Memo", "Status")
    End With

    lastRowCur = GetLastDataRow(curWS)

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
    selectSQL = "SELECT Job_Number, PO_Number, Customer_Name, Part_description, Part_Number, Job_Quantity, Delivery_Required_Date FROM jobs"
    result = SQLite3PrepareV2(dbHandle, selectSQL, stmtHandle)
    If result = 0 Then
        Do While SQLite3Step(stmtHandle) = 100
            jobNum = Trim(SQLite3ColumnText(stmtHandle, 0))
            If jobNum <> "" Then
                Dim jobData(6) As Variant
                jobData(0) = SQLite3ColumnText(stmtHandle, 0) ' Job_Number
                jobData(1) = SQLite3ColumnText(stmtHandle, 1) ' PO_Number
                jobData(2) = SQLite3ColumnText(stmtHandle, 2) ' Customer_Name
                jobData(3) = SQLite3ColumnText(stmtHandle, 3) ' Part_Description
                jobData(4) = SQLite3ColumnText(stmtHandle, 4) ' Part_Number
                jobData(5) = SQLite3ColumnText(stmtHandle, 5) ' Job_Quantity
                jobData(6) = SQLite3ColumnText(stmtHandle, 6) ' Delivery_Required_Date
                dbDict(jobNum) = jobData
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

                ' 获取 Shipped 表的最后一行
                Dim shipLastRow As Long: shipLastRow = GetLastDataRow(shippedWS) + 1

                ' 移动
                curWS.Rows(r).Resize(totalRowsToMove).Copy shippedWS.Rows(shipLastRow)

                ' 删除
                curWS.Rows(r).Resize(totalRowsToMove).Delete

                ' 更新行数
                movedCount = movedCount + totalRowsToMove
                lastRowCur = GetLastDataRow(curWS)
                ' LastRowCur = shippedLastRow

                ' !!!!! 重要: 同步 r 的值, 避免跳过行 !!!!!
                r = r - 1 ' 因为删除了, 本行要重新检查
                If r < 2 Then r = 2 ' 防止r小于表头

             Else
                 r = r + 1 '跳到下个job开始位置
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

    ' -------- 8. 数据库有但本地 Priority Sheet 无的条目，添加至 Priority Sheet，同时插入空行、并设置格式  --------
    Dim rg As Range
    Dim assemblySQL As String, assemblyStmtHandle As LongPtr, drawingNumber As String
    Dim jobInfoSQL As String, jobInfoStmtHandle As LongPtr
    Dim partDescription As String, jobQuantity As String, drawingRelease As String
    For Each k In dbDict.Keys
        If Not curDict.Exists(k) Then
            ' 在 Priority Sheet 追加记录
            lastRowCur = lastRowCur + 1

            ' 从数据库提取内容，并按字段添加新数据
            With curWS
                .Cells(lastRowCur, 1).Value = dbDict(k)(0) 'JOB #
                .Cells(lastRowCur, 2).Value = dbDict(k)(1) 'PO #
                .Cells(lastRowCur, 3).Value = dbDict(k)(2) 'Customer
                .Cells(lastRowCur, 4).Value = dbDict(k)(3) 'Description
                .Cells(lastRowCur, 5).Value = dbDict(k)(4) 'Part #
                .Cells(lastRowCur, 6).Value = dbDict(k)(5) 'Qty.
                .Cells(lastRowCur, 7).Value = dbDict(k)(6) 'Ship Date
            End With

            ' 调用CreateSingleHyperlink添加超链接
            Call CreateSingleHyperlink(curWS.Cells(lastRowCur, 5), dbHandle)

            ' 设置数据行区域：A-G列，设为橙色
            Set rg = curWS.Range(curWS.Cells(lastRowCur, 1), curWS.Cells(lastRowCur, 7))
            With rg
                .Interior.Color = RGB(255, 199, 44) ' 橙色
            End With

            ' 设置灰色背景 for 空行
            lastRowCur = lastRowCur + 1
            
            Set rg = curWS.Range(curWS.Cells(lastRowCur, 1), curWS.Cells(lastRowCur, 7))
            With rg
                .Interior.Color = RGB(242, 242, 242) ' 淡灰色
            End With

            ' 查询 assemblies 表获取 drawing_number
            Dim partNumber As String: partNumber = dbDict(k)(4) ' Part #
            assemblySQL = "SELECT drawing_number FROM assemblies WHERE part_number = '" & partNumber & "'"
            result = SQLite3PrepareV2(dbHandle, assemblySQL, assemblyStmtHandle)
            If result = 0 Then
                Dim firstDrawing As Boolean: firstDrawing = True
                Do While SQLite3Step(assemblyStmtHandle) = 100
                    drawingNumber = Trim(SQLite3ColumnText(assemblyStmtHandle, 0))
                    If drawingNumber <> "" Then
                        ' 根据 drawing_number 查询 jobs 表获取 part_description 和 job_quantity
                        jobInfoSQL = "SELECT Part_description, Job_Quantity, Drawing_Release FROM jobs WHERE Part_Number = '" & drawingNumber & "'"
                        Dim jobInfoResult As Long: jobInfoResult = SQLite3PrepareV2(dbHandle, jobInfoSQL, jobInfoStmtHandle)
                        If jobInfoResult = 0 Then
                            If SQLite3Step(jobInfoStmtHandle) = 100 Then
                                partDescription = Trim(SQLite3ColumnText(jobInfoStmtHandle, 0))
                                jobQuantity = Trim(SQLite3ColumnText(jobInfoStmtHandle, 1))
                                drawingRelease = Trim(SQLite3ColumnText(jobInfoStmtHandle, 2))

                                If firstDrawing Then
                                    ' 第一次找到 drawing_number，直接在当前灰色背景行进行操作
                                    curWS.Cells(lastRowCur, 5).Value = drawingNumber ' E列: drawing_number
                                    curWS.Cells(lastRowCur, 4).Value = partDescription ' D列: part_description
                                    curWS.Cells(lastRowCur, 6).Value = jobQuantity ' F列: job_quantity
                                    curWS.Cells(lastRowCur, 7).Value = drawingRelease ' G列: drawing_release
                                    firstDrawing = False
                                Else
                                    ' 后续找到的 drawing_number，先添加一个灰色行
                                    lastRowCur = lastRowCur + 1
                                    Set rg = curWS.Range(curWS.Cells(lastRowCur, 1), curWS.Cells(lastRowCur, 7))
                                    With rg
                                        .Interior.Color = RGB(242, 242, 242) ' 淡灰色
                                    End With

                                    ' 然后将查询到的信息填入
                                    curWS.Cells(lastRowCur, 5).Value = drawingNumber ' E列: drawing_number
                                    curWS.Cells(lastRowCur, 4).Value = partDescription ' D列: part_description
                                    curWS.Cells(lastRowCur, 6).Value = jobQuantity ' F列: job_quantity
                                    curWS.Cells(lastRowCur, 7).Value = drawingRelease ' G列: drawing_release
                                End If
                            End If
                            SQLite3Finalize jobInfoStmtHandle
                        End If
                    End If
                Loop
                SQLite3Finalize assemblyStmtHandle
            End If
        End If
    Next

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
    Dim lastRowD As Long, lastRowE As Long
    
    CountParts = 0

    ' 获取D列和E列的最后一行数据行
    lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    ' 取较大者作为基准
    Dim lastRow As Long
    If lastRowD >= lastRowE Then
        lastRow = lastRowD
    Else
        lastRow = lastRowE
    End If

    Debug.Print "CountParts: startRow = " & startRow & ", lastRowD = " & lastRowD & ", lastRowE = " & lastRowE

    If startRow >= lastRow Then
      lastRow = startRow + 1
    End If

    r = startRow + 1 ' 从下一行开始

    Do While r <= lastRow
        If Trim(ws.Cells(r, 1).Value) <> "" Then
            Debug.Print "CountParts: 在第 " & r & " 行遇到下一个Job，Part信息结束"
            Exit Do ' 遇到下一个 Job,退出
        End If
        
        CountParts = CountParts + 1
        r = r + 1 ' 移动到下一行
    Loop

    Debug.Print "CountParts: Found " & CountParts & " parts for Job at row " & startRow
    CountParts = CountParts
End Function

Function GetLastDataRow(ws As Worksheet) As Long
  ' 获取 Priority Sheet 数据的最后一行(综合A,D,E列)

  Dim lastRowA As Long, lastRowD As Long, lastRowE As Long

  ' 获取A,D,E列的最后一行
  lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
  lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
  lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

  ' 比较 D 和 E 的结果
  Dim lastRowDE As Long
  If lastRowD >= lastRowE Then
     lastRowDE = lastRowD
  Else
     lastRowDE = lastRowE
  End If

  ' 若 D 或 E 的结果大于 A, 则取 D/E 中的较大值，否则+1
  If lastRowDE > lastRowA Then
     GetLastDataRow = lastRowDE
  Else
     If lastRowA < 2 Then
       GetLastDataRow = 1
     Else
       GetLastDataRow = lastRowA + 1 ' 确保新行
     End If
  End If

  Debug.Print "GetLastDataRow: lastRowA=" & lastRowA & ", lastRowD=" & lastRowD & ", lastRowE=" & lastRowE & ", Result=" & GetLastDataRow
End Function


