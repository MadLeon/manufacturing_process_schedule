Option Explicit

' ==================== 配置变量 ====================
' 数据库文件路径
Const DB_PATH As String = "C:\Users\ee\manufacturing_process_schedule\jobs.db" ' 更新为实际的网络路径
' 订单输入日志文件路径
Const OE_ENTRY_PATH As String = "C:\Users\ee\manufacturing_process_schedule\order entry log.xlsm" ' 更新为实际的网络路径

' ==================== 主同步函数 ====================
' 功能：将订单输入日志中的数据同步到数据库，确保两者保持一致
' 目标：jobs.db == oeentry(DELIVERY SCHEDULE)，当前文件(DELIVERY SCHEDULE) ≈ jobs.db
Sub SyncJobsDBAndOrderEntryLog()
    Dim dbPath As String, dbHandle As LongPtr, result As Long, stmtHandle As LongPtr
    Dim jobsDBExists As Boolean
    Dim srcEntryApp As Object, srcEntryBook As Workbook, srcEntryWS As Worksheet
    Dim curBook As Workbook, curWS As Worksheet, shippedWS As Worksheet
    Dim lastRowEntry As Long, lastRowCur As Long
    Dim entryDict As Object, dbDict As Object, curDict As Object
    Dim r As Long, i As Integer
    Dim insertSQL As String, selectSQL As String, deleteSQL As String
    Dim k As Variant

    dbPath = DB_PATH
    jobsDBExists = (Dir(dbPath) <> "")

    ' 步骤1：初始化 SQLite3 DLL
    result = SQLite3Initialize(ThisWorkbook.Path)
    If result <> SQLITE_INIT_OK Then
        MsgBox "SQLite3 初始化失败": Exit Sub
    End If

    ' 步骤2：打开或创建数据库
    result = SQLite3Open(dbPath, dbHandle)
    If result <> SQLITE_OK Then
        MsgBox "无法打开数据库: " & SQLite3ErrMsg(dbHandle): SQLite3Free: Exit Sub
    End If

    ' 步骤3：创建表（如果不存在），所有字段类型为 TEXT
    If Not jobsDBExists Then
        Dim sqlCreate As String
        sqlCreate = "CREATE TABLE IF NOT EXISTS jobs (" & _
            "job_id INTEGER PRIMARY KEY AUTOINCREMENT, " & _
            "oe_number TEXT, job_number TEXT, customer_name TEXT, job_quantity TEXT, " & _
            "part_number TEXT, revision TEXT, customer_contact TEXT, drawing_release TEXT, line_number TEXT, " & _
            "part_description TEXT, unit_price TEXT, po_number TEXT, packing_slip TEXT, packing_quantity TEXT, " & _
            "invoice_number TEXT, delivery_required_date TEXT, delivery_shipped_date TEXT, " & _
            "create_timestamp TEXT DEFAULT (datetime('now','localtime')), last_modified TEXT)"

        result = SQLite3PrepareV2(dbHandle, sqlCreate, stmtHandle)
        If result = SQLITE_OK Then SQLite3Step stmtHandle: SQLite3Finalize stmtHandle
        Debug.Print "数据库 jobs.db 已创建，jobs 表已初始化（所有字段为 TEXT 类型）"
    Else
        Debug.Print "数据库 jobs.db 已存在"
    End If

    ' 步骤4：打开订单日志中的"发货计划"工作表（只读模式，后台打开）
    Set srcEntryApp = CreateObject("Excel.Application")
    srcEntryApp.Visible = False
    srcEntryApp.DisplayAlerts = False
    Set srcEntryBook = srcEntryApp.Workbooks.Open(OE_ENTRY_PATH, ReadOnly:=True)
    Set srcEntryWS = srcEntryBook.Sheets("DELIVERY SCHEDULE")
    lastRowEntry = srcEntryWS.Cells(srcEntryWS.rows.Count, 1).End(-4162).row
    Debug.Print "订单日志中的数据行数: ", lastRowEntry - 3

    ' 步骤5：构建字典：以 Job_Number（工作单号）为键的订单数据字典
    Set entryDict = CreateObject("Scripting.Dictionary")
    For r = 4 To lastRowEntry
        If Trim(srcEntryWS.Cells(r, 2).Value) <> "" Then
            entryDict(Trim(srcEntryWS.Cells(r, 2).Value)) = r
        End If
    Next
    Debug.Print "订单日志中的工作单号数量: ", entryDict.Count

    ' 步骤6：构建字典：数据库中现有的 Job_Number（工作单号）
    Set dbDict = CreateObject("Scripting.Dictionary")
    selectSQL = "SELECT Job_Number FROM jobs"
    result = SQLite3PrepareV2(dbHandle, selectSQL, stmtHandle)
    If result = SQLITE_OK Then
        Do While SQLite3Step(stmtHandle) = 100
            dbDict(Trim(SQLite3ColumnText(stmtHandle, 0))) = 1
        Loop
        SQLite3Finalize stmtHandle
    End If
    Debug.Print "数据库中的工作单号数量: ", dbDict.Count

    ' ==================== 步骤7：同步数据库和订单日志（强一致性） ====================
    ' A. 新增操作：订单日志中有，但数据库中没有 ==> 插入到数据库
    ' B. 删除操作：数据库中有，但订单日志中没有 ==> 从数据库中删除
    insertSQL = "INSERT INTO jobs (OE_Number, Job_Number, Customer_Name, Job_Quantity, Part_Number, Revision, Customer_Contact, " & _
      "Drawing_Release, Line_Number, Part_Description, Unit_Price, PO_Number, Packing_Slip, Packing_Quantity, Invoice_Number, Delivery_Required_Date, Delivery_Shipped_Date, Last_Modified) " & _
      "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

    ' (1) 新增数据：将在订单日志中但不在数据库中的记录插入到数据库
    Dim addedToDB As Long
    addedToDB = 0
    For Each k In entryDict.Keys
        If Not dbDict.Exists(k) Then
            r = entryDict(k)
            result = SQLite3PrepareV2(dbHandle, insertSQL, stmtHandle)
            If result = SQLITE_OK Then
                For i = 1 To 17
                    SQLite3BindText stmtHandle, i, Trim(srcEntryWS.Cells(r, i).Text)
                Next
                SQLite3BindText stmtHandle, 18, Format(Now, "yyyy-mm-dd HH:nn:ss")
                SQLite3Step stmtHandle
                SQLite3Finalize stmtHandle
                ' Debug.Print "已添加到数据库: Job_Number=" & k
                addedToDB = addedToDB + 1
            End If
        End If
    Next
    Debug.Print "本次添加到数据库的记录数: ", addedToDB

    ' (2) 删除数据：将在数据库中但不在订单日志中的记录从数据库中删除
    Dim deletedFromDB As Long
    deletedFromDB = 0
    For Each k In dbDict.Keys
        If Not entryDict.Exists(k) Then
            deleteSQL = "DELETE FROM jobs WHERE Job_Number = ?"
            result = SQLite3PrepareV2(dbHandle, deleteSQL, stmtHandle)
            If result = SQLITE_OK Then
                SQLite3BindText stmtHandle, 1, k
                SQLite3Step stmtHandle
                SQLite3Finalize stmtHandle
                Debug.Print "已从数据库中删除: Job_Number=" & k
                deletedFromDB = deletedFromDB + 1
            End If
        End If
    Next
    Debug.Print "本次从数据库中删除的记录数: ", deletedFromDB

    ' 关闭订单日志文件
    srcEntryBook.Close False
    srcEntryApp.Quit
    Set srcEntryBook = Nothing
    Set srcEntryWS = Nothing
    Set srcEntryApp = Nothing
    
    ' 关闭数据库连接并释放资源
    SQLite3Close dbHandle
    SQLite3Free
    
    Debug.Print "订单日志和数据库同步完成！"
End Sub
