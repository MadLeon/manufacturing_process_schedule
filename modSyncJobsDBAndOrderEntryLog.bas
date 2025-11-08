Option Explicit

' sync oe entry to jobs.db, then sync current workbook to jobs.db

Sub SyncJobsDBAndOrderEntryLog()
    ' --- 目标: jobs.db == oeentry(DELIVERY SCHEDULE), 本文件(DELIVERY SCHEDULE) ≈ jobs.db ---
    Dim dbPath As String, dbHandle As LongPtr, result As Long, stmtHandle As LongPtr
    Dim jobsDBExists As Boolean
    Dim srcEntryApp As Object, srcEntryBook As Workbook, srcEntryWS As Worksheet
    Dim curBook As Workbook, curWS As Worksheet, shippedWS As Worksheet
    Dim lastRowEntry As Long, lastRowCur As Long
    Dim entryDict As Object, dbDict As Object, curDict As Object
    Dim r As Long, i As Integer
    Dim insertSQL As String, selectSQL As String, deleteSQL As String
    Dim k As Variant

    dbPath = ThisWorkbook.Path & "\jobs.db"
    jobsDBExists = (Dir(dbPath) <> "")

    ' 1. 初始化DLL
    result = SQLite3Initialize(ThisWorkbook.Path)
    If result <> SQLITE_INIT_OK Then
        MsgBox "SQLite3初始化失败": Exit Sub
    End If

    ' 2. 打开/新建数据库
    result = SQLite3Open(dbPath, dbHandle)
    If result <> SQLITE_OK Then
        MsgBox "无法打开数据库: " & SQLite3ErrMsg(dbHandle): SQLite3Free: Exit Sub
    End If

    ' 3. 如无则建表，所有字段均为TEXT类型
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
        Debug.Print "数据库jobs.db已新建并初始化jobs表 (所有字段均为TEXT)"
    Else
        Debug.Print "数据库jobs.db已存在"
    End If

    ' 4. 打开 oe entry log.xlsm 的 DELIVERY SCHEDULE sheet(只读后台)
    Set srcEntryApp = CreateObject("Excel.Application")
    srcEntryApp.Visible = False
    srcEntryApp.DisplayAlerts = False
    Set srcEntryBook = srcEntryApp.Workbooks.Open(ThisWorkbook.Path & "\order entry log.xlsm", ReadOnly:=True)
    Set srcEntryWS = srcEntryBook.Sheets("DELIVERY SCHEDULE")
    lastRowEntry = srcEntryWS.Cells(srcEntryWS.Rows.Count, 1).End(-4162).Row
    Debug.Print "oe entry 行数: "; lastRowEntry - 3

    ' 5. 生成字典：oeentry(Job_Number为key)
    Set entryDict = CreateObject("Scripting.Dictionary")
    For r = 4 To lastRowEntry
        If Trim(srcEntryWS.Cells(r, 2).Value) <> "" Then
            entryDict(Trim(srcEntryWS.Cells(r, 2).Value)) = r
        End If
    Next
    Debug.Print "oe entry Job_Number 条数: "; entryDict.Count

    ' 6. 生成数据库当前Job_Number字典
    Set dbDict = CreateObject("Scripting.Dictionary")
    selectSQL = "SELECT Job_Number FROM jobs"
    result = SQLite3PrepareV2(dbHandle, selectSQL, stmtHandle)
    If result = SQLITE_OK Then
        Do While SQLite3Step(stmtHandle) = 100
            dbDict(Trim(SQLite3ColumnText(stmtHandle, 0))) = 1
        Loop
        SQLite3Finalize stmtHandle
    End If
    Debug.Print "数据库 jobs.db Job_Number 条数: "; dbDict.Count

    ' 7. ============ A. 同步数据库与oe entry（强一致） =============
    '   - 1) 新增 oe entry有(db无) ==> db插入
    '   - 2) 删除 db有(oe entry无) ==> db删除
    insertSQL = "INSERT INTO jobs (OE_Number, Job_Number, Customer_Name, Job_Quantity, Part_Number, Revision, Customer_Contact, " & _
      "Drawing_Release, Line_Number, Part_Description, Unit_Price, PO_Number, Packing_Slip, Packing_Quantity, Invoice_Number, Delivery_Required_Date, Delivery_Shipped_Date, Last_Modified) " & _
      "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

    ' (1) 新增，全部用SQLite3BindText
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
                ' Debug.Print "新增到数据库: Job_Number=" & k
                addedToDB = addedToDB + 1
            End If
        End If
    Next
    Debug.Print "本次共新增到数据库条目数: "; addedToDB

    ' (2) 删除
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
                Debug.Print "已从数据库删除: Job_Number=" & k
                deletedFromDB = deletedFromDB + 1
            End If
        End If
    Next
    Debug.Print "本次共从数据库删除条目数: "; deletedFromDB

    srcEntryBook.Close False
    srcEntryApp.Quit
    Set srcEntryBook = Nothing
    Set srcEntryWS = Nothing
    Set srcEntryApp = Nothing
    
    SQLite3Close dbHandle
    SQLite3Free
    
    Debug.Print "同步 oe entry 与 jobs.db 完成！"
End Sub

