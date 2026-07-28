Option Explicit

' =========================================================
' Module: Function_DatabaseSync
' Purpose: On every save, keep the job-management database (record.db) in
'          sync with this DIR file's part_attachment row: refresh its
'          updated_at if the row already exists, otherwise insert it.
'          Called from Workbook_BeforeSave (see dir/ThisWorkBook.bas).
'          An explicit Ctrl+S (see Data_SharedValues.IsCtrlSSave, set by
'          dir_signature/Handler_SaveLock.HandleCtrlSSave) additionally marks
'          the attachment's status as 'completed' - unless it's already
'          'reviewed', which is a step further along and shouldn't regress.
' =========================================================

' TODO: point back to the production database ("\\rtdnas2\OE\record.db", see
' .github/skills/database) before deployment - currently pointed at a local
' dev copy for testing.
Private Const DB_PATH As String = "C:\Users\ee\drawing_tree\data\record.db"

' SQLite3.dll / SQLite3_StdCall.dll cannot be looked up next to ThisWorkbook.Path
' because DIR files are copied out to many different customer folders (see
' dir log/mod_CheckSelectedCell.bas). They must instead live next to the master
' "DIR Template.xlsm", which every copy is made from.
Private Const DLL_DIR As String = "\\rtdnas2\QCReports\FINAL REPORTS\"

' Header cell locations on the DIR sheet (see Issue #2 data mapping table).
Private Const CELL_JOB_NUMBER As String = "H6"
Private Const CELL_DRAWING_REVISION As String = "C7"

' Public entry point. Never lets a database problem block the save:
' any failure is logged to the Immediate window and swallowed.
Public Sub SyncDIRAttachment()
    On Error GoTo Failed
    SyncDIRAttachmentCore
    Exit Sub

Failed:
    Debug.Print "SyncDIRAttachment failed: " & Err.Description
    CloseSQLite
End Sub

Private Sub SyncDIRAttachmentCore()
    ' Consume the flag immediately so it can never leak into a later save
    ' (e.g. if something below exits early or errors out).
    Dim markCompleted As Boolean
    markCompleted = Data_SharedValues.IsCtrlSSave
    Data_SharedValues.IsCtrlSSave = False
    Debug.Print "[SyncDIRAttachmentCore] markCompleted = " & markCompleted

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)

    Dim filePath As String, fileName As String
    filePath = ThisWorkbook.FullName
    fileName = BaseFileName(ThisWorkbook.Name)

    Dim jobNumber As String, drawingNumber As String
    jobNumber = Trim$(CStr(ws.Range(CELL_JOB_NUMBER).Value))
    drawingNumber = ExtractDrawingNumber(Trim$(CStr(ws.Range(CELL_DRAWING_REVISION).Value)))

    ' Header not filled in yet (e.g. a blank template) - nothing to sync.
    If jobNumber = "" Or drawingNumber = "" Then
        Debug.Print "[SyncDIRAttachmentCore] jobNumber or drawingNumber empty - exiting. jobNumber=[" & jobNumber & "] drawingNumber=[" & drawingNumber & "]"
        Exit Sub
    End If

    If Not InitializeSQLite(DB_PATH, DLL_DIR) Then
        Debug.Print "[SyncDIRAttachmentCore] InitializeSQLite failed for DB_PATH=" & DB_PATH & " DLL_DIR=" & DLL_DIR
        Exit Sub
    End If

    Dim existingId As Long
    existingId = FindAttachmentIdByPath(filePath)
    Debug.Print "[SyncDIRAttachmentCore] filePath=[" & filePath & "] existingId=" & existingId

    If existingId > 0 Then
        Dim updateSql As String, updateOk As Boolean
        If markCompleted Then
            updateSql = "UPDATE part_attachment SET " & _
                "status = CASE WHEN status = 'reviewed' THEN status ELSE 'completed' END, " & _
                "updated_at = datetime('now', 'localtime') WHERE id = " & existingId
        Else
            updateSql = "UPDATE part_attachment SET updated_at = datetime('now', 'localtime') WHERE id = " & existingId
        End If
        Debug.Print "[SyncDIRAttachmentCore] running: " & updateSql
        updateOk = ExecuteNonQuery(updateSql)
        Debug.Print "[SyncDIRAttachmentCore] ExecuteNonQuery result = " & updateOk
    Else
        Debug.Print "[SyncDIRAttachmentCore] no existing row - inserting"
        InsertAttachment jobNumber, drawingNumber, fileName, filePath, markCompleted
    End If

    CloseSQLite
End Sub

' Looks up an existing part_attachment row by its unique file_path.
' Returns 0 if no matching row exists.
Private Function FindAttachmentIdByPath(filePath As String) As Long
    Dim result As Variant
    result = ExecuteQuery("SELECT id FROM part_attachment WHERE file_path = '" & EscapeSql(filePath) & "'")

    If IsNull(result) Then
        FindAttachmentIdByPath = 0
    Else
        FindAttachmentIdByPath = CLng(result(0)(0))
    End If
End Function

' Resolves job_number/drawing_number to a job/part pair and inserts the
' part_attachment row. part_id / order_item_id are left NULL when they
' cannot be resolved (per Issue #2, this must not block or warn the user -
' it just means the underlying job/part/order_item data isn't in the DB yet).
' A new row has no prior 'reviewed' status to protect, so markCompleted maps
' straight to the initial status.
Private Sub InsertAttachment(jobNumber As String, drawingNumber As String, fileName As String, filePath As String, markCompleted As Boolean)
    Dim jobId As Long, partId As Long, orderItemId As Long
    jobId = FindJobId(jobNumber)
    partId = FindLatestPartId(drawingNumber)
    orderItemId = FindOrderItemId(jobId, partId)

    Dim partIdSql As String, orderItemIdSql As String
    partIdSql = IIf(partId > 0, CStr(partId), "NULL")
    orderItemIdSql = IIf(orderItemId > 0, CStr(orderItemId), "NULL")

    Dim statusValue As String
    statusValue = IIf(markCompleted, "completed", "in progress")

    Dim sql As String
    sql = "INSERT INTO part_attachment (part_id, order_item_id, file_type, file_name, file_path, is_active, status) VALUES (" & _
        partIdSql & ", " & orderItemIdSql & ", 'dir', '" & EscapeSql(fileName) & "', '" & EscapeSql(filePath) & "', 1, '" & statusValue & "')"

    Debug.Print "[InsertAttachment] jobId=" & jobId & " partId=" & partId & " orderItemId=" & orderItemId & " statusValue=" & statusValue
    Debug.Print "[InsertAttachment] running: " & sql
    Debug.Print "[InsertAttachment] ExecuteNonQuery result = " & ExecuteNonQuery(sql)
End Sub

Private Function FindJobId(jobNumber As String) As Long
    Dim result As Variant
    result = ExecuteQuery("SELECT id FROM job WHERE job_number = '" & EscapeSql(jobNumber) & "'")

    If IsNull(result) Then
        FindJobId = 0
    Else
        FindJobId = CLng(result(0)(0))
    End If
End Function

' A drawing number can have several part rows (one per revision, see Issue #2
' "part 冲突处理"); always take the newest revision.
Private Function FindLatestPartId(drawingNumber As String) As Long
    Dim result As Variant
    result = ExecuteQuery("SELECT id FROM part WHERE drawing_number = '" & EscapeSql(drawingNumber) & _
        "' ORDER BY revision DESC LIMIT 1")

    If IsNull(result) Then
        FindLatestPartId = 0
    Else
        FindLatestPartId = CLng(result(0)(0))
    End If
End Function

' The part on the DIR file may not be the top-level part actually on the order -
' it can be a sub-component of an assembly several levels down in part_tree
' (BOM). So this doesn't just match order_item.part_id = partId directly; it
' walks partId up through part_tree (child_id -> parent_id) collecting every
' ancestor, and matches an order_item whose part is partId itself or any of
' those ancestors.
Private Function FindOrderItemId(jobId As Long, partId As Long) As Long
    If jobId = 0 Or partId = 0 Then
        FindOrderItemId = 0
        Exit Function
    End If

    Dim sql As String
    sql = "WITH RECURSIVE ancestor_part(id) AS (" & _
        "SELECT " & partId & " " & _
        "UNION " & _
        "SELECT pt.parent_id FROM part_tree pt JOIN ancestor_part a ON pt.child_id = a.id" & _
        ") " & _
        "SELECT id FROM order_item WHERE job_id = " & jobId & _
        " AND part_id IN (SELECT id FROM ancestor_part) LIMIT 1"

    Dim result As Variant
    result = ExecuteQuery(sql)

    If IsNull(result) Then
        FindOrderItemId = 0
    Else
        FindOrderItemId = CLng(result(0)(0))
    End If
End Function

' C7 holds "{drawing_number} REV. {revision}" (see dir log/mod_CheckSelectedCell.bas).
' Only the drawing number is needed - see FindLatestPartId for why the revision
' text itself is ignored.
Private Function ExtractDrawingNumber(c7 As String) As String
    Dim pos As Long
    pos = InStr(1, c7, " REV.", vbTextCompare)

    If pos > 0 Then
        ExtractDrawingNumber = Trim$(Left$(c7, pos - 1))
    Else
        ExtractDrawingNumber = c7
    End If
End Function

' Strips the file extension from a workbook name, e.g. "Foo @123.xlsm" -> "Foo @123".
Private Function BaseFileName(workbookName As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(workbookName, ".")

    If dotPos > 0 Then
        BaseFileName = Left$(workbookName, dotPos - 1)
    Else
        BaseFileName = workbookName
    End If
End Function

Private Function EscapeSql(value As String) As String
    EscapeSql = Replace(value, "'", "''")
End Function
