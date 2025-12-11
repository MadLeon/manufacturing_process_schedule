' VBA
Option Explicit

' This sub is triggered by clicking add new record button in a sheet called input form.
Sub AddNewJobToDB()
    ' --- Goal: Directly add new job entry from order entry log to jobs.db ---
    Dim dbPath As String
    Dim curBook As Workbook, inputWS As Worksheet, deliveryWS As Worksheet
    Dim lastRowDelivery As Long
    Dim jobDict As Object
    Dim r As Long, i As Integer
    Dim insertSQL As String, checkSQL As String, historyInsertSQL As String, deleteSQL As String
    Dim k As Variant
    Dim OE_Number As String, Job_Number As String, Customer_Name As String, Job_Quantity As String, Part_Number As String
    Dim Revision As String, Customer_Contact As String, Drawing_Release As String, Line_Number As String, Part_Description As String
    Dim Unit_Price As String, PO_Number As String, Packing_Slip As String, Packing_Quantity As String, Invoice_Number As String
    Dim Delivery_Required_Date As String, Delivery_Shipped_Date As String
    Dim rs As Variant

    ' 1. Set object variables
    Set curBook = ThisWorkbook
    Set inputWS = curBook.Sheets("Input Form")
    Set deliveryWS = curBook.Sheets("DELIVERY SCHEDULE")

    ' 2. Initialize SQLite
    If Not InitializeSQLite(DB_PATH) Then
        MsgBox "Failed to initialize SQLite. Check the mod_SQLite module.", vbCritical
        Exit Sub
    End If

    ' 3. Get data from Input Form
    OE_Number = Trim(inputWS.Range("OE").Value)
    Job_Number = Trim(inputWS.Range("JobNum").Value)
    Customer_Name = Trim(inputWS.Range("Customer").Value)
    Job_Quantity = Trim(inputWS.Range("qty").Value)
    Part_Number = Trim(inputWS.Range("Parts").Value)
    Revision = Trim(inputWS.Range("Revision").Value)
    Customer_Contact = Trim(inputWS.Range("contact").Value)
    Drawing_Release = Trim(inputWS.Range("od").Value)
    Line_Number = Trim(inputWS.Range("poline").Value)
    Part_Description = Trim(inputWS.Range("desc").Value)
    Unit_Price = Trim(inputWS.Range("price").Value)
    PO_Number = Trim(inputWS.Range("po").Value)
    Delivery_Required_Date = Trim(inputWS.Range("date").Value)

    ' 4. Check if required data exists
    If OE_Number = "" Or Job_Number = "" Then
        Debug.Print "Critical input is missing.", vbCritical
        CloseSQLite
        Exit Sub
    End If
    
    ' 5. Build dictionary of current Job_Number in delivery schedule
    Set jobDict = CreateObject("Scripting.Dictionary")
    lastRowDelivery = LastRow(deliveryWS)
    
        For r = 4 To lastRowDelivery
        Dim jn As String, ln As String, dr As String
        jn = Trim(deliveryWS.Cells(r, 2).Value)   ' Job_Number
        ln = Trim(deliveryWS.Cells(r, 9).Value)   ' Line_Number
        dr = Trim(deliveryWS.Cells(r, 16).Value)  ' Delivery_Required_Date

        If jn <> "" Then
            Dim compositeKey As String
            compositeKey = jn & "|" & ln & "|" & dr
            jobDict(compositeKey) = r
        End If
    Next r

    Debug.Print "Number of unique job entries: "; jobDict.Count
    
    ' 6. Insert new job into jobs table
    ' Generate the unique key
    Dim Unique_Key As String
    Unique_Key = Job_Number & "|" & Line_Number & "|" & Delivery_Required_Date

    insertSQL = "INSERT INTO jobs (OE_Number, Job_Number, Customer_Name, Job_Quantity, Part_Number, Revision, Customer_Contact, " & _
                "Drawing_Release, Line_Number, Part_Description, Unit_Price, PO_Number, Delivery_Required_Date, unique_key, " & _
                "create_timestamp, last_modified) " & _
                "VALUES ('" & Replace(OE_Number, "'", "''") & "', '" & Replace(Job_Number, "'", "''") & "', '" & Replace(Customer_Name, "'", "''") & "', '" & _
                Replace(Job_Quantity, "'", "''") & "', '" & Replace(Part_Number, "'", "''") & "', '" & Replace(Revision, "'", "''") & "', '" & _
                Replace(Customer_Contact, "'", "''") & "', '" & Replace(Drawing_Release, "'", "''") & "', '" & Replace(Line_Number, "'", "''") & "', '" & _
                Replace(Part_Description, "'", "''") & "', '" & Replace(Unit_Price, "'", "''") & "', '" & Replace(PO_Number, "'", "''") & "', '" & _
                Replace(Delivery_Required_Date, "'", "''") & "', '" & Replace(Unique_Key, "'", "''") & "', datetime('now','localtime'), datetime('now','localtime'))"

    If ExecuteNonQuery(insertSQL) Then
       Debug.Print "New job added to jobs table: " & Job_Number
    Else
       Debug.Print "Failed to add new job to jobs table: " & Job_Number
    End If

    On Error GoTo 0

    ' 7. Check for completed jobs (exists in DB but not in delivery schedule) and move to job_history
    Dim jobHistoryExists As Boolean
    jobHistoryExists = TableExists("job_history")

    If Not jobHistoryExists Then
        ' Create job_history table if it doesn't exist
        Dim sqlCreateHistory As String
        sqlCreateHistory = "CREATE TABLE IF NOT EXISTS job_history (" & _
                           "job_id INTEGER PRIMARY KEY AUTOINCREMENT, " & _
                           "oe_number TEXT, job_number TEXT, customer_name TEXT, job_quantity TEXT, " & _
                           "part_number TEXT, revision TEXT, customer_contact TEXT, drawing_release TEXT, line_number TEXT, " & _
                           "part_description TEXT, unit_price TEXT, po_number TEXT, packing_slip TEXT, packing_quantity TEXT, " & _
                           "invoice_number TEXT, delivery_required_date TEXT, delivery_shipped_date TEXT, " & _
                           "create_timestamp TEXT, last_modified TEXT, completed_timestamp TEXT DEFAULT (datetime('now','localtime')))"
        If Not ExecuteNonQuery(sqlCreateHistory) Then
            MsgBox "Failed to create job_history table. Check debug log.", vbCritical
            CloseSQLite
            Exit Sub
        End If
        Debug.Print "job_history table created"
    End If
    
    ' 8. Move completed jobs to job_history and delete from jobs
    'Dim rs As Variant, jobID As Long
    'checkSQL = "SELECT job_id, OE_Number, Job_Number, Customer_Name, Job_Quantity, Part_Number, Revision, Customer_Contact, " & _
    '           "Drawing_Release, Line_Number, Part_Description, Unit_Price, PO_Number, Packing_Slip, Packing_Quantity, Invoice_Number, Delivery_Required_Date, Delivery_Shipped_Date, " & _
    '           "create_timestamp, last_modified FROM jobs"

    'rs = ExecuteSQL(checkSQL)
    'If Not IsEmpty(rs) Then
    '    For r = LBound(rs) To UBound(rs)
    '        Dim jobNumber As String, lineNumber As String, deliveryDate As String
    '        jobNumber = rs(r)(2)
    '        lineNumber = rs(r)(9)
    '        deliveryDate = rs(r)(16)

    '        ' Check if the job exists in DELIVERY SCHEDULE with the same Job_Number, Line_Number, and Delivery_Required_Date
    '        Dim foundInDeliverySchedule As Boolean
    '        foundInDeliverySchedule = False

    '        If jobDict.Exists(jobNumber) Then
    '            ' Check Line_Number and Delivery_Required_Date
    '            If Trim(deliveryWS.Cells(jobDict(jobNumber), 9).Value) = lineNumber And Trim(deliveryWS.Cells(jobDict(jobNumber), 16).Value) = deliveryDate Then
    '                foundInDeliverySchedule = True
    '            End If
    '        End If

    '        If Not foundInDeliverySchedule Then
    '            jobID = rs(r)(0)
    '            OE_Number = rs(r)(1)
    '            Job_Number = rs(r)(2)
    '            Customer_Name = rs(r)(3)
    '            Job_Quantity = rs(r)(4)
    '            Part_Number = rs(r)(5)
    '            Revision = rs(r)(6)
    '            Customer_Contact = rs(r)(7)
    '            Drawing_Release = rs(r)(8)
    '            Line_Number = rs(r)(9)
    '            Part_Description = rs(r)(10)
    '            Unit_Price = rs(r)(11)
    '            PO_Number = rs(r)(12)
    '            Delivery_Required_Date = rs(r)(16)
                    
    '        historyInsertSQL = "INSERT INTO job_history (OE_Number, Job_Number, Customer_Name, Job_Quantity, Part_Number, Revision, Customer_Contact, " & _
    '                            "Drawing_Release, Line_Number, Part_Description, Unit_Price, PO_Number, Delivery_Required_Date, " & _
    '                            "create_timestamp, last_modified) VALUES ('" & Replace(OE_Number, "'", "''") & "', '" & Replace(Job_Number, "'", "''") & "', '" & Replace(Customer_Name, "'", "''") & "', '" & _
    '                            Replace(Job_Quantity, "'", "''") & "', '" & Replace(Part_Number, "'", "''") & "', '" & Replace(Revision, "'", "''") & "', '" & _
    '                            Replace(Customer_Contact, "'", "''") & "', '" & Replace(Drawing_Release, "'", "''") & "', '" & Replace(Line_Number, "'", "''") & "', '" & _
    '                            Replace(Part_Description, "'", "''") & "', '" & Replace(Unit_Price, "'", "''") & "', '" & Replace(PO_Number, "'", "''") & "', '" & _
    '                            Replace(Delivery_Required_Date, "'", "''") & "', '" & rs(r)(18) & "', '" & rs(r)(19) & "')" ' Use existing timestamps

    '        If ExecuteNonQuery(historyInsertSQL) Then
    '            Debug.Print "Job moved to job_history: " & Job_Number
    '            'Now delete from jobs table
    '            deleteSQL = "DELETE FROM jobs WHERE job_id = " & jobID
    '            If ExecuteNonQuery(deleteSQL) Then
    '                Debug.Print "Job deleted from jobs table: " & Job_Number
    '            Else
    '                Debug.Print "Failed to delete job from jobs table: " & Job_Number
    '            End If
    '        Else
    '            Debug.Print "Failed to move job to job_history: " & Job_Number
    '        End If
    '    End If
    'Next r
    'End If
    
    ' 9. Close SQLite
    CloseSQLite
    Exit Sub
    
    'Debug.Print "AddNewJobToDB completed!"

HandleDuplicate:
    If Err.Number <> 0 Then
        Debug.Print "A record with the same Job Number, Line Number, and Delivery Required Date already exists.", vbExclamation
        Err.Clear
    End If
    CloseSQLite
    Exit Sub

End Sub


' Helper function to get the last row with data in a worksheet
Function LastRow(ws As Worksheet) As Long
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
End Function

' Helper function to check if a table exists in the database
Function TableExists(tableName As String) As Boolean
    Dim sql As String, rs As Variant
    sql = "SELECT name FROM sqlite_master WHERE type='table' AND name='" & tableName & "';"
    rs = ExecuteSQL(sql)
    TableExists = Not IsEmpty(rs)
End Function

