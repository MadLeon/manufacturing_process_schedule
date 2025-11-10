Option Explicit

' DON'T USE THIS!!!
' This function use cached db method, it's very slow for large data sets
Sub AddHyperlinkToCell(cell As Range, drawingsDict As Object)
    ' Adds hyperlink to cell using cached data
    Dim drawingNumber As String, fileLocation As String
    Dim key As Variant

    drawingNumber = Trim(cell.Value)

    If drawingNumber <> "" Then
        ' Check if cell already has hyperlink
        On Error Resume Next
        Dim hyp As Hyperlink
        Set hyp = cell.Worksheet.Hyperlinks(cell.Address)
        On Error GoTo 0

        If Not hyp Is Nothing Then
            ' Skip if already hyperlinked
            Set hyp = Nothing
        Else
            fileLocation = "" ' Reset for each drawingNumber

            ' 1. First check drawing_number in cache
            For Each key In drawingsDict.Keys
                If drawingsDict(key).Item("drawing_number") = drawingNumber Then
                    fileLocation = drawingsDict(key).Item("file_location")
                    Exit For
                End If
            Next key

            ' 2. If not found, check drawing_name with customer match
            If fileLocation = "" Then
                Dim customerName As String
                customerName = Trim(cell.Worksheet.Cells(cell.Row, 3).Value) ' Column C is Customer

                For Each key In drawingsDict.Keys
                    If InStr(1, key, drawingNumber, vbTextCompare) > 0 And _
                       InStr(1, drawingsDict(key).Item("file_location"), customerName, vbTextCompare) > 0 Then
                        fileLocation = drawingsDict(key).Item("file_location")
                        Exit For
                    End If
                Next key
            End If

            ' 3. Add hyperlink if file found
            If fileLocation <> "" Then
                On Error Resume Next
                cell.Worksheet.Hyperlinks.Add Anchor:=cell, Address:=fileLocation, TextToDisplay:=drawingNumber
                On Error GoTo 0

                ' Set font style
                With cell.Font
                    .Name = "Cambria"
                    .Size = 16
                End With

                ' Update drawing_number in database
                If Not modSQLite.ExecuteNonQuery("UPDATE drawings SET drawing_number = '" & drawingNumber & "' WHERE drawing_name = '" & key & "'") Then
                    Debug.Print "Error updating drawing number"
                End If
            End If
        End If
    End If
End Sub

' USE THIS INSTEAD!!!
Sub CreateSingleHyperlink(cell As Range)
    ' Adds hyperlink to single cell by querying database directly
    Dim drawingNumber As String, fileLocation As String, dName As String
    Dim results As Variant

    drawingNumber = Trim(cell.Value)

    If drawingNumber <> "" Then
        ' Check if cell already has hyperlink
        On Error Resume Next
        Dim hyp As Hyperlink
        Set hyp = cell.Worksheet.Hyperlinks(cell.Address)
        On Error GoTo 0

        If Not hyp Is Nothing Then
            ' Skip if already hyperlinked
            Set hyp = Nothing
        Else
            fileLocation = "" ' Reset for each drawingNumber

            ' 1. First query drawing_number
            results = modSQLite.ExecuteSQL("SELECT file_location FROM drawings WHERE drawing_number = '" & drawingNumber & "'")
            If Not IsNull(results) Then
                fileLocation = Trim(results(0)(0))
            End If

            ' 2. If not found, query drawing_name with customer match
            If fileLocation = "" Then
                Dim customerName As String, folderAlias As String
                customerName = Trim(cell.Worksheet.Cells(cell.Row, 3).Value) ' Column C is Customer
                folderAlias = ""

                ' Try to get folder alias from customer_folder_map
                Dim mapResults As Variant
                mapResults = modSQLite.ExecuteSQL("SELECT folder_name FROM customer_folder_map WHERE customer_name = '" & Replace(customerName, "'", "''") & "'")
                If Not IsNull(mapResults) Then
                    folderAlias = Trim(mapResults(0)(0))
                End If

                If folderAlias <> "" Then
                    ' Use mapped folder name for matching
                    results = modSQLite.ExecuteSQL("SELECT file_location, drawing_name FROM drawings WHERE drawing_name LIKE '%" & drawingNumber & "%' AND file_location LIKE '%" & folderAlias & "%'")
                Else
                    ' Fallback to original customerName
                    results = modSQLite.ExecuteSQL("SELECT file_location, drawing_name FROM drawings WHERE drawing_name LIKE '%" & drawingNumber & "%' AND file_location LIKE '%" & customerName & "%'")
                End If
                If Not IsNull(results) Then
                    fileLocation = Trim(results(0)(0))
                    dName = Trim(results(0)(1))
                End If
            End If

            ' 3. Add hyperlink if file found
            If fileLocation <> "" Then
                On Error Resume Next
                cell.Worksheet.Hyperlinks.Add Anchor:=cell, Address:=fileLocation, TextToDisplay:=drawingNumber
                On Error GoTo 0

                ' Set font style
                With cell.Font
                    .Name = "Cambria"
                    .Size = 16
                End With

                ' Update drawing_number in database
                Dim updateSQL As String
                If dName <> "" Then
                    updateSQL = "UPDATE drawings SET drawing_number = '" & drawingNumber & "' WHERE drawing_name = '" & dName & "'"
                Else
                    updateSQL = "UPDATE drawings SET drawing_number = '" & drawingNumber & "' WHERE drawing_number = '" & drawingNumber & "'"
                End If
                
                If Not modSQLite.ExecuteNonQuery(updateSQL) Then
                    Debug.Print "Error updating drawing number"
                End If
            End If
        End If
    End If
End Sub

' DON'T USE THIS!!!
' This function use cached db method, it's very slow for large data sets
Sub CreateAllHyperlinks()
    ' Creates hyperlinks for all applicable cells in Priority Sheet using caching
    Dim curBook As Workbook, curWS As Worksheet
    Dim lastRow As Long, r As Long
    Dim drawingsDict As Object
    Dim results As Variant, row As Variant

    ' 1. Initialize workbook and worksheet
    Set curBook = ThisWorkbook
    On Error Resume Next
    Set curWS = curBook.Sheets("Priority Sheet")
    If curWS Is Nothing Then
        MsgBox "Worksheet 'Priority Sheet' not found!", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' 2. Initialize SQLite connection
    If Not modSQLite.InitializeSQLite(ThisWorkbook.Path & "\jobs.db") Then
        MsgBox "Failed to initialize database!", vbCritical
        Exit Sub
    End If

    ' 3. Cache drawings table data
    Set drawingsDict = CreateObject("Scripting.Dictionary")
    results = modSQLite.ExecuteSQL("SELECT drawing_name, drawing_number, file_location FROM drawings")
    
    If Not IsNull(results) Then
        For Each row In results
            Dim dName As String, dNumber As String, fLocation As String
            dName = Trim(row(0))
            dNumber = Trim(row(1))
            fLocation = Trim(row(2))

            If Not drawingsDict.Exists(dName) Then
                Set drawingsDict(dName) = CreateObject("Scripting.Dictionary")
            End If

            With drawingsDict(dName)
                .Item("drawing_number") = dNumber
                .Item("file_location") = fLocation
            End With
        Next
    End If

    ' 4. Get last row in Priority Sheet
    lastRow = GetLastDataRow(curWS)

    ' 5. Process column E (Part #) from row 2
    For r = 2 To lastRow
        Call AddHyperlinkToCell(curWS.Cells(r, 5), drawingsDict)
    Next r

    ' 6. Clean up
    modSQLite.CloseSQLite
    Set drawingsDict = Nothing
    MsgBox "Hyperlinks created successfully!", vbInformation
End Sub

' USE THIS INSTEAD!!!
Sub CreateHyperlinks()
    ' Creates hyperlinks for selected cells in column E without caching
    Dim curBook As Workbook, curWS As Worksheet
    Dim selectedRange As Range, cell As Range
    Dim eCells As Collection
    Dim c As Variant

    ' 1. Initialize workbook and worksheet
    Set curBook = ThisWorkbook
    On Error Resume Next
    Set curWS = curBook.Sheets("Priority Sheet")
    If curWS Is Nothing Then
        MsgBox "Worksheet 'Priority Sheet' not found!", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' 2. Initialize SQLite connection
    If Not modSQLite.InitializeSQLite(ThisWorkbook.Path & "\jobs.db") Then
        MsgBox "Failed to initialize database!", vbCritical
        Exit Sub
    End If

    ' 3. Get selected range and collect column E cells
    Set selectedRange = Selection
    Set eCells = New Collection

    For Each cell In selectedRange
        If cell.Column = 5 And cell.Row > 1 Then ' Column E and skip header
            eCells.Add cell
        End If
    Next cell

    ' 4. Process each selected cell
    For Each c In eCells
        Set cell = c
        Call CreateSingleHyperlink(cell)
    Next c

    ' 5. Clean up
    modSQLite.CloseSQLite
    Set eCells = Nothing
    MsgBox "Hyperlinks created successfully!", vbInformation
End Sub