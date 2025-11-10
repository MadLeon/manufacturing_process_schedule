Private Sub Worksheet_Change(ByVal Target As Range)
    Dim mapWS As Worksheet
    Set mapWS = Me ' This code assumes it is in the Map sheet's code module

    Dim dbPath As String
    Dim rowNum As Long
    Dim folderName As String, customerName As String
    Dim sqlInsert As String

    ' Only process changes in column B, from row 4 down
    If Not Intersect(Target, mapWS.Range("B4:B" & mapWS.rows.Count)) Is Nothing Then
        For Each cell In Intersect(Target, mapWS.Range("B4:B" & mapWS.rows.Count))
            rowNum = cell.row
            folderName = Trim(mapWS.Cells(rowNum, 1).Value)
            customerName = Trim(cell.Value)
            If folderName <> "" And customerName <> "" Then
                dbPath = ThisWorkbook.Path & "\jobs.db"
                If InitializeSQLite(dbPath) Then
                    Dim sqlUpdate As String, sqlCheck As String
                    Dim updateResult As Boolean
                    sqlUpdate = "UPDATE customer_folder_map SET customer_name = '" & Replace(customerName, "'", "''") & "' WHERE folder_name = '" & Replace(folderName, "'", "''") & "'"
                    updateResult = ExecuteNonQuery(sqlUpdate)
                    ' If no row was updated, insert new
                    sqlCheck = "SELECT COUNT(*) FROM customer_folder_map WHERE folder_name = '" & Replace(folderName, "'", "''") & "'"
                    Dim results As Variant
                    results = ExecuteSQL(sqlCheck)
                    If IsNull(results) Or results(0)(0) = 0 Then
                        sqlInsert = "INSERT INTO customer_folder_map (folder_name, customer_name) VALUES ('" & Replace(folderName, "'", "''") & "', '" & Replace(customerName, "'", "''") & "')"
                        Call ExecuteNonQuery(sqlInsert)
                    End If
                    CloseSQLite
                End If
            End If
        Next cell
    End If
End Sub

Sub ListGDriveFoldersToMapSheet()
    Dim mapWS As Worksheet
    Dim folderPath As String
    Dim fso As Object, folder As Object, subFolder As Object
    Dim rowNum As Long

    ' Set the path to G:\
    folderPath = "G:\"

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    ' Get or create the Map sheet
    On Error Resume Next
    Set mapWS = ThisWorkbook.Sheets("Map")
    If mapWS Is Nothing Then
        Set mapWS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        mapWS.Name = "Map"
    End If
    On Error GoTo 0

    ' Start from row 4
    rowNum = 4

    ' Clear previous folder names in column A
    mapWS.Range("A4:A" & mapWS.rows.Count).ClearContents

    ' List all subfolders in G:\
    For Each subFolder In folder.SubFolders
        mapWS.Cells(rowNum, 1).Value = subFolder.Name
        rowNum = rowNum + 1
    Next subFolder

    ' Sort the folder names in column A (from row 4 down)
    Dim lastRow As Long
    lastRow = mapWS.Cells(mapWS.rows.Count, 1).End(xlUp).row
    If lastRow >= 4 Then
        mapWS.Range("A4:A" & lastRow).Sort Key1:=mapWS.Range("A4"), Order1:=xlAscending, Header:=xlNo
    End If

    MsgBox "Folders from G:\ have been listed and sorted in the Map sheet.", vbInformation
End Sub
