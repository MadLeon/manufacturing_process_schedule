' Parse the workbook name to extract parts separated by " @"
' Store the results in C7 and H6 of the current sheet
Public Sub HandleFileNameParsing(ws As Worksheet)
    Dim fname As String, base As String
    Dim pos As Long
    Dim x As String, y As String
    Dim dotPos As Long

    ' Get the current workbook name (including extension)
    fname = ThisWorkbook.Name
    
    ' Remove the file extension to get the base name
    dotPos = InStrRev(fname, ".")
    If dotPos > 0 Then
        base = Left$(fname, dotPos - 1)
    Else
        base = fname
    End If

    ' Look for the separator " @"
    pos = InStr(1, base, " @", vbBinaryCompare)

    If pos > 0 Then
        ' If found, split into x (before " @") and y (after " @")
        x = Trim$(Left$(base, pos - 1))
        y = Trim$(Mid$(base, pos + 2))
    Else
        ' If not found, assign the full base name to x and leave y empty
        x = base
        y = ""
    End If

    ' Write the results to the appropriate cells
    ws.Range("C7").Value = x
    ws.Range("H6").Value = y
End Sub
