" Shipped Sheet event handler"
Private Sub Worksheet_Change(ByVal Target As Range)
    ' If column J (10) is changed and not the header row
    If Target.Column = 10 And Target.row > 1 Then
        Application.EnableEvents = False  ' Disable events to prevent recursion
        Dim wsPri As Worksheet
        Set wsPri = ThisWorkbook.Sheets("Priority Sheet")
        Dim rowCopy As Long: rowCopy = Target.row
        Dim jobNum As String

        ' Get Job_Number (column 1)
        jobNum = Trim(Me.Cells(rowCopy, 1).Value)

        Select Case Target.Value
            Case "Return"
                ' 1. Copy row to Priority Sheet
                If jobNum <> "" Then
                    Dim priRow As Long

                    priRow = GetLastDataRow(wsPri) + 1 ' Use GetLastDataRow to find next row
                    wsPri.Range(wsPri.Cells(priRow, 1), wsPri.Cells(priRow, 7)).Value = _
                        Me.Range(Me.Cells(rowCopy, 1), Me.Cells(rowCopy, 7)).Value 'Copy columns A:G

                    ' --- 2. Format the pasted row (highlight, font, borders, alignment) ---
                    Dim rg As Range
                    Set rg = wsPri.Range(wsPri.Cells(priRow, 1), wsPri.Cells(priRow, 7)) ' Columns A to G
                    With rg
                        .Interior.Color = RGB(255, 199, 44)
                        .Font.Name = "Cambria"
                        .Font.Size = 16
                        With .Borders
                            .LineStyle = xlContinuous
                            .Color = vbBlack
                            .Weight = xlThin
                        End With
                        ' Center align columns A, B, C, E, F, G
                        With .Columns(1)  ' A
                            .HorizontalAlignment = xlCenter
                        End With
                        With .Columns(2)  ' B
                            .HorizontalAlignment = xlCenter
                        End With
                        With .Columns(3)  ' C
                            .HorizontalAlignment = xlCenter
                        End With
                        With .Columns(5)  ' E
                            .HorizontalAlignment = xlCenter
                        End With
                        With .Columns(6)  ' F
                            .HorizontalAlignment = xlCenter
                        End With
                        With .Columns(7)  ' G
                            .HorizontalAlignment = xlCenter
                        End With
                        ' Left align column D
                        With .Columns(4)
                            .HorizontalAlignment = xlLeft
                        End With
                    End With
                End If

                ' 3. Delete the row from Shipped sheet
                Me.rows(rowCopy).Delete
                Debug.Print "Moved to Priority Sheet", vbInformation

            Case "Delete"
                ' Confirm deletion
                If MsgBox("Are you sure you want to delete this row?", vbYesNo + vbQuestion, "Delete Confirmation") = vbYes Then
                    Dim rowsToDelete As Long
                    rowsToDelete = CountParts(Me, rowCopy) ' Use CountParts to get number of rows
                    Me.rows(rowCopy).Resize(rowsToDelete).Delete ' Delete the rows
                    Debug.Print "Row(s) deleted", vbInformation
                Else
                    Target.Value = ""  ' Reset value if not deleted
                End If
        End Select

        Application.EnableEvents = True
    End If
End Sub



