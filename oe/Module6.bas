Attribute VB_Name = "Module6"
Sub Report()

Application.ScreenUpdating = False
Workbooks("order entry log.xlsm").Sheets("report").Select
Range("A:A").Value = ""



    Dim startdate As Date, enddate As Date
    Dim rng As Range, destRow As Long
    Dim shtSrc As Worksheet, shtDest As Worksheet
    Dim c As Range

    Set shtSrc = Workbooks("order entry log.xlsm").Sheets("Parts Shipped")
    Set shtDest = Workbooks("order entry log.xlsm").Sheets("report")

    destRow = 25 'start copying to this row
On Error GoTo ReEnter
    startdate = CDate(InputBox("Begining Date"))
    enddate = CDate(InputBox("End Date"))
    'Range("d16").Value = startdate
    'Range("d17").Value = enddate
    Range("c1").Value = "Performance Report From  " & startdate & "  To  " & enddate
 
    
ReEnter:
    MsgBox "If you don't get a result," & vbNewLine & "Please try to Re-Enter with a new date" & vbNewLine & vbNewLine & "" & vbNewLine & "    For Example" & vbNewLine & " 'M/30/Y' or 'M/28/Y'"


    'don't scan the entire column...
    Set rng = Application.Intersect(shtSrc.Range("P:P"), shtSrc.UsedRange)

    For Each c In rng.Cells
        If c.Value >= startdate And c.Value <= enddate Then

            'c.Offset(0, 2).Resize(1, 5).Copy _
                          'shtDest.Cells(destRow, 1)
                          c.Offset(0, 2).Resize(1, 5).Copy
                          shtDest.Cells(destRow, 1).PasteSpecial Paste:=xlPasteValues
                          
                          
            destRow = destRow + 1

        End If
    Next

End Sub



