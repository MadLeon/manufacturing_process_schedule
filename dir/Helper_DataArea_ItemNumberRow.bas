' =========================================================
' Module: Helper_DataArea_ItemNumberRow
' Purpose: Dynamically adjust the first item title row (row 15)
'          based on the value in H9. Handles merging/unmerging
'          and data validation synchronization.
' =========================================================

Option Explicit

Public Sub HandleItemNumber(ws As Worksheet)
    Dim ctrlVal As String
    ctrlVal = Trim(CStr(ws.Range("H9").Value)) ' Read the control value stored in cell H9
    
    ' Synchronize and update H10 with the same value as H9
    ws.Range("H10").Value = ctrlVal
    
    ' Get the title text from H6
    Dim HeaderText As String
    HeaderText = Trim(CStr(ws.Range("H6").Value))
    
    ' If ctrlVal = "1" ? split the merged header cells A15:I15
    If ctrlVal = "1" Then
        If ws.Range("A15:I15").MergeCells Then
            ' Clear and unmerge A15:I15
            ws.Range("A15:I15").ClearContents
            ws.Range("A15:I15").UnMerge
            
            ' Copy the formatting from row 16 to row 15
            ws.Rows(16).Copy
            ws.Rows(15).PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
            
            ' Copy data validation rules from row 16 to row 15 (columns A to I)
            Dim c As Long, vType As Long
            For c = 1 To 9   ' Columns A through I
                On Error Resume Next
                ws.Cells(15, c).Validation.Delete ' Remove existing validation (if any)
                On Error GoTo 0
                
                vType = -1
                On Error Resume Next
                vType = ws.Cells(16, c).Validation.Type ' Read validation type from row 16
                On Error GoTo 0
                
                ' If the source cell has a validation rule (not input-only)
                If vType <> -1 And vType <> xlValidateInputOnly Then
                    With ws.Cells(16, c).Validation
                        ' Duplicate the validation settings from row 16
                        ws.Cells(15, c).Validation.Add _
                            Type:=.Type, _
                            AlertStyle:=.AlertStyle, _
                            Operator:=.Operator, _
                            Formula1:=.Formula1, _
                            Formula2:=.Formula2
                        
                        ' Copy all validation properties
                        ws.Cells(15, c).Validation.IgnoreBlank = .IgnoreBlank
                        ws.Cells(15, c).Validation.InCellDropdown = .InCellDropdown
                        ws.Cells(15, c).Validation.ErrorTitle = .ErrorTitle
                        ws.Cells(15, c).Validation.ErrorMessage = .ErrorMessage
                        ws.Cells(15, c).Validation.InputTitle = .InputTitle
                        ws.Cells(15, c).Validation.InputMessage = .InputMessage
                        ws.Cells(15, c).Validation.ShowError = .ShowError
                        ws.Cells(15, c).Validation.ShowInput = .ShowInput
                    End With
                End If
            Next c
        End If
    
    ' If ctrlVal > 1 ? merge A15:I15 and append “-1” to the header
    ElseIf IsNumeric(ctrlVal) Then
        If val(ctrlVal) > 1 Then
            ' Merge A15:I15 if not already merged
            If ws.Range("A15:I15").MergeCells = False Then
                ws.Range("A15:I15").Merge
                ' Center the content horizontally
                ws.Range("A15:I15").HorizontalAlignment = xlCenter
                ws.Range("A15:I15").Font.Bold = True
                ' Set background color: Aqua, Accent 5, Lighter 80%
                ws.Range("A15:I15").Interior.Color = RGB(218, 238, 243)
            End If
            
            ' Update the header text with a “-1” suffix if available
            If ws.Index = 1 Then
                If HeaderText <> "" Then
                    ws.Range("A15").Value = HeaderText & "-1"
                Else
                    ws.Range("A15").ClearContents
                End If
            End If
        End If
    End If
End Sub


