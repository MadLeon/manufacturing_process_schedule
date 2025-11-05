' Shipped Sheet 模块代码里
Private Sub Worksheet_Change(ByVal Target As Range)
    ' 仅监控 J 列（第10列，且不含表头）
    If Target.Column = 10 And Target.Row > 1 Then
        Application.EnableEvents = False  ' 避免事件循环
        Dim wsPri As Worksheet
        Set wsPri = ThisWorkbook.Sheets("Priority Sheet")
        Dim rowCopy As Long: rowCopy = Target.Row
        Dim jobNum As String

        ' 提取 Job_Number （第1列）
        jobNum = Trim(Me.Cells(rowCopy, 1).Value)

        Select Case Target.Value
            Case "Return"
                ' 1. 移动数据到 Priority Sheet
                If jobNum <> "" Then
                    Dim priRow As Long
                    priRow = wsPri.Cells(wsPri.Rows.Count, 1).End(xlUp).Row + 1
                    wsPri.Range(wsPri.Cells(priRow, 1), wsPri.Cells(priRow, 7)).Value = _
                        Me.Range(Me.Cells(rowCopy, 1), Me.Cells(rowCopy, 7)).Value 'A:G数据

                    ' --- 2. 格式化新插入行 (***重点，在返回后设置格式***) ---
                    Dim rg As Range
                    Set rg = wsPri.Range(wsPri.Cells(priRow, 1), wsPri.Cells(priRow, 7)) ' A 到 G 列，只对 A 到 G 设置样式
                    With rg
                        .Interior.Color = RGB(255, 199, 44)
                        .Font.Name = "Cambria"
                        .Font.Size = 16
                        With .Borders
                            .LineStyle = xlContinuous
                            .Color = vbBlack
                            .Weight = xlThin
                        End With
                        ' 水平居中 A, B, C, F, G 列
                        With .Columns(1)  ' A 列
                            .HorizontalAlignment = xlCenter
                        End With
                        With .Columns(2)  ' B 列
                            .HorizontalAlignment = xlCenter
                        End With
                        With .Columns(3)  ' C 列
                            .HorizontalAlignment = xlCenter
                        End With
                        With .Columns(5)  ' E 列
                            .HorizontalAlignment = xlCenter
                        End With
                        With .Columns(6)  ' F 列
                            .HorizontalAlignment = xlCenter
                        End With
                        With .Columns(7)  ' G 列
                            .HorizontalAlignment = xlCenter
                        End With
                        ' 第4列 (D 列) 左对齐
                        With .Columns(4)
                            .HorizontalAlignment = xlLeft
                        End With
                    End With
                End If

                ' 3. 删除 Shipped 表行
                Me.Rows(rowCopy).Delete
                MsgBox "条目已返回到 Priority Sheet。", vbInformation

            Case "Delete"
                ' 确认是否删除
                If MsgBox("确定要删除此条目吗？", vbYesNo + vbQuestion, "确认删除") = vbYes Then
                    Me.Rows(rowCopy).Delete
                    MsgBox "条目已删除。", vbInformation
                Else
                    Target.Value = ""  ' 恢复为空
                End If
        End Select

        Application.EnableEvents = True
    End If
End Sub

