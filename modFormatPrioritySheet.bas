' modFormatPrioritySheet.bas
Sub FormatPrioritySheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Priority Sheet")
    If ws Is Nothing Then
        MsgBox "未找到 Priority Sheet": Exit Sub
    End If
    On Error GoTo 0

    Dim usedRng As Range, lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    lastCol = 9 ' 只处理前9列

    Set usedRng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' 1. 单元格格式：文本, 字号, 字体
    usedRng.NumberFormat = "@"
    usedRng.Font.Name = "Cambria"
    usedRng.Font.Size = 16

    ' 2. 所有单元格垂直居中
    usedRng.VerticalAlignment = xlVAlignCenter

    ' 3. 列宽自动（类似双击列头右边）
    usedRng.Columns.AutoFit

    ' 4. 标题行设置
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Interior.Color = RGB(255, 199, 206) ' 淡粉色
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        ' 加内外所有边框（细线）
        With .Borders
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
    End With

    ' 5. 单元格水平居中和左对齐控制
    ' 第1、2、3、6、7、8、9列全体：水平居中
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 2), ws.Cells(lastRow, 2)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 3), ws.Cells(lastRow, 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 5), ws.Cells(lastRow, 5)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 6), ws.Cells(lastRow, 6)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 7), ws.Cells(lastRow, 7)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 8), ws.Cells(lastRow, 8)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 9), ws.Cells(lastRow, 9)).HorizontalAlignment = xlCenter
    ' 第4，5列标题行外（2到lastRow）左对齐
    If lastRow > 1 Then
        ws.Range(ws.Cells(2, 4), ws.Cells(lastRow, 4)).HorizontalAlignment = xlLeft
    End If

    ' 6. 前9列加排序下拉按钮（自动筛选）
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).AutoFilter

    ' 7. 设置数据区域 A1:G 列边框
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 7)).Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With

    Debug.Print "Priority Sheet格式设置完成!"
End Sub
