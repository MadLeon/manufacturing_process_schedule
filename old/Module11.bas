Attribute VB_Name = "Module11"
Sub MoveData()
    Application.ScreenUpdating = False ' 关闭屏幕刷新，加快宏运行速度
    Dim i As Range
    Dim num As Integer
    num = 1

    ' 捕获异常，防止ShowAllData报错（如没筛选时）
    On Error Resume Next

    ' 取消所有筛选
    ActiveSheet.ShowAllData

    ' 按 H1 列降序排序，“JOB#”为命名区域
    Range("JOB#").Sort Key1:=Range("h1"), Order1:=xlDescending

    ' 多列自动筛选
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=1
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=2
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=3
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=4
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=5
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=6
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=7
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=8
    ActiveSheet.Range("$B$2:$N$1500").AutoFilter Field:=13

    ' 对“DELIVERY SCHEDULE TRACKING”表的H列升序排序
    ActiveWorkbook.Worksheets("DELIVERY SCHEDULE TRACKING").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DELIVERY SCHEDULE TRACKING").AutoFilter.Sort.SortFields.Add _
        Key:=Range("H2:H1500"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DELIVERY SCHEDULE TRACKING").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' 将DELIVERY SCHEDULE TRACKING中最末H列的内容复制到Cal表的A1
    Workbooks("Manufacturing Process Schedule.xlsm").Activate
    Sheets("DELIVERY SCHEDULE TRACKING").Select
    Range("H1500").End(xlUp).Select
    ActiveCell.Copy
    Sheets("Cal").Select
    Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True

    ' 获取Cal表A1的值，作为比较目标
    tgtVal = (Workbooks("Manufacturing Process Schedule.xlsm").Sheets("Cal").Range("a1"))

    ' 打开order entry log，切换至Delivery Schedule
    Workbooks.Open ("\\RTDNAS2\oe\order entry log.xlsm")
    Workbooks("Order Entry Log.xlsm").Activate
    Sheets("Delivery Schedule").Activate

    On Error Resume Next
    ' 取消所有筛选
    ActiveSheet.ShowAllData

    ' 循环遍历B4:B1500区域，筛选大于tgtVal（最新交期）的行复制到Manufacturing Process Schedule.xlsm的Temp工作表末尾
    For Each i In Range("B4:B1500")
        If i.Value > (tgtVal) Then
            i.Select
            ActiveCell.Rows("1:1").EntireRow.Select
            Selection.Copy
            Workbooks("Manufacturing Process Schedule.xlsm").Activate
            Sheets("Temp").Range("a250").End(xlUp).Offset(num, 0).PasteSpecial
            num = num + 1
            Workbooks("Order Entry Log.xlsm").Activate
            Sheets("Delivery Schedule").Activate
        End If
    Next i

    ' 切换到Temp表，删除空行和多余列
    Workbooks("Manufacturing Process Schedule.xlsm").Activate
    Sheets("Temp").Activate

    On Error Resume Next
    Columns("a").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Columns("m:o").EntireColumn.Delete
    Columns("n:p").EntireColumn.Delete
    Columns("g").EntireColumn.Delete
    Columns("h").EntireColumn.Delete
    Columns("i").EntireColumn.Delete

    ' 按新需求调整各列顺序
    ' PO移到B列
    Columns("I:I").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    ' DWG Rel移到C列
    Columns("H:H").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    ' Part#移到D列
    Columns("G:G").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    ' Description移到E列
    Columns("I:I").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    ' Customer移到F列
    Columns("G:G").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    ' QTY移到G列
    Columns("H:H").Select
    Selection.Cut
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight
    ' Due Date移到I列
    Columns("J:J").Select
    Selection.Cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    ' 删除原始J列
    Columns("J").EntireColumn.Delete

    ' 复制Temp当前有效区域到delivery schedule tracking表A2末尾
    Workbooks("Manufacturing Process Schedule.xlsm").Activate
    Sheets("temp").Select
    Range("A1").End(xlDown).Select
    ActiveCell.CurrentRegion.Select
    ActiveCell.CurrentRegion.Copy

    Workbooks("Manufacturing Process Schedule.xlsm").Activate
    Sheets("delivery schedule tracking").Select
    Range("A2").End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.PasteSpecial xlPasteValues

    ' 清空Temp表内容
    Workbooks("Manufacturing Process Schedule.xlsm").Activate
    Sheets("temp").Select
    Range("A1").End(xlDown).Select
    ActiveCell.CurrentRegion.Select
    ActiveCell.CurrentRegion.ClearContents

    Workbooks("Manufacturing Process Schedule.xlsm").Activate
    Sheets("delivery schedule tracking").Select
    Range("A2").End(xlDown).Select

    ' ---------- 移除已发货任务 ----------
    ' 复制Order Entry Log里的DELIVERY SCHEDULE到Shipped
    Windows("Manufacturing process Schedule.xlsm").Activate
    Sheets("Shipped").Select
    Columns("A").EntireColumn.Delete

    Windows("Order Entry Log.xlsm").Activate
    Sheets("DELIVERY SCHEDULE").Select
    Range("B4:B1500").Select
    Selection.Copy
    Windows("Manufacturing process Schedule.xlsm").Activate
    Sheets("Shipped").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    Range("A1").Select
    Columns("a").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

    ' 将delivery schedule tracking的H列复制到List表A列
    Sheets("delivery schedule tracking").Select
    Range("h3:h500").Select
    Selection.Copy
    Sheets("List").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ' 删除List表中与Shipped重复项
    Dim iListCount As Integer
    Dim iCtr As Integer
    iListCount = Sheets("List").Cells(Rows.Count, "A").End(xlUp).Row
    For Each x In Sheets("Shipped").Range("A1:A" & Sheets("shipped").Cells(Rows.Count, "A").End(xlUp).Row)
        For iCtr = iListCount To 1 Step -1
            If x.Value = Sheets("List").Cells(iCtr, 1).Value Then
                Sheets("List").Cells(iCtr, 1).EntireRow.Delete
            End If
        Next iCtr
    Next

    Sheets("DELIVERY SCHEDULE TRACKING").Select
    Range("H1500").End(xlUp).Select

    ' 删除delivery schedule tracking中与List重复项
    iListCount = Sheets("DELIVERY SCHEDULE TRACKING").Cells(Rows.Count, "H").End(xlUp).Row
    For Each x In Sheets("list").Range("A1:A" & Sheets("list").Cells(Rows.Count, "A").End(xlUp).Row)
        For iCtr = iListCount To 3 Step -1
            If x.Value = Sheets("DELIVERY SCHEDULE TRACKING").Cells(iCtr, 8).Value Then
                Sheets("DELIVERY SCHEDULE TRACKING").Cells(iCtr, 1).EntireRow.Delete
            End If
        Next iCtr
    Next

    ' 为表格设定细线边框
    Cells.Select
    Range("B1").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With

    ' 关闭Order Entry Log文件
    Windows("Order Entry Log.xlsm").Activate
    Workbooks("order entry log.xlsm").Close SaveChanges:=False

    ' 定位到I列最后一个有内容的单元格
    Range("I500").End(xlUp).Select

    Application.ScreenUpdating = True ' 恢复屏幕刷新
    MsgBox "DATA UPDATE COMPLETED" ' 操作完成提示

End Sub
