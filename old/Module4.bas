Attribute VB_Name = "Module4"
Sub RemoveShippedItems()
    Attribute RemoveShippedItems.VB_Description = "This Macro will remove jobs that have already shipped."
    Attribute RemoveShippedItems.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False  ' 关闭屏幕刷新，提高速度
    Dim iListCount As Integer           ' 用于记录数据行数的变量
    Dim iCtr As Integer                 ' 用于循环计数

    '
    ' Shipped Macro
    ' 该宏用于移除已经发货的任务
    '

    ' 激活生产排程文件
    Windows("Manufacturing process Schedule.xlsm").Activate
    Sheets("Shipped").Select
    Columns("A").EntireColumn.Delete        ' 删除A列（清理旧数据）

    ' 打开订单日志文件，复制交付计划
    Workbooks.Open ("\\bvserver\oe\order entry log.xlsm")
    Windows("Order Entry Log.xlsm").Activate
    Sheets("DELIVERY SCHEDULE").Select
    Range("B4:B1000").Select
    Selection.Copy

    ' 回到生产排程，粘贴交付计划到Shipped表
    Windows("Manufacturing process Schedule.xlsm").Activate
    Sheets("Shipped").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ' 删除Shipped表中A列的所有空行
    Columns("a").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

    ' 复制delivery schedule tracking表的H列到List表的A列
    Sheets("delivery schedule tracking").Select
    Range("h3:h1000").Select
    Selection.Copy
    Sheets("List").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ' 获取List表实际数据行数
    iListCount = Sheets("List").Cells(Rows.Count, "A").End(xlUp).Row

    ' 从Shipped表A列，逐项与List表A列做比对，若相同则删除List表相应行
    For Each x In Sheets("Shipped").Range("A1:A" & Sheets("shipped").Cells(Rows.Count, "A").End(xlUp).Row)
        For iCtr = iListCount To 1 Step -1
            If x.Value = Sheets("List").Cells(iCtr, 1).Value Then
                Sheets("List").Cells(iCtr, 1).EntireRow.Delete
            End If
        Next iCtr
    Next

    ' 再次获取delivery schedule tracking表的H列数据行数
    iListCount = Sheets("DELIVERY SCHEDULE TRACKING").Cells(Rows.Count, "H").End(xlUp).Row

    ' 从List表A列，逐项与delivery schedule tracking表H列比对，若相同则删除delivery schedule tracking对应行
    For Each x In Sheets("list").Range("A1:A" & Sheets("list").Cells(Rows.Count, "A").End(xlUp).Row)
        For iCtr = iListCount To 3 Step -1
            If x.Value = Sheets("DELIVERY SCHEDULE TRACKING").Cells(iCtr, 8).Value Then
                Sheets("DELIVERY SCHEDULE TRACKING").Cells(iCtr, 1).EntireRow.Delete
            End If
        Next iCtr
    Next

    ' 关闭Order Entry Log文件，不保存更改
    Windows("Order Entry Log.xlsm").Activate
    Workbooks("order entry log.xlsm").Close SaveChanges:=False

    ' 选中delivery schedule tracking表中H列最后有数据的单元格
    Sheets("DELIVERY SCHEDULE TRACKING").Select
    Range("H1500").End(xlUp).Select

    Application.ScreenUpdating = True  ' 重新开启屏幕刷新
    MsgBox "SHIPPED ITEM REMOVAL COMPLETED"  ' 操作完成弹窗

    'Application.ScreenUpdating = True
    'MsgBox "Done!"

End Sub
