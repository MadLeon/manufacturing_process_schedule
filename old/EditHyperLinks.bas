Attribute VB_Name = "EditHyperLinks"

' -----------------------------------------------------------------------------
' 本代码用于批量替换当前工作簿所有工作表里的超链接地址部分。
' 流程说明：
'   1. 用户输入需要查找和替换的“旧地址”(xOld)和“新地址”(xNew)。
'   2. 遍历工作簿内全部工作表与每个表中的全部超链接。
'   3. 对每个超链接的地址(Address)属性，用Replace函数将所有包含的旧地址部分替换为新地址。
' 使用场景举例：批量更换网站域名、服务器路径等超链接地址。
' -----------------------------------------------------------------------------

Sub ReplaceHyperlinks()
    Dim Ws As Worksheet              ' 用于遍历每个工作表
    Dim xHyperlink As Hyperlink      ' 用于遍历工作表内每个超链接
    Dim xOld As String, xNew As String  ' 存储用户输入的旧/新链接内容

    ' 获取用户输入的需要替换的“旧地址”和“新地址”
    xOld = InputBox("Enter the old hyperlink address:")
    xNew = InputBox("Enter the new hyperlink address:")
    
    ' 遍历全部工作表和全部超链接，执行替换操作
    For Each Ws In ThisWorkbook.Worksheets
        For Each xHyperlink In Ws.Hyperlinks
            xHyperlink.Address = Replace(xHyperlink.Address, xOld, xNew)
        Next xHyperlink
    Next Ws
End Sub

