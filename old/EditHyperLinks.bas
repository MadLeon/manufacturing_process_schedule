Attribute VB_Name = "EditHyperLinks"
Sub ReplaceHyperlinks()
    Dim Ws As Worksheet
    Dim xHyperlink As Hyperlink
    Dim xOld As String, xNew As String
    xOld = InputBox("Enter the old hyperlink address:")
    xNew = InputBox("Enter the new hyperlink address:")
    
    For Each Ws In ThisWorkbook.Worksheets
        For Each xHyperlink In Ws.Hyperlinks
            xHyperlink.Address = Replace(xHyperlink.Address, xOld, xNew)
        Next xHyperlink
    Next Ws
End Sub
