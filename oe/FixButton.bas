Attribute VB_Name = "FixButton"
Sub FixButton()


' List of fixes
' Filter not working - Fixed
' Formula dragdown -Fixed
' inputform alignment issue - Fixed
Application.ScreenUpdating = False

    If MsgBox("List of Fixed Problems:" & Chr(10) & "   * AutoFilter" & Chr(10) & "   * Formula" & Chr(10) & "   * Inputform Alignment " & Chr(10) & "Do you want to apply these fixes? ", vbYesNo) = vbNo Then Exit Sub
    

Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("DELIVERY SCHEDULE")
    
    ws.Range("A3:R1000").AutoFilter
    ws.Range("A3:R1000").AutoFilter
    
    ws.Range("R4:R500").FillDown
    
    
    
Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Worksheets("Input Form")
    
        ws2.Activate
        
       On Error Resume Next
       
        ws2.Unprotect
        
    
        ws2.Range("Customer").HorizontalAlignment = xlLeft
        ws2.Range("Customer").Font.Size = 16
        
        ws2.Range("QTY").Font.Size = 16
        ws2.Range("QTY").HorizontalAlignment = xlLeft
        
        ws2.Range("Parts").Font.Size = 16
        ws2.Range("Parts").HorizontalAlignment = xlLeft
        
        ws2.Range("Revision").HorizontalAlignment = xlLeft
        ws2.Range("Revision").Font.Size = 16
        
        ws2.Range("Contact").HorizontalAlignment = xlLeft
        ws2.Range("Contact").Font.Size = 16
        
        ws2.Range("poline").HorizontalAlignment = xlLeft
        ws2.Range("poline").Font.Size = 16
        
        ws2.Range("desc").HorizontalAlignment = xlLeft
        ws2.Range("desc").Font.Size = 16
        
        ws2.Range("price").HorizontalAlignment = xlLeft
        ws2.Range("price").Font.Size = 16
        
        ws2.Range("po").HorizontalAlignment = xlLeft
        ws2.Range("po").Font.Size = 16
        
        ws2.Range("date").HorizontalAlignment = xlLeft
        ws2.Range("date").Font.Size = 16
        
        ws2.Protect
        
        ws2 = Nothing
        
    MsgBox ("Fixes applied")
        ws.Activate
        
Application.ScreenUpdating = True

        
End Sub
