Attribute VB_Name = "stamp"
'STAMP , SINGLE LINE MULTIPLE LABELS
'-----------------------------------------------------------------------------------------
' b-PAC 3.0 Component Sample (Address)
' (C)Copyright Brother Industries, Ltd. 2009
'-----------------------------------------------------------------------------------------
Option Explicit

Const sPath = "M:\EnvLabels\StampTest\Stamp.lbx"
Sub Stamp()
    Dim bRet As Boolean
    Dim ObjDoc As bpac.Document
    Set ObjDoc = CreateObject("bpac.Document")
    
    If MsgBox("Do you want to print multiple labels? ", vbYesNo) = vbNo Then Exit Sub
    

    bRet = ObjDoc.Open(sPath)
    If (bRet <> False) Then
        Dim intTotal As Integer
        intTotal = Selection.Rows.Count
        Dim TotalLbs As Integer
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        TotalLbs = InputBox("How many labels Do you want to print?: ")
        MsgBox (TotalLbs)
        
        If TotalLbs <= 0 Then Exit Sub
        
        
        
        
        
'Error:
        'MsgBox ("Error Please Enter Valid Value")
        'Exit Sub
        
        
        'If TotalLbs = 0 Then Exit Sub
        
        ObjDoc.StartPrint "", bpoDefault
        Dim row As Integer
        If MsgBox("You are printing : " & TotalLbs & " Labels are you sure?", vbYesNo) = vbNo Then Exit Sub

        
    If TotalLbs >= 1 Then
    
        For row = 1 To intTotal
        Do
            Dim LoopInt As Integer
            Dim stng As String
            Dim intRow As Integer
            Dim initials As String
            
            LoopInt = LoopInt + 1
            
            ' O.E
            intRow = Selection.Cells(row, 1).row
            stng = Cells(intRow, 1).Text
            ObjDoc.GetObject("OrderEntry").Text = stng
             ' Date
            stng = Cells(intRow, 8).Text
            ObjDoc.GetObject("Date").Text = stng
            
            ' Job #
            stng = Cells(intRow, 2).Text
            ObjDoc.GetObject("Job#").Text = stng
            
            ' Part Name
            stng = Cells(intRow, 12).Text
            ObjDoc.GetObject("PO#").Text = stng
            
            ' Customer
            stng = Cells(intRow, 3).Text
            ObjDoc.GetObject("Customer").Text = stng
            
            'Delivery Date
            stng = Cells(intRow, 16).Text
            ObjDoc.GetObject("del").Text = stng
            
            'ln Number # out of
            stng = Cells(intRow, 9).Text
            ObjDoc.GetObject("Ln#").Text = stng
            
            'ln Number Total (Try to create the number last before a empty space
            stng = Cells(intRow, 9).End(xlDown).Text
            ObjDoc.GetObject("Ln#Total").Text = stng
            ' Customer Initials
            stng = Cells(intRow, 7)
            initials = Left(stng, 1) & "." & Mid(stng, InStr(stng, " ") + 1, 1) & "."
            ObjDoc.GetObject("InitC").Text = initials
            ' Quantity
            'stng = Cells(intRow, 4).Text
            'ObjDoc.GetObject("Qty").Text = stng
            
            ' Register to print
            ObjDoc.PrintOut 1, bpoDefault
            'MsgBox ("Printed : " & LoopInt)
            
        Loop Until TotalLbs = LoopInt
        
        
        Next
        
        ' Finish Print-Setting.ÅiStart the printing.Åj
            ObjDoc.EndPrint
            MsgBox ("Finished Printing")
            
        ' Close lbx file
            ObjDoc.Close
        Else
            ActiveWorkbook.FollowHyperlink Address:="M:\manufacturing process", NewWindow:=True
            MsgBox ("Opened Drive Rerun The Code")
    End If
End If
    
    
    

    
End Sub


