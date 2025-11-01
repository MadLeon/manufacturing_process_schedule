Attribute VB_Name = "Module9"
'FOR MP FOLDER MULTIPLE LINES SINGLE LABEL
'-----------------------------------------------------------------------------------------
' b-PAC 3.0 Component Sample (Address)
' (C)Copyright Brother Industries, Ltd. 2009
'-----------------------------------------------------------------------------------------
Option Explicit

Const sPath = "M:\EnvLabels\MP single label.lbx"

Sub cmdPrint_Click()
    Dim bRet As Boolean
    Dim ObjDoc As bpac.Document
    Set ObjDoc = CreateObject("bpac.Document")
    
    'Open lbx file
    bRet = ObjDoc.Open(sPath)
    If (bRet <> False) Then
        ' Determine how many rows the user selected
        Dim iTotal As Integer
        iTotal = Selection.Rows.Count

        ' Start Print-Setting
        ObjDoc.StartPrint "", bpoDefault
        Dim r As Integer
        For r = 1 To iTotal
            Dim Str As String
            Dim iRow As Integer
          
            ' Job Number
            iRow = Selection.Cells(r, 1).row
            Str = Cells(iRow, 2).Text
            ObjDoc.GetObject("objJob #:").Text = Str
            
            ' Company
            Str = Cells(iRow, 3).Text
            ObjDoc.GetObject("objCustomer:").Text = Str
            
            ' Order QTY
            Str = Cells(iRow, 4).Text
            ObjDoc.GetObject("objQty.:").Text = Str
            
            ' Part Number
            Str = Cells(iRow, 5).Text
            ObjDoc.GetObject("objParts:").Text = Str
            
            ' Descriptions
            Str = Cells(iRow, 10).Text
            ObjDoc.GetObject("objDescriptions:").Text = Str

            ' Due Date
            Str = Cells(iRow, 16).Text
            ObjDoc.GetObject("objDel. Req'd:").Text = Str
            
            ' Register to print
            ObjDoc.PrintOut 1, bpoDefault
        Next
        
        ' Finish Print-Setting.ÅiStart the printing.Åj
        ObjDoc.EndPrint
        
        ' Close lbx file
        ObjDoc.Close
    End If
    
End Sub

