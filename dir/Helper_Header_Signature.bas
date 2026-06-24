' Display a signature picture based on the selected value in Target cell
Public Sub HandleSignatureSelection(ws As Worksheet, Target As Range)
    Dim wsData As Worksheet
    Dim sigName As String
    Dim sigPic As Shape, newPic As Shape
    
    Set wsData = ThisWorkbook.Sheets("Data") ' Signatures stored in "Data" sheet
    
    ' Determine the signature name based on input value
    Select Case Target.Value
        Case "Franco": sigName = "FrancoSig"
        Case "Devang": sigName = "DevangSig"
        Case "Sajid": sigName = "SajidSig"
        Case Else: sigName = ""
    End Select
    
    ' Remove any existing validation or previous signature picture
    On Error Resume Next
    ws.Range("H57").Validation.Delete
    ws.Shapes("SignaturePic").Delete
    On Error GoTo 0
    
    ' If a signature exists, copy it from the Data sheet
    If sigName <> "" Then
        Set sigPic = wsData.Shapes(sigName)
        Debug.Print "Original picture: " & sigPic.Name & " | Top=" & sigPic.Top & " | Left=" & sigPic.Left
    
        sigPic.Copy
        ws.PasteSpecial
        Dim shpRange As ShapeRange
        Set shpRange = Selection.ShapeRange
        Set newPic = shpRange(1)
        newPic.Name = "SignaturePic"
        Debug.Print "New picture pasted: " & newPic.Name & " | Top=" & newPic.Top & " | Left=" & newPic.Left
    
        ' Move the new picture
        newPic.Top = newPic.Top + 9
        newPic.Left = newPic.Left + 2
        Debug.Print "New picture moved: Top=" & newPic.Top & " | Left=" & newPic.Left
    End If
    
    ' Clear the target cell content (signature text)
    ws.Range("H57").ClearContents
End Sub
